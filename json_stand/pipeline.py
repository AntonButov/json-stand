from __future__ import annotations

import base64
import io
import json
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from json_stand.models import ConvertError, ConvertResult

# LibreOffice writes Calibri 11 back into the sheet and maps one OOXML width
# unit to 5.55 pt (19.579 cm per 100 units). Dividing PDF points by this keeps
# the printed column the same width as the PDF column.
EXCEL_PT_PER_WIDTH = 5.55
EMU_PER_PT = 12700

# ECMA-376 CT_Font child order. openpyxl/LibreOffice often emit name/color
# before sz; openxml-audit then reports Unexpected element 'sz'/'color'.
_FONT_CHILD_ORDER = (
    "b",
    "i",
    "strike",
    "condense",
    "extend",
    "outline",
    "shadow",
    "u",
    "vertAlign",
    "sz",
    "color",
    "name",
    "family",
    "charset",
    "scheme",
)
_SSML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _find_soffice() -> str | None:
    import os

    env = os.environ.get("SOFFICE_PATH")
    if env and Path(env).exists():
        return env
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    mac = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return str(mac) if mac.exists() else None


def _normalize_xlsx(path: Path, timeout_s: int = 60) -> str | None:
    """Round-trip via LibreOffice so openxml-audit / Excel open cleanly."""
    soffice = _find_soffice()
    if not soffice:
        return "libreoffice_missing"

    work = Path(tempfile.mkdtemp(prefix="json-stand-norm-"))
    profile = work / "profile"
    profile.mkdir()
    src = work / "input.xlsx"
    shutil.copy2(path, src)
    try:
        # xlsx → ods → xlsx avoids in-place convert quirks
        proc1 = subprocess.run(
            [
                soffice,
                f"-env:UserInstallation=file://{profile}",
                "--headless",
                "--norestore",
                "--convert-to",
                "ods",
                "--outdir",
                str(work),
                str(src),
            ],
            capture_output=True,
            text=True,
            timeout=timeout_s,
            check=False,
        )
        ods = work / "input.ods"
        if proc1.returncode != 0 or not ods.exists():
            return (proc1.stderr or proc1.stdout or "libreoffice_fail_ods")[:300]

        proc2 = subprocess.run(
            [
                soffice,
                f"-env:UserInstallation=file://{profile}",
                "--headless",
                "--norestore",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(work),
                str(ods),
            ],
            capture_output=True,
            text=True,
            timeout=timeout_s,
            check=False,
        )
        out = work / "input.xlsx"
        if proc2.returncode != 0 or not out.exists() or out.stat().st_size == 0:
            return (proc2.stderr or proc2.stdout or "libreoffice_fail_xlsx")[:300]
        shutil.copy2(out, path)
        return None
    except subprocess.TimeoutExpired:
        return "timeout"
    finally:
        shutil.rmtree(work, ignore_errors=True)


def points_to_excel_width(points: float) -> float:
    return round(points / EXCEL_PT_PER_WIDTH, 2)


def _local(tag: str) -> str:
    if tag.startswith("{"):
        return tag.rsplit("}", 1)[-1]
    return tag


def _canonicalize_ooxml_fonts(path: Path) -> None:
    """Rewrite xl/styles.xml <font> children into schema order."""
    try:
        with zipfile.ZipFile(path, "r") as zin:
            if "xl/styles.xml" not in zin.namelist():
                return
            styles_xml = zin.read("xl/styles.xml")
            others = {name: zin.read(name) for name in zin.namelist() if name != "xl/styles.xml"}
    except zipfile.BadZipFile:
        return

    root = ET.fromstring(styles_xml)
    ns = {"m": _SSML}
    fonts = root.find("m:fonts", ns)
    if fonts is None:
        # default-ns stripped variants
        fonts = root.find(f"{{{_SSML}}}fonts")
    if fonts is None:
        return

    order_index = {name: i for i, name in enumerate(_FONT_CHILD_ORDER)}
    changed = False
    for font in list(fonts):
        if _local(font.tag) != "font":
            continue
        children = list(font)
        ranked: list[tuple[int, int, ET.Element]] = []
        for idx, child in enumerate(children):
            ranked.append((order_index.get(_local(child.tag), 1000), idx, child))
        ranked.sort(key=lambda item: (item[0], item[1]))
        new_order = [item[2] for item in ranked]
        if new_order == children:
            continue
        changed = True
        for child in children:
            font.remove(child)
        for child in new_order:
            font.append(child)

    if not changed:
        return

    # Preserve original namespace prefixes on write.
    ET.register_namespace("", _SSML)
    payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in others.items():
            zout.writestr(name, data)
        zout.writestr("xl/styles.xml", payload)
    path.write_bytes(buf.getvalue())


def _positive_floats(value: object) -> list[float] | None:
    if not isinstance(value, list):
        return None
    numbers: list[float] = []
    for item in value:
        if isinstance(item, bool) or not isinstance(item, (int, float)):
            return None
        numbers.append(float(item))
    return numbers


def _apply_style(cell, style: object) -> None:
    if not isinstance(style, dict):
        return
    size = style.get("size")
    font_name = style.get("font") or "Arial"
    cell.font = Font(
        name=str(font_name),
        size=float(size) if isinstance(size, (int, float)) and not isinstance(size, bool) else None,
        bold=bool(style.get("bold")),
    )
    cell.alignment = Alignment(
        horizontal=style.get("halign") or "left",
        vertical=style.get("valign") or "bottom",
        wrap_text=bool(style.get("wrap")),
    )


def _apply_merges(ws, merges: object) -> None:
    if not isinstance(merges, list):
        return
    for ref in merges:
        if isinstance(ref, str) and ":" in ref:
            ws.merge_cells(ref)


def _restore_row_heights(path: Path, tables: list[dict[str, Any]]) -> None:
    wb = load_workbook(path)
    for ws, table in zip(wb.worksheets, tables):
        heights = _positive_floats(table.get("row_heights_pt"))
        if not heights:
            continue
        for index, points in enumerate(heights, start=1):
            if points > 0:
                ws.row_dimensions[index].height = points
    wb.save(path)


def _pt_to_cell(
    x_pt: float,
    y_pt: float,
    col_widths: list[float],
    row_heights: list[float],
) -> tuple[int, int, int, int]:
    """Return 0-based col/row and EMU offsets inside the cell."""
    x = max(0.0, x_pt)
    y = max(0.0, y_pt)
    col = 0
    for width in col_widths:
        if x < width or col == len(col_widths) - 1:
            break
        x -= width
        col += 1
    if not col_widths:
        col = 0
        x = 0.0
    row = 0
    for height in row_heights:
        if y < height or row == len(row_heights) - 1:
            break
        y -= height
        row += 1
    if not row_heights:
        row = 0
        y = 0.0
    return col, row, int(round(x * EMU_PER_PT)), int(round(y * EMU_PER_PT))


def _ensure_sheet_covers_images(table: dict[str, Any]) -> None:
    """Pad empty rows so every image bbox fits on the sheet grid."""
    images = table.get("images") or []
    if not isinstance(images, list) or not images:
        return
    widths = _positive_floats(table.get("col_widths_pt")) or []
    heights = _positive_floats(table.get("row_heights_pt")) or []
    origin_x = float(table.get("origin_x_pt") or 0)
    origin_y = float(table.get("origin_y_pt") or 0)
    rows = table.get("rows") or []
    styles = table.get("cell_styles") or []
    if not isinstance(rows, list):
        return

    # Images above the text grid: prepend empty rows.
    top_pad = 0.0
    for image in images:
        if not isinstance(image, dict):
            continue
        y = float(image.get("y_pt") or 0) - origin_y
        if y < 0:
            top_pad = max(top_pad, -y)
    if top_pad > 0:
        heights = [round(top_pad, 2)] + heights
        cols = len(rows[0]) if rows and isinstance(rows[0], list) else 1
        rows.insert(0, [None] * cols)
        if isinstance(styles, list):
            styles.insert(0, [None] * cols)
        # Shift merges down by one row.
        merges = table.get("merges") or []
        shifted: list[str] = []
        for ref in merges:
            if not isinstance(ref, str) or ":" not in ref:
                continue
            start, end = ref.split(":", 1)
            def bump(cell: str) -> str:
                letters = "".join(ch for ch in cell if ch.isalpha())
                number = int("".join(ch for ch in cell if ch.isdigit()) or "1")
                return f"{letters}{number + 1}"
            shifted.append(f"{bump(start)}:{bump(end)}")
        table["merges"] = shifted
        origin_y -= top_pad
        table["origin_y_pt"] = round(origin_y, 2)

    # Images below / to the right: extend heights and widths.
    for image in images:
        if not isinstance(image, dict):
            continue
        x1 = float(image.get("x_pt") or 0) - origin_x + float(image.get("width_pt") or 0)
        y1 = float(image.get("y_pt") or 0) - origin_y + float(image.get("height_pt") or 0)
        total_w = sum(widths) if widths else 0.0
        total_h = sum(heights) if heights else 0.0
        if x1 > total_w:
            widths = list(widths) + [round(x1 - total_w, 2)]
        if y1 > total_h:
            extra = round(y1 - total_h, 2)
            heights = list(heights) + [extra]
            cols = len(rows[0]) if rows and isinstance(rows[0], list) else 1
            rows.append([None] * cols)
            if isinstance(styles, list):
                styles.append([None] * cols)

    table["col_widths_pt"] = widths
    table["row_heights_pt"] = heights
    table["rows"] = rows
    table["cell_styles"] = styles


def _place_images(path: Path, tables: list[dict[str, Any]]) -> None:
    wb = load_workbook(path)
    temp_dirs: list[Path] = []
    try:
        for ws, table in zip(wb.worksheets, tables):
            images = table.get("images") or []
            if not isinstance(images, list) or not images:
                continue
            widths = _positive_floats(table.get("col_widths_pt")) or []
            heights = _positive_floats(table.get("row_heights_pt")) or []
            origin_x = float(table.get("origin_x_pt") or 0)
            origin_y = float(table.get("origin_y_pt") or 0)
            for image in images:
                if not isinstance(image, dict):
                    continue
                b64 = image.get("data_b64")
                if not isinstance(b64, str) or not b64:
                    continue
                try:
                    raw = base64.b64decode(b64)
                except Exception:  # noqa: BLE001
                    continue
                folder = Path(tempfile.mkdtemp(prefix="json-stand-img-"))
                temp_dirs.append(folder)
                tmp = folder / "img.png"
                tmp.write_bytes(raw)
                xl_img = XLImage(str(tmp))
                width_pt = float(image.get("width_pt") or 0)
                height_pt = float(image.get("height_pt") or 0)
                x_pt = float(image.get("x_pt") or 0) - origin_x
                y_pt = float(image.get("y_pt") or 0) - origin_y
                col, row, col_off, row_off = _pt_to_cell(x_pt, y_pt, widths, heights)
                marker = AnchorMarker(col=col, colOff=col_off, row=row, rowOff=row_off)
                size = XDRPositiveSize2D(
                    cx=max(1, int(round(width_pt * EMU_PER_PT))),
                    cy=max(1, int(round(height_pt * EMU_PER_PT))),
                )
                xl_img.anchor = OneCellAnchor(_from=marker, ext=size)
                ws.add_image(xl_img)
            for index, points in enumerate(heights, start=1):
                if points > 0:
                    ws.row_dimensions[index].height = points
            for index, points in enumerate(widths, start=1):
                if points > 0:
                    ws.column_dimensions[get_column_letter(index)].width = points_to_excel_width(points)
        wb.save(path)
    finally:
        for folder in temp_dirs:
            shutil.rmtree(folder, ignore_errors=True)


def _apply_geometry(ws, table: dict[str, Any]) -> None:
    widths = _positive_floats(table.get("col_widths_pt"))
    if widths:
        for index, points in enumerate(widths, start=1):
            if points > 0:
                ws.column_dimensions[get_column_letter(index)].width = points_to_excel_width(points)
    heights = _positive_floats(table.get("row_heights_pt"))
    if heights:
        for index, points in enumerate(heights, start=1):
            if points > 0:
                ws.row_dimensions[index].height = points


def _load_payload(raw: dict[str, Any]) -> tuple[list[dict[str, Any]], ConvertError | None]:
    sheets = raw.get("sheets")
    if isinstance(sheets, list) and sheets:
        tables = sheets
        label = "sheet"
    else:
        tables = raw.get("tables")
        label = "table"
    if not isinstance(tables, list) or not tables:
        return [], ConvertError("bad_json", "expected non-empty 'tables' array")
    cleaned: list[dict[str, Any]] = []
    for i, table in enumerate(tables, start=1):
        if not isinstance(table, dict):
            return [], ConvertError("bad_json", f"{label} {i} must be an object")
        rows = table.get("rows")
        if not isinstance(rows, list) or not rows:
            return [], ConvertError("bad_json", f"{label} {i} needs non-empty 'rows'")
        item = {
            "page": table.get("page", i),
            "rows": rows,
            "col_widths_pt": table.get("col_widths_pt"),
            "row_heights_pt": table.get("row_heights_pt"),
            "merges": table.get("merges") or [],
            "cell_styles": table.get("cell_styles") or [],
            "images": table.get("images") or [],
            "origin_x_pt": table.get("origin_x_pt") or 0,
            "origin_y_pt": table.get("origin_y_pt") or 0,
        }
        _ensure_sheet_covers_images(item)
        cleaned.append(item)
    return cleaned, None


def convert_data(
    payload: dict[str, Any],
    out_path: Path,
    *,
    normalize: bool = True,
    include_base64: bool = False,
) -> ConvertResult:
    tables, err = _load_payload(payload)
    if err:
        return ConvertResult(ok=False, errors=[err])

    wb = Workbook()
    # remove default sheet after creating first real one
    default = wb.active
    sheets: list[str] = []
    for index, table in enumerate(tables, start=1):
        name = f"Table{index}"
        ws = wb.create_sheet(title=name) if index > 1 or default is None else default
        if index == 1:
            ws.title = name
        sheets.append(name)
        for r_i, row in enumerate(table["rows"], start=1):
            if not isinstance(row, list):
                return ConvertResult(
                    ok=False,
                    errors=[ConvertError("bad_json", f"{name} row {r_i} must be a list")],
                )
            styles = table.get("cell_styles") or []
            row_styles = styles[r_i - 1] if r_i - 1 < len(styles) else []
            for c_i, cell in enumerate(row, start=1):
                value = "" if cell is None else cell
                written = ws.cell(row=r_i, column=c_i, value=value)
                style = row_styles[c_i - 1] if c_i - 1 < len(row_styles) else None
                _apply_style(written, style)
        _apply_merges(ws, table.get("merges"))
        _apply_geometry(ws, table)
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.25
        ws.page_margins.bottom = 0.25

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)

    if normalize:
        nerr = _normalize_xlsx(out_path)
        if nerr == "libreoffice_missing":
            pass
        elif nerr:
            return ConvertResult(
                ok=False,
                sheets=sheets,
                errors=[ConvertError("normalize_fail", nerr)],
                workbook_path=str(out_path),
            )
        else:
            _restore_row_heights(out_path, tables)
    _place_images(out_path, tables)
    _canonicalize_ooxml_fonts(out_path)

    b64 = None
    if include_base64:
        b64 = base64.b64encode(out_path.read_bytes()).decode("ascii")

    return ConvertResult(
        ok=True,
        sheets=sheets,
        errors=[],
        workbook_path=str(out_path),
        workbook_base64=b64,
    )


def convert_file(
    path: Path | str,
    out_path: Path | str | None = None,
    *,
    normalize: bool = True,
    include_base64: bool = False,
) -> ConvertResult:
    source = Path(path)
    if not source.exists():
        return ConvertResult(ok=False, errors=[ConvertError("not_found", f"file not found: {source}")])
    try:
        payload = json.loads(source.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        return ConvertResult(ok=False, errors=[ConvertError("bad_json", str(exc))])
    if not isinstance(payload, dict):
        return ConvertResult(ok=False, errors=[ConvertError("bad_json", "root must be an object")])

    dest = Path(out_path) if out_path else source.with_suffix(".xlsx")
    return convert_data(payload, dest, normalize=normalize, include_base64=include_base64)
