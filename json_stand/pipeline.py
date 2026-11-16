from __future__ import annotations

import base64
import json
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from json_stand.models import ConvertError, ConvertResult

# LibreOffice writes Calibri 11 back into the sheet and maps one OOXML width
# unit to 5.55 pt (19.579 cm per 100 units). Dividing PDF points by this keeps
# the printed column the same width as the PDF column.
EXCEL_PT_PER_WIDTH = 5.55


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
        cleaned.append(
            {
                "page": table.get("page", i),
                "rows": rows,
                "col_widths_pt": table.get("col_widths_pt"),
                "row_heights_pt": table.get("row_heights_pt"),
                "merges": table.get("merges") or [],
                "cell_styles": table.get("cell_styles") or [],
            }
        )
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

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)

    if normalize:
        nerr = _normalize_xlsx(out_path)
        if nerr == "libreoffice_missing":
            # still ok for API without LO; CI installs soffice
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
