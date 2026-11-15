from __future__ import annotations

import base64
import json
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any

from openpyxl import Workbook

from json_stand.models import ConvertError, ConvertResult


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


def _load_payload(raw: dict[str, Any]) -> tuple[list[dict[str, Any]], ConvertError | None]:
    tables = raw.get("tables")
    if not isinstance(tables, list) or not tables:
        return [], ConvertError("bad_json", "expected non-empty 'tables' array")
    cleaned: list[dict[str, Any]] = []
    for i, table in enumerate(tables, start=1):
        if not isinstance(table, dict):
            return [], ConvertError("bad_json", f"table {i} must be an object")
        rows = table.get("rows")
        if not isinstance(rows, list) or not rows:
            return [], ConvertError("bad_json", f"table {i} needs non-empty 'rows'")
        cleaned.append({"page": table.get("page", i), "rows": rows})
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
            for c_i, cell in enumerate(row, start=1):
                value = "" if cell is None else cell
                ws.cell(row=r_i, column=c_i, value=value)

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
