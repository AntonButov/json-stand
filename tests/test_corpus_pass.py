from __future__ import annotations

import json
import shutil
from pathlib import Path

import pytest

from json_stand.pipeline import convert_file

ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "tests" / "corpus" / "manifest.json"
PASS = ROOT / "tests" / "corpus" / "pass"


def _entries() -> list[dict]:
    return json.loads(MANIFEST.read_text(encoding="utf-8"))["pass"]


@pytest.fixture(scope="session")
def require_soffice():
    if not shutil.which("soffice") and not Path(
        "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    ).exists():
        pytest.skip("LibreOffice soffice not available")


@pytest.fixture(scope="session")
def excel_check():
    try:
        from excel_stand.pipeline import check_file
    except ImportError as exc:
        pytest.skip(f"excel-stand not installed: {exc}")
    return check_file


@pytest.mark.libreoffice
@pytest.mark.parametrize("entry", _entries(), ids=lambda e: e["id"])
def test_corpus_convert_and_excel_stand(entry: dict, require_soffice, excel_check, tmp_path: Path):
    src = PASS / entry["file"]
    assert src.exists()
    out = tmp_path / f"{entry['id']}.xlsx"
    converted = convert_file(src, out_path=out, normalize=True)
    assert converted.ok, converted.to_dict()
    assert out.exists() and out.stat().st_size > 0

    checked = excel_check(out, out_dir=tmp_path / f"check-{entry['id']}", include_base64=False, timeout_s=120)
    assert checked.ok, checked.to_dict()


def test_manifest_has_20():
    assert len(_entries()) == 20
