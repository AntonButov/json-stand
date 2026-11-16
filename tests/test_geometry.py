import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from json_stand.pipeline import convert_data

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _sheet(path: Path) -> ET.Element:
    xml = zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml")
    return ET.fromstring(xml)


def test_sheet_sizes_match_pdf_points(tmp_path: Path):
    out = tmp_path / "sized.xlsx"
    result = convert_data(
        {
            "tables": [
                {
                    "page": 1,
                    "rows": [["A", "B"], ["1", "2"]],
                    "col_widths_pt": [72.0, 37.5],
                    "row_heights_pt": [18.0, 12.0],
                }
            ]
        },
        out,
        normalize=True,
    )
    assert result.ok, result.to_dict()
    root = _sheet(out)
    widths = [float(col.get("width")) for col in root.find("m:cols", NS)]
    heights = [float(row.get("ht")) for row in root.find("m:sheetData", NS)]
    # 1 excel width unit = 5.55 pt under the Calibri 11 sheet LibreOffice writes back.
    assert widths == [12.97, 6.76]
    assert heights == [18.0, 12.0]


def test_sheet_keeps_font_merge_and_exact_row_height(tmp_path: Path):
    out = tmp_path / "styled.xlsx"
    result = convert_data(
        {
            "sheets": [
                {
                    "page": 1,
                    "rows": [["ПАО Сбербанк", "БИК"], [None, "Сч. №"], ["№", "Брелок"]],
                    "col_widths_pt": [100.0, 40.0],
                    "row_heights_pt": [12.0, 18.0, 9.0],
                    "merges": ["A1:A2"],
                    "cell_styles": [
                        [
                            {"size": 8.5, "bold": True, "font": "Arial", "halign": "left", "valign": "top", "wrap": True},
                            {"size": 8.5, "bold": False, "font": "Arial", "halign": "left", "valign": "top", "wrap": False},
                        ],
                        [
                            None,
                            {"size": 8.5, "bold": False, "font": "Arial", "halign": "left", "valign": "top", "wrap": False},
                        ],
                        [
                            {"size": 6.5, "bold": False, "font": "Arial", "halign": "left", "valign": "top", "wrap": False},
                            {"size": 6.5, "bold": False, "font": "Arial", "halign": "left", "valign": "top", "wrap": False},
                        ],
                    ],
                }
            ]
        },
        out,
        normalize=True,
    )
    assert result.ok, result.to_dict()
    root = _sheet(out)
    heights = [float(row.get("ht")) for row in root.find("m:sheetData", NS)]
    assert heights == [12.0, 18.0, 9.0]
    merges = [node.get("ref") for node in root.find("m:mergeCells", NS)]
    assert "A1:A2" in merges
    styles = ET.fromstring(zipfile.ZipFile(out).read("xl/styles.xml"))
    fonts = []
    for font in styles.find("m:fonts", NS):
        name = font.find("m:name", NS)
        size = font.find("m:sz", NS)
        fonts.append(
            (
                None if name is None else name.get("val"),
                None if size is None else float(size.get("val")),
                font.find("m:b", NS) is not None,
            )
        )
    assert ("Arial", 8.5, True) in fonts
    assert ("Arial", 6.5, False) in fonts
