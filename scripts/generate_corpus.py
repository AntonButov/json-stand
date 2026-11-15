#!/usr/bin/env python3
"""Write 20 pdf-stand-shaped JSON fixtures + manifest."""

from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PASS = ROOT / "tests" / "corpus" / "pass"
MANIFEST = ROOT / "tests" / "corpus" / "manifest.json"


def dump(name: str, tables: list[dict]) -> None:
    path = PASS / name
    path.write_text(json.dumps({"tables": tables}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("wrote", path.relative_to(ROOT))


def t(page: int, rows: list[list[str | None]]) -> dict:
    return {"page": page, "rows": rows}


def main() -> int:
    PASS.mkdir(parents=True, exist_ok=True)

    dump("01_simple.json", [t(1, [["Name", "Qty", "Price"], ["Apple", "3", "1.5"], ["Banana", "5", "0.8"], ["Cherry", "2", "3.2"]])])
    dump(
        "02_two_tables.json",
        [
            t(1, [["A", "B"], ["1", "2"]]),
            t(1, [["SKU", "Stock"], ["X1", "100"], ["X2", "50"]]),
        ],
    )
    dump("03_wide.json", [t(1, [[f"Col{i}" for i in range(1, 9)], [str(i) for i in range(1, 9)]])])
    dump("04_tall.json", [t(1, [["id", "value"]] + [[str(i), str(i * 1.5)] for i in range(1, 26)])])
    dump("05_russian.json", [t(1, [["Товар", "Кол", "Цена"], ["Яблоки", "10", "120.5"], ["Молоко", "2", "89"]])])
    dump("06_numbers.json", [t(1, [["Number", "Percent", "Currency"], ["1234.56", "15%", "$99.99"]])])
    dump("07_empty_cells.json", [t(1, [["A", "B", "C"], ["x", "", "z"], ["", "y", ""]])])
    dump(
        "08_invoice.json",
        [t(1, [["Item", "Qty", "Amount"], ["Service A", "1", "1000"], ["Service B", "2", "500"], ["Total", "", "2000"]])],
    )
    dump("09_dates.json", [t(1, [["Date", "Event"], ["2026-11-14", "Launch"], ["2026-12-01", "Review"]])])
    dump("10_styled_header.json", [t(1, [["Product", "Score"], ["Alpha", "9"], ["Beta", "7"]])])
    dump(
        "11_cities.json",
        [t(1, [["City", "Pop", "Region"], ["Moscow", "12", "C"], ["SPb", "5", "NW"], ["Kazan", "1", "V"]])],
    )
    dump(
        "12_links_like.json",
        [t(1, [["Label", "URL"], ["Docs", "https://example.com"], ["Home", "https://example.org"]])],
    )
    dump(
        "13_mixed.json",
        [t(1, [["Label", "Amount", "Note"], ["total", "1000.00", "ok"], ["rate", "25.50%", "quarter"]])],
    )
    dump("14_single_row.json", [t(1, [["Only", "Header"], ["one", "row"]])])
    dump(
        "15_inventory.json",
        [t(1, [["SKU", "Name", "Qty"], ["A-1", "Bolt", "100"], ["A-2", "Nut", "200"], ["B-1", "Washer", "500"]])],
    )
    dump(
        "16_multipage.json",
        [
            t(1, [["Page", "Item"], ["1", "First"]]),
            t(2, [["Page", "Item"], ["2", "Second"], ["2", "Third"]]),
        ],
    )
    dump(
        "17_budget.json",
        [t(1, [["Category", "Plan", "Fact"], ["Dev", "100", "90"], ["Ops", "50", "55"], ["Sales", "200", "180"]])],
    )
    dump("18_employees.json", [t(1, [["Name", "Role"], ["Ann", "Eng"], ["Bob", "PM"], ["Cid", "QA"]])])
    dump("19_metrics.json", [t(1, [["Metric", "Value"], ["CTR", "2.1%"], ["CR", "0.8%"], ["ARPU", "12.5"]])])
    dump(
        "20_orders.json",
        [t(1, [["Order", "Status", "Sum"], ["1001", "paid", "10"], ["1002", "new", "20"], ["1003", "paid", "15"]])],
    )

    entries = []
    for path in sorted(PASS.glob("*.json")):
        entries.append({"id": path.stem.split("_", 1)[0], "file": path.name, "expect": "pass"})
    MANIFEST.write_text(json.dumps({"pass": entries}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("wrote", MANIFEST.relative_to(ROOT), "count", len(entries))
    return 0 if len(entries) == 20 else 1


if __name__ == "__main__":
    raise SystemExit(main())
