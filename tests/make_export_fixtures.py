"""Generate export CSVs from the recovered schemas, containing the real hazards.

Deliberately includes every failure mode the estate exhibits, so the readers
and gates are exercised against them rather than against clean data:

* a text field containing a comma inside quotes — the case ``QuoteStyle.None``
  corrupts,
* matrix item names as ``PARENT : CHILD`` with stray whitespace, plus one
  differing only by case,
* an inventory location absent from ``canonical.locations``,
* quantities written with thousands separators and in accounting parentheses,
* a ``Quantity Remaining`` that disagrees with ordered minus received,
* a blank item name, which must become a null key rather than an empty string.
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path
from typing import Dict, List, Sequence

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from sc.extract.schema_store import SourceSchema, load_schemas

INVENTORY_ROWS: List[Dict[str, object]] = [
    {"Internal ID": 101, "Inventory Location": "US B2B", "Name": "PREDATOR : PRED-BK-01 ",
     "Product Category": "Cues, Playing", "On Hand": "1,204", "Available": "1,100",
     "On Sales Order": 104, "On Purchase Order": 500, "Average Cost": "88.50",
     "Safety Stock Level": 100, "Location Reorder Point": 250, "Lead Time": 90},
    {"Internal ID": 102, "Inventory Location": "HIE", "Name": "predator : pred-bk-01",
     "Product Category": "Cues, Playing", "On Hand": "300", "Available": "300",
     "On Sales Order": 0, "On Purchase Order": 0, "Average Cost": "88.50",
     "Safety Stock Level": 0, "Location Reorder Point": 0, "Lead Time": 90},
    {"Internal ID": 103, "Inventory Location": "US B2B", "Name": "REVO-12.4",
     "Product Category": "Shafts, REVO", "On Hand": "(35)", "Available": "0",
     "On Sales Order": 35, "On Purchase Order": 200, "Average Cost": "310.00",
     "Safety Stock Level": 50, "Location Reorder Point": 75, "Lead Time": 120},
    {"Internal ID": 104, "Inventory Location": "Narnia", "Name": "POISON : PSN-CUE-3",
     "Product Category": "Cues, Poison", "On Hand": "40", "Available": "40",
     "On Sales Order": 0, "On Purchase Order": 0, "Average Cost": "45.00",
     "Safety Stock Level": 10, "Location Reorder Point": 20, "Lead Time": 60},
    {"Internal ID": 105, "Inventory Location": "INT B2C", "Name": "   ",
     "Product Category": "Accessories", "On Hand": "12", "Available": "12",
     "On Sales Order": 0, "On Purchase Order": 0, "Average Cost": "5.00",
     "Safety Stock Level": 0, "Location Reorder Point": 0, "Lead Time": 30},
]

OPEN_PO_ROWS: List[Dict[str, object]] = [
    {"Date": "2026-01-15", "Expected Receipt Date": "2026-04-01", "Original PI Date": "2026-03-15",
     "Legacy PO #": "L-88", "Status": "Pending Receipt", "Document Number": "PO1001",
     "Name": "HIE/Hamson", "Inventory Location": "US B2B", "Line ID": 1,
     "Product Category": "Cues, Playing", "Item": "PREDATOR : PRED-BK-01",
     "Quantity": "500", "Quantity Fulfilled/Received": "200", "Quantity Remaining": "300",
     "Amount": "44,250.00", "$ Remaining": "26550.00", "Closed": "No",
     "Memo (Main)": "Spring wave, expedite", "Item Note": "matte finish"},
    {"Date": "2026-01-15", "Expected Receipt Date": "2026-04-01", "Original PI Date": "2026-03-15",
     "Legacy PO #": "L-88", "Status": "Pending Receipt", "Document Number": "PO1001",
     "Name": "HIE/Hamson", "Inventory Location": "US B2B", "Line ID": 2,
     "Product Category": "Shafts, REVO", "Item": "REVO-12.4",
     "Quantity": "200", "Quantity Fulfilled/Received": "0", "Quantity Remaining": "150",
     "Amount": "62,000.00", "$ Remaining": "62000.00", "Closed": "No",
     "Memo (Main)": "Note with a comma, and more text", "Item Note": ""},
    {"Date": "2025-11-02", "Expected Receipt Date": "2026-01-20", "Original PI Date": "2026-01-05",
     "Legacy PO #": "", "Status": "Closed", "Document Number": "PO1002",
     "Name": "Jingdian", "Inventory Location": "HIE", "Line ID": 1,
     "Product Category": "Cues, Poison", "Item": "POISON : PSN-CUE-3",
     "Quantity": "80", "Quantity Fulfilled/Received": "80", "Quantity Remaining": "0",
     "Amount": "3600.00", "$ Remaining": "0", "Closed": "Yes",
     "Memo (Main)": "", "Item Note": ""},
]

ALLOCATION_ROWS: List[Dict[str, object]] = [
    {"Internal ID": 900, "Name": "PREDATOR : PRED-NP-01", "Description": "New cue, flagship",
     "Base Price": "499.00", "Targeted Launch Date": "2026-10-01",
     "B2C": 120, "Sponsorships": 20, "EMEA": 200, "Americas": 350, "ASIA": 80,
     "Tradeshows": 30, "Marketing": 15},
    {"Internal ID": 901, "Name": "REVO : REVO-NP-12.9", "Description": "Shaft, carbon",
     "Base Price": "379.00", "Targeted Launch Date": "2026-11-15",
     "B2C": 60, "Sponsorships": 10, "EMEA": 90, "Americas": 140, "ASIA": 40,
     "Tradeshows": 12, "Marketing": 8},
]

ROW_SETS: Dict[str, List[Dict[str, object]]] = {
    "netsuite_2600947": INVENTORY_ROWS,
    "netsuite_2600946": OPEN_PO_ROWS,
    "netsuite_2620991": ALLOCATION_ROWS,
}


def write_export(schema: SourceSchema, rows: Sequence[Dict[str, object]], drop_dir: Path) -> Path:
    """Write one export with every declared column present, in declared order."""
    drop_dir.mkdir(parents=True, exist_ok=True)
    target: Path = drop_dir / f"{schema.slug}.csv"
    with target.open("w", encoding="utf-8-sig", newline="") as handle:
        # QUOTE_MINIMAL, so the comma-bearing fields are genuinely quoted and the
        # reader has to honour the quoting to parse them.
        writer = csv.DictWriter(handle, fieldnames=schema.column_names, quoting=csv.QUOTE_MINIMAL)
        writer.writeheader()
        for row in rows:
            writer.writerow({name: row.get(name, "") for name in schema.column_names})
    return target


def build(schema_path: Path, drop_dir: Path) -> List[Path]:
    written: List[Path] = []
    for schema in load_schemas(schema_path):
        rows = ROW_SETS.get(schema.slug)
        if rows is None:
            continue
        written.append(write_export(schema, rows, drop_dir))
    return written


if __name__ == "__main__":
    target_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("data/drop")
    for path in build(Path("sc/export_schemas.yaml"), target_dir):
        print(f"wrote {path}")
