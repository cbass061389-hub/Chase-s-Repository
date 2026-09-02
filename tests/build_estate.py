"""Assemble a simulated OneDrive estate that reproduces the real failure modes.

Deliberately includes: a header band that does not start at row 1, hidden
staging sheets, a Power Query stack reading from a Downloads folder, a
password-protected VBA project, an .xlsb, a NetSuite CSV export, two files
holding the same columns with different row counts, and files that must be
skipped (Archive folder, lock file).
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from tests.make_fixtures import build_vba_project, write_xlsb, write_xlsx

IN_TRANSIT_HEADERS = ["Container #", "Invoice No", "SKU", "Qty", "ETD", "ETA", "POL", "POD", "Region"]

MASHUP_SECTION = r'''section Section1;

shared #"Items on Purchase Order" = let
    Source = Csv.Document(File.Contents("C:\Users\CharlesBass\Downloads\ItemsOnPurchaseOrder.csv"),[Delimiter=","]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    // a semicolon in a comment ; must not split the query
    Filtered = Table.SelectRows(Promoted, each [Status] <> "Closed")
in
    Filtered;

shared InTransit_Lines = let
    Src = Excel.Workbook(File.Contents("\\PGFS01\Supply Chain\In Transit Detail.xlsx"), null, true),
    Data = Src{[Item="Lines",Kind="Sheet"]}[Data],
    Joined = Table.NestedJoin(Data, {"SKU"}, #"Items on Purchase Order", {"SKU"}, "PO", JoinKind.LeftOuter)
in
    Joined;

shared SO_Data = let
    Source = SharePoint.Files("https://predatorgroup.sharepoint.com/sites/SupplyChain", [ApiVersion = 15])
in
    Source;
'''


def build(root: Path) -> Path:
    root.mkdir(parents=True, exist_ok=True)

    transit_rows = [IN_TRANSIT_HEADERS] + [
        [f"MSCU{700000 + i}", f"INV-{2600 + i}", f"PRED-BK-{i % 9:02d}", 100 + i * 7,
         "2026-08-01", "2026-09-18", "Yantian", "Savannah", "Americas"]
        for i in range(140)
    ]
    write_xlsx(
        root / "In Transit" / "In Transit.xlsx",
        [("InTransit", transit_rows, "visible"),
         ("Staging_DoNotTouch", [["helper"], ["x"]], "hidden")],
        mashup_section=MASHUP_SECTION,
        tables={"InTransit": ("tbl_InTransit", f"A1:I{len(transit_rows)}", IN_TRANSIT_HEADERS)},
        connections=[{"name": "NetSuite_ItemsOnPO", "type": "5",
                      "connection_string": r"Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$",
                      "command": "SELECT * FROM [Items on Purchase Order]"}],
        defined_names=[("InTransit_Range", "InTransit!$A$1:$I$140")],
    )

    # Same columns, fewer rows, plus one extra column -> duplicate-truth conflict.
    dashboard_rows = [IN_TRANSIT_HEADERS + ["Days Late"]] + [
        row + [max(0, i - 100)] for i, row in enumerate(transit_rows[1:118])
    ]
    write_xlsx(
        root / "Tables" / "Table In-Transit Dashboard.xlsx",
        [("Dashboard", dashboard_rows, "visible")],
        external_link_target=r"file:///\\PGFS01\Supply Chain\In Transit.xlsx",
    )

    # Header band starts at row 4, under a title block.
    meeting_rows = [
        ["Supply Chain Update Meeting", None, None, None, None, None],
        ["Week of 2026-08-31", None, None, None, None, None],
        [None, None, None, None, None, None],
        ["SKU", "Months Supply", "Forecast 3M", "On Hand", "Worst Flag", "ABC Class"],
    ] + [
        [f"PRED-BK-{i % 9:02d}", 3.5 + i * 0.4, 120 - i, 900 + i * 3,
         "OVERSTOCK" if i % 3 else "CAPITAL TRAP", "A" if i < 12 else "B"]
        for i in range(96)
    ]
    write_xlsx(
        root / "Meetings" / "Supply_Chain_Update_Meeting_Workbook.xlsm",
        [("Inventory Health", meeting_rows, "visible"),
         ("Calc_MonthsSupply", [["SKU", "Calc"]] + [[f"PRED-BK-{i % 9:02d}", i] for i in range(96)], "visible"),
         ("_lookup", [["key", "value"], ["a", 1]], "veryHidden")],
        vba=build_vba_project("SCUpdate", ["mod_Refresh", "mod_Flags", "mod_Export"],
                              ["cls_Allocation"], protected=True),
        formula_cells={
            "Calc_MonthsSupply": [(r, 1, f"VLOOKUP(A{r},'Inventory Health'!$A:$F,2,FALSE)")
                                  for r in range(2, 98)],
            "Inventory Health": [(r, 4, f"IF(B{r}>8,\"OVERSTOCK\",\"MONITOR\")") for r in range(5, 101)],
        },
    )

    write_xlsb(
        root / "Overstock" / "Overstock_Static_Inventory.xlsb",
        [("Overstock", [["SKU", "Warehouse", "On Hand", "Months Supply", "Extended Cost", "UOM"]]
          + [[f"PRED-CUE-{i:03d}", "JAX-01", str(400 + i), str(9.0 + i * 0.2), str(12000 + i * 90), "EA"]
             for i in range(60)], "visible"),
         ("Notes", [["note"], ["reviewed 2026-08"]], "hidden")],
    )

    exports = root / "NetSuite Exports"
    exports.mkdir(parents=True, exist_ok=True)
    (exports / "ItemsOnPurchaseOrder.csv").write_bytes(
        ("\ufeffPO #,Line,Item,Qty Ordered,Qty Received,Committed Date,Vendor\r\n"
         + "".join(
             f"PO{1000 + i},1,PRED-BK-{i % 9:02d},{500 + i},{i * 3},2026-1{i % 2}-05,HIE\r\n"
             for i in range(75))).encode("utf-8")
    )

    prototypes = root / "Prototypes"
    prototypes.mkdir(parents=True, exist_ok=True)
    (prototypes / "build_production_visibility.py").write_text(
        "# existing prototype: production visibility export\n", encoding="utf-8"
    )
    (prototypes / "Build-RealData.ps1").write_text(
        "# writes real-data.js\nWrite-Output 'stub'\n", encoding="utf-8"
    )

    # --- must be skipped by the sweep ---
    write_xlsx(root / "Archive" / "In Transit OLD.xlsx", [("Sheet1", [["x"], ["y"]], "visible")])
    write_xlsx(root / "In Transit" / "~$In Transit.xlsx", [("Sheet1", [["x"], ["y"]], "visible")])
    write_xlsx(root / "Backup" / "Master_Reference.xlsx", [("Sheet1", [["x"], ["y"]], "visible")])
    return root


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("/tmp/estate")
    print(f"built estate at {build(target)}")
