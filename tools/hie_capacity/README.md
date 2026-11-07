# HIE Capacity Actuals pipeline

Builds `reports/HIE_Capacity_Actuals.html` from HIE's weekly "Predator open order list" attachments.

| Step | Script | Input | Output |
|---|---|---|---|
| 1 | `parse_ool.py` | `ool_raw/*.txt` (one text dump per weekly open order list) | `ool_lines.csv`, `ool_plan.csv`, `ool_reports.csv` |
| 2 | `analyze.py` | `ool_lines.csv`, `ool_plan.csv` | `lines_tracked.csv`, `weekly.csv`, `overdue_sku.csv` |
| 3 | `metrics.py` | step 2 outputs + `prices.csv` (sku, category, sell_price, price_basis) | `metrics.json` |
| 4 | `build_html.py` | `metrics.json` | `reports/HIE_Capacity_Actuals.html` |

Run in order from this folder: `python3 parse_ool.py && python3 analyze.py && python3 metrics.py && python3 build_html.py`.
Set `TODAY` in `analyze.py` / `metrics.py` to the refresh date.

Key rules (see the report's "How this was built"):
- Commitment = the line's **original PI date** (first filled). HIE's "Scheduled Finished Date" is internal, and
  "Actually Finished Date" is a forecast on open lines, so neither is used as the commitment.
- A line is finished when good + B-grade + defect reaches order qty. Lines that vanish with nothing produced are
  treated as cancelled / re-split and excluded from slip.
- Revenue = units x selling price (open SO rate; else cost x product-line markup; else SKU-family median).
  "Late revenue" is revenue pushed out, not proven lost sales.

Raw attachment text and intermediate CSVs are not committed (supplier data, ~15 MB).

## HIE price changes (in progress)
`cost.py` reads weekly "Predator warehouse inventory list" text dumps (`wh_raw/*.txt`: ITEM# | PO# | QTY | Unit Price)
plus `../cap/ool_lines.csv` and `../cap/prices.csv`, dates each PO from its first appearance on the open order list
(older POs from PO number), and writes `po_prices.csv` and `cost.json` (unit-weighted, chain-linked YoY price index;
$ impact = price change x units ordered). `build_html.py` adds the price section when `cost.json` is present.
