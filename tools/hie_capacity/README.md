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

## HIE PI ledger and price changes
- `pi_ledger.py` merges parsed HIE PI confirmations and monthly "Pending order list" lines (`../pi/pi_lines_*.tsv`) into
  `pi_ledger.csv` (first and latest quote per PO line), `pi_prices.csv` (quoted price per SKU/PO) and
  `pi_first_dates.json` (PI dates on HIE's first quote). `analyze.py` uses those dates as the commitment where available.
- `cost.py` builds the per-SKU, per-PO price history: HIE's latest quoted price first (what HIE charges), warehouse-list
  price as fallback. Price year = year of the PO's earliest PI date (HIE prices by PI date), from the PI ledger, else the
  open order list, else PO-number interpolation. Future price years are excluded. HIE-announced increases (tables with
  the prior year's final price) are extracted separately as a check. Output `cost.json`: unit-weighted,
  chain-linked YoY index (years with <5 comparable SKUs skipped) and $ impact = price change x units ordered.
- `cancel_lines.tsv` (cancellations, reductions, holds, HIE responses and charges) is rendered as-is.

Run order: `pi_ledger.py` -> `parse_ool.py` -> `analyze.py` -> `metrics.py` -> `cost.py` -> `build_html.py`.
