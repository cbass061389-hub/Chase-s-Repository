# tools

Small maintenance scripts for the CM Open Order Tracker (the tracker files
themselves are live business data and are not committed here).

| Script | Purpose |
|---|---|
| `tracker_new_products.py in.xlsx out.xlsx` | Adds / rebuilds the **New Products** tab: 85% completion status per 2026 launch SKU (mirrors *Launch Status 2026* in the meeting workbook), the `LaunchThreshold` setting on Config, the Summary block, Read Me and Change Log rows. |
| `jax_shipment_allocation.py packing_list.xlsx emea_alloc.xlsx launch_plan.csv out.xlsx [plan_asks.csv] [ask_col]` | JAX inbound launch allocation: per cue SKU on an HIE packing list, what arrived vs what the shipment plan requested for that sailing (remaining JAX need net of receipts and AP share, 85% rule), the launch-plan allowances per Jacksonville channel, and a recommended split (EMEA via JAX first, rest pro-rata). AP and EMEA Direct excluded. `plan_asks.csv` is exported from the Shipment Request workbook's 'JAX Shipments by Date' tab. |
| `tracker_embed_snapshot.py tracker.html tracker.xlsx out.html [label]` | Re-embeds the workbook as the HTML viewer's built-in snapshot after a refresh. |

Recalculate after `tracker_new_products.py` (LibreOffice headless) so cached values
are present for viewers that do not compute formulas.
