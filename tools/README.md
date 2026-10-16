# tools

Small maintenance scripts for the CM Open Order Tracker (the tracker files
themselves are live business data and are not committed here).

| Script | Purpose |
|---|---|
| `tracker_new_products.py in.xlsx out.xlsx` | Adds / rebuilds the **New Products** tab: 85% completion status per 2026 launch SKU (mirrors *Launch Status 2026* in the meeting workbook), the `LaunchThreshold` setting on Config, the Summary block, Read Me and Change Log rows. |
| `jax_shipment_allocation.py packing_list.xlsx emea_alloc.xlsx out.xlsx [npa_export.csv]` | Builds the JAX inbound launch-allocation workbook: per cue SKU, units arriving vs the Jacksonville-fulfilled channel allocations (EMEA via JAX from the EMEA sheet; B2C, Americas, Sponsorship, Tradeshow, Marketing from the NetSuite New Product Allocations export, pasted or passed), days to launch, left for stock / shortfall. EMEA Direct and ASIA/AP excluded. |
| `tracker_embed_snapshot.py tracker.html tracker.xlsx out.html [label]` | Re-embeds the workbook as the HTML viewer's built-in snapshot after a refresh. |

Recalculate after `tracker_new_products.py` (LibreOffice headless) so cached values
are present for viewers that do not compute formulas.
