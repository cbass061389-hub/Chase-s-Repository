# tools

Small maintenance scripts for the CM Open Order Tracker (the tracker files
themselves are live business data and are not committed here).

| Script | Purpose |
|---|---|
| `tracker_new_products.py in.xlsx out.xlsx` | Adds / rebuilds the **New Products** tab: 85% completion status per 2026 launch SKU (mirrors *Launch Status 2026* in the meeting workbook), the `LaunchThreshold` setting on Config, the Summary block, Read Me and Change Log rows. |
| `jax_shipment_allocation.py packing_list.xlsx emea_alloc.xlsx launch_plan.csv out.xlsx` | JAX inbound launch allocation: per cue SKU on an HIE packing list, the launch-plan allowances for the Jacksonville channels (EMEA via JAX from the EMEA sheet; B2C, NA, Trade show, Sponsorship, Marketing from the 2026 Launch Plan export), coverage of plan, and a recommended split of the shipment (EMEA first, rest pro-rata) with what is still owed. AP and EMEA Direct excluded. |
| `tracker_embed_snapshot.py tracker.html tracker.xlsx out.html [label]` | Re-embeds the workbook as the HTML viewer's built-in snapshot after a refresh. |

Recalculate after `tracker_new_products.py` (LibreOffice headless) so cached values
are present for viewers that do not compute formulas.
