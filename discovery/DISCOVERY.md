# Phase 0 — Discovery

Generated 2026-09-02T13:00:03+00:00 in 30.69s from `/home/user/Chase-s-Repository/sc/config.yaml`.
Read-only sweep. No source workbook was opened in Excel, written to, or had a macro executed.

**42 secret(s) redacted** before writing (netsuite_media_token x42). Power Query text embeds credentials — a NetSuite File Cabinet link carries its access token in `h=` — and `queries/` is committed, so tokens are stripped from every written file and from every location in this report. The endpoint stays visible; the credential does not.

## Executive summary

- **3 sources** found across 1 reachable root.
- **3 workbooks carry Power Query** (23 M files now version-controlled under `queries/`).
- **3 workbooks carry VBA** — logic that lives outside the query stack.
- **2 sources rank high or critical** on dependency risk.
- **4 duplicate-truth candidates** (2 of them disagree on row count).
- **14 of 17 expected workbooks were not found.**

### Key risks

1. **Circular refresh dependency — refresh order is undefined.** Supply_Chain_Update_Meeting_Workbook.xlsm -> Shipment_Request_HIE__Updated.xlsm -> Supply_Chain_Update_Meeting_Workbook.xlsm. Whichever workbook refreshes second reads the other's stale output, so the numbers change depending on the order somebody happened to click. This has to be broken before anything downstream can be trusted.
2. **One workbook reading two versions of the same file.** `Supply_Chain_Update_Meeting_Workbook.xlsm` pulls from `Shipment Request HIE - Updated.xlsm`, `Shipment Request HIE .xlsm`. Part of its data is sourced from a superseded file, which is a live wrong-number bug, not a tidiness issue.
3. **5 upstream export(s) feed more than one workbook, each with its own transformation.** Worst: NetSuite File Cabinet export id=2600949 (account 3492685) is consumed by 3 workbooks. This is the root cause of two reports disagreeing — they do not disagree about the data, they disagree about the logic. Consolidating the copies without consolidating the logic fixes nothing.
4. **Two sources claim the same data with different numbers.** Worst case: `Supply_Chain_Update_Meeting_Workbook.xlsm!HIE Production Schedule` (774 rows) vs `Shipment_Request_HIE__Updated.xlsm!HIEProdFlat` (546 rows) — 100% header overlap, 228 row difference. Until one is named the source, every downstream number is arguable.
5. **Fragile paths.** `Supply_Chain_Update_Meeting_Workbook.xlsm` scores 164 (critical): hardcoded per-user local path — breaks on any other machine (2 locations, e.g. C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Chase\Shipment Request HIE .xlsm); SharePoint URL — survives moves better, but needs auth at refresh -> https://predatorgroup.sharepoint.com/Share%20All%20Files/Forecast/Reference/Actuals%20vs%20Forecast%20Archive/Sales%20Actuals%20+%20Forecast%20for%20Actuals%20vs%20Forecast.xlsx
6. **14 expected workbooks not located** — either renamed, in a skipped folder, or outside the configured roots. Named below; each needs a path before Phase 2 can claim coverage.

### Recommended action

1. **Break the refresh cycle.** Name one workbook as upstream and have the other read from the warehouse instead of from it directly. Nothing else here is worth fixing while the graph has a loop in it.
2. **Pull the forked upstreams into one extraction per export.** Each shared export should be read once, transformed once, and reused — the per-workbook query logic is what has to be reconciled, not the output.
3. **Retire the superseded file(s)** listed under version skew, and repoint the queries still reading them.
4. Confirm the domain and grain calls in the sheet-level entity map, especially every key marked `needs confirmation`.
5. Sign off on the canonical model in `SCHEMA.md` — Phase 2 does not start before that.

## Sources, ranked worst-first

`ROLE` is inferred from evidence in the file — formula density, query sources and external links — not from its name. A `COPY` must not be read by Phase 2 when its source is also present.

| File | Domain | Role | Grain (inferred) | Refresh | Risk | Rows | Sheets | Probe |
|---|---|---|---|---|---|---|---|---|
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | inventory | HYBRID | po_number x container x invoice x sku x region_channel x dat | power_query (11 queries) + workbook_connecti | 164 critical | 167,485 | 76 (51 hidden) | ok |
| `Shipment_Request_HIE__Updated.xlsm` | production | COPY | po_number x invoice x sku x region_channel x date | power_query (8 queries) + workbook_connectio | 105 critical | 1,607,405 | 36 (7 hidden) | ok |
| `Direct_Shipments_Consolidator.xlsm` | dsr | COPY | sku x region_channel x date | power_query (1 queries) + workbook_connectio | 28 medium | 3,887 | 13 (0 hidden) | ok |

## Refresh dependency graph

Every query, and the canonical upstream it reads. Upstreams are identified by what they *are* rather than by URL text, so two queries hitting the same export land on the same node even when their URLs differ.

| Workbook | Query | Upstream kind | Upstream | Is another swept file |
|---|---|---|---|---|
| `Direct_Shipments_Consolidator.xlsm` | Open Order | netsuite_file_cabinet | NetSuite File Cabinet export id=2600949 (account 3492685) |  |
| `Shipment_Request_HIE__Updated.xlsm` | HIE Shipment | file | Supply Chain Update Meeting Workbook.xlsm | yes |
| `Shipment_Request_HIE__Updated.xlsm` | Location_Inventory | netsuite_file_cabinet | NetSuite File Cabinet export id=2600947 (account 3492685) |  |
| `Shipment_Request_HIE__Updated.xlsm` | NetSuite Forecast | netsuite_file_cabinet | NetSuite File Cabinet export id=2606291 (account 3492685) |  |
| `Shipment_Request_HIE__Updated.xlsm` | New Prod Alloc | netsuite_file_cabinet | NetSuite File Cabinet export id=2620991 (account 3492685) |  |
| `Shipment_Request_HIE__Updated.xlsm` | Open_Order_Report | netsuite_file_cabinet | NetSuite File Cabinet export id=2600949 (account 3492685) |  |
| `Shipment_Request_HIE__Updated.xlsm` | PO | netsuite_file_cabinet | NetSuite File Cabinet export id=2600946 (account 3492685) |  |
| `Shipment_Request_HIE__Updated.xlsm` | Regional Inv Loc | netsuite_file_cabinet | NetSuite File Cabinet export id=2600947 (account 3492685) |  |
| `Shipment_Request_HIE__Updated.xlsm` | Sales | netsuite_file_cabinet | NetSuite File Cabinet export id=2600950 (account 3492685) |  |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | Current Inventory | netsuite_file_cabinet | NetSuite File Cabinet export id=2600947 (account 3492685) |  |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | Customer Open Orders | netsuite_file_cabinet | NetSuite File Cabinet export id=2600949 (account 3492685) |  |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | Forecast Qty | netsuite_file_cabinet | NetSuite File Cabinet export id=2606291 (account 3492685) |  |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | HIEInv | file | Shipment Request HIE .xlsm |  |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | HIEProdFlat | file | Shipment Request HIE - Updated.xlsm | yes |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | Items On PO | netsuite_file_cabinet | NetSuite File Cabinet export id=2600946 (account 3492685) |  |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | JingdianProdFlat | file | Shipment Request HIE - Updated.xlsm | yes |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | New Product Allocations | netsuite_file_cabinet | NetSuite File Cabinet export id=2620991 (account 3492685) |  |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | PoisonInvFlat | file | Shipment Request HIE - Updated.xlsm | yes |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | Purchase Price | netsuite_file_cabinet | NetSuite File Cabinet export id=2613629 (account 3492685) |  |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | ValidatedSales | sharepoint | https://predatorgroup.sharepoint.com/Share%20All%20Files/Forecast/Reference/Ac |  |

### Circular dependencies

- `Supply_Chain_Update_Meeting_Workbook.xlsm` -> `Shipment_Request_HIE__Updated.xlsm` -> `Supply_Chain_Update_Meeting_Workbook.xlsm`

Refresh order is undefined in a cycle. Whichever side refreshes second is reading the other's previous output.

### Version skew — one workbook, two versions of the same file

- `Supply_Chain_Update_Meeting_Workbook.xlsm`
  - `Shipment Request HIE - Updated.xlsm` <- queries: HIEProdFlat, JingdianProdFlat, PoisonInvFlat
  - `Shipment Request HIE .xlsm` <- queries: HIEInv

### Forked upstreams — one export, many transformations

This is where duplicate numbers come from. One export, read independently by several workbooks, each applying its own logic.

| Upstream | Kind | Workbooks | Consumed by |
|---|---|---|---|
| NetSuite File Cabinet export id=2600949 (account 3492685) | netsuite_file_cabinet | 3 | `Supply_Chain_Update_Meeting_Workbook.xlsm` (Customer Open Orders); `Shipment_Request_HIE__Updated.xlsm` (Open_Order_Report); `Direct_Shipments_Consolidator.xlsm` (Open Order) |
| NetSuite File Cabinet export id=2600947 (account 3492685) | netsuite_file_cabinet | 2 | `Supply_Chain_Update_Meeting_Workbook.xlsm` (Current Inventory); `Shipment_Request_HIE__Updated.xlsm` (Location_Inventory); `Shipment_Request_HIE__Updated.xlsm` (Regional Inv Loc) |
| NetSuite File Cabinet export id=2600946 (account 3492685) | netsuite_file_cabinet | 2 | `Supply_Chain_Update_Meeting_Workbook.xlsm` (Items On PO); `Shipment_Request_HIE__Updated.xlsm` (PO) |
| NetSuite File Cabinet export id=2620991 (account 3492685) | netsuite_file_cabinet | 2 | `Supply_Chain_Update_Meeting_Workbook.xlsm` (New Product Allocations); `Shipment_Request_HIE__Updated.xlsm` (New Prod Alloc) |
| NetSuite File Cabinet export id=2606291 (account 3492685) | netsuite_file_cabinet | 2 | `Supply_Chain_Update_Meeting_Workbook.xlsm` (Forecast Qty); `Shipment_Request_HIE__Updated.xlsm` (NetSuite Forecast) |

### Referenced files that were not in the sweep

- `Supply_Chain_Update_Meeting_Workbook.xlsm` query `HIEInv` -> `C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Chase\Shipment Request HIE .xlsm`
  - referenced workbook was not in the sweep — add its folder to discovery.roots or confirm the path is dead

## Sheet-level entity map

For a workbook with dozens of tabs the sheet, not the file, is the unit of consolidation — each tab is a different entity. Keys marked **?** scored below the confirmation threshold and are guesses; correct them rather than trusting them.

### `Supply_Chain_Update_Meeting_Workbook.xlsm` — 73 populated of 76 sheets

| Sheet | State | Domain | Data rows | Cols | Formulas | Hdr row | Detected keys |
|---|---|---|---|---|---|---|---|
| Forecast Qty |  | demand | 43,193 | 9 | 0 | 1 | sku=Item, location=Sub Location**?**, quantity=Forecast Qty**?**, region_channel=REVO Region**?** |
| ValidatedSales | hidden | unclassified | 38,458 | 14 | 0 | 1 | quantity=Units, sku=SKU, location=Location_lvl_2 |
| Purchase Price | hidden | supplier | 13,912 | 4 | 0 | 1 | sku=SKU, currency=Average Cost**?** |
| BOMMaster | hidden | supplier | 11,789 | 13 | 35,367 | 1 | date=Effective Date**?**, po_line=Member Line**?**, quantity=Member Quantity**?**, sku=Member Item**?** |
| __CleanBO | hidden | unclassified | 9,362 | 7 | 37,451 | 1 | sku=Item, region_channel=Channel (Regional ID), quantity=Live BO Qty**?** |
| Open Order Report | hidden | unclassified | 6,069 | 15 | 0 | 1 | date=Ship Date, po_number=Document Number, sku=Item |
| __Alloc Engine | veryHidden | inventory | 5,136 | 31 | 24,528 | 1 | quantity=On Hand, region_channel=Channel, sku=SKU, currency=Avg Cost**?** |
| Current Inventory | hidden | inventory | 4,814 | 13 | 0 | 1 | sku=SKU, region_channel=REVO Region**?** |
| Allocation Plan | veryHidden | allocation | 4,753 | 10 | 47,534 | 17 | region_channel=Channel, sku=SKU, quantity=Allocated Units**?** |
| Dash Lists |  | unclassified | 3,609 | 7 | 1 | 1 | region_channel=Region, sku=SKU |
| Inventory Review |  | inventory | 2,680 | 15 | 15,984 | 2 | sku=SKU, invoice=Inv Cost, quantity=Qty on PO, currency=Keep Inv Cost**?** |
| Alloc_Hold_List | veryHidden | inventory | 2,301 | 11 | 4,602 | 1 | date=Ship Date, quantity=On Hand, region_channel=Region, sku=SKU, currency=BO Value ($)**?** |
| BP_Inventory | veryHidden | inventory | 2,301 | 12 | 7 | 12 | date=Ship Date, quantity=On Hand, region_channel=Region, sku=SKU, currency=BO Value $**?** |
| Position Engine | veryHidden | inventory | 2,213 | 31 | 59,751 | 1 | quantity=On Hand, region_channel=Region, sku=SKU, currency=Avg Cost**?**, po_number=Inbound PO Qty**?** |
| NP Allocation Table |  | allocation | 1,922 | 7 | 0 | 1 | currency=Value, date=Targeted Launch Date**?** |
| Items on PO | hidden | open_po | 1,820 | 19 | 0 | 1 | po_number=Document Number, currency=Amount, date=Date, quantity=Quantity, sku=Item, po_line=Line ID, location=Inventory Location**?** |
| BP_Sales | veryHidden | unclassified | 1,630 | 10 | 14,677 | 12 | region_channel=Region, sku=SKU, currency=BO Value $**?** |
| __AllocMix | veryHidden | demand | 1,188 | 19 | 21,386 | 3 | region_channel=Americas, sku=SKU |
| Projected Inventory |  | inventory | 1,127 | 37 | 0 | 3 | sku=SKU, quantity=On-Hand Units, currency=Cost/Sales Spread, invoice=Inv Cost ($), po_number=PO Inbound (12M) |
| HIE Production Schedule | veryHidden | production | 773 | 6 | 0 | 1 | date=Date, po_number=PO, quantity=Qty, sku=SKU |
| ARC Build Live |  | demand | 720 | 8 | 3,603 | 5 | sku=Variant SKU (BOM parent)**?** |
| Allocated New Products | hidden | unclassified | 660 | 6 | 0 | 2 | region_channel=B2C |
| SO_Closure_List | hidden | open_po | 563 | 11 | 1,690 | 2 | date=Ship Date, quantity=On Hand, region_channel=Region, sku=SKU, currency=BO Value ($)**?**, po_number=Open PO Qty**?** |
| BP_Procurement | veryHidden | supplier | 504 | 10 | 4,517 | 12 | po_number=PO #, sku=Item, supplier=Vendor, quantity=Qty Remaining, date=Expected Date**?** |
| PO_Late_Report | veryHidden | supplier | 501 | 13 | 582 | 1 | po_number=PO #, sku=Item, supplier=Vendor, quantity=Qty Remaining, date=Expected Date**?** |
| ARC Buildable |  | inventory | 477 | 32 | 839 | 1 | — |
| BP_Logistics | veryHidden | unclassified | 416 | 8 | 8 | 12 | quantity=Units, region_channel=Channel, sku=SKU |
| INV \| Apparel | hidden | inventory | 386 | 13 | 0 | 6 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| Sales_Hygiene_Flags | hidden | unclassified | 351 | 17 | 1,053 | 1 | date=Ship Date, region_channel=Region, sku=SKU, currency=BO Value ($)**?**, quantity=BO Qty**?** |
| JAX Container Validation |  | in_transit | 254 | 31 | 7,112 | 1 | date=ETA, sku=SKU, container=Container ID / Group, currency=Price Data Flag, invoice=Inv Months, quantity=Qty Avail (Asia), region_channel=Americas On |
| __Justify_Container Shipment | veryHidden | in_transit | 206 | 15 | 0 | 1 | region_channel=Region, sku=SKU, quantity=Allocated Qty**?** |
| Cue Shipment |  | unclassified | 200 | 20 | 185 | 6 | region_channel=Americas, sku=ITEM, quantity=Ship QTY**?** |
| JingdianProdFlat | hidden | supplier | 176 | 8 | 0 | 1 | date=Date, po_number=PO, quantity=Qty, sku=SKU |
| ARC Parts Live |  | inventory | 170 | 9 | 1,362 | 5 | quantity=On Hand (Current Inv), po_number=On PO Remaining**?**, sku=Part SKU**?** |
| BK Dist Plan |  | unclassified | 166 | 10 | 0 | 7 | date=Launch Date**?**, quantity=Butt Qty**?**, sku=Butt SKU**?** |
| HIEInv |  | supplier | 141 | 2 | 0 | 1 | sku=SKU, invoice=HIE Inv**?** |
| __Justify_HIE Shipment | veryHidden | supplier | 134 | 15 | 0 | 1 | region_channel=Region, sku=SKU, quantity=Allocated Qty**?** |
| INV \| Gloves | hidden | inventory | 132 | 13 | 0 | 6 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| INV \| Other | hidden | inventory | 132 | 13 | 0 | 6 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| PoisonInvFlat | hidden | unclassified | 104 | 6 | 0 | 1 | currency=Amount, po_number=PO, quantity=Qty, sku=SKU |
| Cue Sales Xcheck |  | unclassified | 92 | 12 | 858 | 1 | sku=SKU, date=Committed Lines (Available Qty), quantity=Units Sold Post-Launch |
| INV \| PRE Accessories | hidden | inventory | 87 | 13 | 0 | 2 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| Container Shipment |  | in_transit | 76 | 10 | 177 | 6 | region_channel=Americas, sku=ITEM, quantity=Ship QTY**?** |
| INV \| Playing Cues | hidden | inventory | 68 | 13 | 0 | 2 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| Inventory Health Review |  | inventory | 60 | 19 | 296 | 6 | quantity=Units, invoice=Inv Cost ($) |
| REVO Production |  | production | 59 | 32 | 867 | 1 | sku=SKU, region_channel=B2C INV |
| INV \| Utility Cues |  | inventory | 54 | 13 | 0 | 2 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| INV \| Components | hidden | inventory | 49 | 13 | 0 | 6 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| INV \| Table Accessories | hidden | inventory | 47 | 13 | 0 | 2 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| In-Stock Reference |  | inventory | 46 | 14 | 2 | 5 | quantity=On Hand, date=Committed, sku=Part SKU**?** |
| S&OP Summary | veryHidden | demand | 44 | 10 | 159 | 6 | region_channel=Region |
| INV \| Cases | hidden | inventory | 44 | 13 | 0 | 6 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| S&OP Meeting Agenda |  | demand | 41 | 8 | 0 | 8 | — |
| INV \| Accessories | hidden | inventory | 35 | 13 | 0 | 6 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| Poison Production | hidden | production | 34 | 19 | 0 | 5 | — |
| Shipped by Region |  | unclassified | 30 | 14 | 304 | 9 | region_channel=Americas |
| Chalk Updates |  | unclassified | 29 | 22 | 5 | 1 | po_number=Document Number, date=Date, sku=Item, quantity=1080 (Predator) Qty**?** |
| INV \| Wood Shafts | hidden | inventory | 29 | 13 | 0 | 2 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| S&OP 6.3.2026 | hidden | demand | 26 | 8 | 0 | 22 | quantity=Corrective Action Available**?**, region_channel=Department/Region - Person Issuing CAR**?** |
| Poison Finished Inv | hidden | production | 25 | 15 | 0 | 3 | location=production warehouse**?**, quantity=Sum of Qty**?** |
| INV \| REVO Shafts | hidden | production | 25 | 13 | 0 | 2 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| CORE Shaft Allocation |  | allocation | 22 | 20 | 174 | 5 | uom=Per Month (High) |
| BO Waterfall | hidden | unclassified | 22 | 5 | 24 | 4 | currency=$ at Selling Price**?** |
| Decision | veryHidden | demand | 21 | 17 | 57 | 10 | region_channel=Channel |
| Tariffs | hidden | tariff | 18 | 22 | 0 | 4 | — |
| Sensitivity | veryHidden | unclassified | 14 | 4 | 0 | 5 | currency=Value |
| PE Config | hidden | unclassified | 13 | 8 | 4 | 13 | — |
| INV \| Carom Shafts | hidden | inventory | 9 | 13 | 0 | 6 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| INV \| Balls | hidden | inventory | 7 | 13 | 0 | 6 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| INV \| Tables | hidden | inventory | 7 | 13 | 0 | 2 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| INV \| Chalk | hidden | inventory | 6 | 13 | 0 | 6 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |
| HIE Production |  | production | 4 | 9 | 0 | 3 | — |
| INV \| Tips | hidden | inventory | 3 | 13 | 0 | 2 | sku=SKU, quantity=On-Hand Units, invoice=Inv Cost ($), currency=% of Category Cost**?** |

### `Shipment_Request_HIE__Updated.xlsm` — 36 populated of 36 sheets

| Sheet | State | Domain | Data rows | Cols | Formulas | Hdr row | Detected keys |
|---|---|---|---|---|---|---|---|
| HIE_Defect |  | quality | 1,048,575 | 3 | 0 | 1 | po_number=PO#, sku=Item#, quantity=Defective Qty **?** |
| Sales |  | unclassified | 255,899 | 5 | 0 | 1 | date=Date, quantity=Quantity, sku=Item |
| Regional Inv Loc |  | inventory | 176,349 | 4 | 0 | 1 | quantity=Available, sku=SKU, region_channel=Region Inv Loc, location=Inventory Location**?** |
| Availability |  | inventory | 41,354 | 18 | 0 | 1 | sku=SKU, invoice=Inv Alloc, quantity=Qty Available |
| NetSuite Forecast | hidden | demand | 41,026 | 5 | 0 | 1 | sku=Item, quantity=Forecast Qty**?** |
| Avail Audit | hidden | unclassified | 26,233 | 8 | 0 | 1 | invoice=? Inv, sku=SKU\|Year\|Month |
| Forecast |  | demand | 5,460 | 10 | 0 | 1 | sku=Item Name, date=Date Forecasted, location=Inventory Location**?** |
| Open Order Report | hidden | unclassified | 3,550 | 7 | 0 | 1 | date=Ship Date, sku=Item, location=dFulfillment Location**?** |
| New Prod Alloc | hidden | unclassified | 1,886 | 7 | 0 | 1 | currency=Value, date=Targeted Launch Date**?** |
| PO |  | unclassified | 1,820 | 15 | 0 | 1 | po_number=Document Number, quantity=Qty, sku=Item, po_line=Line ID, date=Expected Receipt Date**?** |
| Location Inventory |  | inventory | 1,522 | 4 | 0 | 1 | quantity=Available, location=Inventory Location**?** |
| HIEProdFlat |  | supplier | 545 | 6 | 0 | 1 | date=Date, po_number=PO, quantity=Qty, sku=SKU |
| HIEInv |  | supplier | 451 | 5 | 0 | 1 | po_number=PO, quantity=Qty, sku=SKU |
| PO HIE Staging |  | supplier | 315 | 4 | 0 | 1 | po_number=PO, sku=SKU, date=Max Prod Date**?** |
| Shipment Request |  | production | 254 | 13 | 1 | 2 | quantity=Qty Needed for 3 Mo JAX (from E), invoice=HIE Inv After Prod**?** |
| Cue Completion Schedule |  | inventory | 228 | 9 | 912 | 4 | po_number=PO, sku=SKU, quantity=Base Supply (On Hand + At Supplier)**?**, date=Prod Date**?** |
| Sheet1 |  | unclassified | 207 | 24 | 1 | 3 | — |
| JingdianProdFlat |  | supplier | 176 | 8 | 0 | 1 | date=Date, po_number=PO, quantity=Qty, sku=SKU |
| HIEWarehouse |  | supplier | 161 | 5 | 161 | 1 | currency=Amount, po_number=PO#, quantity=QTY, sku=ITEM# |
| Defect Fallout to Date |  | quality | 161 | 5 | 0 | 4 | sku=Item, date=Fallout to Date (Units)**?**, quantity=Fallout 2026 YTD (Units)**?** |
| HIE Shipment |  | supplier | 127 | 2 | 0 | — | — |
| PoisonInvFlat |  | unclassified | 104 | 6 | 0 | 1 | currency=Amount, po_number=PO, quantity=Qty, sku=SKU |
| Cost Dashboard |  | unclassified | 95 | 35 | 18 | 1 | currency=Cost, sku=SKU, quantity=Qty Needed |
| AVANT GARDE Dashboard |  | unclassified | 90 | 8 | 241 | 5 | — |
| JAX Shipments by Date |  | unclassified | 88 | 10 | 267 | 13 | sku=SKU, quantity=Qty 8/24, currency=Est. Cost**?**, date=Launch Date**?** |
| Launch Status 2026 |  | supplier | 88 | 21 | 1,777 | 4 | sku=SKU, po_number=PO Exists, date=Last HIE Prod Date**?**, quantity=Launch Qty (Planned Alloc)**?**, region_channel=In Jacksonville (INT/US B2B+B2C)* |
| Shipment Request HIE to PG JAX |  | production | 85 | 29 | 1 | 10 | po_number=PO, sku=SKU, po_line=Line Built?, quantity=Missing Qty**?** |
| New Product Readiness |  | inventory | 85 | 42 | 3,069 | 9 | sku=SKU, quantity=On Hand (Received), date=Date Target Met (helper), region_channel=B2C Alloc, supplier=At Supplier (Completed)**?** |
| PLW Completion |  | open_po | 41 | 16 | 566 | 8 | sku=Model, po_number=PO Due, quantity=Open PO Qty**?** |
| Sheet4 | hidden | unclassified | 36 | 4 | 0 | 2 | — |
| Launch Checkpoints |  | supplier | 34 | 23 | 567 | 5 | sku=SKU, date=L-14 Date**?**, quantity=CUE minus HIEProdFlat at L-7 (units)**?** |
| Launch Location Detail |  | in_transit | 22 | 26 | 507 | 5 | sku=SKU, po_number=PO number(s), date=Last Prod Date (HIE)**?**, invoice=HIE — paid inv**?**, quantity=% Finished vs Launch Qty**?**, region_channel=J |
| Sheet5 | hidden | unclassified | 17 | 2 | 0 | 3 | currency=Sum of Value**?** |
| Helpers | hidden | dsr | 11 | 10 | 0 | 2 | region_channel=INT B2C**?** |
| China 888 US Overflow |  | inventory | 10 | 13 | 49 | 11 | sku=SKU, quantity=Finished Now (On Hand + At Supplier)**?** |
| Alloc Overrides |  | in_transit | 2 | 6 | 0 | 4 | sku=SKU, region_channel=B2C Alloc |

### `Direct_Shipments_Consolidator.xlsm` — 12 populated of 13 sheets

| Sheet | State | Domain | Data rows | Cols | Formulas | Hdr row | Detected keys |
|---|---|---|---|---|---|---|---|
| DSR_ReleaseAudit |  | dsr | 793 | 13 | 0 | 1 | quantity=Qty, sku=SKU (Canon), invoice=JAX Inv**?** |
| DSR_Data |  | dsr | 793 | 14 | 0 | 2 | quantity=Qty, region_channel=Region, sku=SKU |
| NP_Flat |  | allocation | 772 | 9 | 0 | 1 | region_channel=Region, invoice=Current Inv**?** |
| Remit_Audit |  | unclassified | 563 | 16 | 0 | 4 | date=ShipDate, quantity=Qty, region_channel=Region, sku=SKU |
| DSR_ReleasePacket |  | dsr | 538 | 5 | 0 | 1 | quantity=Qty, sku=SKU (Canon) |
| DSR_ImportLog |  | dsr | 260 | 4 | 0 | 2 | — |
| DSR_ReqAudit |  | dsr | 49 | 18 | 0 | 1 | — |
| DSR_Dashboard |  | dsr | 45 | 10 | 0 | 9 | quantity=Qty |
| EMEA_Dist |  | allocation | 28 | 4 | 0 | 1 | sku=Item |
| DSR_CompanyMap |  | dsr | 13 | 7 | 0 | 3 | region_channel=Channel |
| Email |  | unclassified | 4 | 2 | 0 | 1 | — |
| DSR_Overrides |  | dsr | 1 | 6 | 1 | 1 | sku=SKU (Canon) |

## Duplicate truth — two sources, same columns

Matched on normalized header signature, so this catches duplicates whose filenames look nothing alike. A row-count difference means at least one is stale or filtered.

| Source A | Rows A | Source B | Rows B | Header overlap | Row delta | Domain | Shared columns |
|---|---|---|---|---|---|---|---|
| `Supply_Chain_Update_Meeting_Workbook.xlsm`!HIE Production Schedule | 774 | `Shipment_Request_HIE__Updated.xlsm`!HIEProdFlat | 546 | 100% | +228 | inventory\|production | B Grade, Date, PO, Qty, SKU, Status |
| `Supply_Chain_Update_Meeting_Workbook.xlsm`!NP Allocation Table | 1,923 | `Shipment_Request_HIE__Updated.xlsm`!New Prod Alloc | 1,887 | 100% | +36 | inventory\|production | Attribute, Base Price, Description, Internal ID, Name, Targeted Launch Date |
| `Supply_Chain_Update_Meeting_Workbook.xlsm`!PoisonInvFlat | 105 | `Shipment_Request_HIE__Updated.xlsm`!PoisonInvFlat | 105 | 100% | 0 | inventory\|production | Amount, PO, Qty, SKU, Source Sheet, Unit Price |
| `Supply_Chain_Update_Meeting_Workbook.xlsm`!JingdianProdFlat | 177 | `Shipment_Request_HIE__Updated.xlsm`!JingdianProdFlat | 177 | 88% | 0 | inventory\|production | Chinese Item, Date, PO, Qty, SKU, Source Sheet |

## Expected workbooks — confirmed vs missing

| Status | Expected asset | Domain | Found at / searched for |
|---|---|---|---|
| **MISSING** | BK Plan forecast | demand | `*BK*Plan*` |
| **MISSING** | Build-RealData.ps1 | platform | `Build-RealData.ps1` |
| **MISSING** | In Transit workbook | in_transit | `*In*Transit*.xls*`, `*InTransit*.xls*` |
| **MISSING** | Master_Reference.xlsx (SO_Data, Item master, Item_Pricing) | reference | `Master_Reference.xls*`, `*Master*Reference*.xls*` |
| **MISSING** | NP Go/No-Go tool + allocation table | allocation | `*Go*No*Go*.xls*`, `*NP*Allocation*.xls*` |
| **MISSING** | NetSuite Items on Purchase Order export | open_po | `*Items*on*Purchase*Order*`, `*ItemsOnPurchaseOrder*` |
| **MISSING** | Overstock / static inventory workbook | inventory | `*Overstock*.xls*`, `*Static*Inventory*.xls*` |
| **MISSING** | Pre-shipment inspection / reject / CAR logs | quality | `*Inspection*.xls*`, `*Quality*.xls*`, `*CAR*Log*.xls*` |
| **MISSING** | REVO production + projected inventory model | production | `*REVO*.xls*` |
| **MISSING** | Shipment_Request_Consolidation | production | `Shipment_Request_Consolidation*`, `*Shipment*Request*Consol*` |
| **MISSING** | Tariff workbook by trade lane and category | tariff | `*Tariff*.xls*` |
| **MISSING** | build_production_visibility.py | production | `build_production_visibility.py` |
| **MISSING** | export_dashboard_data.py | platform | `export_dashboard_data.py` |
| **MISSING** | sc_analysis.py companion script | inventory | `sc_analysis.py` |
| found | Direct Shipment Request tool | dsr | `Direct_Shipments_Consolidator.xlsm` |
| found | Shipment_Request_HIE | production | `Shipment_Request_HIE__Updated.xlsm` |
| found | Supply Chain Update Meeting Workbook | inventory | `Supply_Chain_Update_Meeting_Workbook.xlsm` |

## External dependency map — what breaks if a path moves

| Dependent file | Kind | Via | Location |
|---|---|---|---|
| `Shipment_Request_HIE__Updated.xlsm` | file | PQ:HIE Shipment | `C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Update Meeting Workbook.xlsm` |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | file | PQ:HIEInv | `C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Chase\Shipment Request HIE .xl` |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | file | PQ:HIEProdFlat | `C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Chase\Shipment Request HIE - U` |
| `Direct_Shipments_Consolidator.xlsm` | web | PQ:Open Order | `https://3492685.app.netsuite.com/core/media/media.nl?id=2600949&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Shipment_Request_HIE__Updated.xlsm` | web | PQ:Open_Order_Report | `https://3492685.app.netsuite.com/core/media/media.nl?id=2600949&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Shipment_Request_HIE__Updated.xlsm` | web | PQ:Location_Inventory | `https://3492685.app.netsuite.com/core/media/media.nl?id=2600947&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Shipment_Request_HIE__Updated.xlsm` | web | PQ:PO | `https://3492685.app.netsuite.com/core/media/media.nl?id=2600946&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Shipment_Request_HIE__Updated.xlsm` | web | PQ:NetSuite Forecast | `https://3492685.app.netsuite.com/core/media/media.nl?id=2606291&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Shipment_Request_HIE__Updated.xlsm` | web | PQ:New Prod Alloc | `https://3492685.app.netsuite.com/core/media/media.nl?id=2620991&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Shipment_Request_HIE__Updated.xlsm` | web | PQ:Sales | `https://3492685.app.netsuite.com/core/media/media.nl?id=2600950&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | web | PQ:Current Inventory | `https://3492685.app.netsuite.com/core/media/media.nl?id=2600947&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | web | PQ:Customer Open Orders | `https://3492685.app.netsuite.com/core/media/media.nl?id=2600949&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | web | PQ:Items On PO | `https://3492685.app.netsuite.com/core/media/media.nl?id=2600946&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | web | PQ:New Product Allocations | `https://3492685.app.netsuite.com/core/media/media.nl?id=2620991&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | web | PQ:Purchase Price | `https://3492685.app.netsuite.com/core/media/media.nl?id=2613629&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | web | PQ:Forecast Qty | `https://3492685.app.netsuite.com/core/media/media.nl?id=2606291&c=3492685&h=<REDACTED>&_xt=.csv` |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | web | PQ:ValidatedSales | `https://predatorgroup.sharepoint.com/Share%20All%20Files/Forecast/Reference/Actuals%20vs%20Forec` |

## Power Query inventory

### `Supply_Chain_Update_Meeting_Workbook.xlsm` — 11 queries

| Query | Lines | External source kinds | Depends on queries |
|---|---|---|---|
| Current Inventory | 18 | web | — |
| Customer Open Orders | 28 | web | — |
| Items On PO | 13 | web | — |
| New Product Allocations | 7 | web | — |
| Purchase Price | 12 | web | — |
| Forecast Qty | 154 | web | — |
| HIEInv | 10 | file | — |
| HIEProdFlat | 8 | file | — |
| ValidatedSales | 13 | web | — |
| JingdianProdFlat | 7 | file | — |
| PoisonInvFlat | 7 | file | — |

### `Shipment_Request_HIE__Updated.xlsm` — 8 queries

| Query | Lines | External source kinds | Depends on queries |
|---|---|---|---|
| Open_Order_Report | 11 | web | — |
| Location_Inventory | 14 | web | — |
| PO | 10 | web | — |
| NetSuite Forecast | 15 | web | — |
| Regional Inv Loc | 11 | web | — |
| New Prod Alloc | 8 | web | — |
| Sales | 10 | web | — |
| HIE Shipment | 7 | file | — |

### `Direct_Shipments_Consolidator.xlsm` — 1 queries

| Query | Lines | External source kinds | Depends on queries |
|---|---|---|---|
| Open Order | 7 | web | — |

## VBA present

Module names only — nothing was executed. A password-protected project cannot be reviewed or ported, which makes it a hard blocker rather than a risk.

| File | Project | Protected | Components | Code modules |
|---|---|---|---|---|
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | VBAProject | **yes** | 3 code / 68 sheet | InvReview, PythonRun, Alloc_Engine |
| `Shipment_Request_HIE__Updated.xlsm` | VBAProject | **yes** | 1 code / 26 sheet | Module1 |
| `Direct_Shipments_Consolidator.xlsm` | VBAProject | **yes** | 18 code / 14 sheet | modDSR_RequestRollup, modDSR_ReleaseGate, modDSR_Overrides, modDSR_Master, modDSR_ImportPipeline, modDSR_Core, modDSR_AuditDashboard, Debugger, modDSR_Stage3_ReleasePacket, modDSR_Stage35_Release, Build_AutoPacket, TheOperator, DashboardQueues, DenialEmail, RemitEmail, modDSR_RemitAudit, BuildRemitTable, MarkStatusAsREMITTED |

## Probe failures and warnings

Clean sweep — no probe errors.

## Next steps

| # | Action | Owner |
|---|---|---|
| 1 | Name the source of record for each duplicate-truth pair | Chase |
| 2 | Correct any wrong domain/grain call in the source table | Chase |
| 3 | Supply paths for the missing expected workbooks | Chase |
| 4 | Sign off `SCHEMA.md` | Chase |
| 5 | Build Phase 2 readers against the confirmed sources only | Claude |

Phase 2 is blocked on items 1-4 by design. Extraction code written before the model is agreed is code that gets rewritten.
