# Query reconciliation — where the forked logic diverges

Generated 2026-09-02T17:50:20+00:00 from the committed M source in `queries/`. No refresh, no Excel.

## Executive summary

- **5 upstream export(s) are read by more than one workbook.**
- **6 blocking conflict(s)** — the numbers cannot agree until these are resolved.
- 25 divergence(s) that change scope or grain, and 4 naming inconsistency(ies).

### The blocking conflicts

**1. Join key built differently — key column "SKU"**

- Upstream: NetSuite File Cabinet export id=2600947 (account 3492685)
- `Shipment_Request_HIE__Updated.xlsm!Regional Inv Loc` -> `if Text.Contains([Name],":") then Text.AfterDelimiter([Name],":") else [Name]`
- `Supply_Chain_Update_Meeting_Workbook.xlsm!Current Inventory` -> `if Text.Contains([Item], ":") then Text.Trim(Text.AfterDelimiter([Item], ":")) else Text.Trim([Item])`
- Only one side applies Text.Trim. The same NetSuite value therefore produces two different key strings, so rows that should join do not, and the same item can appear twice in a combined view.
- **Fix:** Normalise the key once, in the canonical layer, and never per query.

**2. Same column name, different measure — output column "Available"**

- Upstream: NetSuite File Cabinet export id=2600947 (account 3492685)
- `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` -> `List.Sum([On Hand])`
- `Shipment_Request_HIE__Updated.xlsm!Regional Inv Loc` -> `List.Sum([Available])`
- Both queries emit a column called "Available" from the same export, but they aggregate different source columns. Any comparison, join or roll-up across the two is comparing different measures under one name — and nothing in either workbook shows that.
- **Fix:** Define this measure once in the canonical layer and give each variant a distinct, honest name.

**3. Same file, different parser — Csv.Document option QuoteStyle**

- Upstream: NetSuite File Cabinet export id=2600947 (account 3492685)
- `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` -> `QuoteStyle.Csv`
- `Shipment_Request_HIE__Updated.xlsm!Regional Inv Loc` -> `QuoteStyle.None`
- QuoteStyle.None ignores CSV quoting, so any field containing a comma splits across columns and silently shifts every value after it on that row. The two queries are reading the same bytes into different tables.
- **Fix:** Read the export once, with QuoteStyle.Csv, and reuse it.

**4. Same file, different parser — Csv.Document option QuoteStyle**

- Upstream: NetSuite File Cabinet export id=2600947 (account 3492685)
- `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` -> `QuoteStyle.Csv`
- `Supply_Chain_Update_Meeting_Workbook.xlsm!Current Inventory` -> `QuoteStyle.None`
- QuoteStyle.None ignores CSV quoting, so any field containing a comma splits across columns and silently shifts every value after it on that row. The two queries are reading the same bytes into different tables.
- **Fix:** Read the export once, with QuoteStyle.Csv, and reuse it.

**5. Same file, different parser — Csv.Document option QuoteStyle**

- Upstream: NetSuite File Cabinet export id=2600949 (account 3492685)
- `Direct_Shipments_Consolidator.xlsm!Open Order` -> `QuoteStyle.None`
- `Shipment_Request_HIE__Updated.xlsm!Open_Order_Report` -> `QuoteStyle.Csv`
- QuoteStyle.None ignores CSV quoting, so any field containing a comma splits across columns and silently shifts every value after it on that row. The two queries are reading the same bytes into different tables.
- **Fix:** Read the export once, with QuoteStyle.Csv, and reuse it.

**6. Same file, different parser — Csv.Document option QuoteStyle**

- Upstream: NetSuite File Cabinet export id=2600949 (account 3492685)
- `Shipment_Request_HIE__Updated.xlsm!Open_Order_Report` -> `QuoteStyle.Csv`
- `Supply_Chain_Update_Meeting_Workbook.xlsm!Customer Open Orders` -> `QuoteStyle.None`
- QuoteStyle.None ignores CSV quoting, so any field containing a comma splits across columns and silently shifts every value after it on that row. The two queries are reading the same bytes into different tables.
- **Fix:** Read the export once, with QuoteStyle.Csv, and reuse it.

## All findings

| Severity | Finding | Upstream | Subject | Query A | A value | Query B | B value |
|---|---|---|---|---|---|---|---|
| blocking | Join key built differently | NetSuite File Cabinet export id=2600947 (acc | key column "SKU" | `Shipment_Request_HIE__Updated.xlsm!Regional Inv Loc` | if Text.Contains([Name],":") then Text.AfterDelimiter([Name],":") else | `Supply_Chain_Update_Meeting_Workbook.xlsm!Current Inventory` | if Text.Contains([Item], ":") then Text.Trim(Text.AfterDelimiter([Item |
| blocking | Same column name, different measure | NetSuite File Cabinet export id=2600947 (acc | output column "Available" | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | List.Sum([On Hand]) | `Shipment_Request_HIE__Updated.xlsm!Regional Inv Loc` | List.Sum([Available]) |
| blocking | Same file, different parser | NetSuite File Cabinet export id=2600947 (acc | Csv.Document option QuoteStyle | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | QuoteStyle.Csv | `Shipment_Request_HIE__Updated.xlsm!Regional Inv Loc` | QuoteStyle.None |
| blocking | Same file, different parser | NetSuite File Cabinet export id=2600947 (acc | Csv.Document option QuoteStyle | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | QuoteStyle.Csv | `Supply_Chain_Update_Meeting_Workbook.xlsm!Current Inventory` | QuoteStyle.None |
| blocking | Same file, different parser | NetSuite File Cabinet export id=2600949 (acc | Csv.Document option QuoteStyle | `Direct_Shipments_Consolidator.xlsm!Open Order` | QuoteStyle.None | `Shipment_Request_HIE__Updated.xlsm!Open_Order_Report` | QuoteStyle.Csv |
| blocking | Same file, different parser | NetSuite File Cabinet export id=2600949 (acc | Csv.Document option QuoteStyle | `Shipment_Request_HIE__Updated.xlsm!Open_Order_Report` | QuoteStyle.Csv | `Supply_Chain_Update_Meeting_Workbook.xlsm!Customer Open Orders` | QuoteStyle.None |
| warning | Different grain | NetSuite File Cabinet export id=2600946 (acc | group-by keys | `Shipment_Request_HIE__Updated.xlsm!PO` | Item x Expected Receipt Date x Document Number x Line ID x Original PI | `Supply_Chain_Update_Meeting_Workbook.xlsm!Items On PO` | (no grouping) |
| warning | Join key derived on one side only | NetSuite File Cabinet export id=2600946 (acc | key column "SKU" | `Shipment_Request_HIE__Updated.xlsm!PO` | (absent) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Items On PO` | if Text.Contains([SKU], ":") then Text.Trim(Text.AfterDelimiter([SKU], |
| warning | Different row scope | NetSuite File Cabinet export id=2600947 (acc | row filters | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | ([Custom] = "JAX") and ([Available] <> null) | `Shipment_Request_HIE__Updated.xlsm!Regional Inv Loc` | (no filter) |
| warning | Different row scope | NetSuite File Cabinet export id=2600947 (acc | row filters | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | ([Custom] = "JAX") and ([Available] <> null) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Current Inventory` | ([InvQty] <> null) |
| warning | Different row scope | NetSuite File Cabinet export id=2600947 (acc | row filters | `Shipment_Request_HIE__Updated.xlsm!Regional Inv Loc` | (no filter) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Current Inventory` | ([InvQty] <> null) |
| warning | Join key derived on one side only | NetSuite File Cabinet export id=2600947 (acc | key column "Name" | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | if Text.Contains([Name], ":") then Text.AfterDelimiter([Name], ":") el | `Shipment_Request_HIE__Updated.xlsm!Regional Inv Loc` | (absent) |
| warning | Join key derived on one side only | NetSuite File Cabinet export id=2600947 (acc | key column "SKU" | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | (absent) | `Shipment_Request_HIE__Updated.xlsm!Regional Inv Loc` | if Text.Contains([Name],":") then Text.AfterDelimiter([Name],":") else |
| warning | Join key derived on one side only | NetSuite File Cabinet export id=2600947 (acc | key column "Name" | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | if Text.Contains([Name], ":") then Text.AfterDelimiter([Name], ":") el | `Supply_Chain_Update_Meeting_Workbook.xlsm!Current Inventory` | (absent) |
| warning | Join key derived on one side only | NetSuite File Cabinet export id=2600947 (acc | key column "Regional ID" | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | (absent) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Current Inventory` | if [InvLoc] = "INT Tradeshows" then "Tradeshow" else if [InvLoc] = "IN |
| warning | Join key derived on one side only | NetSuite File Cabinet export id=2600947 (acc | key column "SKU" | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | (absent) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Current Inventory` | if Text.Contains([Item], ":") then Text.Trim(Text.AfterDelimiter([Item |
| warning | Join key derived on one side only | NetSuite File Cabinet export id=2600947 (acc | key column "Regional ID" | `Shipment_Request_HIE__Updated.xlsm!Regional Inv Loc` | (absent) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Current Inventory` | if [InvLoc] = "INT Tradeshows" then "Tradeshow" else if [InvLoc] = "IN |
| warning | Same file, different parser | NetSuite File Cabinet export id=2600947 (acc | Csv.Document option Columns | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | (not set) | `Shipment_Request_HIE__Updated.xlsm!Regional Inv Loc` | 24 |
| warning | Same file, different parser | NetSuite File Cabinet export id=2600947 (acc | Csv.Document option Columns | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | (not set) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Current Inventory` | 24 |
| warning | Different grain | NetSuite File Cabinet export id=2600949 (acc | group-by keys | `Direct_Shipments_Consolidator.xlsm!Open Order` | Company Name x Sales Rep x Item | `Shipment_Request_HIE__Updated.xlsm!Open_Order_Report` | Ship Date x Territory x Item |
| warning | Different grain | NetSuite File Cabinet export id=2600949 (acc | group-by keys | `Direct_Shipments_Consolidator.xlsm!Open Order` | Company Name x Sales Rep x Item | `Supply_Chain_Update_Meeting_Workbook.xlsm!Customer Open Orders` | Item x Ship Date x Regional ID x Document Number x Company Name x Comm |
| warning | Different grain | NetSuite File Cabinet export id=2600949 (acc | group-by keys | `Shipment_Request_HIE__Updated.xlsm!Open_Order_Report` | Ship Date x Territory x Item | `Supply_Chain_Update_Meeting_Workbook.xlsm!Customer Open Orders` | Item x Ship Date x Regional ID x Document Number x Company Name x Comm |
| warning | Join key derived on one side only | NetSuite File Cabinet export id=2600949 (acc | key column "Regional ID" | `Direct_Shipments_Consolidator.xlsm!Open Order` | (absent) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Customer Open Orders` | if [Territory] = "B2B Asia Pacific" then "AP" else if [Territory] = "B |
| warning | Join key derived on one side only | NetSuite File Cabinet export id=2600949 (acc | key column "SKU" | `Direct_Shipments_Consolidator.xlsm!Open Order` | (absent) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Customer Open Orders` | if Text.Contains([SKU], ":") then Text.Trim(Text.AfterDelimiter([SKU], |
| warning | Join key derived on one side only | NetSuite File Cabinet export id=2600949 (acc | key column "Regional ID" | `Shipment_Request_HIE__Updated.xlsm!Open_Order_Report` | (absent) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Customer Open Orders` | if [Territory] = "B2B Asia Pacific" then "AP" else if [Territory] = "B |
| warning | Join key derived on one side only | NetSuite File Cabinet export id=2600949 (acc | key column "SKU" | `Shipment_Request_HIE__Updated.xlsm!Open_Order_Report` | (absent) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Customer Open Orders` | if Text.Contains([SKU], ":") then Text.Trim(Text.AfterDelimiter([SKU], |
| warning | Same file, different parser | NetSuite File Cabinet export id=2600949 (acc | Csv.Document option Columns | `Direct_Shipments_Consolidator.xlsm!Open Order` | 25 | `Shipment_Request_HIE__Updated.xlsm!Open_Order_Report` | (not set) |
| warning | Same file, different parser | NetSuite File Cabinet export id=2600949 (acc | Csv.Document option Columns | `Shipment_Request_HIE__Updated.xlsm!Open_Order_Report` | (not set) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Customer Open Orders` | 25 |
| warning | Different row scope | NetSuite File Cabinet export id=2606291 (acc | row filters | `Shipment_Request_HIE__Updated.xlsm!NetSuite Forecast` | ([Forecast Qty] <> null and [Forecast Qty] <> 0) | `Supply_Chain_Update_Meeting_Workbook.xlsm!Forecast Qty` | [Forecast Qty] <> 0 and [Forecast Qty] <> null |
| warning | Different grain | NetSuite File Cabinet export id=2606291 (acc | group-by keys | `Shipment_Request_HIE__Updated.xlsm!NetSuite Forecast` | Item x Allocation Group x Forecast Mnth x Product Category | `Supply_Chain_Update_Meeting_Workbook.xlsm!Forecast Qty` | (no grouping) |
| warning | Different row scope | NetSuite File Cabinet export id=2620991 (acc | row filters | `Shipment_Request_HIE__Updated.xlsm!New Prod Alloc` | ([Name] <> "BC8 PRE P3 BLU LL" and [Name] <> "BC8 PRE P3 BLU NW" and [ | `Supply_Chain_Update_Meeting_Workbook.xlsm!New Product Allocations` | (no filter) |
| info | Same measure, different name | NetSuite File Cabinet export id=2600947 (acc | List.Sum([On Hand]) | `Shipment_Request_HIE__Updated.xlsm!Location_Inventory` | Available | `Supply_Chain_Update_Meeting_Workbook.xlsm!Current Inventory` | InvQty |
| info | Same measure, different name | NetSuite File Cabinet export id=2600949 (acc | List.Sum([Back Order Qty]) | `Direct_Shipments_Consolidator.xlsm!Open Order` | Qty_BO | `Shipment_Request_HIE__Updated.xlsm!Open_Order_Report` | BO |
| info | Same measure, different name | NetSuite File Cabinet export id=2600949 (acc | List.Sum([Back Order Qty]) | `Direct_Shipments_Consolidator.xlsm!Open Order` | Qty_BO | `Supply_Chain_Update_Meeting_Workbook.xlsm!Customer Open Orders` | TtlBO |
| info | Same measure, different name | NetSuite File Cabinet export id=2600949 (acc | List.Sum([Back Order Qty]) | `Shipment_Request_HIE__Updated.xlsm!Open_Order_Report` | BO | `Supply_Chain_Update_Meeting_Workbook.xlsm!Customer Open Orders` | TtlBO |

## Per-upstream detail

### NetSuite File Cabinet export id=2600947 (account 3492685) — **FORKED**

| Workbook | Query | Steps | QuoteStyle | Grain | Measures (what they really aggregate) | Filters |
|---|---|---|---|---|---|---|
| `Shipment_Request_HIE__Updated.xlsm` | Location_Inventory | 11 | QuoteStyle.Csv | Inventory Location x Name | Available = List.Sum([On Hand]) | ([Custom] = "JAX") and ([Available] <> null) |
| `Shipment_Request_HIE__Updated.xlsm` | Regional Inv Loc | 8 | QuoteStyle.None | Name x Inventory Location | Available = List.Sum([Available]) | — |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | Current Inventory | 13 | QuoteStyle.None | Name x Inventory Location | InvQty = List.Sum([On Hand]) | ([InvQty] <> null) |

### NetSuite File Cabinet export id=2600949 (account 3492685) — **FORKED**

| Workbook | Query | Steps | QuoteStyle | Grain | Measures (what they really aggregate) | Filters |
|---|---|---|---|---|---|---|
| `Direct_Shipments_Consolidator.xlsm` | Open Order | 4 | QuoteStyle.None | Company Name x Sales Rep x Item | Back Ordered = List.Sum([Back Order Amount]); Qty_BO = List.Sum([Back Order Qty]); Qty_RTS = List.Sum([Qty Committed]); Rate = List.Average([Item Rate]); Ready to Ship = List.Sum([Amount Ready To Ship]) | — |
| `Shipment_Request_HIE__Updated.xlsm` | Open_Order_Report | 8 | QuoteStyle.Csv | Ship Date x Territory x Item | BO = List.Sum([Back Order Qty]) | — |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | Customer Open Orders | 12 | QuoteStyle.None | Item x Ship Date x Regional ID x Document Number x Company Name x Commit x Sales Rep x Balance x Credit Limit x Credit Hold x Status x Item Rate | TtlBO = List.Sum([Back Order Qty]) | — |

### Shipment Request HIE - Updated.xlsm

| Workbook | Query | Steps | QuoteStyle | Grain | Measures (what they really aggregate) | Filters |
|---|---|---|---|---|---|---|
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | HIEProdFlat | 5 | — | — | — | — |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | JingdianProdFlat | 4 | — | — | — | — |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | PoisonInvFlat | 4 | — | — | — | — |

### NetSuite File Cabinet export id=2600946 (account 3492685) — **FORKED**

| Workbook | Query | Steps | QuoteStyle | Grain | Measures (what they really aggregate) | Filters |
|---|---|---|---|---|---|---|
| `Shipment_Request_HIE__Updated.xlsm` | PO | 7 | QuoteStyle.None | Item x Expected Receipt Date x Document Number x Line ID x Original PI Date | Qty = List.Sum([Quantity Remaining]) | — |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | Items On PO | 8 | QuoteStyle.None | — | — | — |

### NetSuite File Cabinet export id=2606291 (account 3492685) — **FORKED**

| Workbook | Query | Steps | QuoteStyle | Grain | Measures (what they really aggregate) | Filters |
|---|---|---|---|---|---|---|
| `Shipment_Request_HIE__Updated.xlsm` | NetSuite Forecast | 12 | QuoteStyle.None | Item x Allocation Group x Forecast Mnth x Product Category | Forecast Qty = List.Sum([Value]) | ([Forecast Qty] <> null and [Forecast Qty] <> 0) |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | Forecast Qty | 18 | QuoteStyle.None | — | — | [Forecast Qty] <> 0 and [Forecast Qty] <> null |

### NetSuite File Cabinet export id=2620991 (account 3492685) — **FORKED**

| Workbook | Query | Steps | QuoteStyle | Grain | Measures (what they really aggregate) | Filters |
|---|---|---|---|---|---|---|
| `Shipment_Request_HIE__Updated.xlsm` | New Prod Alloc | 5 | QuoteStyle.None | — | — | ([Name] <> "BC8 PRE P3 BLU LL" and [Name] <> "BC8 PRE P3 BLU |
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | New Product Allocations | 4 | QuoteStyle.None | — | — | — |

### NetSuite File Cabinet export id=2600950 (account 3492685)

| Workbook | Query | Steps | QuoteStyle | Grain | Measures (what they really aggregate) | Filters |
|---|---|---|---|---|---|---|
| `Shipment_Request_HIE__Updated.xlsm` | Sales | 7 | QuoteStyle.None | Date x Item x Territory | Quantity = List.Sum([Sum of Quantity]) | ([Regional ID] <> null) |

### NetSuite File Cabinet export id=2613629 (account 3492685)

| Workbook | Query | Steps | QuoteStyle | Grain | Measures (what they really aggregate) | Filters |
|---|---|---|---|---|---|---|
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | Purchase Price | 7 | QuoteStyle.None | Name x Item Parent Sub x Product Category (no hierarchy) | Average Cost = List.Max([Average Cost]) | — |

### Shipment Request HIE .xlsm

| Workbook | Query | Steps | QuoteStyle | Grain | Measures (what they really aggregate) | Filters |
|---|---|---|---|---|---|---|
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | HIEInv | 7 | — | SKU | HIE Inv = List.Sum([Ttl Qty]); Ttl Qty = List.Sum([Qty]) | — |

### Supply Chain Update Meeting Workbook.xlsm

| Workbook | Query | Steps | QuoteStyle | Grain | Measures (what they really aggregate) | Filters |
|---|---|---|---|---|---|---|
| `Shipment_Request_HIE__Updated.xlsm` | HIE Shipment | 4 | — | — | — | — |

### https://predatorgroup.sharepoint.com/Share%20All%20Files/Forecast/Reference/Actuals%20vs%20Forecast%20Archive/Sales%20Actuals%20+%20Forecast%20for%20Actuals%20vs%20Forecast.xlsx

| Workbook | Query | Steps | QuoteStyle | Grain | Measures (what they really aggregate) | Filters |
|---|---|---|---|---|---|---|
| `Supply_Chain_Update_Meeting_Workbook.xlsm` | ValidatedSales | 10 | — | Location_lvl_2 x YYYYMM x Name x Year | Units = List.Sum([Sum of Sum of Quantity]) | ([Year] = 2025 or [Year] = 2026); ([Location_lvl_2] <> "Cros |

## What to do with this

Each forked export should be read **once** into the canonical layer, then reused. The reconciliation above is the specification for that single extraction: for every conflict, one side is right, and the answer belongs in `SCHEMA.md` rather than in seven copies of a query.

Order of work:

1. Resolve the blocking conflicts — they are cases where two workbooks are already reporting different numbers under the same label.
2. Decide, per forked export, whether differing grain and scope are intentional. Where they are, both variants derive from one extraction.
3. Settle the naming, once.
