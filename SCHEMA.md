# Canonical Model — proposal for sign-off

**Status: PROPOSED. Phase 2 does not start until you sign this off.**

This is the contract between the workbook estate and everything downstream —
warehouse, `SC_Reference.xlsx`, HTML engine. Change it after Phase 2 exists and
you rewrite readers, so it gets argued about now instead.

Read [What the sweep found](#what-the-sweep-found) first — it changes the shape
of the answer — then the four rules, then
[Decisions I need from you](#decisions-i-need-from-you).

---

## What the sweep found

Discovery has now run against `Supply_Chain_Update_Meeting_Workbook.xlsm`,
`Shipment_Request_HIE__Updated.xlsm` and `Direct_Shipments_Consolidator.xlsm`.
Three things it turned up change the plan.

### 1. The system of record is NetSuite, not the workbooks

All three workbooks refresh from the **same eight NetSuite File Cabinet
exports**, addressed as tokenised `media.nl?id=...` URLs. Five of those eight are
read by more than one workbook, each applying its own transformation:

| Export id | Read by | Under the query name | Canonical entity |
|---|---|---|---|
| 2600946 | Meeting, HIE | Items On PO / PO | `open_po` |
| 2600947 | Meeting, HIE (x2) | Current Inventory / Location_Inventory / Regional Inv Loc | `inventory_onhand` |
| 2600949 | Meeting, HIE, DSR | Customer Open Orders / Open_Order_Report / Open Order | `demand` (`open_so`) |
| 2600950 | HIE | Sales | `demand` (`actual_shipped`) |
| 2606291 | Meeting, HIE | Forecast Qty / NetSuite Forecast | `demand` (`forecast`) |
| 2613629 | Meeting | Purchase Price | `item` (cost) |
| 2620991 | Meeting, HIE | New Product Allocations / New Prod Alloc | `allocation` |
| SharePoint | Meeting | ValidatedSales (Actuals vs Forecast) | `demand` confidence basis |

**This is why two reports disagree.** They are not disagreeing about the data —
they read the identical export. They are disagreeing about the *logic* applied
after it lands. Consolidating the workbooks without consolidating the query logic
would move the problem, not solve it.

So the canonical data layer reads those eight exports **once each**, applies one
agreed transformation per entity, and the workbooks become consumers. That is the
consolidation.

### 2. There is a circular refresh dependency

```
Supply_Chain_Update_Meeting_Workbook.xlsm
    -> reads HIEInv, HIEProdFlat, JingdianProdFlat, PoisonInvFlat
       from Shipment Request HIE
Shipment_Request_HIE__Updated.xlsm
    -> reads "HIE Shipment"
       from Supply Chain Update Meeting Workbook
```

Refresh order is undefined. Whichever refreshes second reads the other's previous
output, so the numbers depend on click order. This has to be broken before
anything downstream is trustworthy, and the fix is structural: both read the
warehouse, neither reads the other.

### 3. One of those reads points at a superseded file

The meeting workbook's `HIEInv` query reads
`...\Supply Chain Files\Chase\Shipment Request HIE .xlsm` — note the space
before the extension — while `HIEProdFlat`, `JingdianProdFlat` and
`PoisonInvFlat` read `...\Shipment Request HIE - Updated.xlsm`. Two different
files, one asset. HIE inventory in the meeting workbook is currently sourced from
the older file. That is a live wrong-number bug, and the sweep could not find the
older file at all — so it may already be gone, in which case that query is
failing silently or serving a cached result.

### 4. Scale, and what it means for the model

| | Meeting | HIE | DSR |
|---|---|---|---|
| Sheets | 76 (51 hidden) | 36 (7 hidden) | 13 (0 hidden) |
| Power Query | 11 | 8 | 1 |
| Workbook connections | 14 | 8 | 1 |
| VBA code modules | 3 (`InvReview`, `PythonRun`, `Alloc_Engine`) | 1 (`Module1`) | 18 (`modDSR_*`) |
| VBA protected | yes | yes | yes |
| Power Pivot data model | 22 parts | — | — |
| Formula cells | ~277,000 | ~8,000 | 1 |

Two consequences:

- **The sheet, not the workbook, is the unit of consolidation.** A 76-tab
  workbook is not one entity; each tab is a different one. The sheet-level entity
  map in `discovery/DISCOVERY.md` is the artifact to review, not the file table.
- **`Position Engine` (59,751 formulas), `Allocation Plan` (47,534),
  `__CleanBO` (37,451), `BOMMaster` (35,367) and `__Alloc Engine` (24,528) are
  calculation engines, not data.** They are Phase 2 *logic* to port, not tables
  to read. All are hidden or veryHidden, and the VBA that drives them is
  password-protected — see decision 5.

You already have a `SKU (Canon)` column in the DSR workbook, which is the
crosswalk concept below under a different name. I have kept the name
`item_crosswalk` for the table and will map `SKU (Canon)` onto it.

---

## The four rules

**1. SKU is the join key everywhere.**
One `item` table owns the canonical SKU. Supplier item codes never join directly —
they resolve through `item_crosswalk` first. A row whose SKU does not resolve
becomes an exception; it is never dropped and never guessed at.

**2. Every row carries its lineage.**
`source_file`, `source_sheet`, `source_row`, `extracted_at`, `data_as_of` on every
fact row, no exceptions. This is what makes a number in the HTML engine traceable
back to a cell. Without it the outputs are unauditable and nobody trusts them.

**3. Units and currency are declared, never inferred.**
Every quantity field names its UOM field. Every money field names its currency
field and carries `fx_rate_to_usd` plus `fx_as_of`. A quantity with no declared
UOM fails validation rather than defaulting to "each".

**4. Conflicts are data.**
When two sources disagree, both values are written to `exceptions` with both
lineages and the delta. Nothing picks a winner silently. A source-of-record
precedence list (which you set, per domain) is the *only* thing allowed to
resolve a conflict, and when it does, the losing value still lands in
`exceptions` for the record.

### Missing is missing

No interpolation, no carry-forward, no zero-fill. A null on-hand means the source
did not report it, and it renders as "—", not as `0`. `0` means the source said
zero. This costs coverage and buys the only thing that matters: a number nobody
has to second-guess.

---

## Lineage columns (on every fact table)

| Field | Type | Notes |
|---|---|---|
| `source_file` | str | absolute path as swept |
| `source_sheet` | str | sheet or `(flat file)` for CSV |
| `source_row` | int \| null | 1-based physical row; null for aggregates |
| `source_id` | str | stable 12-char id from the discovery manifest |
| `extracted_at` | datetime (UTC) | when the reader ran |
| `data_as_of` | date | **the vintage of the data itself**, not the file mtime |

`data_as_of` vs `extracted_at` is the distinction that matters. A file modified
this morning can hold a snapshot from three weeks ago. Where a source states its
own as-of date, that wins. Where it does not, `data_as_of` is null and the
`Data Vintage` tab shows it as unknown — which is itself a finding.

---

## Entities

Grain is stated as the column set that must be unique. A duplicate on that set
is a blocking validation failure.

### `item` — the SKU spine
**Grain:** `sku`

| Field | Type | Unit | Notes |
|---|---|---|---|
| `sku` | str | — | canonical key, uppercase-trimmed |
| `description` | str \| null | — | |
| `family` | str \| null | — | |
| `product_line` | str \| null | — | |
| `brand` | str \| null | — | Predator / Poison / REVO / BK |
| `status` | enum | — | `active`, `phase_out`, `discontinued`, `new_product` |
| `abc_class` | enum \| null | — | `A`/`B`/`C`, **by trailing-12-month revenue**, never units |
| `standard_cost` | decimal \| null | `cost_currency` | |
| `cost_currency` | str | ISO 4217 | |
| `base_uom` | str | — | required; no default |

### `item_crosswalk` — supplier codes to canonical SKU
**Grain:** `supplier_id + supplier_item_code`

Deliberately separate from `item`. One SKU has many supplier codes, and a
supplier code can be reused across SKUs over time, so `valid_from`/`valid_to`
are on this table.

### `inventory_onhand`
**Grain:** `sku + location_id + ownership + snapshot_date`

| Field | Type | Unit | Notes |
|---|---|---|---|
| `qty_on_hand` | decimal | `uom` | |
| `qty_available` | decimal \| null | `uom` | on-hand less committed |
| `qty_committed` | decimal \| null | `uom` | |
| `ownership` | enum | — | `predator_paid`, `supplier_held_unpaid`, `consignment`, `in_bond` |
| `extended_value` | decimal \| null | `currency` | |
| `snapshot_date` | date | — | part of the grain |

`ownership` is a first-class dimension, not a flag. Supplier-held-unpaid stock
is a different financial and availability animal from paid stock in Jacksonville,
and every downstream number that sums them together is wrong.

### `open_po`
**Grain:** `po_number + po_line`

`qty_ordered`, `qty_received`, `qty_remaining` (derived, not read),
`committed_date`, `revised_date`, `days_late` (derived), `supplier_id`,
`unit_cost` + `cost_currency`, `status`.

`qty_remaining = qty_ordered - qty_received`, floored at 0, computed here and
never read from a source column — sources disagree on whether it nets
cancellations.

### `in_transit_shipment` / `in_transit_line`
**Header grain:** `shipment_id` (container # where present, else invoice #)
**Line grain:** `shipment_id + sku + channel`

Header: `container_number`, `invoice_number`, `bl_number`, `pol`, `pod`,
`etd`, `eta`, `revised_eta`, `mode` (`ocean`/`air`/`truck`/`courier`),
`status`, `supplier_id`, `vessel`, `carrier`.
Line: `qty`, `uom`, `channel`, `allocation_ref`, `unit_cost`, `cost_currency`.

Split into two tables because channel splits live at line level while ETA lives
at header level. Flattening them double-counts ETA in every "arriving this week"
figure.

**A shipment with no container and no ETA is a blocking validation failure**, not
a null. That row cannot answer "what is going to happen?", which is the whole
point of the table.

### `allocation`
**Grain:** `sku + region_channel + effective_date`

`region_channel` enum: `americas`, `emea`, `apac`, `b2c`, `tradeshow`,
`sponsorship`, `marketing`, `hq_reserve`.
Fields: `qty_allocated`, `qty_consumed`, `qty_remaining` (derived),
`np_launch_control` (bool), `hq_lock_qty`, `lock_expires`, `approver`.

**Validation gate:** total allocated per SKU must not exceed on-hand plus
in-transit plus open PO. Breach writes to `exceptions` with the shortfall in
units and dollars — that is an over-promise, and it is the single most
expensive thing this system can catch early.

### `dsr_request` / `dsr_line`
**Header grain:** `request_id` · **Line grain:** `request_id + sku`

Header: `requester`, `region`, `requested_at`, `decision`
(`approved`/`denied`/`overridden`/`pending`), `decision_reason`,
`decided_by`, `decided_at`.
Line: `qty_requested`, `qty_approved`, `uom`, plus **`inventory_snapshot_json`** —
the availability position at the moment of the decision.

That snapshot is the point of the table. Without it you can see a decision but
never reconstruct whether it was the right call on the information available,
which makes override patterns unarguable-with.

### `production`
**Grain:** `production_order_id + sku + stage`

`supplier_id`, `qty_planned`, `qty_completed`, `qty_b_grade`, `qty_rework`,
`qty_scrap`, `uom`, `scheduled_completion`, `actual_completion`,
`schedule_variance_days` (derived), `stage`, `status`.

B-grade, rework and scrap are separate quantity columns, not a single "loss"
column. They have different dispositions and different root causes.

### `demand`
**Grain:** `sku + period_start + demand_type`

`demand_type`: `forecast`, `open_so`, `actual_shipped`, `budget`.
`qty`, `uom`, `period_start`, `period_end`, `forecast_version`,
`confidence_score`, `confidence_basis`.

One table with a `demand_type` discriminator rather than four tables — forecast
versus actual comparison is the most common question asked of it, and separate
tables make that a join every single time.

### `quality`
**Grain:** `inspection_id + sku + defect_code`

`inspection_date`, `inspection_type` (`pre_shipment`/`incoming`/`in_process`/`final`),
`supplier_id`, `qty_inspected`, `qty_rejected`, `reject_rate` (derived),
`defect_code`, `defect_description`, `severity` (`critical`/`major`/`minor`),
`car_number`, `car_status`, `disposition`.

### `supplier`
**Grain:** `supplier_id`

Known set: HIE/Hamson, Jingdian, Yalin/Shandong BCN, Box Billiards, Apollo.
`nominal_lead_time_days`, `actual_lead_time_p50`, `actual_lead_time_p90`
(both derived from `open_po` receipt history), `open_spend`, `spend_currency`,
`country`, `status`.

P50 and P90 rather than an average. An average lead time hides the tail, and the
tail is what breaks a launch date.

### `tariff`
**Grain:** `trade_lane + hts_category + effective_date`

`rate_pct`, `additional_duty_pct`, `basis` (`fob`/`cif`/`customs_value`),
`effective_date`, `expires_date`, `authority`, `notes`.
Date-effective and never overwritten — a landed-cost figure from March has to
stay reproducible after a rate change.

### `exceptions` — the trust layer
**Grain:** `exception_id`

`exception_type` (`duplicate_key`, `orphan_sku`, `negative_qty`,
`stale_snapshot`, `missing_container_or_eta`, `allocation_over_supply`,
`uom_mismatch`, `source_conflict`), `severity` (`blocking`/`warning`/`info`),
`entity`, `entity_key`, `left_value`, `right_value`, `delta`,
`delta_dollars`, `left_lineage`, `right_lineage`, `detected_at`.

Sorted by `delta_dollars` descending in every output. Dollars rank the queue;
counts do not.

### `source_vintage` — populated from the discovery manifest
**Grain:** `source_id + run_id`

`path`, `file_modified_at`, `data_as_of`, `row_count`, `refresh_status`,
`probe_status`, `risk_band`. This drives the `Data Vintage` tab.

---

## Derived metrics — your existing logic, carried forward verbatim

Thresholds live in `sc/config.yaml`, not in code and not in formulas.

| Metric | Definition |
|---|---|
| `months_supply` | on-hand ÷ average monthly forward demand (forecast where present, else trailing 3-month actual) |
| `overstock` | `months_supply > 8.0` |
| `dead_stock` | forward forecast = 0 across the horizon |
| `forecast_confidence` | 3-month lookback: forecast vs actual variance |
| `abc_class` | trailing-12-month revenue, A ≤ 80% cumulative, B ≤ 95%, C beyond |
| `days_late` | today − `committed_date` for unreceived PO lines |
| `worst_flag` | first match down the hierarchy below |

**Worst Flag hierarchy** (evaluated top-down, first match wins):

1. `CAPITAL TRAP`
2. `MOSTLY DEAD`
3. `AT RISK`
4. `OVERSTOCK`
5. `LOW CONFIDENCE`
6. `OVER-FORECASTED`
7. `MONITOR`

I have carried the hierarchy and thresholds across as-is. I have **not** ported
the exact boolean test for `CAPITAL TRAP`, `MOSTLY DEAD`, `AT RISK` and
`OVER-FORECASTED`, because those live in the meeting workbook's formulas and VBA
and the discovery sweep flags that project as password-protected. See decision 5.

---

## Validation gates — these block a publish

Non-zero CLI exit, exception list in the log, no output written.

| Gate | Blocks on |
|---|---|
| Duplicate keys | any entity's stated grain not unique |
| Orphan SKUs | fact-table SKU absent from `item` after crosswalk |
| Negative on-hand | `qty_on_hand < 0` |
| Stale snapshot | `data_as_of` older than `snapshot_stale_days` (7) |
| In-transit completeness | line with no container **and** no ETA |
| Allocation over supply | allocated > on-hand + in-transit + open PO |
| UOM mismatch | quantity field with no declared UOM, or UOM conflict on one SKU |
| Currency undeclared | money field with no currency |
| Unresolved conflict | `source_conflict` with no precedence rule to resolve it |

Warnings that do *not* block: unknown `data_as_of`, low forecast confidence,
`item` attributes missing.

---

## Decisions I need from you

Phase 2 is blocked on these. They are narrower than they were before the sweep —
the workbooks answered most of the earlier questions themselves.

1. **Break the cycle — which direction?** My recommendation: the HIE workbook is
   upstream for production and supplier-held inventory; the meeting workbook
   consumes it and never the reverse. So `HIE Shipment` in the HIE workbook stops
   reading the meeting workbook and reads the warehouse instead. Confirm, or tell
   me it is the other way round.

2. **The eight exports — one transformation each.** For every forked export
   above, the per-workbook query logic differs. I need to know, per export, which
   workbook's version is *correct* — or that they are deliberately different
   (e.g. `Location_Inventory` and `Regional Inv Loc` both read export 2600947 but
   at different grains, which looks intentional). Answer per row of that table;
   this is the single biggest item.

3. **`Shipment Request HIE .xlsm` — dead or alive?** The sweep could not find it.
   If it is gone, `HIEInv` in the meeting workbook is broken right now and I need
   to know what it should read. If it exists, it is in a folder I did not sweep.

4. **`ownership` for HIE stock.** The HIE sheets distinguish
   `HIE total (finished at supplier)` from JAX inventory. Is
   supplier-held-unpaid the right label for the former, and is
   `predator_paid` / `supplier_held_unpaid` sufficient, or do you also need
   `consignment` and `in_bond`?

5. **The protected VBA.** Three projects, all password-protected, holding
   `Alloc_Engine`, `InvReview`, `PythonRun` and 18 `modDSR_*` modules. Those
   modules *are* the allocation and release logic — the Worst Flag tests, the
   go/no-go gate, the release packet rules. I cannot port what I cannot read, and
   I will not infer thresholds from output values and get them quietly wrong.
   Either unprotect them for the port, or state the rules. `PythonRun` suggests
   you already shell out to Python from VBA; if that script is on disk, point me
   at it and it may answer this for free.

6. **Allocation grain.** `NP Allocation Table` (1,923 rows, meeting) and
   `New Prod Alloc` (1,887 rows, HIE) both derive from export 2620991 and differ
   by 36 rows. Which is right, and is allocation keyed by
   `sku + region_channel + effective_date` as proposed, or is there a launch/wave
   dimension I am missing?

7. **`data_as_of`.** The exports are pulled live at refresh, so `data_as_of` is
   the refresh timestamp — fine. But `Projected Inventory`, `BK Dist Plan` and the
   `S&OP 6.3.2026` sheet are point-in-time. Where do those record their own
   as-of date, or should I take file modified time and label it as an estimate?

8. **Forecast horizon** for `months_supply` — 3, 6, or 12 months forward? It
   changes every overstock and dead-stock call. `Forecast Qty` is 154 lines of M
   and already does period bucketing, so tell me which bucket to divide by and I
   will match the existing logic rather than inventing a parallel one.
