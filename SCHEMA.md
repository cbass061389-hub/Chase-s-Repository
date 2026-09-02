# Canonical Model — proposal for sign-off

**Status: PROPOSED. Phase 2 does not start until you sign this off.**

This is the contract between the workbook estate and everything downstream —
warehouse, `SC_Reference.xlsx`, HTML engine. Change it after Phase 2 exists and
you rewrite readers, so it gets argued about now instead.

Read the four rules, skim the entities, then go to
[Decisions I need from you](#decisions-i-need-from-you) at the bottom. That
section is the actual ask.

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

Phase 2 is blocked on these. Everything else I will decide and tell you about.

1. **Source of record per domain.** For each domain with more than one candidate
   in `discovery/manifest.json`, which file wins? The sweep flags the pairs; it
   cannot know your intent. In-transit is the live example — the In Transit
   workbook and the Table In-Transit Dashboard hold the same columns with
   different row counts.

2. **`ownership` values.** Is the four-way split
   (`predator_paid` / `supplier_held_unpaid` / `consignment` / `in_bond`)
   right, or does the estate only really distinguish paid from unpaid?

3. **`shipment_id`.** Container # as the key, falling back to invoice # when
   absent — correct? It breaks for LCL consolidations where one container carries
   several invoices. Is that real in your lanes?

4. **`region_channel` list.** I have `americas`, `emea`, `apac`, `b2c`,
   `tradeshow`, `sponsorship`, `marketing`, `hq_reserve`. Missing or extra?

5. **The Worst Flag boolean tests.** The exact conditions for `CAPITAL TRAP`,
   `MOSTLY DEAD`, `AT RISK` and `OVER-FORECASTED`. Either state them, or
   unprotect the VBA project so I can read them — I will not reverse-engineer
   thresholds from output values and quietly get them slightly wrong.

6. **`data_as_of` per source.** Which sources state their own as-of date in a
   cell or a filename, and where? Everything else falls back to unknown, and a
   `Data Vintage` tab full of "unknown" is a tab nobody trusts.

7. **Forecast horizon.** How many months forward does `months_supply` divide by —
   3, 6, or 12? It changes every overstock and dead-stock call in the system.

