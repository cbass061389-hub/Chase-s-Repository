# FLOOR-MKR — shared floor app for REVO, TABLE and WOOD

Published artifact: https://claude.ai/artifact/4SZKc2AgpTqKzUdVcWpXsA (private until shared from its Share menu).

One page, PG Apps UI design system, shared database. Everyone who opens it sees the same
board; every tick, QC event, hold and plan cell is stored centrally and appears for
everyone within seconds. It replaces the per-line workbook floors for day-to-day work; the
workbooks stay as the reporting and planning layer until the warehouse reads the app's data.

## Sections (binder tabs)

| Tab | For whom | What |
|---|---|---|
| Floor | operators | one row per serialised unit (or cart on REVO), station ticks, rest gates, status, QC button |
| Holds | operators, leads | open holds with station and reason; return a unit to the floor |
| Log | leads, quality | every QC event: who, when, station, reason, notes, quantity split; CSV export |
| Plan | planner | planned starts per SKU per week |
| Lines | manufacturing manager | planned vs released, FPY, reject and rework rates against targets, 13-week trends, reject Pareto, targets |

Preferences (gear) holds theme (Auto / Dark / Light / Blue), density, and for editors the
Setup panel: line names, stations in order with rest days, learned reasons with categories.

## Roles (set in the artifact's Share menu)

| Share level | Can |
|---|---|
| Viewer / Commenter | see everything |
| Contributor | tick stations, run QC, place and release holds, edit the plan |
| Editor / Owner | all of the above plus add units, change setup and targets, delete log events |

## Data model (artifact database)

| Collection | One document per | Key fields |
|---|---|---|
| `config/lines` | the three lines | name, unitNoun, color, stations[{id,name,zone,restDays}] |
| `settings/targets` | — | attainment, fpy, reject, rework |
| `units/{LINE_UNITID}` | serialised unit or cart | line, unitId, wo, sku, serial, qty, batch, target, stations{}, stationDates{}, status, statusDate, order |
| `events/{auto}` | QC submit | ts, date, line, unitId, serial, outcome, station, reason, notes, qtyReleased/Rework/BGrade/Reject/Hold, by |
| `holds/{auto}` | hold | line, unitId, station, reason, dateHeld, released, releasedAt, by |
| `plan/{LINE_WEEK_SKU}` | planned starts | line, weekStart (Monday), sku, qty |
| `reasons/{LINE_text}` | learned reason | line, text, timesUsed, firstSeen, lastUsed, active, category |

Seeded on 8 Nov 2026 from the TABLE Floor (70 tables, WEEK 1-9, six fully ticked) and the
WOOD batch plan (84 cues, three batches, 15 plan weeks), plus twelve starter wood reasons.

REVO reseeded on 13 Nov 2026 from `Revo_Production - New.xlsm`: the 15-op sequence in floor
order (IC, 10, 20, 30, 100, ASSYM, 40, 110, 50, 60, 70, 120, 130, 160, FQC; cure after 100 and
40), the Quality_Codes lists (defect type, location, root cause, corrective action), 39 carts
from the REVO Floor with their op ticks and released-to-date, 18 holds (2 open), the Production
Plan from the week of 9 Nov (Friday week-ends converted to Monday week-starts), and 187 release
passes since 10 Aug with the Reject Log descriptions attached. `tools/seed_revo.py` regenerates it.

### Cart lines (REVO)

A line flagged `partialRelease` treats a unit as a cart that releases in passes, mirroring the
REVO release form: release qty for the pass, then rework / B-grade / reject out of it, to
inventory computed, left-on-cart after the pass. When the line has quality codes, every
disposition is explained line by line (qty, defect type, location, detected op, root cause,
corrective action, note) and submit is blocked until the lines add up to each disposition.
Events carry `details[]`; the Lines Pareto counts coded reject lines by defect type.
Units carry `qtyReleased` and `qtyOut`; the cart closes when nothing is left.

`tools/seed_from_workbooks.py` regenerates the seed documents; they are written with the
ArtifactData tool, never from the page.

## Decisions

- **Dark-first PG Apps UI, no React bundle.** Hand-written CSS on the system's tokens and
  fonts (copied from the design-system artifact) so the page is one file with no dependency.
- **Ticks write stationDates.** A rest gate needs the date a stage finished; ticking a
  station records it, so nobody types dates.
- **Events are append-only.** QC never edits a row; a correction is a new event, and only an
  editor can delete one. Reasons are free text that the list learns, as in the workbooks.
- **Carts and units share one model.** A REVO cart has qty > 1 and splits by quantity at QC;
  a table or cue has qty 1 and picks an outcome. Same log, same dashboard maths.
- **Rules:** everyone reads; Contributors write shared data; `config` and `settings` need
  Editor. The Log's delete is admin-only in the UI; the store does not enforce it.

## Limits

- The Log and Lines read the latest 1,000 events. At REVO volumes archive older events
  (export CSV, then delete) once a quarter, or move reporting onto the warehouse.
- Phone is view-friendly; ticking is meant for a tablet or desktop at the station.
- Names come from claude.ai profiles; only ids are stored.
