# Wood Production — Panthera Manufacturing cue butt line

**What this is:** the third line workbook, built to the same pattern as `Table_Production.xlsm`
(and the REVO workbook it was cloned from) so the floor runs the same way on every line and
the manager's `Manufacturing_Lines` workbook can read all three identically.

**Source of truth for the process:** `JAX Cue Build - Project Tracker v2` (kickoff 4 Nov 2026,
Jacoby tour 21 Oct). Brandon's nine-step plain-butt scope, two-week rests between turning
stages, the inspector-is-never-the-builder rule, batches of 25 with a 10 % overrun, 32-40 % RH,
and the three build milestones (1 prototype 18 Dec, 25 broomsticks 20 Jan, 50 sellable 11 Jun).

## Files

| Path | What |
|---|---|
| `tools/build_wood_production.py` | generates `work/Wood_Production.xlsx` from scratch (openpyxl). Re-run after changing stations, batches or placeholders |
| `vba/modWoodQuality.bas` | the QC / release flow (import) |
| `vba/frmWoodQuality.frm.txt` | the self-building QC form (paste into a blank UserForm named `frmWoodQuality`) |
| `work/Wood_Production.xlsx` | the built workbook (git-ignored; attach the VBA and save as `.xlsm`) |

## Install (one-time)

1. Open `Wood_Production.xlsx`, **Alt+F11**, File > Import File… → `vba/modWoodQuality.bas`.
2. Insert > UserForm, set **(Name)** = `frmWoodQuality`, View Code, paste `vba/frmWoodQuality.frm.txt` over the default lines.
3. Developer > Insert > Button (Form Control) at the top of `WOOD Floor`, assign macro `WoodQC_Process`, caption it **QC / RELEASE CUE**.
4. Debug > Compile VBAProject, then **Save As `.xlsm`**.

## Sheets

| Sheet | Mirrors (TABLE) | Wood-specific |
|---|---|---|
| `WOOD Floor` | `TABLE Floor` — one row per serialised unit, station ticks, Release? / Hold? / Next Station / PriorityKey, QC Status greys the row | 7 stations WS01–WS07; `Rough Turn Date` / `Taper Date` with computed `Rest 1 Until` / `Rest 2 Until`; `Rest OK?` shows **REST until d mmm** while a rest gate is closed; `Wood Lot` ties a cue to its receiving lot |
| `Plan Config` | station capacity table, plan clock, zone roll-up, bottleneck | `Rest After (days)` and `Inspector ≠ Builder?` per station; wood parameters (RH window, moisture range, overrun %, batch size); **BATCH PLAN** that drives Production Plan; calendar lead time = touch days + rest days |
| `Holds`, `Release Log`, `Reject Log`, `Rework Tracker`, `QC Log`, `QC Reasons` | identical column sets and table names (`tbl_ReleaseLog` …) | `QC Reasons` pre-seeded with twelve wood reasons from the Jacoby notes; still free text |
| `Production Plan` | row 1 ISO week, row 2 Monday, row 3 capacity, row 4 totals, SKUs from row 5 | cells are formulas over the BATCH PLAN, not typed numbers; totals turn red when they exceed the weakest station |
| `Weekly Reporting`, `Summary` | same forecast / actual / reject formulas by family via the classification sheet | first pass yield column; live floor status block (resting, on hold, ready for QC …) |
| `Cue Classification` | `Table Classification` | two placeholder SKUs until INV2 (SKU convention) and STR3 (sellable design) land |
| `Wood Receiving` | — (REVO has incoming inspection; TABLE does not) | QA1: grade A/B/C, moisture, RH at receipt, disposition, Lacey Act reference, moisture-in-range check |
| `Humidity Log` | — | QA3: readings per location, EXCURSION flag against 32-40 % RH, action taken |

## Decisions worth knowing

- **Stations are grouped, not Brandon's nine steps one-to-one.** Nine ticks per cue is noise on a
  board; seven stations keep the turning / machining / assembly / inspection boundaries that
  matter for rest gates and for the inspector rule. The station table is config: add a row on
  Plan Config and a matching header on the floor and the VBA picks it up (it reads station ids
  from `Plan Config!A6:A12`).
- **Rest is tracked with dates, not ticks.** A tick says a step happened; a rest gate needs to
  know *when*. Two date columns and two formulas give the board "what is going to happen" for
  wood, which is the thing the TABLE board never had to answer.
- **Every minute on Plan Config is a placeholder** (yellow, with a cell note). Chad and Steven's
  setup sheets after visit 1 (P1.3) replace them. Nothing downstream changes.
- **Planned starts include the 10 % overrun.** Jacoby overruns every order for movement and
  mistakes. Planning 25 and hoping for 25 good is how you miss M6; the board seeds 28 and 55.
- **No Daily Shift / Scheduler / Component Tracker.** One builder and one part-timer do not
  need a shift scheduler. Add those sheets when the second builder is trained (P2.3).
- **Finish is out of scope.** Interim finishing is at Jacoby (P2.4); a booth is the M7 gate.
  `Release` on this board means *built, inspected, ready for finish / inventory*.

## Next steps

1. After visit 1: replace station minutes, confirm the station grouping against the setup sheets, set the moisture range (QA1) and the humidity excursion rule (QA3).
2. When INV2 lands: replace the placeholder SKUs on `Cue Classification`, `Plan Config` BATCH PLAN and `WOOD Floor`.
3. First prototype week: run the QC form once end-to-end with Chad or Steven in the room; fix wording, not logic.
