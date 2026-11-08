# Manufacturing Lines — throughput & quality across REVO, TABLE and WOOD

**What this is:** the manufacturing manager's one view of the three lines. It reads each line's
own workbook through Power Query and answers, per line and for all three: what happened
(released vs planned, week / month / year to date), why (reject Pareto), the risk (first pass
yield and reject % trend over 13 weeks) and what to do (flags against targets you set).

Nobody types in this file. The three line workbooks stay the systems of record; this file
refreshes from them.

## Files

| Path | What |
|---|---|
| `tools/build_manufacturing_lines.py` | generates `work/Manufacturing_Lines.xlsx` (`--sample` adds fake rows to prove the formulas and charts; never ship that build) |
| `queries/*.m` | the Power Query text, one file per query |
| `work/Manufacturing_Lines.xlsx` | the built workbook (git-ignored) |

## Install (one-time, ~10 minutes)

1. On `Config`, set the three workbook paths in the `Lines` table (OneDrive-synced local paths work; SharePoint URLs need `Web.Contents` instead of `File.Contents`, change it in the `fnLine*` queries if you go that way).
2. Data > Get Data > From Other Sources > **Blank Query** > Advanced Editor. Paste each `.m` file and name the query **exactly** as the file:
   - `Lines`, `fnLineReleases`, `fnLineRejects`, `fnLinePlan` → Close & Load To… > **Only Create Connection**
   - `Releases` → Close & Load To… > Existing worksheet > `Data_Releases!$A$1`
   - `Rejects` → `Data_Rejects!$A$1`, `Plan` → `Data_Plan!$A$1`, `RejectPareto` → `Data_Pareto!$A$1` (accept the overwrite of the placeholder header row)
3. Data > **Refresh All**. Set the query properties to refresh on open if you want it hands-off.

## What each sheet does

| Sheet | Content |
|---|---|
| `Dashboard` | WTD / MTD / YTD blocks per line: planned, released, rework, B-grade, reject, attainment, first pass yield, reject %, rework %, a throughput flag and a quality flag. Below: released vs planned per week per line (three small charts), reject % by line and FPY by line over 13 weeks |
| `Weekly` | the 13-week numbers behind the charts |
| `Pareto` | top ten reject reasons per line, last 13 weeks, with share |
| `Config` | paths, targets (attainment 95 %, FPY 90 %, reject ≤ 5 %, rework ≤ 10 %), period definitions |
| `Data_*` | landing tables for the queries, one common shape per entity |

Flags are icon + word (`● On plan`, `▲ Watch`, `■ Behind`; `● Good`, `▲ Watch`, `■ Act`), coloured with the
status palette, so they read in print and for colour-blind readers.

## How the three line formats are reconciled

| | REVO | TABLE | WOOD |
|---|---|---|---|
| Release log | sheet `Release Log`: Date (text m/d/yyyy), SKU, Cart, ReleasedQty, WO, Rework, B Grade, Reject | table `tbl_ReleaseLog` | table `tbl_ReleaseLog` |
| Reject log | sheet `Reject Log`: the old macro wrote SKU into the WO column and WO into SKU; the query un-swaps them | `tbl_RejectLog` | `tbl_RejectLog` |
| Plan | `Production Plan` grid: A = SKU, row 2 = Monday dates, data from row 5 | same | same (formula-driven from the batch plan) |
| Unit | cart | serialised table | serialised cue |

The REVO layout is taken from the REVO macros that still live inside `Table_Production.xlsm`
(`REVO_Ops.bas`), not from the REVO workbook itself, which was not available. Verify the sheet
names and the date format against the live file on the first refresh.

## Decisions worth knowing

- **Power Query, not links.** External-link formulas break the moment a file is renamed and
  give no error you can act on. A query that cannot find a file fails the refresh loudly, by
  design: a line showing zero is the worst outcome.
- **Whole-column SUMIFS over the landing sheets.** The dashboard formulas never reference the
  query tables by name, so reloading or re-creating a query cannot break them.
- **FPY counts rework and B-grade against yield**, not just reject. A unit that needed a second
  touch did not pass first time; that is the number that tells you whether the process is in
  control, and it is the same definition on all three lines.
- **13 weeks of trend.** Long enough to see a shift, short enough to stay readable at a glance.
  Change `TREND_WEEKS` in the builder if you want a quarter plus.
- **Pareto on the reason text operators typed.** Until the lines categorise reasons (the
  `Category (later)` column on each QC Reasons sheet), this is as good as the data gets, and it
  is what tells you which categories to create.

## Next steps

1. Confirm the three paths and run the first refresh against real files; fix any sheet-name mismatch in the `fnLine*` queries (REVO is the one to watch).
2. Agree the four targets on `Config` with Osvaldo; they drive every flag.
3. When the lines have categorised reasons, change `RejectPareto` to group on category instead of text.
4. Longer term this is the shape of the `production` and `quality` entities in `SCHEMA.md`; the `sc` extraction layer can read the same three workbooks and land the warehouse without Power Query.
