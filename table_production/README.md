# Table Production — QC / release flow for TABLE Floor

**What this adds:** a Quality form that captures a serial number, an outcome
(Release / B-Grade / Rework / Reject / Hold), a free-text reason and notes for
one serialised table, logs every event, and greys the table's row on
`TABLE Floor`.

**Why it is a new flow, not an edit:** the 16 stacked `RELEASE CARTS` buttons on
`TABLE Floor` all ran `ReleaseToShippingAndReceiving` in `REVO_Ops`. That macro
is the REVO shaft flow: it reads a sheet called `REVO Floor` (which this workbook
does not have — the first thing it would do is create an empty one), works at
cart level with a quantity split, and writes dates as text. `TABLE Floor` is a
different shape: one row = one serialised table, station ticks, a `Release?`
formula, a `Serial` column, and log tables that already have `Serial` /
`Unit ID` / `Station` columns waiting to be used. So the table flow is built
for that layout and the REVO code is left untouched.

## Install (one-time, ~2 minutes)

1. Open `work/Table_Production.xlsm` (the patched copy — see below) and press
   **Alt+F11**.
2. **File > Import File…** → `table_production/vba/modTableQuality.bas`.
3. **Insert > UserForm.** In the Properties window set **(Name)** to
   `frmTableQuality`. Right-click it > **View Code**, paste the whole of
   `table_production/vba/frmTableQuality.frm.txt`, replacing the two default
   lines. (The form builds its own controls at runtime, so the design surface
   stays blank — that is expected.)
4. Save as `.xlsm`.

Nothing else to wire: the buttons already point at `TableQC_Process`.

## How it works

**Button: `QC / RELEASE TABLE`** (top of `TABLE Floor`) → `TableQC_Process`

| You have selected | What happens |
|---|---|
| a cell on a table's row | that one table is QC'd |
| anything else | every table with `Release? = TRUE` and no `QC Status` is QC'd one after another (confirmed first; Skip / Cancel all on the form) |

The form shows Unit ID, WO#, SKU, target date, zone and next station, and asks for:

| Field | Rule |
|---|---|
| Serial | scan or type; required to **Release** (serial is applied at ST05), optional otherwise; upper-cased |
| Outcome | Release / B-Grade / Rework / Reject / Hold. Defaults to Release when every station is ticked |
| Station | ST01–ST05 from `Plan Config!A6:A10` plus PreAssym; defaults to the row's Next Station |
| Reason | free text. The dropdown offers every reason used before; a new one is remembered. Required for anything but Release |
| Notes | free text |

On submit, one event writes to:

| Where | What |
|---|---|
| `TABLE Floor` | `Serial` (col D), `QC Status` (col R), `QC Date` (col S). Conditional formatting greys A:Q when QC Status is set; REJECTED also strikes through. Hold writes no status — the existing `Hold?` formula shows ON HOLD from the Holds sheet |
| `QC Log` → `tbl_QCLog` (created on first use) | the audit row: timestamp, user, WO#, SKU, Unit ID, serial, outcome, station, reason, notes, zone, target, floor row |
| `tbl_ReleaseLog` | one row per table with the 1-unit split (Qty Released / Rework / B Grade / Reject), real date, serial. `Summary` and `Weekly Reporting` read this |
| `tbl_RejectLog` | on Reject: date, SKU, WO#, Unit ID, qty 1, description (reason — notes), serial, station |
| `tbl_ReworkTracker` | on Rework: the row plus the ST01–ST05 ticks as they stood |
| `Holds` sheet | on Hold: Unit ID, station, reason, date, Released? = FALSE |
| `QC Reasons` → `tbl_QCReasons` (created on first use) | reason, times used, first/last seen, Active, and an empty **Category (later)** column |

**Undo:** select the row(s) and run `TableQC_ClearStatus` (Alt+F8). It clears
QC Status / QC Date so the row un-greys; log rows are left for you to delete,
because silently removing audit rows is worse than a stray one.

## Design decisions

- **Columns are resolved by header text on row 3**, never by letter. Insert a
  column on `TABLE Floor` and nothing breaks. If `QC Status` / `QC Date` are
  missing (an un-patched copy of the sheet) the macro creates them.
- **Greying is conditional formatting keyed on `QC Status`**, not paint applied
  by the macro. Clearing the status is the only undo needed, and the rule
  survives sorting and row moves.
- **Reasons are free text that the list learns from.** You said you are still
  learning the operation; forcing a pick-list now would produce "Other" as the
  top reason. When the data is there, fill `Category (later)` on `QC Reasons`
  and set `Active = FALSE` on the noise — the form picks that up with no code
  change. Reporting by category is then a lookup from `QC Log[Reason]`.
- **Dates are written as dates.** The REVO flow wrote `Format(Date, "m/d/yyyy")`
  text; the table reports compare `tbl_ReleaseLog[Date]` to dates with
  SUMPRODUCT, so text would silently drop every row.
- **The form builds its own controls.** A `.frm` without its `.frx` will not
  import, and the `.frx` cannot be produced outside the VBE. A self-building
  form is pasted in once and is diffable in git like any other code.
- **The REVO macros are not deleted.** `REVO_Ops`, `frmReleaseDetails`,
  `frmReworkRelease`, `Tube_Dash` etc. still reference REVO sheets that are not
  in this workbook. They do nothing until something calls them. Removing them is
  a separate cleanup once you are sure nothing else does.

## Files

| Path | What |
|---|---|
| `vba/modTableQuality.bas` | the flow (import this) |
| `vba/frmTableQuality.frm.txt` | the form code (paste this) |
| `tools/patch_workbook.py` | re-applies the sheet patch to a fresh copy of the workbook: headers, conditional formatting, button macro + caption, column widths. Zip-level; every other part is byte-identical |
| `vba_export/` | the workbook's VBA as it was in the uploaded file, one file per module, so future changes show up as a diff |
| `work/Table_Production.xlsm` | the patched workbook (git-ignored; regenerate with the tool) |

## Known gaps / next steps

1. **Holds data validation** on `Holds!B2:B50` still lists the REVO ops
   (`IC,10,20,…`). The macro writes ST01–ST05 regardless; change the list to
   `PreAssym,ST01,ST02,ST03,ST04,ST05` when convenient.
2. **No email on release yet.** The REVO flow mailed a summary; you did not ask
   for it here. Easy to add from `QC Log` once the recipients are known.
3. **Rework return path.** A reworked table is greyed with status REWORK. When
   it comes back, clear its status, untick the stations it must redo, and QC it
   again; the second event is logged as a separate row.
4. **Compile check.** This was built without Excel. After pasting, run
   **Debug > Compile VBAProject** once; anything it flags is a one-line fix.
