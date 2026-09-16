# Install

Ten minutes. Do it on a copy first.

## 0. Save a copy

`Revo_Production.xlsm` → `Revo_Production_BACKUP_<date>.xlsm`. The install adds
sheets and a column and rebuilds a form. It does not delete history, but take
the copy anyway.

## 1. Turn on VBA project access (one checkbox, one time)

`File → Options → Trust Center → Trust Center Settings → Macro Settings`

☑ **Trust access to the VBA project object model**

This lets `REVO_FormBuilder` create the UserForm. A form's layout is a binary
stream inside the VBA project — it cannot be delivered as a text file you
paste in, so it gets built from code instead. That also means the form is
version-controlled and one command to rebuild.

If IT will not allow this, skip it and use
[`MANUAL_FORM_BUILD.md`](MANUAL_FORM_BUILD.md) — about ten minutes of
drag-and-drop. Everything else installs either way.

## 2. Put the source files beside the workbook

Copy the whole `revo/vba/` folder next to `Revo_Production.xlsm`, or leave the
repo checked out beside it. The form builder looks in both places and will
prompt with a file picker if it finds neither.

## 3. Import the modules

`Alt+F11` → `File → Import File…` and import each of these in order:

1. `REVO_Core.bas`
2. `REVO_Quality.bas`
3. `REVO_Release.bas`
4. `REVO_QualityDash.bas`
5. `REVO_CartVelocity.bas`
6. `REVO_Scorecard.bas`
7. `REVO_FormBuilder.bas`
8. `REVO_Install.bas`

Do **not** import `frmReleaseDetails.code.vb` — the builder reads it from disk.

Leave the existing `REVO_Ops`, `Update_Analysis` and the rest alone. Nothing
here collides with them: the new entry point is `REVO_ReleaseCarts`, not
`ReleaseToShippingAndReceiving`.

## 4. Run it

`Alt+F8` → **`REVO_Install`** → Run.

It will:

- create `Quality Log`, `Quality_Codes`, `Quality Dashboard`, `REVO Scorecard`,
  `Cart Velocity`, `REVO Audit`, `REVO Self Test`
- seed the defect and root-cause taxonomy from your own reject history
- code the 439 historical reject descriptions into `Quality Log`
- rebuild `frmReleaseDetails` with the quality capture fields
- re-point the REVO Floor button to `REVO_ReleaseCarts`
- re-point the dead Release Log "undo" button to `REVO_ResetCart`
- add a `Release Status` column to REVO Floor
- build the dashboards and run the self test

## 5. Read the self test

**Do this before trusting any number.** The `REVO Self Test` sheet checks every
assumption the code makes about your sheet layouts and says PASS / WARN / FAIL
with what to do about each one.

Expect on a first run:

| Check | Expected | Why |
|---|---|---|
| Release? column holds only checkboxes | **FAIL**, 2 text cells | Rows 13 and 20 — the original bug. Clear the contents and the checkbox comes back. |
| REVO Floor agrees with Release Log | **FAIL**, ~10+ mismatches | Column Y was written by both a formula and VBA. Step 6 settles it. |
| Rows that would skip the form silently | **FAIL**, ~10 | Rows 21–29, 39, 42 — released-to-date 100 of 100 with nothing in the log. |
| Floor Log snapshots | **WARN**, 2 snapshots | One interval. Velocity measures but will not apply. Expected. |

## 6. Reconcile the board

`Alt+F8` → **`REVO_ReconcileFloor`**

It lists every row where REVO Floor disagrees with Release Log, then asks
before writing. Release Log is what actually shipped, so it wins.

Until this is done the scorecard is working from numbers nobody can vouch for.

## 7. Clear the two stranded cells

On REVO Floor, select `W13` and `W20` and clear contents (`Delete`). The
checkbox format is a *cell format*, so the checkbox returns the moment the text
is gone. Cart 34 gets its remaining 48 units back.

## 8. Try a release

Tick a `Release?` checkbox and press the button. You should get the new form
with rework / B grade / reject quantities, each with defect type, location,
detected-at op, root cause and action.

Enter a rework or reject quantity and the matching panel wakes up. Submit
without a root cause and it will stop you — "Not Determined" is an acceptable
answer, blank is not.

## Daily

Point your plan-update button at **`REVO_DailyUpdate`**, or run it by hand. It
refreshes cart velocity, the quality dashboard and the scorecard in that order.

Keep running `LOG FLOOR` daily. Cart velocity needs 8 clean intervals before it
will move a capacity number, and every snapshot is one more.

## Changing the taxonomy

Edit the `Quality_Codes` sheet. Column A defect types, B locations, C root
causes, D ops, E corrective actions. The form reads them every time it opens —
no rebuild needed.

Columns G:I are the classifier map used to code historical free text and to
pre-select on the form. Longest keyword wins.

## Before handing modules to anyone

```
python3 revo/tools/check_vba.py revo/vba
```

Catches, without needing Excel: statements over VBA's 24-continuation limit,
identifiers colliding with reserved words (`cDate`, `dO`, `eNum` all do),
unbalanced blocks, error handlers that read `Err` too late or fall off without
`Resume`, and early-bound UserForm references. Every one of those shipped at
least once before this existed. Exit code is non-zero on failure, so it can gate
a commit.

## Troubleshooting

**"Programmatic access to the VBA project is not trusted"** — step 1.

**Form builder cannot find the code file** — it searches, in order: any path you
pass it, the workbook's own folder, `revo\vba` beneath it, Downloads, Desktop,
the current directory, and a copy cached inside the workbook. Failing all of
that it opens a file picker — point it at `frmReleaseDetails.code.vb`.

**If the workbook lives on OneDrive or SharePoint**, `ThisWorkbook.Path` is an
`https://` URL, not a folder, so the code file cannot be found beside it. The
install says so and searches Downloads and Desktop instead. Keep the `revo\vba`
folder in Downloads and it resolves on its own.

After the first successful build the form code is cached on a veryHidden sheet
inside the workbook, so later rebuilds need no file at all — including on a
machine that never had the repo.

**Button still runs the old macro** — run `REVO_RewireButtons`, or right-click
the button → Assign Macro → `REVO_ReleaseCarts`.

**Something behaves oddly after a sheet layout change** — run
`REVO_RunSelfTest`. It is the fastest way to see what moved.
