# Why the release UserForm stopped appearing

Diagnosis against `Revo_Production.xlsm` as supplied. Every figure below was
read out of the file.

## The immediate cause

Column W on REVO Floor — `Release?` — held **220 `FALSE`, one `"PARTIAL"`, one
`"RELEASED"`, and zero `TRUE`**.

`ReleaseToShippingAndReceiving` only opens the form for rows where
`Boolish(W)` is true. Nothing qualified, so it ran the whole loop, jumped to
`SendMailAndExit`, and quit — no form, no message, no error.

## Four defects behind that

### 1. The macro destroyed its own trigger

W is an Excel 365 **cell-checkbox** column. Confirmed in the file: the workbook
carries a `Checkbox` feature-property bag, and the W column styles (`138`,
`785`) carry the `xfComplement` that renders the checkbox. A checkbox cell only
renders as a checkbox while its value is boolean.

The macro wrote text into it:

```vba
wsSource.Cells(r, COL_RELEASE).Value = "RELEASED"
wsSource.Cells(r, COL_RELEASE).Value = "PARTIAL"
wsSource.Cells(r, COL_RELEASE).Value = "NO DATA"
```

Once that happens the checkbox is gone and the operator cannot click it again.
**Row 13** proves it: WO2889, cart 34, `W13 = "PARTIAL"`, `Z13 = 48` units still
remaining. Permanently stranded.

### 2. Ten rows skip the form silently

```vba
remaining = qty - priorRel
If remaining <= 0 Then
    wsSource.Cells(r, COL_RELEASE).Value = "RELEASED"
    GoTo NextRow                    ' form never shows
End If
```

Rows **21, 22, 23, 25, 26, 27, 28, 29, 39, 42** all carry Released-To-Date =
Qty = 100 with `W = FALSE`. Tick any of them and the checkbox flips to the word
`RELEASED` with no form — exactly the reported symptom.

### 3. Formulas and VBA fighting over the same two columns

`Y` and `Z` were built as formulas:

```
=IF(COUNTA($B11:$D11)=0,"",IF($W11="RELEASED",$E11,0))
```

Rows 11–19 still carry them. From row 20 down they are hardcoded numbers,
because the VBA writes `.Value` into Y and Z. The macro wiped the formulas,
then a fill-down carried `100 / 0` into rows that were never released. Two
owners, so neither was right.

### 4. The form was shown with screen updating off

`Application.ScreenUpdating = False` is set before the loop; `frm.Show` is
inside it. A modal UserForm shown that way can paint blank, or open behind the
Excel window.

## Also found

- **`FIRST_DATA_ROW = 2`** in the sub. The header is row 9 and data starts at
  row 11. Harmless only because W2:W8 happen to be empty.
- **`UndoReleaseForCart` does not exist.** The Release Log button points at a
  macro that is not in the project — the escape hatch for a stranded row was
  already gone. (`Revo_ShowSelectedSkuDashboard` and
  `BoxPlan_RemoveFromUtilization_Archive` are missing too.)
- **`BuildBOMHierarchyView` is Public in both `REVO_Ops` and
  `Update_Analysis`.** Nothing calls it, so it still compiles, but it appears
  twice in the macro list and will throw *Ambiguous name detected* the first
  time anything calls it unqualified.
- **A dead guard on the form.** `Prime` set `mLoading = True`, then the first
  `Me.Controls(...)` reference triggered `UserForm_Initialize`, which ended with
  `mLoading = False`. Every `_Change` handler then fired during setup.

## What was ruled out

- `txtRelease` is a **Label**, not a TextBox (ClsidCacheIndex `0x15`), so
  `Me.txtRelease.Caption` is valid.
- All nineteen controls the form's code addresses exist.
- No duplicate declarations, and no ambiguous name is actually called, so the
  project does compile.
- REVO Floor is not protected.
- The REVO Floor button is correctly wired to
  `[0]!ReleaseToShippingAndReceiving`.
