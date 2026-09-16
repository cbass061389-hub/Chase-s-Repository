# Building frmReleaseDetails by hand

Only needed if IT will not allow **Trust access to the VBA project object
model**. Otherwise `REVO_BuildReleaseForm` does all of this in one click.

About ten minutes.

## 1. Create the form

`Alt+F11` → `Insert → UserForm`. In Properties set:

| Property | Value |
|---|---|
| `(Name)` | `frmReleaseDetails` |
| `Caption` | `Release cart` |
| `Width` | `664` |
| `Height` | `500` |

## 2. Add the controls

Names must match exactly — the release macro and the form's code both address
them by name. Positions are cosmetic; get them roughly right.

### Header — labels straight on the form

| Name | Type | Caption | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `lblHdr1` | Label | `SKU` | 12 | 8 | 34 | 14 |
| `lblSKU` | Label | | 48 | 8 | 268 | 14 |
| `lblHdr2` | Label | `Cart` | 324 | 8 | 30 | 14 |
| `lblCart` | Label | | 356 | 8 | 74 | 14 |
| `lblHdr3` | Label | `WO` | 438 | 8 | 26 | 14 |
| `lblWO` | Label | | 466 | 8 | 90 | 14 |
| `lblHdr4` | Label | `Cart qty` | 12 | 26 | 50 | 14 |
| `lblCartQty` | Label | | 64 | 26 | 50 | 14 |
| `lblHdr5` | Label | `Already released` | 130 | 26 | 92 | 14 |
| `lblPriorRel` | Label | | 224 | 26 | 50 | 14 |
| `lblHdr6` | Label | `Remaining` | 290 | 26 | 60 | 14 |
| `lblRemaining` | Label | | 352 | 26 | 50 | 14 |

### Release quantity

| Name | Type | Caption | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `lblHdr7` | Label | `Release qty this pass` | 12 | 58 | 122 | 16 |
| `txtCartQtyRelease` | TextBox | | 138 | 56 | 56 | 20 |
| `lblHdr8` | Label | `To inventory` | 216 | 58 | 72 | 16 |
| `lblToInventory` | Label | `0` | 292 | 54 | 66 | 22 |
| `lblAfterThisPass` | Label | | 372 | 58 | 274 | 16 |
| `lblWarn` | Label | | 12 | 78 | 634 | 14 |

### Disposition quantities

| Name | Type | Caption | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `lblHdr9` | Label | `Rework` | 12 | 100 | 46 | 16 |
| `txtRework` | TextBox | | 60 | 98 | 46 | 20 |
| `lblHdr10` | Label | `B Grade` | 128 | 100 | 50 | 16 |
| `txtBGrade` | TextBox | | 180 | 98 | 46 | 20 |
| `lblHdr11` | Label | `Reject` | 248 | 100 | 40 | 16 |
| `txtReject` | TextBox | | 290 | 98 | 46 | 20 |

### Frame `fraRework` — caption `Rework detail`, at (8, 122), 644 × 92

Positions below are **relative to the frame**.

| Name | Type | Caption | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `lblRwD` | Label | `Defect` | 8 | 16 | 44 | 14 |
| `cboReworkDefect` | ComboBox | | 54 | 14 | 128 | 18 |
| `lblRwL` | Label | `Location` | 190 | 16 | 48 | 14 |
| `cboReworkLocation` | ComboBox | | 240 | 14 | 104 | 18 |
| `lblRwO` | Label | `At op` | 352 | 16 | 32 | 14 |
| `cboReworkOp` | ComboBox | | 386 | 14 | 68 | 18 |
| `lblRwR` | Label | `Root cause` | 8 | 40 | 60 | 14 |
| `cboReworkRoot` | ComboBox | | 70 | 38 | 128 | 18 |
| `lblRwA` | Label | `Action` | 206 | 40 | 40 | 14 |
| `cboReworkAction` | ComboBox | | 248 | 38 | 128 | 18 |
| `lblRwN` | Label | `Notes` | 8 | 64 | 40 | 14 |
| `txtReworkNotes` | TextBox | | 54 | 62 | 574 | 18 |

### Frame `fraBGrade` — caption `B grade detail`, at (8, 220), 644 × 70

| Name | Type | Caption | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `lblBgD` | Label | `Defect` | 8 | 16 | 44 | 14 |
| `cboBGradeDefect` | ComboBox | | 54 | 14 | 128 | 18 |
| `lblBgL` | Label | `Location` | 190 | 16 | 48 | 14 |
| `cboBGradeLocation` | ComboBox | | 240 | 14 | 104 | 18 |
| `lblBgR` | Label | `Root cause` | 352 | 16 | 60 | 14 |
| `cboBGradeRoot` | ComboBox | | 414 | 14 | 128 | 18 |
| `lblBgN` | Label | `Notes` | 8 | 40 | 40 | 14 |
| `txtBGradeNotes` | TextBox | | 54 | 38 | 574 | 18 |

### Frame `fraReject` — caption `Reject detail`, at (8, 296), 644 × 92

| Name | Type | Caption | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `lblRjD` | Label | `Defect` | 8 | 16 | 44 | 14 |
| `cboRejectDefect` | ComboBox | | 54 | 14 | 128 | 18 |
| `lblRjL` | Label | `Location` | 190 | 16 | 48 | 14 |
| `cboRejectLocation` | ComboBox | | 240 | 14 | 104 | 18 |
| `lblRjO` | Label | `At op` | 352 | 16 | 32 | 14 |
| `cboRejectOp` | ComboBox | | 386 | 14 | 68 | 18 |
| `lblRjR` | Label | `Root cause` | 8 | 40 | 60 | 14 |
| `cboRejectRoot` | ComboBox | | 70 | 38 | 128 | 18 |
| `lblRjA` | Label | `Action` | 206 | 40 | 40 | 14 |
| `cboRejectAction` | ComboBox | | 248 | 38 | 128 | 18 |
| `lblRjN` | Label | `Notes` | 8 | 64 | 40 | 14 |
| `txtRejectNotes` | TextBox | | 54 | 62 | 574 | 18 |

### Buttons — on the form, not in a frame

| Name | Type | Caption | Left | Top | Width | Height | Other |
|---|---|---|---|---|---|---|---|
| `btnSubmit` | CommandButton | `Submit release` | 430 | 400 | 108 | 28 | `Default` = True |
| `btnCancel` | CommandButton | `Cancel` | 548 | 400 | 100 | 28 | `Cancel` = True |

## 3. Paste the code

Right-click the form → `View Code`. Delete anything there, then paste the whole
of `revo/vba/frmReleaseDetails.code.vb`.

## 4. Check it

`Alt+F8` → `REVO_RunSelfTest`. The "Release form has its controls" row will
tell you if any name is wrong or missing.
