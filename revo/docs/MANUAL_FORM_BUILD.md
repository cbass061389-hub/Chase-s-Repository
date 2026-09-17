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
| `Width` | `684` |
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
| `lblHdr9` | Label | `Rework` | 12 | 94 | 46 | 16 |
| `txtRework` | TextBox | | 60 | 92 | 46 | 20 |
| `lblHdr10` | Label | `B Grade` | 128 | 94 | 50 | 16 |
| `txtBGrade` | TextBox | | 180 | 92 | 46 | 20 |
| `lblHdr11` | Label | `Reject` | 248 | 94 | 40 | 16 |
| `txtReject` | TextBox | | 290 | 92 | 46 | 20 |
| `lblAllocation` | Label | | 348 | 94 | 314 | 16 |

### Frame `fraDetail` — at (8, 118), 656 × 118

Caption: `Quality detail  -  one line per cause. Split a disposition across as
many causes as it needs.` Positions are **relative to the frame**.

| Name | Type | Caption | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `lblEnD` | Label | `Disposition` | 8 | 18 | 62 | 14 |
| `cboDisposition` | ComboBox | | 72 | 16 | 86 | 18 |
| `lblEnQ` | Label | `Qty` | 168 | 18 | 22 | 14 |
| `txtLineQty` | TextBox | | 192 | 16 | 44 | 20 |
| `lblEnDef` | Label | `Defect` | 250 | 18 | 40 | 14 |
| `cboDefect` | ComboBox | | 292 | 16 | 126 | 18 |
| `lblEnL` | Label | `Location` | 428 | 18 | 48 | 14 |
| `cboLocation` | ComboBox | | 478 | 16 | 104 | 18 |
| `lblEnO` | Label | `At op` | 8 | 44 | 34 | 14 |
| `cboOp` | ComboBox | | 72 | 42 | 86 | 18 |
| `lblEnR` | Label | `Root cause` | 168 | 44 | 62 | 14 |
| `cboRoot` | ComboBox | | 232 | 42 | 136 | 18 |
| `lblEnA` | Label | `Action` | 380 | 44 | 40 | 14 |
| `cboAction` | ComboBox | | 424 | 42 | 158 | 18 |
| `lblEnN` | Label | `Notes` | 8 | 70 | 34 | 14 |
| `txtLineNotes` | TextBox | | 72 | 68 | 510 | 18 |
| `btnAddLine` | CommandButton | `Add line` | 8 | 92 | 86 | 22 |
| `btnRemoveLine` | CommandButton | `Remove selected` | 100 | 92 | 110 | 22 |
| `lblEnHint` | Label | (see below) | 220 | 96 | 420 | 14 |

`lblEnHint` caption: `Example: reject 10 = 3 collet crush, 5 chip at butt,
2 breakthrough. Add each as its own line.`

### The lines list — on the form, not in a frame

| Name | Type | Caption | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `lblHdrList` | Label | `Lines recorded for this release` | 12 | 244 | 200 | 14 |
| `lstLines` | ListBox | | 8 | 260 | 656 | 150 |

On `lstLines` set `ColumnCount` = **7** and `ColumnWidths` =
`58 pt;30 pt;90 pt;62 pt;38 pt;96 pt;120 pt`.

### Buttons — on the form, not in a frame

| Name | Type | Caption | Left | Top | Width | Height | Other |
|---|---|---|---|---|---|---|---|
| `btnSubmit` | CommandButton | `Submit release` | 446 | 420 | 108 | 28 | `Default` = True |
| `btnCancel` | CommandButton | `Cancel` | 564 | 420 | 100 | 28 | `Cancel` = True |

## 3. Paste the code

Right-click the form → `View Code`. Delete anything there, then paste the whole
of `revo/vba/frmReleaseDetails.code.vb`.

## 4. Check it

`Alt+F8` → `REVO_RunSelfTest`. The "Release form has its controls" row will
tell you if any name is wrong or missing.
