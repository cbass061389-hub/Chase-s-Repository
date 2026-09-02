# Calculation engine map

Generated 2026-09-02T18:05:45+00:00. Sheets with at least 1,000 formulas, read directly from the workbook XML — no Excel, no macro execution.

## Executive summary

- **15 calculation sheets** hold 266,269 formulas between them.
- **13,849 formula cells carry a destroyed reference** (`#REF!` and friends) across 5 column(s). These are not results that failed to calculate — the references themselves are gone, so those rows are producing nothing and have been since whatever delete broke them.
- **21 column(s) contain a mix of formulas** where one pattern dominates. In a formula column that is normally a hand-edited cell, and it is invisible in the grid.
- 1 sheet(s) use volatile or dependency-breaking functions (INDIRECT, OFFSET, TODAY, NOW...).

Formulas are collapsed by replacing row numbers with `{r}`, so a column of 5,000 formulas shows as the one pattern it actually is. A column with a single pattern is a computed field and the pattern is its definition — that is what makes this logic portable even though the VBA around it is password-protected.

### Destroyed references — fix these first

| Sheet | Column | Broken cells | Total cells | Share | Broken pattern |
|---|---|---|---|---|---|
| `__CleanBO` | A — Item | 2,909 | 9,362 | 31% | 'Open Order Report'!#REF! |
| `__CleanBO` | B — Channel (Regional ID) | 2,909 | 9,362 | 31% | 'Open Order Report'!#REF! |
| `__CleanBO` | C — Live BO Qty | 2,909 | 9,362 | 31% | IF(AND('Open Order Report'!#REF!<>"",N('Open Order Report'!# |
| `__CleanBO` | D — Live BO $ | 2,909 | 9,362 | 31% | C{r}*N('Open Order Report'!#REF!) |
| `Position Engine` | P — Monthly Run Rate | 2,213 | 2,213 | 100% | (SUMIFS(#REF!,#REF!,$A{r},#REF!,$AA{r},#REF!,">="&EDATE('PE  |

A `#REF!` inside the formula text means the range it pointed at was deleted. Every one of these cells is silently contributing nothing to whatever depends on it, and no total anywhere shows a gap.

### Columns that need a look

| Sheet | Column | Cells | Patterns | Consistency | Exception |
|---|---|---|---|---|---|
| `__CleanBO` | A — Item | 9,362 | 2 | 68.9% | 2,909 cell(s) differ |
| `__CleanBO` | B — Channel (Regional ID) | 9,362 | 2 | 68.9% | 2,909 cell(s) differ |
| `__CleanBO` | C — Live BO Qty | 9,362 | 2 | 68.9% | 2,909 cell(s) differ |
| `__CleanBO` | D — Live BO $ | 9,362 | 2 | 68.9% | 2,909 cell(s) differ |
| `Launch Status 2026` | C — (no header) | 88 | 4 | 96.6% | 1 cell(s) differ |
| `Launch Status 2026` | D — (no header) | 88 | 4 | 96.6% | 1 cell(s) differ |
| `Launch Status 2026` | E — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | F — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | G — (no header) | 88 | 3 | 96.6% | 2 cell(s) differ |
| `Launch Status 2026` | H — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | I — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | J — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | K — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | L — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | M — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | N — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | Q — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | R — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | S — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | T — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |
| `Launch Status 2026` | U — (no header) | 88 | 2 | 96.6% | 3 cell(s) differ |

Each of these is either a deliberate two-block structure or a formula somebody overwrote. Both are worth knowing before the logic is ported.

## Sheets, by size

| Sheet | State | Formulas | Formula cols | Lookups | Volatile | Whole-col refs | Broken refs | Reads from |
|---|---|---|---|---|---|---|---|---|
| `Position Engine` | veryHidden | 59,751 | 27 | VLOOKUPx19,917 | — | — | 2,213 | Current Inventory, Forecast Qty, Items on PO, Open Order Report, PE Config |
| `Allocation Plan` | veryHidden | 47,534 | 10 | — | — | — | — | __Alloc Engine |
| `__CleanBO` | hidden | 37,451 | 5 | — | — | — | 11,636 | Alloc_Hold_List, Open Order Report, REF, SO_Closure_List, Sales_Hygiene_Flags |
| `BOMMaster` | hidden | 35,367 | 3 | — | — | — | — | (self-contained) |
| `__Alloc Engine` | veryHidden | 24,528 | 5 | — | — | — | — | (self-contained) |
| `__AllocMix` | veryHidden | 21,386 | 19 | — | — | — | — | Forecast Qty, PE Config, Position Engine, __CleanBO |
| `Inventory Review` |  | 15,984 | 13 | XLOOKUPx2,642 | — | — | — | Inventory Health Review |
| `JAX Container Validation` |  | 7,112 | 28 | INDEXx762, MATCHx762 | — | — | — | Container Shipment, Position Engine |
| `Alloc_Hold_List` | veryHidden | 4,602 | 2 | — | — | — | — | Current Inventory |
| `ARC Build Live` |  | 3,603 | 7 | INDEXx1,440, MATCHx1,440, XLOOKUPx720 | — | — | — | ARC Buildable, ARC Parts Live |
| `New Product Readiness` |  | 3,069 | 41 | INDEXx255, MATCHx255 | TODAYx255 | 935 | — | Alloc Overrides, Cue Completion Schedule, Defect Fallout to Date, HIEInv, HIEProdFlat |
| `Launch Status 2026` |  | 1,777 | 21 | — | — | 264 | — | New Product Readiness, PO |
| `SO_Closure_List` | hidden | 1,690 | 3 | — | — | — | — | Current Inventory, Items on PO |
| `ARC Parts Live` |  | 1,362 | 9 | — | — | — | — | ARC Buildable |
| `Sales_Hygiene_Flags` | hidden | 1,053 | 3 | — | — | — | — | PE Config |

## Recovered definitions

Per sheet, the formula columns and what each one computes. This is the specification for porting the logic to the canonical layer.

### `Position Engine` [veryHidden] — 59,751 formulas

Reads from: `Current Inventory`, `Forecast Qty`, `Items on PO`, `Open Order Report`, `PE Config`, `Purchase Price`, `REF`, `__AllocMix`

**C — Owner** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(VLOOKUP($B{r},'PE Config'!$A{r}:$B{r},2,FALSE),"")
```

**E — Product Category** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(VLOOKUP($A{r},'Purchase Price'!$A{r}:$C{r},3,FALSE),"")
```

**F — Avg Cost** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(VLOOKUP($A{r},'Purchase Price'!$A{r}:$D{r},4,FALSE),"")
```

**G — On Hand** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Current Inventory'!$C{r}:$C{r},'Current Inventory'!$A{r}:$A{r},$A{r},'Current Inventory'!$D{r}:$D{r},$B{r})
```

**H — Live BO** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Open Order Report'!$M{r}:$M{r},'Open Order Report'!$A{r}:$A{r},$A{r},'Open Order Report'!$C{r}:$C{r},$AA{r},'Open Order Report'!$B{r}:$B{r},">="&DATE(2020,1,1))+IF($AB{r}="",0,SUMIFS('Open Order Report'!$M{r}:$M{r},'Open Order Report'!$A{r}:$A{r},$A{r},'Open Order Report'!$C{r}:$C{r},$AB{r},'Open Order Report'!$B{r}:$B{r},">="&DATE(2020,1,1)))
```

**I — Net Available** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
G{r}-H{r}
```

**J — Inbound PO Qty** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Items on PO'!$N{r}:$N{r},'Items on PO'!$K{r}:$K{r},$A{r},'Items on PO'!$B{r}:$B{r},">40",'Items on PO'!$F{r}:$F{r},"<>TO*",'Items on PO'!$G{r}:$G{r},"<>*Clawson*",'Items on PO'!$G{r}:$G{r},"<>*Predator Group*")
```

**K — Inbound PO Value** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Items on PO'!$P{r}:$P{r},'Items on PO'!$K{r}:$K{r},$A{r},'Items on PO'!$B{r}:$B{r},">40",'Items on PO'!$F{r}:$F{r},"<>TO*",'Items on PO'!$G{r}:$G{r},"<>*Clawson*",'Items on PO'!$G{r}:$G{r},"<>*Predator Group*")
```

**L — Next PO Due** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
_xlfn.MINIFS('Items on PO'!$B{r}:$B{r},'Items on PO'!$K{r}:$K{r},$A{r},'Items on PO'!$B{r}:$B{r},">"&'PE Config'!$B{r},'Items on PO'!$F{r}:$F{r},"<>TO*",'Items on PO'!$G{r}:$G{r},"<>*Clawson*",'Items on PO'!$G{r}:$G{r},"<>*Predator Group*")
```

**M — Fwd 3Mo Fcst** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Forecast Qty'!$B{r}:$B{r},'Forecast Qty'!$A{r}:$A{r},$A{r},'Forecast Qty'!$D{r}:$D{r},">="&'PE Config'!$B{r},'Forecast Qty'!$D{r}:$D{r},"<="&'PE Config'!$B{r},'Forecast Qty'!$C{r}:$C{r},$AA{r})+IF($AB{r}="",0,SUMIFS('Forecast Qty'!$B{r}:$B{r},'Forecast Qty'!$A{r}:$A{r},$A{r},'Forecast Qty'!$D{r}:$D{r},">="&'PE Config'!$B{r},'Forecast Qty'!$D{r}:$D{r},"<="&'PE Config'!$B{r},'Forecast Qty'!$C{r}:$C{r},$AB{r}))
```

**N — Fwd 12Mo Fcst** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Forecast Qty'!$B{r}:$B{r},'Forecast Qty'!$A{r}:$A{r},$A{r},'Forecast Qty'!$D{r}:$D{r},">="&'PE Config'!$B{r},'Forecast Qty'!$D{r}:$D{r},"<="&'PE Config'!$B{r},'Forecast Qty'!$C{r}:$C{r},$AA{r})+IF($AB{r}="",0,SUMIFS('Forecast Qty'!$B{r}:$B{r},'Forecast Qty'!$A{r}:$A{r},$A{r},'Forecast Qty'!$D{r}:$D{r},">="&'PE Config'!$B{r},'Forecast Qty'!$D{r}:$D{r},"<="&'PE Config'!$B{r},'Forecast Qty'!$C{r}:$C{r},$AB{r}))
```

**O — Zero Fcst Mo** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
MAX(0,12-COUNTIFS('Forecast Qty'!$A{r}:$A{r},$A{r},'Forecast Qty'!$C{r}:$C{r},$AA{r},'Forecast Qty'!$D{r}:$D{r},">="&'PE Config'!$B{r},'Forecast Qty'!$D{r}:$D{r},"<="&'PE Config'!$B{r},'Forecast Qty'!$B{r}:$B{r},">0"))
```

**P — Monthly Run Rate** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
(SUMIFS(#REF!,#REF!,$A{r},#REF!,$AA{r},#REF!,">="&EDATE('PE Config'!$B{r},-'PE Config'!$B{r}))+IF($AB{r}="",0,SUMIFS(#REF!,#REF!,$A{r},#REF!,$AB{r},#REF!,">="&EDATE('PE Config'!$B{r},-'PE Config'!$B{r}))))/'PE Config'!$B{r}
```

**Q — Avg Lead Time (d)** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
'PE Config'!$B{r}
```

**R — Lead Time (mo)** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
Q{r}/30
```

**S — Months of Supply** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
IF(N{r}>0,G{r}/(N{r}/12),IF(P{r}>0,G{r}/P{r},IF(G{r}>0,999,0)))
```

**T — Coverage vs LT** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
S{r}-R{r}
```

**U — AllocMix B2C%** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(VLOOKUP($A{r},__AllocMix!$A{r}:$E{r},2,FALSE),"")
```

**V — AllocMix Amer%** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(VLOOKUP($A{r},__AllocMix!$A{r}:$E{r},3,FALSE),"")
```

**W — AllocMix EMEA%** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(VLOOKUP($A{r},__AllocMix!$A{r}:$E{r},4,FALSE),"")
```

**X — AllocMix AP%** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(VLOOKUP($A{r},__AllocMix!$A{r}:$E{r},5,FALSE),"")
```

**Y — Position Status** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
IF(AND($N{r}=0,$P{r}=0),"NO DEMAND",IF(AND($N{r}=0,$P{r}>0),"NO FORECAST",IF($S{r}>Thr_Excess,"EXCESS",IF($T{r}<Thr_Critical,"CRITICAL",IF($T{r}<Thr_AtRisk,"AT RISK",IF($T{r}<Thr_Watch,"WATCH","HEALTHY"))))))
```

**AA — (no header)** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(VLOOKUP($B{r},'PE Config'!$A{r}:$D{r},3,FALSE),"")
```

**AB — (no header)** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(VLOOKUP($B{r},'PE Config'!$A{r}:$D{r},4,FALSE),"")
```

**AC — BO$** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
H{r}*IF(ISNUMBER(F{r}),F{r},0)
```

**AD — Key** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
A{r}&"|"&B{r}
```

**AE — BO$ RankKey** · 2,213 cells · rows 2-2214 · 1 pattern(s) · 100.0% consistent

```excel
IF(AC{r}>0,AC{r}+ROW()/1000000000,"")
```

### `Allocation Plan` [veryHidden] — 47,534 formulas

Reads from: `__Alloc Engine`

**A — Allocation Plan — Margin-Max by Delivery (Full Pool)** · 4,752 cells · rows 18-4769 · 1 pattern(s) · 100.0% consistent

```excel
'__Alloc Engine'!A{r}
```

**B — (no header)** · 4,762 cells · rows 4-4769 · 11 pattern(s) · 99.8% consistent

```excel
'__Alloc Engine'!B{r}
```

10 other pattern(s) in this column, largest with 1 cell(s):

```excel
SUM('__Alloc Engine'!$W{r}:$W{r})   /* x1 */
SUM('__Alloc Engine'!$Y{r}:$Y{r})   /* x1 */
```

**C — (no header)** · 4,752 cells · rows 18-4769 · 1 pattern(s) · 100.0% consistent

```excel
'__Alloc Engine'!C{r}
```

**D — (no header)** · 4,752 cells · rows 18-4769 · 1 pattern(s) · 100.0% consistent

```excel
'__Alloc Engine'!D{r}
```

**E — (no header)** · 4,753 cells · rows 18-4770 · 2 pattern(s) · 99.9% consistent

```excel
'__Alloc Engine'!P{r}
```

1 other pattern(s) in this column, largest with 1 cell(s):

```excel
SUM(E{r}:E{r})   /* x1 */
```

**F — (no header)** · 4,753 cells · rows 18-4770 · 2 pattern(s) · 99.9% consistent

```excel
'__Alloc Engine'!S{r}
```

1 other pattern(s) in this column, largest with 1 cell(s):

```excel
SUM(F{r}:F{r})   /* x1 */
```

**G — (no header)** · 4,752 cells · rows 18-4769 · 1 pattern(s) · 100.0% consistent

```excel
'__Alloc Engine'!U{r}
```

**H — (no header)** · 4,752 cells · rows 18-4769 · 1 pattern(s) · 100.0% consistent

```excel
'__Alloc Engine'!V{r}
```

**I — (no header)** · 4,753 cells · rows 18-4770 · 2 pattern(s) · 99.9% consistent

```excel
'__Alloc Engine'!W{r}
```

1 other pattern(s) in this column, largest with 1 cell(s):

```excel
SUM(I{r}:I{r})   /* x1 */
```

**J — (no header)** · 4,753 cells · rows 18-4770 · 2 pattern(s) · 99.9% consistent

```excel
'__Alloc Engine'!Z{r}
```

1 other pattern(s) in this column, largest with 1 cell(s):

```excel
SUM(J{r}:J{r})   /* x1 */
```

### `__CleanBO` [hidden] — 37,451 formulas

Reads from: `Alloc_Hold_List`, `Open Order Report`, `REF`, `SO_Closure_List`, `Sales_Hygiene_Flags`

**A — Item** · 9,362 cells · rows 2-9363 · 2 pattern(s) · 68.9% consistent

```excel
'Open Order Report'!A{r}
```

1 other pattern(s) in this column, largest with 2,909 cell(s):

```excel
'Open Order Report'!#REF!   /* x2,909 */
```

**B — Channel (Regional ID)** · 9,362 cells · rows 2-9363 · 2 pattern(s) · 68.9% consistent

```excel
'Open Order Report'!C{r}
```

1 other pattern(s) in this column, largest with 2,909 cell(s):

```excel
'Open Order Report'!#REF!   /* x2,909 */
```

**C — Live BO Qty** · 9,362 cells · rows 2-9363 · 2 pattern(s) · 68.9% consistent

```excel
IF(AND('Open Order Report'!$A{r}<>"",N('Open Order Report'!$M{r})>0,ISERROR(SEARCH("Clawson",'Open Order Report'!$D{r})),ISERROR(SEARCH("Carrinho",'Open Order Report'!$D{r})),ISERROR(SEARCH("Beckmann",'Open Order Report'!$D{r})),ISERROR(SEARCH("Predator Group",'Open Order Report'!$D{r})),IF('Open Order Report'!$B{r}="",TRUE,'Open Order Report'!$B{r}>=DATE(2020,1,1)),COUNTIFS(SO_Closure_List!$A{r}:$A{r},$A{r},SO_Closure_List!$C{r}:$C{r},'Open Order Report'!$L{r},SO_Closure_List!$F{r}:$F{r},IF('Open Order Report'!$B{r}="","=",'Open Order Report'!$B{r}))=0,COUNTIFS(Alloc_Hold_List!$A{r}:$A{r},$A{r},Alloc_Hold_List!$C{r}:$C{r},'Open Order Report'!$L{r},Alloc_Hold_List!$K{r}:$K{r},IF('Open Order Report'!$B{r}="","=",'Open Order Report'!$B{r}))=0,COUNTIFS(Sales_Hygiene_Flags!$B{r}:$B{r},$A{r},Sales_Hygiene_Flags!$D{r}:$D{r},'Open Order Report'!$L{r},Sales_Hygiene_Flags!$G{r}:$G{r},IF('Open Order Report'!$B{r}="","=",'Open Order Report'!$B{r}))=0),'Open Order Report'!$M{r},0)
```

1 other pattern(s) in this column, largest with 2,909 cell(s):

```excel
IF(AND('Open Order Report'!#REF!<>"",N('Open Order Report'!#REF!)>0,ISERROR(SEARCH("Clawson",'Open Order Report'!#REF!)),ISERROR(SEARCH("Carrinho",'Open Order Report'!#REF!)),ISERROR(SEARCH("Beckmann",'Open Order Report'!#REF!)),ISERROR(SEARCH("Predator Group",'Open Order Report'!#REF!)),IF('Open Order Report'!#REF!="",TRUE,'Open Order Report'!#REF!>=DATE(2020,1,1)),COUNTIFS(SO_Closure_List!$A{r}:$A{r},$A{r},SO_Closure_List!$C{r}:$C{r},'Open Order Report'!#REF!,SO_Closure_List!$F{r}:$F{r},IF('Open Order Report'!#REF!="","=",'Open Order Report'!#REF!))=0,COUNTIFS(Alloc_Hold_List!$A{r}:$A{r},$A{r},Alloc_Hold_List!$C{r}:$C{r},'Open Order Report'!#REF!,Alloc_Hold_List!$K{r}:$K{r},IF('Open Order Report'!#REF!="","=",'Open Order Report'!#REF!))=0,COUNTIFS(Sales_Hygiene_Flags!$B{r}:$B{r},$A{r},Sales_Hygiene_Flags!$D{r}:$D{r},'Open Order Report'!#REF!,Sales_Hygiene_Flags!$G{r}:$G{r},IF('Open Order Report'!#REF!="","=",'Open Order Report'!#REF!))=0),'Open Order Report'!#REF!,0)   /* x2,909 */
```

**D — Live BO $** · 9,362 cells · rows 2-9363 · 2 pattern(s) · 68.9% consistent

```excel
C{r}*N('Open Order Report'!$K{r})
```

1 other pattern(s) in this column, largest with 2,909 cell(s):

```excel
C{r}*N('Open Order Report'!#REF!)   /* x2,909 */
```

### `BOMMaster` [hidden] — 35,367 formulas

**K — Parent SKU** · 11,789 cells · rows 2-11790 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(_xlfn.TEXTAFTER($B{r}," : ",-1),$B{r})
```

**L — Member SKU** · 11,789 cells · rows 2-11790 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(_xlfn.TEXTAFTER($E{r}," : ",-1),$E{r})
```

**M — Is ARC Table** · 11,789 cells · rows 2-11790 · 1 pattern(s) · 100.0% consistent

```excel
IF(AND(LEFT($K{r},3)="TP ",ISNUMBER(SEARCH("ARC",$K{r}))),1,0)
```

### `__Alloc Engine` [veryHidden] — 24,528 formulas

**AA — Base Alloc (pre-sweep)** · 4,752 cells · rows 2-4753 · 1 pattern(s) · 100.0% consistent

```excel
FLOOR(MIN(Q{r},MAX(0,L{r}-R{r})),1)
```

**AB — Unmet_c** · 4,752 cells · rows 2-4753 · 1 pattern(s) · 100.0% consistent

```excel
MAX(0,P{r}-AA{r})
```

**AC — PriorUnmet** · 5,136 cells · rows 2-5137 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS($AB{r}:$AB{r},$A{r}:$A{r},$A{r},$C{r}:$C{r},"<"&$C{r})
```

**AD — SKU Leftover** · 5,136 cells · rows 2-5137 · 1 pattern(s) · 100.0% consistent

```excel
MAX(0,L{r}-SUMIFS($AA{r}:$AA{r},$A{r}:$A{r},$A{r}))
```

**AE — Sweep_c** · 4,752 cells · rows 2-4753 · 1 pattern(s) · 100.0% consistent

```excel
MIN(AB{r},MAX(0,AD{r}-AC{r}))
```

### `__AllocMix` [veryHidden] — 21,386 formulas

Reads from: `Forecast Qty`, `PE Config`, `Position Engine`, `__CleanBO`

**B — As of** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
IF($R{r}=0,0,N{r}/$R{r})
```

**C — (no header)** · 1,189 cells · rows 1-1191 · 2 pattern(s) · 99.9% consistent

```excel
IF($R{r}=0,0,O{r}/$R{r})
```

1 other pattern(s) in this column, largest with 1 cell(s):

```excel
'PE Config'!$B{r}   /* x1 */
```

**D — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
IF($R{r}=0,0,P{r}/$R{r})
```

**E — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
IF($R{r}=0,0,Q{r}/$R{r})
```

**F — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(__CleanBO!$C{r}:$C{r},__CleanBO!$A{r}:$A{r},$A{r},__CleanBO!$B{r}:$B{r},N{r})
```

**G — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(__CleanBO!$C{r}:$C{r},__CleanBO!$A{r}:$A{r},$A{r},__CleanBO!$B{r}:$B{r},O{r})
```

**H — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(__CleanBO!$C{r}:$C{r},__CleanBO!$A{r}:$A{r},$A{r},__CleanBO!$B{r}:$B{r},P{r})
```

**I — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(__CleanBO!$C{r}:$C{r},__CleanBO!$A{r}:$A{r},$A{r},__CleanBO!$B{r}:$B{r},Q{r})
```

**J — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Forecast Qty'!$B{r}:$B{r},'Forecast Qty'!$A{r}:$A{r},$A{r},'Forecast Qty'!$C{r}:$C{r},N{r},'Forecast Qty'!$D{r}:$D{r},">="&'PE Config'!$B{r},'Forecast Qty'!$D{r}:$D{r},"<="&'PE Config'!$B{r})
```

**K — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Forecast Qty'!$B{r}:$B{r},'Forecast Qty'!$A{r}:$A{r},$A{r},'Forecast Qty'!$C{r}:$C{r},O{r},'Forecast Qty'!$D{r}:$D{r},">="&'PE Config'!$B{r},'Forecast Qty'!$D{r}:$D{r},"<="&'PE Config'!$B{r})
```

**L — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Forecast Qty'!$B{r}:$B{r},'Forecast Qty'!$A{r}:$A{r},$A{r},'Forecast Qty'!$C{r}:$C{r},P{r},'Forecast Qty'!$D{r}:$D{r},">="&'PE Config'!$B{r},'Forecast Qty'!$D{r}:$D{r},"<="&'PE Config'!$B{r})
```

**M — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Forecast Qty'!$B{r}:$B{r},'Forecast Qty'!$A{r}:$A{r},$A{r},'Forecast Qty'!$C{r}:$C{r},Q{r},'Forecast Qty'!$D{r}:$D{r},">="&'PE Config'!$B{r},'Forecast Qty'!$D{r}:$D{r},"<="&'PE Config'!$B{r})
```

**N — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
F{r}+J{r}
```

**O — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
G{r}+K{r}
```

**P — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
H{r}+L{r}
```

**Q — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
I{r}+M{r}
```

**R — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
SUM(N{r}:Q{r})
```

**S — (no header)** · 1,188 cells · rows 4-1191 · 1 pattern(s) · 100.0% consistent

```excel
IF($R{r}=0,"",SUM(B{r}:E{r}))
```

### `Inventory Review` — 15,984 formulas

Reads from: `Inventory Health Review`

**J — (no header)** · 2,642 cells · rows 3-2669 · 1 pattern(s) · 100.0% consistent

```excel
_xlfn.XLOOKUP(B{r},'Inventory Health Review'!$P{r}:$P{r},'Inventory Health Review'!$Q{r}:$Q{r},"Other")
```

**K — (no header)** · 2,667 cells · rows 3-2669 · 1 pattern(s) · 100.0% consistent

```excel
_xlfn.LET(_xlpm.tot,SUMIFS(Current_Inventory[InvQty],Current_Inventory[SKU],$A{r}),_xlpm.bb,SUMIFS(Current_Inventory[InvQty],Current_Inventory[SKU],$A{r},Current_Inventory[InvLoc],"Billard Beckmann")+SUMIFS(Current_Inventory[InvQty],Current_Inventory[SKU],$A{r},Current_Inventory[InvLoc],"Showroom - BB"),IF(_xlpm.tot=0,1,MEDIAN(0,(_xlpm.tot-_xlpm.bb)/_xlpm.tot,1)))
```

**L — (no header)** · 2,667 cells · rows 3-2669 · 1 pattern(s) · 100.0% consistent

```excel
ROUND($E{r}*$K{r},0)
```

**M — (no header)** · 2,667 cells · rows 3-2669 · 1 pattern(s) · 100.0% consistent

```excel
$H{r}*$K{r}
```

**N — (no header)** · 2,667 cells · rows 3-2669 · 1 pattern(s) · 100.0% consistent

```excel
$G{r}*$K{r}
```

**O — (no header)** · 2,667 cells · rows 3-2669 · 1 pattern(s) · 100.0% consistent

```excel
IF(AND(LEFT(A{r},1)<>"P",LEFT(A{r},8)<>"A PRE JP",K{r}>0),1,0)
```

### `JAX Container Validation` — 7,112 formulas

Reads from: `Container Shipment`, `Position Engine`

**D — BO Qty - Americas** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(Customer_Open_Orders[TtlBO],Customer_Open_Orders[Item],$A{r},Customer_Open_Orders[Regional ID],"Americas")
```

**E — BO Qty - All Regions** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
SUM(SUMIFS(Customer_Open_Orders[TtlBO],Customer_Open_Orders[Item],$A{r}))
```

**F — JAX Inv** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(SUMIFS(tblInventoryReview[JAX Inv],tblInventoryReview[SKU],$A{r}),0)
```

**G — Americas On-Hand** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(Current_Inventory[InvQty],Current_Inventory[SKU],$A{r},Current_Inventory[Regional ID],"Americas")
```

**H — Qty on PO** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(SUMIFS(tblInventoryReview[Qty on PO],tblInventoryReview[SKU],$A{r}),0)
```

**I — Fcst Aug-26 (Amer)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(Forecast_Qty[Forecast Qty],Forecast_Qty[Item],$A{r},Forecast_Qty[Regional ID],"Americas",Forecast_Qty[YearMonth],202608)
```

**J — Fcst Sep-26 (Amer)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(Forecast_Qty[Forecast Qty],Forecast_Qty[Item],$A{r},Forecast_Qty[Regional ID],"Americas",Forecast_Qty[YearMonth],202609)
```

**K — Fcst Oct-26 (Amer)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(Forecast_Qty[Forecast Qty],Forecast_Qty[Item],$A{r},Forecast_Qty[Regional ID],"Americas",Forecast_Qty[YearMonth],202610)
```

**L — Fcst 3-Mo Total** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
SUM(I{r}:K{r})
```

**M — Forecast 12m** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(SUMIFS(tblInventoryReview[Forecast 12m],tblInventoryReview[SKU],$A{r}),0)
```

**N — Inv Months** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(SUMIFS(tblInventoryReview[Inv_Months],tblInventoryReview[SKU],$A{r}),"")
```

**O — In-Transit: Sunray 6/24 (ETA 7/27)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Container Shipment'!$B{r}:$B{r},'Container Shipment'!$A{r}:$A{r},$A{r})
```

**P — In-Transit: YueQin 6/24 (ETA 7/27)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Container Shipment'!$B{r}:$B{r},'Container Shipment'!$A{r}:$A{r},$A{r})
```

**Q — Container ID / Group** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IF(O{r}>0,"Sunray — Predator Container (shipped 6/24)",IF(P{r}>0,"YueQin — Predator Container (shipped 6/24)",""))
```

**R — ETA** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IF(O{r}+P{r}>0,DATE(2026,7,27),"")
```

**S — Total In-Transit (Container Shipment tab)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
O{r}+P{r}
```

**T — Open PO Remaining, Exp. Receipt Aug-Oct 26** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(Items_On_PO[Quantity Remaining],Items_On_PO[Item],$A{r},Items_On_PO[Expected Receipt Date],">="&DATE(2026,8,1),Items_On_PO[Expected Receipt Date],"<="&DATE(2026,10,31))
```

**U — Unit Cost (Purchase Price tab)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(INDEX(Purchase_Price[Average Cost],MATCH($A{r},Purchase_Price[SKU],0)),"")
```

**V — Avg Cost (Position Engine)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(INDEX('Position Engine'!$F{r}:$F{r},MATCH($A{r},'Position Engine'!$A{r}:$A{r},0)),"")
```

**W — Base Price / List (NP Allocation)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(INDEX(New_Product_Allocations[Base Price],MATCH("* : "&$A{r},New_Product_Allocations[Name],0)),"")
```

**X — Realized Rate MIN (open SOs, Amer)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(_xlfn.MINIFS(Customer_Open_Orders[Item Rate],Customer_Open_Orders[Item],$A{r},Customer_Open_Orders[Regional ID],"Americas",Customer_Open_Orders[Item Rate],">0"),"")
```

**Y — Realized Rate MAX (open SOs, Amer)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(_xlfn.MAXIFS(Customer_Open_Orders[Item Rate],Customer_Open_Orders[Item],$A{r},Customer_Open_Orders[Regional ID],"Americas",Customer_Open_Orders[Item Rate],">0"),"")
```

**Z — # Open SO Lines w/ Rate** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
COUNTIFS(Customer_Open_Orders[Item],$A{r},Customer_Open_Orders[Regional ID],"Americas",Customer_Open_Orders[Item Rate],">0")
```

**AA — Price Data Flag** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IF(AND(U{r}="",V{r}="",W{r}="",AE{r}="",AD{r}=0),"NO PRICE OR COST DATA",IF(AD{r}=0,"No realized price - cost/list only",IF(AND(U{r}="",V{r}="",AE{r}=""),"No cost data","OK")))
```

**AB — Realized Rate MIN (open SOs, all regions)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(_xlfn.MINIFS(Customer_Open_Orders[Item Rate],Customer_Open_Orders[Item],$A{r},Customer_Open_Orders[Item Rate],">0"),"")
```

**AC — Realized Rate MAX (open SOs, all regions)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(_xlfn.MAXIFS(Customer_Open_Orders[Item Rate],Customer_Open_Orders[Item],$A{r},Customer_Open_Orders[Item Rate],">0"),"")
```

**AD — # Open SO Lines w/ Rate (all regions)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
COUNTIFS(Customer_Open_Orders[Item],$A{r},Customer_Open_Orders[Item Rate],">0")
```

**AE — Implied Unit Cost (Inv Cost / Total Inv)** · 254 cells · rows 2-255 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(SUMIFS(tblInventoryReview[Inv Cost],tblInventoryReview[SKU],$A{r})/SUMIFS(tblInventoryReview[Total Inv],tblInventoryReview[SKU],$A{r}),"")
```

### `Alloc_Hold_List` [veryHidden] — 4,602 formulas

Reads from: `Current Inventory`

**F — On Hand** · 2,301 cells · rows 2-2302 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Current Inventory'!$C{r}:$C{r},'Current Inventory'!$A{r}:$A{r},$A{r})
```

**I — BO Value ($)** · 2,301 cells · rows 2-2302 · 1 pattern(s) · 100.0% consistent

```excel
G{r}*H{r}
```

### `ARC Build Live` — 3,603 formulas

Reads from: `ARC Buildable`, `ARC Parts Live`

**B — (no header)** · 721 cells · rows 4-725 · 2 pattern(s) · 99.9% consistent

```excel
IF($A{r}="","",IF(LEFT($A{r},10)="TP PRE ARC","New (boxed)","Legacy (discrete)"))
```

1 other pattern(s) in this column, largest with 1 cell(s):

```excel
COUNTA(_xlfn.ANCHORARRAY(A{r}))   /* x1 */
```

**C — (no header)** · 720 cells · rows 6-725 · 1 pattern(s) · 100.0% consistent

```excel
IF($A{r}="","",COUNTIFS(BOM_Master[Parent SKU],$A{r},BOM_Master[Is ARC Table],1))
```

**D — (no header)** · 720 cells · rows 6-725 · 1 pattern(s) · 100.0% consistent

```excel
IF($A{r}="","",_xlfn.LET(_xlpm.p,$A{r},_xlpm.cond,(BOM_Master[Parent SKU]=_xlpm.p)*(BOM_Master[Is ARC Table]=1),_xlpm.ln,_xlfn._xlws.FILTER(BOM_Master[Member SKU],_xlpm.cond),_xlpm.qt,_xlfn._xlws.FILTER(BOM_Master[Member Quantity],_xlpm.cond),_xlpm.av,_xlfn.XLOOKUP(_xlpm.ln,'ARC Parts Live'!$A{r}:$A{r},'ARC Parts Live'!$G{r}:$G{r},0),_xlpm.per,IF(IFERROR(_xlpm.qt,0)>0,FLOOR(_xlpm.av/_xlpm.qt,1),9999),_xlpm.b,MIN(_xlpm.per),_xlfn.HSTACK(_xlpm.b,INDEX(_xlpm.ln,MATCH(_xlpm.b,_xlpm.per,0)),INDEX(_xlpm.av,MATCH(_xlpm.b,_xlpm.per,0)))))
```

**G — (no header)** · 720 cells · rows 6-725 · 1 pattern(s) · 100.0% consistent

```excel
IF($A{r}="","",SUMIFS(Customer_Open_Orders[TtlBO],Customer_Open_Orders[Item],$A{r}))
```

**H — (no header)** · 720 cells · rows 6-725 · 1 pattern(s) · 100.0% consistent

```excel
IF($A{r}="","",SUMIFS(Forecast_Qty[Forecast Qty],Forecast_Qty[Item],$A{r},Forecast_Qty[Regional ID],"Americas",Forecast_Qty[YearMonth],">="&'ARC Buildable'!$B{r},Forecast_Qty[YearMonth],"<="&'ARC Buildable'!$B{r}))
```

### `New Product Readiness` — 3,069 formulas

Reads from: `Alloc Overrides`, `Cue Completion Schedule`, `Defect Fallout to Date`, `HIEInv`, `HIEProdFlat`, `JAX Shipments by Date`, `Location Inventory`, `NetSuite Forecast`, `Open Order Report`, `Sales`

**D — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF(C{r}="","TBD",C{r}-TODAY())
```

**F — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
SUMPRODUCT((TRIM('Location Inventory'!$B{r}:$B{r})=$A{r})*IFERROR('Location Inventory'!$C{r}:$C{r}+0,0))
```

**G — (no header)** · 86 cells · rows 5-94 · 2 pattern(s) · 98.8% consistent

```excel
IF(LEFT($A{r},3)<>"BCP","N/A",SUMIFS(HIEInv!$B:$B,HIEInv!$A:$A,$A{r},HIEInv!$C:$C,"regular inventory"))
```

1 other pattern(s) in this column, largest with 1 cell(s):

```excel
COUNTIF(M{r}:M{r},"Behind")   /* x1 */
```

**H — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF(LEFT($A{r},3)<>"BCP","N/A",IF($C{r}="",SUMIFS(HIEProdFlat!$D:$D,HIEProdFlat!$B:$B,$A{r}),SUMIFS(HIEProdFlat!$D:$D,HIEProdFlat!$B:$B,$A{r},HIEProdFlat!$C:$C,"<="&$C{r})))
```

**I — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('JAX Shipments by Date'!$D{r}:$D{r},'JAX Shipments by Date'!$A{r}:$A{r},$A{r})*(('JAX Shipments by Date'!$A{r}+10)<=$C{r})+SUMIFS('JAX Shipments by Date'!$E{r}:$E{r},'JAX Shipments by Date'!$A{r}:$A{r},$A{r})*(('JAX Shipments by Date'!$A{r}+10)<=$C{r})+SUMIFS('JAX Shipments by Date'!$F{r}:$F{r},'JAX Shipments by Date'!$A{r}:$A{r},$A{r})*(('JAX Shipments by Date'!$A{r}+10)<=$C{r})
```

**J — (no header)** · 86 cells · rows 5-94 · 2 pattern(s) · 98.8% consistent

```excel
SUM(F{r}:I{r})
```

1 other pattern(s) in this column, largest with 1 cell(s):

```excel
COUNTIF(M{r}:M{r},"N/A")+COUNTIF(M{r}:M{r},"TBD Date")   /* x1 */
```

**K — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF(OR($E{r}="",$E{r}=0),"N/A",$J{r}/$E{r})
```

**L — (no header)** · 86 cells · rows 5-94 · 2 pattern(s) · 98.8% consistent

```excel
IF(LEFT($A{r},3)<>"BCP","N/A",IF(COUNTIF(HIEProdFlat!$B:$B,$A{r})=0,"None",_xlfn.MAXIFS(HIEProdFlat!$C:$C,HIEProdFlat!$B:$B,$A{r})))
```

1 other pattern(s) in this column, largest with 1 cell(s):

```excel
COUNTIF(S{r}:S{r},"Yes*")   /* x1 */
```

**M — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF(LEFT($A{r},3)<>"BCP","N/A",IF($C{r}="","TBD Date",IF(OR($E{r}="",$E{r}=0),"N/A",IF($K{r}<0.85,"Behind",IF(OR(NOT(ISNUMBER($L{r})),$L{r}<=$C{r}-45),"Ahead of Schedule","On Track")))))
```

**O — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF(OR($E{r}="",$E{r}=0),"",ROUND($E{r}*$N{r},0))
```

**Q — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF($AN{r}<>"",$AN{r},IF($AC{r}>0,SUMIFS(New_Prod_Alloc[Value],New_Prod_Alloc[Name],"*"&$A{r},New_Prod_Alloc[Attribute],"B{r}C"),$AD{r}))
```

**R — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF($AO{r}<>"",$AO{r},IF($AC{r}>0,$AC{r}-$Q{r},$AE{r}-$AD{r}))
```

**S — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF($Q{r}+$R{r}=0,"No Alloc Data",IF($R{r}=0,"Yes — B{r}C only",IF(AND(ISNUMBER($AL{r}),$AL{r}>0.85),"Yes — completed >85%","No — hold in China")))
```

**T — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Defect Fallout to Date'!$B{r}:$B{r},'Defect Fallout to Date'!$A{r}:$A{r},$A{r})
```

**U — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Defect Fallout to Date'!$C{r}:$C{r},'Defect Fallout to Date'!$A{r}:$A{r},$A{r})
```

**V — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF($O{r}="","",$T{r}-$O{r})
```

**W — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Cue Completion Schedule'!$C{r}:$C{r},'Cue Completion Schedule'!$A{r}:$A{r},$A{r})
```

**X — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF(LEFT($A{r},3)<>"BCP","N/A",$F{r}+$G{r}+$W{r})
```

**Y — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF(LEFT($A{r},3)<>"BCP","N/A",MAX(0,$E{r}-$X{r}))
```

**Z — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
_xlfn.MINIFS('Cue Completion Schedule'!$B{r}:$B{r},'Cue Completion Schedule'!$A{r}:$A{r},$A{r},'Cue Completion Schedule'!$I{r}:$I{r},"Yes")
```

**AA — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF(LEFT($A{r},3)<>"BCP","N/A",IF($F{r}+$G{r}>=$E{r},"Complete now",IF($Z{r}=0,"Not scheduled to complete",$Z{r})))
```

**AB — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF(AND(ISNUMBER($AA{r}),ISNUMBER($C{r})),$AA{r}-$C{r},"")
```

**AC — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(New_Prod_Alloc[Value],New_Prod_Alloc[Name],"*"&$A{r})
```

**AD — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('NetSuite Forecast'!$D{r}:$D{r},'NetSuite Forecast'!$A{r}:$A{r},"*"&$A{r},'NetSuite Forecast'!$B{r}:$B{r},"B{r}C")
```

**AE — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('NetSuite Forecast'!$D{r}:$D{r},'NetSuite Forecast'!$A{r}:$A{r},"*"&$A{r})
```

**AF — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF($AP{r}<>"",$AP{r},IF($AC{r}>0,"New Prod Alloc",IF($AE{r}>0,"NetSuite Forecast (fallback)","None")))
```

**AG — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(Sales!$D{r}:$D{r},Sales!$B{r}:$B{r},"*"&$A{r})
```

**AH — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF(NOT(ISNUMBER($C{r})),"",SUMIFS(Sales!$D{r}:$D{r},Sales!$B{r}:$B{r},"*"&$A{r},Sales!$A{r}:$A{r},"<"&$C{r}))
```

**AI — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS(Sales!$D{r}:$D{r},Sales!$B{r}:$B{r},"*"&$A{r},Sales!$A{r}:$A{r},">="&TODAY()-90)
```

**AJ — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF($AG{r}<=0,"No sales",IF(AND(ISNUMBER($C{r}),$C{r}>TODAY()),"SOLD PRE-LAUNCH","Selling (launch passed)"))
```

**AK — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF(LEFT($A{r},3)<>"BCP","N/A",MAX(0,$E{r}-$AG{r}))
```

**AL — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IF(LEFT($A{r},3)<>"BCP","N/A",IF(OR($E{r}="",$E{r}=0),"N/A",($F{r}+$G{r})/$E{r}))
```

**AM — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Open Order Report'!$D{r}:$D{r},'Open Order Report'!$C{r}:$C{r},"*"&$A{r})
```

**AN — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(INDEX('Alloc Overrides'!$B{r}:$B{r},MATCH($A{r},'Alloc Overrides'!$A{r}:$A{r},0)),"")
```

**AO — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(INDEX('Alloc Overrides'!$C{r}:$C{r},MATCH($A{r},'Alloc Overrides'!$A{r}:$A{r},0)),"")
```

**AP — (no header)** · 85 cells · rows 10-94 · 1 pattern(s) · 100.0% consistent

```excel
IFERROR(INDEX('Alloc Overrides'!$D{r}:$D{r},MATCH($A{r},'Alloc Overrides'!$A{r}:$A{r},0)),"")
```

### `Launch Status 2026` — 1,777 formulas

Reads from: `New Product Readiness`, `PO`

**A — Launch Production Status — 2026 launches (HIE)** · 88 cells · rows 5-92 · 4 pattern(s) · 38.6% consistent

```excel
"Flag: below 85% complete"
```

3 other pattern(s) in this column, largest with 28 cell(s):

```excel
"No completion data (non-BCP)"   /* x28 */
"Requested item"   /* x17 */
```

**B — (no header)** · 17 cells · rows 5-92 · 7 pattern(s) · 23.5% consistent

```excel
"P{r} Maple Teal"
```

6 other pattern(s) in this column, largest with 3 cell(s):

```excel
"Panthera 10"   /* x3 */
"Air Rush Rose Gold"   /* x2 */
```

**C — (no header)** · 88 cells · rows 5-92 · 4 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!A{r}
```

3 other pattern(s) in this column, largest with 1 cell(s):

```excel
"BCP PRE SE PAN 10-1 NW"   /* x1 */
"BCP PRE SE PAN 10-2 LEA"   /* x1 */
```

**D — (no header)** · 88 cells · rows 5-92 · 4 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!B{r}
```

3 other pattern(s) in this column, largest with 1 cell(s):

```excel
"Panthera 10 - Birdseye Maple Grip / No Wrap"   /* x1 */
"Panthera 10 - Leather Handle / Lizard Wrap"   /* x1 */
```

**E — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!C{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
DATE(2026,9,15)   /* x3 */
```

**F — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!D{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
19   /* x3 */
```

**G — (no header)** · 88 cells · rows 5-92 · 3 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!E{r}
```

2 other pattern(s) in this column, largest with 2 cell(s):

```excel
140   /* x2 */
60   /* x1 */
```

**H — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!F{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
"N/A"   /* x3 */
```

**I — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!G{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
"N/A"   /* x3 */
```

**J — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r})),H{r}+I{r},"N/A")
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
"N/A"   /* x3 */
```

**K — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!H{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
"N/A"   /* x3 */
```

**L — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!I{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
"N/A"   /* x3 */
```

**M — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!AL{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
"N/A"   /* x3 */
```

**N — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!K{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
"N/A"   /* x3 */
```

**O — (no header)** · 88 cells · rows 5-92 · 1 pattern(s) · 100.0% consistent

```excel
IF(COUNTIF(PO!$A:$A,"*"&$C{r})>0,"Yes","No")
```

**P — (no header)** · 88 cells · rows 5-92 · 1 pattern(s) · 100.0% consistent

```excel
SUMIF(PO!$A:$A,"*"&$C{r},PO!$C:$C)
```

**Q — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!M{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
"Not tracked — not made by HIE"   /* x3 */
```

**R — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!S{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
"Not tracked — not made by HIE"   /* x3 */
```

**S — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!L{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
""   /* x3 */
```

**T — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!AA{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
""   /* x3 */
```

**U — (no header)** · 88 cells · rows 5-92 · 2 pattern(s) · 96.6% consistent

```excel
'New Product Readiness'!AM{r}
```

1 other pattern(s) in this column, largest with 3 cell(s):

```excel
""   /* x3 */
```

### `SO_Closure_List` [hidden] — 1,690 formulas

Reads from: `Current Inventory`, `Items on PO`

**I — (no header)** · 564 cells · rows 1-565 · 2 pattern(s) · 99.8% consistent

```excel
G{r}*H{r}
```

1 other pattern(s) in this column, largest with 1 cell(s):

```excel
SUBTOTAL(9,I{r}:I{r})   /* x1 */
```

**J — (no header)** · 563 cells · rows 3-565 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Current Inventory'!$C{r}:$C{r},'Current Inventory'!$A{r}:$A{r},$A{r})
```

**K — (no header)** · 563 cells · rows 3-565 · 1 pattern(s) · 100.0% consistent

```excel
SUMIFS('Items on PO'!$N{r}:$N{r},'Items on PO'!$K{r}:$K{r},$A{r},'Items on PO'!$F{r}:$F{r},"<>TO*",'Items on PO'!$G{r}:$G{r},"<>*Clawson*",'Items on PO'!$G{r}:$G{r},"<>*Predator Group*")
```

### `ARC Parts Live` — 1,362 formulas

Reads from: `ARC Buildable`

**B — (no header)** · 171 cells · rows 4-175 · 2 pattern(s) · 99.4% consistent

```excel
IF($A{r}="","",IFERROR("BX"&_xlfn.TEXTAFTER($A{r}," BX"),"—"))
```

1 other pattern(s) in this column, largest with 1 cell(s):

```excel
COUNTA(_xlfn.ANCHORARRAY(A{r}))   /* x1 */
```

**C — (no header)** · 170 cells · rows 6-175 · 1 pattern(s) · 100.0% consistent

```excel
IF($A{r}="","",COUNTIFS(BOM_Master[Member SKU],$A{r},BOM_Master[Is ARC Table],1))
```

**D — (no header)** · 170 cells · rows 6-175 · 1 pattern(s) · 100.0% consistent

```excel
IF($A{r}="","",SUMIFS(Current_Inventory[InvQty],Current_Inventory[SKU],$A{r}))
```

**E — (no header)** · 170 cells · rows 6-175 · 1 pattern(s) · 100.0% consistent

```excel
IF($A{r}="","",SUMIFS(Customer_Open_Orders[TtlBO],Customer_Open_Orders[Item],$A{r}))
```

**F — (no header)** · 170 cells · rows 6-175 · 1 pattern(s) · 100.0% consistent

```excel
IF($A{r}="","",SUMIFS(Items_On_PO[Quantity Remaining],Items_On_PO[Item],$A{r}))
```

**G — (no header)** · 170 cells · rows 6-175 · 1 pattern(s) · 100.0% consistent

```excel
IF($A{r}="","",MAX(0,$D{r}-$E{r}))
```

**H — (no header)** · 170 cells · rows 6-175 · 1 pattern(s) · 100.0% consistent

```excel
IF($A{r}="","",SUMIFS('ARC Buildable'!$T{r}:$T{r},'ARC Buildable'!$M{r}:$M{r},$A{r}))
```

**I — (no header)** · 170 cells · rows 6-175 · 1 pattern(s) · 100.0% consistent

```excel
IF($A{r}="","",$D{r}-$H{r})
```

### `Sales_Hygiene_Flags` [hidden] — 1,053 formulas

Reads from: `PE Config`

**J — BO Value ($)** · 351 cells · rows 2-352 · 1 pattern(s) · 100.0% consistent

```excel
H{r}*I{r}
```

**P — Order Age (d)** · 351 cells · rows 2-352 · 1 pattern(s) · 100.0% consistent

```excel
'PE Config'!$B{r}-G{r}
```

**Q — Removal Step** · 351 cells · rows 2-352 · 1 pattern(s) · 100.0% consistent

```excel
IF(P{r}>Overdue_Cutoff,"6 Severely overdue",IF(ISNUMBER(SEARCH("Pending approval",A{r})),"7 Pending approval",IF(ISNUMBER(SEARCH("Over credit limit",A{r})),"8 Over credit limit",IF(ISNUMBER(SEARCH("Zero item rate",A{r})),"9 Zero item rate",IF(ISNUMBER(SEARCH("Credit hold",A{r})),"10 Credit hold ON","Kept (no removal rule)")))))
```
