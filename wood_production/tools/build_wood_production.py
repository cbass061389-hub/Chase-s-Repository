"""Build Wood_Production.xlsx — the Panthera Manufacturing (cue butt) line workbook.

Mirrors the TABLE Production workbook's core: Plan Config -> Floor board -> QC /
release logs -> Production Plan -> Weekly Reporting / Summary, with the sheets a
wood line needs that a table line does not: receiving inspection (grade and
moisture), a humidity log against the 32-40 % RH window, and rest gates between
turning stages.

Everything that is a guess is a yellow input cell with a note saying so. Station
minutes in particular are placeholders until Chad and Steven write the setup
sheets after visit 1 (tracker P1.3).

Usage:
    python wood_production/tools/build_wood_production.py <out.xlsx>
"""

from __future__ import annotations

import math
import sys
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
TITLE_FONT = Font(name=FONT, size=14, bold=True)
BOLD = Font(name=FONT, bold=True)
NORMAL = Font(name=FONT)
MUTED = Font(name=FONT, italic=True, color="6E6E6E")
INPUT_FONT = Font(name=FONT, color="0000FF")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

LAST_FLOOR_ROW = 500
LAST_LOG_ROW = 5000
PLAN_WEEKS = 40

# ---------------------------------------------------------------------------
# Process definition — Brandon's nine-step plain-butt scope, grouped into seven
# board stations. Minutes are placeholders (see module docstring).
# ---------------------------------------------------------------------------
STATIONS = [
    # id, name, zone, low_min, high_min, rest_after_days, inspector_not_builder
    ("WS01", "Receive, inspect & cut blanks (grade, moisture)", "Pre-Build", 15, 25, 0, "N"),
    ("WS02", "Rough turn forearm & handle to 1.300 straight", "Turning", 30, 45, 14, "N"),
    ("WS03", "Rough taper & forearm coring", "Turning", 40, 60, 14, "N"),
    ("WS04", "Cut hardware; machine A-joint (F/M), weight cartridge hole, butt cap profile", "Machining", 60, 90, 0, "N"),
    ("WS05", "Joint pin install; preform glue-up (forearm to handle)", "Assembly", 30, 45, 1, "N"),
    ("WS06", "Final taper 1.350 -> 1.250 in 0.025 steps", "Turning", 45, 75, 0, "N"),
    ("WS07", "Independent inspection, robot test, clear-coat prep", "Finishing", 20, 30, 0, "Y"),
]
STATION_IDS = [s[0] for s in STATIONS]
FIRST_ST_ROW = 6
LAST_ST_ROW = FIRST_ST_ROW + len(STATIONS) - 1          # 12

# Batch plan from the project tracker (P1.6, P2.6, P3.4). 10 % allowance = Brandon's overrun.
BATCHES = [
    # id, label, sku, good_units, allowance, start (Mon), end (Mon of last week), target
    ("B1", "Prototype — Broomstick 1", "PAN-BRM-001", 1, 0.00, date(2026, 12, 14), date(2026, 12, 14), date(2026, 12, 18)),
    ("B2", "25 identical broomsticks (plain butt)", "PAN-BRM-001", 25, 0.10, date(2027, 1, 4), date(2027, 1, 18), date(2027, 1, 20)),
    ("B3", "50 sellable cues (first sellable design, STR3)", "PAN-SELL-001", 50, 0.10, date(2027, 3, 29), date(2027, 6, 7), date(2027, 6, 11)),
]
PLAN_START = date(2026, 12, 14)

STARTER_REASONS = [
    "Warpage / movement after cut",
    "Moisture content out of range",
    "Runout — not straight",
    "Dimension out of tolerance",
    "Chatter / tool marks",
    "Grain defect / check / crack",
    "Joint fit (A-joint)",
    "Glue line / glue-up failure",
    "Pin dot / surface blemish",
    "Weight cartridge hole off",
    "Operator error / wrong setup sheet",
    "Material shortage (hardware)",
]

CLASSIFICATION = [
    ("PAN-BRM-001", "BROOMSTICK", "PLAIN BUTT", "PLACEHOLDER",
     "SKU convention pending (tracker INV2). Broomstick 1: two 18in battens, A-joint radial pin, PETP tube, bumper (INV4).",
     "Project tracker v2", "Plain, no wrap, no inlay", "Forearm-to-handle, cored"),
    ("PAN-SELL-001", "SELLABLE", "TBD (STR3)", "PLACEHOLDER",
     "First sellable design is owned by Parker (STR3); ready by M6. Replace this row when the spec lands.",
     "Project tracker v2", "TBD", "TBD"),
]


def monday(d: date) -> date:
    return d - timedelta(days=d.weekday())


def style_header(ws, row: int, first_col: int, last_col: int) -> None:
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = BOLD
        cell.fill = HDR_FILL
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def set_input(cell, value, note: str | None = None) -> None:
    cell.value = value
    cell.fill = INPUT_FILL
    cell.font = INPUT_FONT
    if note:
        cell.comment = Comment(note, "Wood_Production build")


def add_table(ws, name: str, ref: str, style: str = "TableStyleMedium2") -> None:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
    ws.add_table(table)


def widths(ws, mapping: dict[str, float]) -> None:
    for col, width in mapping.items():
        ws.column_dimensions[col].width = width


def base_font(ws) -> None:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and cell.font == Font():
                cell.font = NORMAL


# ---------------------------------------------------------------------------
# Plan Config
# ---------------------------------------------------------------------------
def build_plan_config(wb: Workbook):
    ws = wb.active
    ws.title = "Plan Config"
    ws["A1"] = "Plan Configuration — Panthera Manufacturing cue butt cell"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Mirrors TABLE Plan Config  |  Yellow cells are inputs  |  Station minutes are PLACEHOLDERS until the visit-1 setup sheets (tracker P1.3)"
    ws["A2"].font = MUTED

    ws["A4"] = "PER-STATION CAPACITY"; ws["A4"].font = BOLD
    headers = ["Station", "Station Name", "Zone", "Low (min)", "High (min)", "Avg (min)", "Parallel?",
               "Hours/Cue", "Cues/day per team", "People Req", "Machines Req", "Weekday People",
               "Weekday Machines", "Weekday Cap/day", "Sun People", "Sun Machines", "Sun Cap/day",
               "Day-only Wk Cap", "2-Shift Wk Cap", "Binding", "Rest After (days)", "Inspector ≠ Builder?"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=5, column=i, value=h)
    style_header(ws, 5, 1, len(headers))

    for i, (sid, name, zone, lo, hi, rest, insp) in enumerate(STATIONS):
        r = FIRST_ST_ROW + i
        ws.cell(row=r, column=1, value=sid).font = BOLD
        ws.cell(row=r, column=2, value=name)
        ws.cell(row=r, column=3, value=zone)
        set_input(ws.cell(row=r, column=4), lo, "Placeholder. Replace with the setup-sheet time after visit 1 (P1.3).")
        set_input(ws.cell(row=r, column=5), hi, "Placeholder. Replace with the setup-sheet time after visit 1 (P1.3).")
        ws.cell(row=r, column=6, value=f"=AVERAGE(D{r}:E{r})")
        set_input(ws.cell(row=r, column=7), "N")
        ws.cell(row=r, column=8, value=f"=F{r}/60")
        ws.cell(row=r, column=9, value=f"=IF(H{r}=0,0,ROUND($Y$8/H{r},2))")
        set_input(ws.cell(row=r, column=10), 1)
        set_input(ws.cell(row=r, column=11), 1 if zone in ("Turning", "Machining") else 0)
        ws.cell(row=r, column=12, value=f"=J{r}")
        ws.cell(row=r, column=13, value=f"=K{r}")
        ws.cell(row=r, column=14, value=f"=ROUND(I{r}*L{r},2)")
        set_input(ws.cell(row=r, column=15), 0)
        set_input(ws.cell(row=r, column=16), 0)
        ws.cell(row=r, column=17, value=f"=ROUND(I{r}*O{r},2)")
        ws.cell(row=r, column=18, value=f"=ROUND(N{r}*$Y$9,2)")
        ws.cell(row=r, column=19, value=f"=ROUND(N{r}*$Y$9+Q{r}*$Y$10,2)")
        ws.cell(row=r, column=20, value=f'=IF(K{r}=0,"Labor",IF(M{r}/MAX(K{r},1)<L{r}/MAX(J{r},1),"Machines","Labor"))')
        set_input(ws.cell(row=r, column=21), rest,
                  "Wood rests a minimum of two weeks between turning stages (Brandon). Glue cure after WS05." if rest else None)
        set_input(ws.cell(row=r, column=22), insp,
                  "The inspector is never the builder (tracker P2.5)." if insp == "Y" else None)

    # Plan clock (X:Y), same roles as TABLE's V:W
    ws["X4"] = "PLAN CLOCK"; ws["X4"].font = BOLD
    ws["X5"] = "Today:";            ws["Y5"] = "=TODAY()"
    ws["X6"] = "Plan start:";       set_input(ws["Y6"], PLAN_START, "First prototype build week (tracker P1.6). Drives Production Plan column B.")
    ws["X7"] = "Board start (roll):"; ws["Y7"] = "=TODAY()"
    ws["X8"] = "Workday Hours:";    set_input(ws["Y8"], 8)
    ws["X9"] = "Day days/wk:";      set_input(ws["Y9"], 5)
    ws["X10"] = "Night days/wk:";   set_input(ws["Y10"], 0)
    for r in range(5, 11):
        ws[f"X{r}"].font = BOLD
        ws[f"Y{r}"].number_format = "m/d/yyyy" if r in (5, 6, 7) else "0"

    # Wood parameters
    ws["A14"] = "WOOD PARAMETERS"; ws["A14"].font = BOLD
    params = [
        (15, "Rest minimum between turning stages (days)", 14, "Brandon: minimum two weeks between stages whenever possible."),
        (16, "RH low (%)", 32, "Jacoby holds 32-40 % RH year round (tracker P1.5)."),
        (17, "RH high (%)", 40, None),
        (18, "Moisture content low (%) — set by QA1", 6, "PLACEHOLDER. Receiving inspection (QA1) sets the accept range."),
        (19, "Moisture content high (%) — set by QA1", 8, "PLACEHOLDER. Receiving inspection (QA1) sets the accept range."),
        (20, "Scrap / overrun allowance (%)", 0.10, "Jacoby overruns every order ~10 % for movement and mistakes (tracker P2.5)."),
        (21, "Standard batch size", 25, "Batches of 25, as HIE runs (tracker INV4)."),
        (22, "Humidity excursion limit (minutes) — set by QA3", None, "PLACEHOLDER. QA3 defines how far out, for how long, and the hold rule."),
        (23, "Finish location (interim)", "Jacoby (per P2.4 / P3.2 gate)", None),
    ]
    for r, label, val, note in params:
        ws[f"A{r}"] = label
        set_input(ws[f"B{r}"], val, note)
    ws["B20"].number_format = "0%"

    # Batch plan — drives Production Plan
    ws["A26"] = "BATCH PLAN  (drives Production Plan; good units + allowance = planned starts)"; ws["A26"].font = BOLD
    bh = ["Batch", "Description", "SKU", "Good units required", "Allowance %", "Planned starts",
          "Start week (Mon)", "End week (Mon)", "Weeks", "Base / week", "Remainder (last week)", "Target date"]
    for i, h in enumerate(bh, start=1):
        ws.cell(row=27, column=i, value=h)
    style_header(ws, 27, 1, len(bh))
    for i, (bid, label, sku, good, allow, start, end, target) in enumerate(BATCHES):
        r = 28 + i
        ws.cell(row=r, column=1, value=bid).font = BOLD
        set_input(ws.cell(row=r, column=2), label)
        set_input(ws.cell(row=r, column=3), sku)
        set_input(ws.cell(row=r, column=4), good)
        set_input(ws.cell(row=r, column=5), allow)
        ws.cell(row=r, column=5).number_format = "0%"
        ws.cell(row=r, column=6, value=f"=ROUNDUP(ROUND(D{r}*(1+E{r}),6),0)")
        set_input(ws.cell(row=r, column=7), start); ws.cell(row=r, column=7).number_format = "m/d/yyyy"
        set_input(ws.cell(row=r, column=8), end);   ws.cell(row=r, column=8).number_format = "m/d/yyyy"
        ws.cell(row=r, column=9, value=f"=(H{r}-G{r})/7+1")
        ws.cell(row=r, column=10, value=f"=INT(F{r}/I{r})")
        ws.cell(row=r, column=11, value=f"=F{r}-J{r}*I{r}")
        set_input(ws.cell(row=r, column=12), target); ws.cell(row=r, column=12).number_format = "m/d/yyyy"
    ws["A32"] = "Start/End are Mondays. Planned starts spread evenly; the remainder lands in the last week. Add a batch on row 31 if the plan grows (Production Plan reads rows 28:31)."
    ws["A32"].font = MUTED

    # Zone roll-up and lead time
    ws["A35"] = "ZONE ROLL-UP"; ws["A35"].font = BOLD
    for i, h in enumerate(["Zone", "Stations", "Total Avg (min)", "Hrs"], start=1):
        ws.cell(row=36, column=i, value=h)
    style_header(ws, 36, 1, 4)
    zones = ["Pre-Build", "Turning", "Machining", "Assembly", "Finishing"]
    for i, z in enumerate(zones):
        r = 37 + i
        ws.cell(row=r, column=1, value=z)
        ws.cell(row=r, column=2, value=f"=COUNTIF($C${FIRST_ST_ROW}:$C${LAST_ST_ROW},A{r})")
        ws.cell(row=r, column=3, value=f"=SUMIF($C${FIRST_ST_ROW}:$C${LAST_ST_ROW},A{r},$F${FIRST_ST_ROW}:$F${LAST_ST_ROW})")
        ws.cell(row=r, column=4, value=f"=C{r}/60")
    r = 37 + len(zones)
    ws.cell(row=r, column=1, value="TOTAL").font = BOLD
    ws.cell(row=r, column=2, value=f"=COUNTA($A${FIRST_ST_ROW}:$A${LAST_ST_ROW})")
    ws.cell(row=r, column=3, value=f"=SUM(C37:C{r-1})")
    ws.cell(row=r, column=4, value=f"=C{r}/60")
    ws.cell(row=r + 1, column=1, value="LESS PARALLEL (stations flagged Y)")
    ws.cell(row=r + 1, column=3, value=f'=SUMIF($G${FIRST_ST_ROW}:$G${LAST_ST_ROW},"Y",$F${FIRST_ST_ROW}:$F${LAST_ST_ROW})')
    ws.cell(row=r + 1, column=4, value=f"=C{r+1}/60")
    ws.cell(row=r + 2, column=1, value="TOUCH TIME / CUE (single builder, min)").font = BOLD
    ws.cell(row=r + 2, column=3, value=f"=C{r}-C{r+1}")
    ws.cell(row=r + 2, column=4, value=f"=C{r+2}/60")
    ws.cell(row=r + 3, column=1, value="REST DAYS / CUE (sum of Rest After)").font = BOLD
    ws.cell(row=r + 3, column=3, value=f"=SUM($U${FIRST_ST_ROW}:$U${LAST_ST_ROW})")
    ws.cell(row=r + 4, column=1, value="CALENDAR LEAD TIME / CUE (days, touch days + rest days)").font = BOLD
    ws.cell(row=r + 4, column=3, value=f"=ROUNDUP(C{r+2}/60/$Y$8,0)+C{r+3}")
    ws.cell(row=r + 5, column=1, value="BOTTLENECK STATION (line model)").font = BOLD
    ws.cell(row=r + 5, column=2, value=f"=INDEX($A${FIRST_ST_ROW}:$A${LAST_ST_ROW},MATCH(MAX($F${FIRST_ST_ROW}:$F${LAST_ST_ROW}),$F${FIRST_ST_ROW}:$F${LAST_ST_ROW},0))")
    ws.cell(row=r + 5, column=3, value=f"=MAX($F${FIRST_ST_ROW}:$F${LAST_ST_ROW})")
    ws.cell(row=r + 5, column=4, value=f"=C{r+5}/60")

    widths(ws, {"A": 46, "B": 44, "C": 14, "D": 10, "E": 10, "F": 10, "G": 9, "H": 10, "I": 11, "J": 10,
                "K": 11, "L": 11, "M": 12, "N": 12, "O": 10, "P": 11, "Q": 11, "R": 13, "S": 13, "T": 10,
                "U": 12, "V": 14, "W": 3, "X": 18, "Y": 14})
    ws.freeze_panes = "A6"
    return ws


# ---------------------------------------------------------------------------
# WOOD Floor
# ---------------------------------------------------------------------------
FLOOR_HEADERS = ["WO#", "SKU", "Unit ID", "Serial", "Qty", "Wood Lot", "Current Zone", "Target Date",
                 *STATION_IDS, "Rough Turn Date", "Rest 1 Until", "Taper Date", "Rest 2 Until", "Rest OK?",
                 "Release?", "Hold?", "Next Station", "PriorityKey", "QC Status", "QC Date"]
COL = {h: get_column_letter(i) for i, h in enumerate(FLOOR_HEADERS, start=1)}
ST_FIRST = COL[STATION_IDS[0]]
ST_LAST = COL[STATION_IDS[-1]]


def floor_formulas(r: int) -> dict[str, str]:
    c = COL
    st = f"${ST_FIRST}{r}:${ST_LAST}{r}"
    n_st = f"COUNTA('Plan Config'!$A${FIRST_ST_ROW}:$A${LAST_ST_ROW})"
    zone_rng = f"'Plan Config'!$C${FIRST_ST_ROW}:$C${LAST_ST_ROW}"
    id_rng = f"'Plan Config'!$A${FIRST_ST_ROW}:$A${LAST_ST_ROW}"
    unit = f"${c['Unit ID']}{r}"
    return {
        c["Current Zone"]: f'=IF({unit}="","",IF(COUNTIF({st},TRUE)=0,"Not started",INDEX({zone_rng},COUNTIF({st},TRUE))))',
        c["Rest 1 Until"]: f'=IF(${c["Rough Turn Date"]}{r}="","",${c["Rough Turn Date"]}{r}+\'Plan Config\'!$U$7)',
        c["Rest 2 Until"]: f'=IF(${c["Taper Date"]}{r}="","",${c["Taper Date"]}{r}+\'Plan Config\'!$U$8)',
        c["Rest OK?"]: (f'=IF({unit}="","",IF(AND(${c["Next Station"]}{r}="WS03",${c["Rest 1 Until"]}{r}<>"",TODAY()<${c["Rest 1 Until"]}{r}),'
                        f'"REST until "&TEXT(${c["Rest 1 Until"]}{r},"d mmm"),'
                        f'IF(AND(${c["Next Station"]}{r}="WS04",${c["Rest 2 Until"]}{r}<>"",TODAY()<${c["Rest 2 Until"]}{r}),'
                        f'"REST until "&TEXT(${c["Rest 2 Until"]}{r},"d mmm"),"OK")))'),
        c["Release?"]: f'=IF({unit}="","",COUNTIF({st},TRUE)>={n_st})',
        c["Hold?"]: f'=IF({unit}="","",IF(COUNTIFS(Holds!$A$2:$A$500,{unit},Holds!$E$2:$E$500,FALSE)>0,"ON HOLD",""))',
        c["Next Station"]: f'=IF({unit}="","",IF(COUNTIF({st},TRUE)>={n_st},"COMPLETE",INDEX({id_rng},COUNTIF({st},TRUE)+1)))',
        c["PriorityKey"]: f'=IF({unit}="","",IF(${c["Release?"]}{r}=TRUE,"",IF(${c["Hold?"]}{r}="ON HOLD","",N(${c["Target Date"]}{r})+ROW()/100000)))',
    }


def build_floor(wb: Workbook):
    ws = wb.create_sheet("WOOD Floor", 0)
    ws["A1"] = "WOOD Floor — live build board (one row = one serialised cue butt)"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Tick a station as it completes (TRUE). Unit ID is the planning key; Serial is applied at WS07. "
                "Enter Rough Turn Date (after WS02) and Taper Date (after WS03): the board shows when the rest period clears. "
                "QC / release through the Wood QC button (frmWoodQuality); QC Status greys the row.")
    ws["A2"].font = MUTED
    for i, h in enumerate(FLOOR_HEADERS, start=1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, 1, len(FLOOR_HEADERS))

    r = 4
    unit_no = 1
    for bid, label, sku, good, allow, start, end, target in BATCHES:
        starts = math.ceil(round(good * (1 + allow), 6))      # 50 x 1.1 is 55.00000000000001 in floating point
        first_unit_row = r + 1
        head = f'="{bid}   —   {label}   —   {starts} cue(s)' + (f' ({good} + {starts - good} overrun)' if starts > good else "") + \
               f'   —   Target: "&TEXT($H{first_unit_row},"ddd d mmm yyyy")'
        ws.cell(row=r, column=1, value=head).font = BOLD
        ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor="EDEDED")
        r += 1
        for _ in range(starts):
            ws.cell(row=r, column=1, value=f"WO-PAN-{bid[1:].zfill(3)}")
            ws.cell(row=r, column=2, value=sku)
            ws.cell(row=r, column=3, value=f"C-{unit_no:04d}")
            ws.cell(row=r, column=5, value=1)
            ws.cell(row=r, column=8, value=target).number_format = "m/d/yyyy"
            for sid in STATION_IDS:
                ws[f"{COL[sid]}{r}"] = False
            for col, formula in floor_formulas(r).items():
                ws[f"{col}{r}"] = formula
            for h in ("Rough Turn Date", "Rest 1 Until", "Taper Date", "Rest 2 Until", "QC Date"):
                ws[f"{COL[h]}{r}"].number_format = "m/d/yyyy"
            unit_no += 1
            r += 1
    last_seeded = r - 1

    # Formulas pre-filled below the seeded rows so new cues can be typed straight in
    for rr in range(last_seeded + 1, LAST_FLOOR_ROW + 1):
        for col, formula in floor_formulas(rr).items():
            ws[f"{col}{rr}"] = formula

    # Validation: station ticks are TRUE/FALSE
    dv = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{ST_FIRST}4:{ST_LAST}{LAST_FLOOR_ROW}")

    # Conditional formatting (same mechanism as the TABLE patch)
    grey = PatternFill("solid", bgColor="D9D9D9")
    grey_font = Font(color="808080")
    strike_font = Font(color="808080", strike=True)
    rng = f"A4:{COL['PriorityKey']}{LAST_FLOOR_ROW}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${COL["QC Status"]}4="REJECTED"'], fill=grey, font=strike_font, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${COL["QC Status"]}4<>""'], fill=grey, font=grey_font))
    ws.conditional_formatting.add(f"{COL['Hold?']}4:{COL['Hold?']}{LAST_FLOOR_ROW}",
                                  FormulaRule(formula=[f'${COL["Hold?"]}4="ON HOLD"'], fill=PatternFill("solid", bgColor="F8CBAD"), font=Font(bold=True, color="9C0006")))
    ws.conditional_formatting.add(f"{COL['Rest OK?']}4:{COL['Rest OK?']}{LAST_FLOOR_ROW}",
                                  FormulaRule(formula=[f'LEFT(${COL["Rest OK?"]}4,4)="REST"'], fill=PatternFill("solid", bgColor="FFE699"), font=Font(color="7F6000")))
    ws.conditional_formatting.add(f"{ST_FIRST}4:{ST_LAST}{LAST_FLOOR_ROW}",
                                  FormulaRule(formula=[f"{ST_FIRST}4=TRUE"], fill=PatternFill("solid", bgColor="C6EFCE"), font=Font(color="006100")))

    widths(ws, {"A": 14, "B": 16, "C": 10, "D": 12, "E": 5, "F": 12, "G": 13, "H": 12,
                **{COL[s]: 7 for s in STATION_IDS},
                COL["Rough Turn Date"]: 12, COL["Rest 1 Until"]: 12, COL["Taper Date"]: 12, COL["Rest 2 Until"]: 12,
                COL["Rest OK?"]: 16, COL["Release?"]: 9, COL["Hold?"]: 10, COL["Next Station"]: 12,
                COL["PriorityKey"]: 12, COL["QC Status"]: 12, COL["QC Date"]: 11})
    ws.freeze_panes = "D4"
    ws.auto_filter.ref = f"A3:{COL['QC Date']}{last_seeded}"
    return ws


# ---------------------------------------------------------------------------
# Holds, logs, reasons
# ---------------------------------------------------------------------------
def build_holds(wb: Workbook):
    ws = wb.create_sheet("Holds")
    for i, h in enumerate(["Unit ID", "Station", "Reason (rework)", "Date Held", "Released?", "Status"], start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, 1, 6)
    ws["H1"] = "How to use: enter the Unit ID and station where the cue is held. Set Released? = TRUE to return it to the floor. The Wood QC form's Hold outcome writes here."
    ws["H1"].font = MUTED
    dv_st = DataValidation(type="list", formula1=f"='Plan Config'!$A${FIRST_ST_ROW}:$A${LAST_ST_ROW}", allow_blank=True)
    dv_rel = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    ws.add_data_validation(dv_st); ws.add_data_validation(dv_rel)
    dv_st.add("B2:B500"); dv_rel.add("E2:E500")
    for r in range(2, 501):
        ws[f"D{r}"].number_format = "m/d/yyyy"
    widths(ws, {"A": 10, "B": 10, "C": 40, "D": 12, "E": 10, "F": 12, "H": 60})
    return ws


def build_log(wb: Workbook, title: str, table: str, headers: list[str], date_cols: tuple[str, ...] = ("A",),
              extra_row2: dict[str, str] | None = None):
    ws = wb.create_sheet(title)
    for i, h in enumerate(headers, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, 1, len(headers))
    for col, formula in (extra_row2 or {}).items():
        ws[f"{col}2"] = formula
    for col in date_cols:
        for r in range(2, 60):
            ws[f"{col}{r}"].number_format = "m/d/yyyy"
    add_table(ws, table, f"A1:{get_column_letter(len(headers))}2")
    for i in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = max(11, min(28, len(headers[i - 1]) + 4))
    return ws


def build_reasons(wb: Workbook):
    ws = wb.create_sheet("QC Reasons")
    headers = ["Reason", "Times Used", "First Seen", "Last Used", "Active", "Category (later)"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, 1, 6)
    for i, reason in enumerate(STARTER_REASONS, start=2):
        ws.cell(row=i, column=1, value=reason)
        ws.cell(row=i, column=2, value=0)
        ws.cell(row=i, column=5, value=True)
        ws[f"C{i}"].number_format = "m/d/yyyy"; ws[f"D{i}"].number_format = "m/d/yyyy"
    add_table(ws, "tbl_QCReasons", f"A1:F{len(STARTER_REASONS) + 1}")
    ws["H1"] = ("Starter list from the Jacoby notes — free text is still accepted on the form and new reasons are added here "
                "automatically. Once there is data, fill 'Category (later)' and set Active = FALSE on the noise.")
    ws["H1"].font = MUTED
    widths(ws, {"A": 38, "B": 11, "C": 12, "D": 12, "E": 8, "F": 18, "H": 70})
    return ws


def build_receiving(wb: Workbook):
    ws = wb.create_sheet("Wood Receiving")
    headers = ["Date", "Lot", "Supplier", "Species", "Grade", "Qty (blanks)", "Moisture %", "RH % at receipt",
               "Temp F", "Disposition", "Inspector", "Certs / Lacey Act ref", "Notes", "Moisture in range?"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, 1, len(headers))
    for r in range(2, 200):
        ws[f"A{r}"].number_format = "m/d/yyyy"
        ws[f"N{r}"] = f'=IF(G{r}="","",AND(G{r}>=\'Plan Config\'!$B$18,G{r}<=\'Plan Config\'!$B$19))'
    add_table(ws, "tbl_WoodReceiving", "A1:N2")
    dv_grade = DataValidation(type="list", formula1='"A,B,C"', allow_blank=True)
    dv_disp = DataValidation(type="list", formula1='"Accept,Reject,Hold"', allow_blank=True)
    dv_species = DataValidation(type="list", formula1='"Maple,Ebony,Rosewood,Other"', allow_blank=True)
    for dv, rng in ((dv_grade, "E2:E500"), (dv_disp, "J2:J500"), (dv_species, "D2:D500")):
        ws.add_data_validation(dv); dv.add(rng)
    ws["P1"] = ("Raw-wood receiving inspection (tracker QA1): grade A/B/C and moisture with a meter at receiving. Maple first; ebony and "
                "rosewood later. Imported maple needs the Lacey Act declaration reference (LOG1). Moisture accept range lives on Plan Config.")
    ws["P1"].font = MUTED
    widths(ws, {"A": 11, "B": 12, "C": 16, "D": 10, "E": 7, "F": 11, "G": 11, "H": 13, "I": 8, "J": 11, "K": 12, "L": 20, "M": 30, "N": 16, "P": 70})
    return ws


def build_humidity(wb: Workbook):
    ws = wb.create_sheet("Humidity Log")
    headers = ["Date/Time", "Location", "RH %", "Temp F", "In Range?", "Action Taken", "Logged By"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, 1, len(headers))
    for r in range(2, 400):
        ws[f"A{r}"].number_format = "m/d/yyyy h:mm"
        ws[f"E{r}"] = f'=IF(C{r}="","",IF(AND(C{r}>=\'Plan Config\'!$B$16,C{r}<=\'Plan Config\'!$B$17),"OK","EXCURSION"))'
    add_table(ws, "tbl_HumidityLog", "A1:G2")
    dv_loc = DataValidation(type="list", formula1='"Cue cell,Wood store,Side room"', allow_blank=True)
    ws.add_data_validation(dv_loc); dv_loc.add("B2:B500")
    ws.conditional_formatting.add("E2:E400", FormulaRule(formula=['$E2="EXCURSION"'], fill=PatternFill("solid", bgColor="F8CBAD"), font=Font(bold=True, color="9C0006")))
    ws["I1"] = ("Target 32-40 % RH (Plan Config). HIE logs humidity and does not act on it, and 85 % of HIE defects are warpage: "
                "this log exists so an excursion gets an action. QA3 sets how long an excursion is tolerable and when material goes on hold.")
    ws["I1"].font = MUTED
    widths(ws, {"A": 16, "B": 12, "C": 8, "D": 8, "E": 12, "F": 40, "G": 14, "I": 70})
    return ws


def build_classification(wb: Workbook):
    ws = wb.create_sheet("Cue Classification")
    ws["A1"] = "Cue Classification — finished-good SKU attribute master (placeholder until the SKU convention lands, tracker INV2)"
    ws["A1"].font = TITLE_FONT
    headers = ["Item", "Category", "Model token", "Confidence", "Note", "Source", "Design", "Construction"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=h)
    style_header(ws, 2, 1, len(headers))
    for i, row in enumerate(CLASSIFICATION, start=3):
        for j, v in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=v)
    widths(ws, {"A": 16, "B": 14, "C": 14, "D": 13, "E": 70, "F": 20, "G": 24, "H": 26})
    return ws


# ---------------------------------------------------------------------------
# Production Plan (mirrors TABLE: row 1 ISO week, row 2 Monday, row 3 capacity, row 4 totals, SKUs from row 5)
# ---------------------------------------------------------------------------
def build_production_plan(wb: Workbook):
    ws = wb.create_sheet("Production Plan")
    ws["A1"] = "Production Week"; ws["A2"] = "Cue"; ws["A3"] = "Est. Capacity"; ws["A4"] = "Weekly Totals"
    for c in ("A1", "A2", "A3", "A4"):
        ws[c].font = BOLD
    last_col = 1 + PLAN_WEEKS
    skus = [r[0] for r in CLASSIFICATION]
    first_sku_row, last_sku_row = 5, 5 + 60   # room for 60 SKUs, like TABLE's long list
    for i in range(PLAN_WEEKS):
        col = get_column_letter(2 + i)
        ws[f"{col}1"] = f"=_xlfn.ISOWEEKNUM({col}2)"
        ws[f"{col}2"] = "='Plan Config'!$Y$6-WEEKDAY('Plan Config'!$Y$6,3)" if i == 0 else f"={get_column_letter(1 + i)}2+7"
        ws[f"{col}2"].number_format = "d-mmm"
        ws[f"{col}3"] = f"=MIN('Plan Config'!$R${FIRST_ST_ROW}:$R${LAST_ST_ROW})"
        ws[f"{col}4"] = f"=SUM({col}{first_sku_row}:{col}{last_sku_row})"
        ws[f"{col}2"].font = BOLD
        ws.column_dimensions[col].width = 7.5
    style_header(ws, 2, 1, last_col)
    bp = "'Plan Config'!"
    for i, sku in enumerate(skus):
        r = first_sku_row + i
        ws[f"A{r}"] = sku
        for j in range(PLAN_WEEKS):
            col = get_column_letter(2 + j)
            ws[f"{col}{r}"] = (f"=SUMPRODUCT(({bp}$C$28:$C$31=$A{r})*({col}$2>={bp}$G$28:$G$31)*({col}$2<={bp}$H$28:$H$31)"
                               f"*({bp}$J$28:$J$31+({col}$2={bp}$H$28:$H$31)*{bp}$K$28:$K$31))")
    ws["A67"] = ("Planned starts per week per SKU, driven by the BATCH PLAN on Plan Config (good units x (1 + allowance)). "
                 "Row 3 is the weakest station's weekly capacity; row 4 should stay at or under it.")
    ws["A67"].font = MUTED
    ws.conditional_formatting.add(f"B4:{get_column_letter(last_col)}4",
                                  FormulaRule(formula=["B4>B3"], fill=PatternFill("solid", bgColor="F8CBAD"), font=Font(bold=True, color="9C0006")))
    ws.column_dimensions["A"].width = 18
    ws.freeze_panes = "B5"
    return ws


# ---------------------------------------------------------------------------
# Weekly Reporting and Summary (same formulas as TABLE, INDEX/MATCH instead of XLOOKUP)
# ---------------------------------------------------------------------------
PLAN_RNG = "'Production Plan'!$B$5:$AO$65"
PLAN_DATES = "'Production Plan'!$B$2:$AO$2"
CLASS_SKU = "'Cue Classification'!$A$3:$A$500"
CLASS_CAT = "'Cue Classification'!$B$3:$B$500"
REL = "'Release Log'!"
REJ = "'Reject Log'!"
REW = "'Rework Tracker'!"


def family_of(sku_rng: str) -> str:
    return f'IFERROR(INDEX({CLASS_CAT},MATCH({sku_rng},{CLASS_SKU},0)),"")'


def build_weekly_reporting(wb: Workbook):
    ws = wb.create_sheet("Weekly Reporting")
    ws["A1"] = "Weekly Reporting — Cues"; ws["A1"].font = TITLE_FONT
    ws["B2"] = "Start Date"; ws["C2"] = "End Date"
    ws["A3"] = "Date Range"; ws["B3"] = "=TODAY()-WEEKDAY(TODAY(),3)"; ws["C3"] = "=B3+6"
    ws["B3"].number_format = ws["C3"].number_format = "m/d/yyyy"
    heads = ["Family", "Forecast", "Actual", "Variance", "Variance %", "Reject", "Rework", "B-Grade", "First Pass Yield"]
    for i, h in enumerate(heads, start=1):
        ws.cell(row=5, column=i, value=h)
    style_header(ws, 5, 1, len(heads))
    fams = ["BROOMSTICK", "SELLABLE"]
    for i, fam in enumerate(fams):
        r = 6 + i
        ws[f"A{r}"] = fam
        ws[f"B{r}"] = f"=SUMPRODUCT(({CLASS_CAT}=$A{r})*({PLAN_RNG})*({PLAN_DATES}>=$B$3)*({PLAN_DATES}<=$C$3))".replace(
            f"({CLASS_CAT}=$A{r})", f"('Cue Classification'!$B$3:$B$63=$A{r})")
        ws[f"C{r}"] = f"=SUMPRODUCT(({REL}$A$2:$A${LAST_LOG_ROW}>=$B$3)*({REL}$A$2:$A${LAST_LOG_ROW}<=$C$3)*({family_of(REL + '$B$2:$B$' + str(LAST_LOG_ROW))}=$A{r})*{REL}$D$2:$D${LAST_LOG_ROW})"
        ws[f"D{r}"] = f"=$C{r}-$B{r}"
        ws[f"E{r}"] = f'=IF($B{r}=0,"",$D{r}/$B{r})'
        ws[f"F{r}"] = f"=SUMPRODUCT(({REL}$A$2:$A${LAST_LOG_ROW}>=$B$3)*({REL}$A$2:$A${LAST_LOG_ROW}<=$C$3)*({family_of(REL + '$B$2:$B$' + str(LAST_LOG_ROW))}=$A{r})*{REL}$H$2:$H${LAST_LOG_ROW})"
        ws[f"G{r}"] = f"=SUMPRODUCT(({REL}$A$2:$A${LAST_LOG_ROW}>=$B$3)*({REL}$A$2:$A${LAST_LOG_ROW}<=$C$3)*({family_of(REL + '$B$2:$B$' + str(LAST_LOG_ROW))}=$A{r})*{REL}$F$2:$F${LAST_LOG_ROW})"
        ws[f"H{r}"] = f"=SUMPRODUCT(({REL}$A$2:$A${LAST_LOG_ROW}>=$B$3)*({REL}$A$2:$A${LAST_LOG_ROW}<=$C$3)*({family_of(REL + '$B$2:$B$' + str(LAST_LOG_ROW))}=$A{r})*{REL}$G$2:$G${LAST_LOG_ROW})"
        ws[f"I{r}"] = f'=IF(($C{r}+$F{r}+$G{r}+$H{r})=0,"",$C{r}/($C{r}+$F{r}+$G{r}+$H{r}))'
        ws[f"E{r}"].number_format = ws[f"I{r}"].number_format = "0.0%"
    t = 6 + len(fams)
    ws[f"A{t}"] = "TOTAL"; ws[f"A{t}"].font = BOLD
    for col in "BCDFGH":
        ws[f"{col}{t}"] = f"=SUM({col}6:{col}{t-1})"
    ws[f"E{t}"] = f'=IF($B{t}=0,"",$D{t}/$B{t})'
    ws[f"I{t}"] = f'=IF(($C{t}+$F{t}+$G{t}+$H{t})=0,"",$C{t}/($C{t}+$F{t}+$G{t}+$H{t}))'
    ws[f"E{t}"].number_format = ws[f"I{t}"].number_format = "0.0%"
    # the Production Plan SKU column classification needs a row-aligned lookup; simplest robust form:
    for i, fam in enumerate(fams):
        r = 6 + i
        ws[f"B{r}"] = (f"=SUMPRODUCT((IFERROR(INDEX({CLASS_CAT},MATCH('Production Plan'!$A$5:$A$65,{CLASS_SKU},0)),\"\")=$A{r})"
                       f"*({PLAN_RNG})*({PLAN_DATES}>=$B$3)*({PLAN_DATES}<=$C$3))")
    ws["A11"] = ("Forecast = planned starts this week (Production Plan). Actual = Qty Released (Release Log). First Pass Yield = released / "
                 "(released + reject + rework + B-grade) for the week.")
    ws["A11"].font = MUTED
    widths(ws, {"A": 16, "B": 12, "C": 12, "D": 11, "E": 11, "F": 9, "G": 9, "H": 9, "I": 15})
    return ws


def build_summary(wb: Workbook):
    ws = wb.create_sheet("Summary")
    ws["B1"] = "Week Start Date"; ws["D1"] = "=TODAY()-WEEKDAY(TODAY(),3)"
    ws["B2"] = "Week Finish Date"; ws["D2"] = "=D1+6"
    ws["D1"].number_format = ws["D2"].number_format = "m/d/yyyy"
    ws["C3"] = "Week"; ws["I3"] = "Month to date"
    ws["C3"].font = ws["I3"].font = BOLD
    heads = ["Family", "Fcast", "Actual", "# Var", "% Var", "# Reject", "% Reject", "Fcast", "Actual", "# Var", "% Var", "# Reject", "% Reject"]
    for i, h in enumerate(heads, start=2):
        ws.cell(row=4, column=i, value=h)
    style_header(ws, 4, 2, 1 + len(heads))
    fams = ["BROOMSTICK", "SELLABLE"]
    rel_sku = REL + f"$B$2:$B${LAST_LOG_ROW}"
    rej_sku = REJ + f"$B$2:$B${LAST_LOG_ROW}"
    for i, fam in enumerate(fams):
        r = 5 + i
        ws[f"B{r}"] = fam
        plan_fam = f"(IFERROR(INDEX({CLASS_CAT},MATCH('Production Plan'!$A$5:$A$65,{CLASS_SKU},0)),\"\")=$B{r})"
        ws[f"C{r}"] = f"=SUMPRODUCT({plan_fam}*({PLAN_RNG})*({PLAN_DATES}>=$D$1)*({PLAN_DATES}<=$D$2))"
        ws[f"D{r}"] = f"=SUMPRODUCT(({REL}$A$2:$A${LAST_LOG_ROW}>=$D$1)*({REL}$A$2:$A${LAST_LOG_ROW}<=$D$2)*({family_of(rel_sku)}=$B{r})*{REL}$D$2:$D${LAST_LOG_ROW})"
        ws[f"E{r}"] = f"=D{r}-C{r}"
        ws[f"F{r}"] = f'=IF(C{r}=0,"",E{r}/C{r})'
        ws[f"G{r}"] = f"=SUMPRODUCT(({REJ}$A$2:$A${LAST_LOG_ROW}>=$D$1)*({REJ}$A$2:$A${LAST_LOG_ROW}<=$D$2)*({family_of(rej_sku)}=$B{r})*{REJ}$E$2:$E${LAST_LOG_ROW})"
        ws[f"H{r}"] = f'=IF(D{r}=0,"",G{r}/D{r})'
        ws[f"I{r}"] = f"=SUMPRODUCT({plan_fam}*({PLAN_RNG})*({PLAN_DATES}>=EOMONTH(TODAY(),-1)+1)*({PLAN_DATES}<=TODAY()))"
        ws[f"J{r}"] = f"=SUMPRODUCT(({REL}$A$2:$A${LAST_LOG_ROW}>=EOMONTH(TODAY(),-1)+1)*({REL}$A$2:$A${LAST_LOG_ROW}<=TODAY())*({family_of(rel_sku)}=$B{r})*{REL}$D$2:$D${LAST_LOG_ROW})"
        ws[f"K{r}"] = f"=J{r}-I{r}"
        ws[f"L{r}"] = f'=IF(I{r}=0,"",K{r}/I{r})'
        ws[f"M{r}"] = f"=SUMPRODUCT(({REJ}$A$2:$A${LAST_LOG_ROW}>=EOMONTH(TODAY(),-1)+1)*({REJ}$A$2:$A${LAST_LOG_ROW}<=TODAY())*({family_of(rej_sku)}=$B{r})*{REJ}$E$2:$E${LAST_LOG_ROW})"
        ws[f"N{r}"] = f'=IF(J{r}=0,"",M{r}/J{r})'
        for col in "FHLN":
            ws[f"{col}{r}"].number_format = "0.0%"
    t = 5 + len(fams)
    ws[f"B{t}"] = "TOTAL"; ws[f"B{t}"].font = BOLD
    for col in "CDEGIJKM":
        ws[f"{col}{t}"] = f"=SUM({col}5:{col}{t-1})"
    ws[f"F{t}"] = f'=IF(C{t}=0,"",E{t}/C{t})'; ws[f"H{t}"] = f'=IF(D{t}=0,"",G{t}/D{t})'
    ws[f"L{t}"] = f'=IF(I{t}=0,"",K{t}/I{t})'; ws[f"N{t}"] = f'=IF(J{t}=0,"",M{t}/J{t})'
    for col in "FHLN":
        ws[f"{col}{t}"].number_format = "0.0%"

    # Floor status block — what is happening right now
    ws["B10"] = "FLOOR STATUS (live)"; ws["B10"].font = BOLD
    fl = "'WOOD Floor'!"
    rows = [
        ("Cues on the board", f"=COUNTA({fl}$C$4:$C${LAST_FLOOR_ROW})"),
        ("Released", f'=COUNTIF({fl}${COL["QC Status"]}$4:${COL["QC Status"]}${LAST_FLOOR_ROW},"RELEASED")'),
        ("Rejected", f'=COUNTIF({fl}${COL["QC Status"]}$4:${COL["QC Status"]}${LAST_FLOOR_ROW},"REJECTED")'),
        ("In rework", f'=COUNTIF({fl}${COL["QC Status"]}$4:${COL["QC Status"]}${LAST_FLOOR_ROW},"REWORK")'),
        ("On hold", f'=COUNTIF({fl}${COL["Hold?"]}$4:${COL["Hold?"]}${LAST_FLOOR_ROW},"ON HOLD")'),
        ("Resting (rest gate not cleared)", f'=COUNTIF({fl}${COL["Rest OK?"]}$4:${COL["Rest OK?"]}${LAST_FLOOR_ROW},"REST*")'),
        ("Ready for QC (all stations ticked, no status)", f'=COUNTIFS({fl}${COL["Release?"]}$4:${COL["Release?"]}${LAST_FLOOR_ROW},TRUE,{fl}${COL["QC Status"]}$4:${COL["QC Status"]}${LAST_FLOOR_ROW},"")'),
        ("Not started", f'=COUNTIF({fl}${COL["Current Zone"]}$4:${COL["Current Zone"]}${LAST_FLOOR_ROW},"Not started")'),
    ]
    for i, (label, formula) in enumerate(rows):
        ws[f"B{11+i}"] = label
        ws[f"D{11+i}"] = formula
    widths(ws, {"A": 3, "B": 40, "C": 9, "D": 9, "E": 8, "F": 8, "G": 9, "H": 9, "I": 9, "J": 9, "K": 8, "L": 8, "M": 9, "N": 9})
    return ws


def build_readme(wb: Workbook):
    ws = wb.create_sheet("Read me", 0)
    ws["B1"] = "Wood_Production — Panthera Manufacturing cue butt line"; ws["B1"].font = TITLE_FONT
    rows = [
        ("Purpose", "Floor board, QC/release logging and weekly reporting for the cue butt cell, built to the same pattern as TABLE Production and REVO so the manager's Manufacturing_Lines workbook can read all three the same way."),
        ("Process", "Brandon's plain-butt scope grouped into seven stations WS01-WS07 (Plan Config). Station minutes are placeholders until Chad and Steven's setup sheets (P1.3)."),
        ("Rest gates", "Wood rests two weeks after rough turning (WS02) and after taper/coring (WS03). Enter the dates on WOOD Floor; 'Rest OK?' shows REST until <date> while the gate is closed."),
        ("Batches", "Plan Config BATCH PLAN holds the tracker's three builds: 1 prototype (Dec 18), 25 broomsticks (Jan 20), 50 sellable (Jun 11). Planned starts include the 10 % overrun allowance and flow to Production Plan."),
        ("WOOD Floor", "One row per serialised cue. Tick stations TRUE as they complete. Release? turns TRUE when every station is ticked. QC/release through the Wood QC button (VBA: modWoodQuality + frmWoodQuality, see repo wood_production/vba)."),
        ("QC", "The form captures serial, outcome (Release/B-Grade/Rework/Reject/Hold), station, free-text reason and notes. It writes QC Log, Release Log, Reject Log, Rework Tracker, Holds and learns reasons into QC Reasons. QC Status greys the row."),
        ("Wood Receiving", "QA1 receiving inspection: grade, moisture, RH at receipt, disposition, Lacey Act reference. Moisture accept range on Plan Config (placeholder until QA1)."),
        ("Humidity Log", "QA3: spot readings per location; EXCURSION flags against 32-40 % RH. Record the action taken."),
        ("Reporting", "Weekly Reporting and Summary read Production Plan (forecast) and Release Log / Reject Log (actual) by family via Cue Classification, exactly as TABLE does."),
        ("Placeholders", "SKUs PAN-BRM-001 / PAN-SELL-001 and the classification rows are placeholders until the SKU convention (INV2) and the sellable design (STR3) exist. Yellow cells are inputs."),
        ("Not included", "Daily shift scheduler, capacity planner and component tracker from TABLE: add when the cell runs more than one builder. Finish is external (Jacoby) until the booth gate (P3.2)."),
    ]
    for i, (k, v) in enumerate(rows, start=3):
        ws[f"B{i}"] = k; ws[f"B{i}"].font = BOLD
        ws[f"C{i}"] = v; ws[f"C{i}"].alignment = Alignment(wrap_text=True, vertical="top")
    widths(ws, {"A": 2, "B": 18, "C": 120})
    return ws


def build(out: Path) -> Path:
    wb = Workbook()
    build_plan_config(wb)
    build_floor(wb)
    build_holds(wb)
    build_log(wb, "Release Log", "tbl_ReleaseLog",
              ["Date", "SKU", "Unit ID", "Qty Released", "Work Order", "Qty Rework", "Qty B Grade", "Qty Reject",
               "Recalled", "Serial", "Lane (Design)", "Bucket (Category)"],
              extra_row2={
                  "K": f'=IF(B2="","",IFERROR(INDEX(\'Cue Classification\'!$G$3:$G$500,MATCH(B2,{CLASS_SKU},0)),"UNMAPPED"))',
                  "L": f'=IF(B2="","",IFERROR(INDEX({CLASS_CAT},MATCH(B2,{CLASS_SKU},0)),"UNMAPPED"))',
              })
    build_log(wb, "Reject Log", "tbl_RejectLog",
              ["Date", "SKU", "WO#", "Unit ID", "Quantity", "Description", "Serial", "Station"])
    build_log(wb, "Rework Tracker", "tbl_ReworkTracker",
              ["Date", "WO", "SKU", "Unit ID", "Qty", "Status", "Previous Release Status", *STATION_IDS, "Serial"])
    build_log(wb, "QC Log", "tbl_QCLog",
              ["Timestamp", "Date", "User", "WO#", "SKU", "Unit ID", "Serial", "Outcome", "Station", "Reason", "Notes",
               "Zone at QC", "Target Date", "Floor Row"], date_cols=("B", "M"))
    build_reasons(wb)
    build_receiving(wb)
    build_humidity(wb)
    build_classification(wb)
    build_production_plan(wb)
    build_weekly_reporting(wb)
    build_summary(wb)
    build_readme(wb)
    for ws in wb.worksheets:
        base_font(ws)
    # sheet order: Read me, WOOD Floor, Plan Config, Holds, logs..., Production Plan, Weekly Reporting, Summary
    order = ["Read me", "WOOD Floor", "Summary", "Weekly Reporting", "Production Plan", "Plan Config", "Holds",
             "Release Log", "Reject Log", "Rework Tracker", "QC Log", "QC Reasons", "Wood Receiving", "Humidity Log",
             "Cue Classification"]
    wb._sheets = [wb[name] for name in order]
    wb.active = 1
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("work/Wood_Production.xlsx")
    print(f"wrote {build(target)}")
