"""Build Manufacturing_Lines.xlsx — the manager's throughput & quality view across REVO, TABLE and WOOD.

The workbook does no data entry. Power Query (manufacturing_lines/queries/*.m) reads
each line's Release Log, Reject Log and Production Plan from its own workbook and
lands them on the Data_* sheets in one common shape; every number on Dashboard,
Weekly and Pareto is a SUMIFS / INDEX-MATCH over those sheets. Refresh All is the
only action.

Usage:
    python manufacturing_lines/tools/build_manufacturing_lines.py <out.xlsx>
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
SECTION_FILL = PatternFill("solid", fgColor="EDEDED")
TITLE_FONT = Font(name=FONT, size=14, bold=True)
BOLD = Font(name=FONT, bold=True)
NORMAL = Font(name=FONT)
MUTED = Font(name=FONT, italic=True, color="6E6E6E")
INPUT_FONT = Font(name=FONT, color="0000FF")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

LINES = ["REVO", "TABLE", "WOOD"]
# dataviz reference palette, categorical slots 1-3 (validated adjacent + all-pairs for three series)
LINE_COLOR = {"REVO": "2A78D6", "TABLE": "EB6834", "WOOD": "1BAF7A"}
PLAN_COLOR = "C3C2B7"
STATUS = {"good": "0CA30C", "warning": "FAB219", "critical": "D03B3B"}

TREND_WEEKS = 13

# Data sheet layouts — must match what the M queries emit (see queries/README)
REL_COLS = ["Line", "Date", "SKU", "Unit ID", "Qty Released", "Qty Rework", "Qty B Grade", "Qty Reject", "Serial", "Work Order"]
REJ_COLS = ["Line", "Date", "SKU", "WO", "Unit ID", "Quantity", "Description", "Serial", "Station"]
PLAN_COLS = ["Line", "Week Start", "SKU", "Planned Qty"]
PARETO_COLS = ["Line", "Reason", "Qty", "Rank", "Key"]

# column letters on the Data sheets
R = {h: get_column_letter(i) for i, h in enumerate(REL_COLS, start=1)}
J = {h: get_column_letter(i) for i, h in enumerate(REJ_COLS, start=1)}
P = {h: get_column_letter(i) for i, h in enumerate(PLAN_COLS, start=1)}
Q = {h: get_column_letter(i) for i, h in enumerate(PARETO_COLS, start=1)}

# Config cells
CFG_WTD, CFG_MTD, CFG_YTD, CFG_END = "Config!$B$21", "Config!$B$22", "Config!$B$23", "Config!$B$24"
T_ATTAIN, T_FPY, T_REJECT, T_REWORK = "Config!$B$14", "Config!$B$15", "Config!$B$16", "Config!$B$17"


def style_header(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = BOLD; cell.fill = HDR_FILL; cell.border = BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def set_input(cell, value, fmt=None):
    cell.value = value; cell.fill = INPUT_FILL; cell.font = INPUT_FONT
    if fmt:
        cell.number_format = fmt


def sumifs(sheet, sum_col, line_ref, date_col, start_ref, end_ref):
    return (f"SUMIFS({sheet}!${sum_col}:${sum_col},{sheet}!$A:$A,{line_ref},"
            f"{sheet}!${date_col}:${date_col},\">=\"&{start_ref},{sheet}!${date_col}:${date_col},\"<=\"&{end_ref})")


# ---------------------------------------------------------------------------
def build_config(wb):
    ws = wb.create_sheet("Config")
    ws["A1"] = "Manufacturing Lines — configuration"; ws["A1"].font = TITLE_FONT
    ws["A2"] = "Yellow cells are inputs. Paths are read by Power Query (table 'Lines'); targets drive the Dashboard flags."; ws["A2"].font = MUTED
    ws["A3"] = "As of"; ws["B3"] = "=TODAY()"; ws["B3"].number_format = "ddd d mmm yyyy"

    ws["A5"] = "LINE WORKBOOKS  (table 'Lines' — Power Query reads this)"; ws["A5"].font = BOLD
    heads = ["Line", "Workbook path", "Enabled", "Layout", "Notes"]
    for i, h in enumerate(heads, start=1):
        ws.cell(row=6, column=i, value=h)
    style_header(ws, 6, 1, len(heads))
    rows = [
        ("REVO", r"C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\REVO Production.xlsm", True, "REVO",
         "Plain sheets: Release Log (Date, SKU, Cart, ReleasedQty, WO, Rework, B Grade, Reject), Reject Log, Production Plan. Confirm the path and that its Release Log dates are text m/d/yyyy."),
        ("TABLE", r"C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Table_Production.xlsm", True, "TABLE",
         "Tables tbl_ReleaseLog / tbl_RejectLog, Production Plan grid."),
        ("WOOD", r"C:\Users\CharlesBass\OneDrive - Predator Group\Supply Chain Files\Wood_Production.xlsm", True, "TABLE",
         "Same table layout as TABLE (built from the same pattern)."),
    ]
    for i, (line, path, enabled, layout, note) in enumerate(rows, start=7):
        ws.cell(row=i, column=1, value=line).font = BOLD
        set_input(ws.cell(row=i, column=2), path)
        set_input(ws.cell(row=i, column=3), enabled)
        set_input(ws.cell(row=i, column=4), layout)
        ws.cell(row=i, column=5, value=note).font = MUTED
    t = Table(displayName="Lines", ref="A6:E9")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    ws["A10"] = "Layout = REVO (plain sheets, text dates) or TABLE (tbl_* tables, real dates). Set Enabled = FALSE to drop a line from the refresh."
    ws["A10"].font = MUTED

    ws["A12"] = "TARGETS"; ws["A12"].font = BOLD
    ws["A13"] = "Metric"; ws["B13"] = "Target"; style_header(ws, 13, 1, 2)
    targets = [("Attainment (released / planned) at least", 0.95), ("First pass yield at least", 0.90),
               ("Reject % at most", 0.05), ("Rework % at most", 0.10)]
    for i, (label, val) in enumerate(targets, start=14):
        ws.cell(row=i, column=1, value=label)
        set_input(ws.cell(row=i, column=2), val, "0%")

    ws["A19"] = "PERIODS (computed; weeks run Monday to Sunday)"; ws["A19"].font = BOLD
    ws["A20"] = "Period"; ws["B20"] = "Start"; style_header(ws, 20, 1, 2)
    ws["A21"] = "Week to date";  ws["B21"] = "=TODAY()-WEEKDAY(TODAY(),3)"
    ws["A22"] = "Month to date"; ws["B22"] = "=EOMONTH(TODAY(),-1)+1"
    ws["A23"] = "Year to date";  ws["B23"] = "=DATE(YEAR(TODAY()),1,1)"
    ws["A24"] = "Period end";    ws["B24"] = "=TODAY()"
    for r in range(21, 25):
        ws[f"B{r}"].number_format = "m/d/yyyy"

    ws["A27"] = "DEFINITIONS"; ws["A27"].font = BOLD
    defs = [
        ("Planned", "Sum of Production Plan starts for weeks whose Monday falls in the period."),
        ("Released", "Sum of Qty Released in the line's Release Log."),
        ("First pass yield", "Released / (Released + Rework + B-Grade + Reject) over the period. Units that needed anything other than a straight release count against it."),
        ("Reject %", "Reject / (Released + Rework + B-Grade + Reject). Rework % likewise."),
        ("Attainment", "Released / Planned. Blank when nothing was planned."),
        ("Flags", "● on target, ▲ within 10 points (throughput) or above reject target but under 2x (quality), ■ beyond that. Icon and word together, never colour alone."),
    ]
    for i, (k, v) in enumerate(defs, start=28):
        ws[f"A{i}"] = k; ws[f"A{i}"].font = BOLD; ws[f"B{i}"] = v
    widths(ws, {"A": 42, "B": 70, "C": 9, "D": 9, "E": 90})
    return ws


def build_data_sheet(wb, name, cols, note):
    ws = wb.create_sheet(name)
    for i, h in enumerate(cols, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, 1, len(cols))
    ws.cell(row=1, column=len(cols) + 2, value=note).font = MUTED
    for i in range(1, len(cols) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 14
    ws.column_dimensions[get_column_letter(len(cols) + 2)].width = 90
    return ws


# ---------------------------------------------------------------------------
def build_weekly(wb):
    ws = wb.create_sheet("Weekly")
    ws["A1"] = f"Weekly trend — last {TREND_WEEKS} weeks by line (feeds the Dashboard charts)"; ws["A1"].font = TITLE_FONT
    metrics = ["Planned", "Released", "Rework", "B-Grade", "Reject", "FPY", "Reject %"]
    ws["A4"] = "Week start"; style_header(ws, 4, 1, 1)
    first_row, last_row = 5, 5 + TREND_WEEKS - 1
    for i in range(TREND_WEEKS):
        r = first_row + i
        ws[f"A{r}"] = f"={CFG_WTD}-7*{TREND_WEEKS - 1 - i}"
        ws[f"A{r}"].number_format = "d-mmm"
    col = 2
    blocks = {}
    for line in LINES:
        c1 = col
        ws.cell(row=3, column=c1, value=line).font = BOLD
        ws.cell(row=3, column=c1).fill = PatternFill("solid", fgColor=LINE_COLOR[line])
        ws.cell(row=3, column=c1).font = Font(name=FONT, bold=True, color="FFFFFF")
        ws.merge_cells(start_row=3, start_column=c1, end_row=3, end_column=c1 + len(metrics) - 1)
        for j, m in enumerate(metrics):
            ws.cell(row=4, column=c1 + j, value=m)
        style_header(ws, 4, c1, c1 + len(metrics) - 1)
        L = {m: get_column_letter(c1 + j) for j, m in enumerate(metrics)}
        blocks[line] = L
        for i in range(TREND_WEEKS):
            r = first_row + i
            s, e = f"$A{r}", f"$A{r}+6"
            lq = f'"{line}"'
            ws[f"{L['Planned']}{r}"] = "=" + sumifs("Data_Plan", P["Planned Qty"], lq, P["Week Start"], s, e)
            ws[f"{L['Released']}{r}"] = "=" + sumifs("Data_Releases", R["Qty Released"], lq, R["Date"], s, e)
            ws[f"{L['Rework']}{r}"] = "=" + sumifs("Data_Releases", R["Qty Rework"], lq, R["Date"], s, e)
            ws[f"{L['B-Grade']}{r}"] = "=" + sumifs("Data_Releases", R["Qty B Grade"], lq, R["Date"], s, e)
            ws[f"{L['Reject']}{r}"] = "=" + sumifs("Data_Releases", R["Qty Reject"], lq, R["Date"], s, e)
            tot = f"({L['Released']}{r}+{L['Rework']}{r}+{L['B-Grade']}{r}+{L['Reject']}{r})"
            ws[f"{L['FPY']}{r}"] = f'=IF({tot}=0,"",{L["Released"]}{r}/{tot})'
            ws[f"{L['Reject %']}{r}"] = f'=IF({tot}=0,"",{L["Reject"]}{r}/{tot})'
            ws[f"{L['FPY']}{r}"].number_format = ws[f"{L['Reject %']}{r}"].number_format = "0.0%"
        col += len(metrics) + 1
    ws[f"A{last_row + 2}"] = "Each week is Monday to Sunday. The last row is the current week to date."
    ws[f"A{last_row + 2}"].font = MUTED
    ws.column_dimensions["A"].width = 11
    for c in range(2, col):
        ws.column_dimensions[get_column_letter(c)].width = 9
    ws.freeze_panes = "B5"
    return ws, blocks, first_row, last_row


def build_pareto(wb):
    ws = wb.create_sheet("Pareto")
    ws["A1"] = f"Reject reasons — last {TREND_WEEKS} weeks, by line (from Data_Pareto)"; ws["A1"].font = TITLE_FONT
    ws["A2"] = "Why it happened. The reason text is what the QC form captured; once the lines categorise reasons, group here by category."; ws["A2"].font = MUTED
    col = 1
    for line in LINES:
        c1 = col
        ws.cell(row=4, column=c1, value=line).font = Font(name=FONT, bold=True, color="FFFFFF")
        ws.cell(row=4, column=c1).fill = PatternFill("solid", fgColor=LINE_COLOR[line])
        ws.merge_cells(start_row=4, start_column=c1, end_row=4, end_column=c1 + 3)
        for j, h in enumerate(["Rank", "Reason", "Qty", "Share"]):
            ws.cell(row=5, column=c1 + j, value=h)
        style_header(ws, 5, c1, c1 + 3)
        rank_c, reason_c, qty_c, share_c = (get_column_letter(c1 + k) for k in range(4))
        tot = f'SUMIFS(Data_Pareto!${Q["Qty"]}:${Q["Qty"]},Data_Pareto!$A:$A,"{line}")'
        for k in range(1, 11):
            r = 5 + k
            ws[f"{rank_c}{r}"] = k
            key = f'"{line}|"&{rank_c}{r}'
            ws[f"{reason_c}{r}"] = f'=IFERROR(INDEX(Data_Pareto!${Q["Reason"]}:${Q["Reason"]},MATCH({key},Data_Pareto!${Q["Key"]}:${Q["Key"]},0)),"")'
            ws[f"{qty_c}{r}"] = f'=IFERROR(INDEX(Data_Pareto!${Q["Qty"]}:${Q["Qty"]},MATCH({key},Data_Pareto!${Q["Key"]}:${Q["Key"]},0)),"")'
            ws[f"{share_c}{r}"] = f'=IF(OR({qty_c}{r}="",{tot}=0),"",{qty_c}{r}/{tot})'
            ws[f"{share_c}{r}"].number_format = "0%"
        ws.column_dimensions[rank_c].width = 6
        ws.column_dimensions[reason_c].width = 38
        ws.column_dimensions[qty_c].width = 7
        ws.column_dimensions[share_c].width = 7
        col += 5
    return ws


# ---------------------------------------------------------------------------
def kpi_block(ws, top, title, start_ref):
    ws.cell(row=top, column=2, value=title).font = BOLD
    for c in range(2, 14):
        ws.cell(row=top, column=c).fill = SECTION_FILL
    heads = ["Line", "Planned", "Released", "Rework", "B-Grade", "Reject", "Attainment", "First pass yield",
             "Reject %", "Rework %", "Throughput", "Quality"]
    for i, h in enumerate(heads, start=2):
        ws.cell(row=top + 1, column=i, value=h)
    style_header(ws, top + 1, 2, 1 + len(heads))
    first = top + 2
    for i, line in enumerate(LINES + ["TOTAL"]):
        r = first + i
        ws[f"B{r}"] = line; ws[f"B{r}"].font = BOLD
        if line != "TOTAL":
            lq = f"$B{r}"
            ws[f"C{r}"] = "=" + sumifs("Data_Plan", P["Planned Qty"], lq, P["Week Start"], start_ref, CFG_END)
            ws[f"D{r}"] = "=" + sumifs("Data_Releases", R["Qty Released"], lq, R["Date"], start_ref, CFG_END)
            ws[f"E{r}"] = "=" + sumifs("Data_Releases", R["Qty Rework"], lq, R["Date"], start_ref, CFG_END)
            ws[f"F{r}"] = "=" + sumifs("Data_Releases", R["Qty B Grade"], lq, R["Date"], start_ref, CFG_END)
            ws[f"G{r}"] = "=" + sumifs("Data_Releases", R["Qty Reject"], lq, R["Date"], start_ref, CFG_END)
        else:
            for col in "CDEFG":
                ws[f"{col}{r}"] = f"=SUM({col}{first}:{col}{r-1})"
        tot = f"($D{r}+$E{r}+$F{r}+$G{r})"
        ws[f"H{r}"] = f'=IF($C{r}=0,"",$D{r}/$C{r})'
        ws[f"I{r}"] = f'=IF({tot}=0,"",$D{r}/{tot})'
        ws[f"J{r}"] = f'=IF({tot}=0,"",$G{r}/{tot})'
        ws[f"K{r}"] = f'=IF({tot}=0,"",$E{r}/{tot})'
        ws[f"L{r}"] = (f'=IF($H{r}="","—",IF($H{r}>={T_ATTAIN},"● On plan",IF($H{r}>={T_ATTAIN}-0.1,"▲ Watch","■ Behind")))')
        ws[f"M{r}"] = (f'=IF($I{r}="","—",IF(AND($I{r}>={T_FPY},$J{r}<={T_REJECT},$K{r}<={T_REWORK}),"● Good",'
                       f'IF(OR($J{r}>2*{T_REJECT},$I{r}<{T_FPY}-0.1),"■ Act","▲ Watch")))')
        for col in "HIJK":
            ws[f"{col}{r}"].number_format = "0.0%"
        if line == "TOTAL":
            for col in "BCDEFGHIJKLM":
                ws[f"{col}{r}"].font = BOLD
    return first + len(LINES) + 1


def add_flag_formats(ws, rng):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({rng.split(":")[0]},1)="●"'],
                                  fill=PatternFill("solid", bgColor=STATUS["good"]), font=Font(bold=True, color="FFFFFF")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({rng.split(":")[0]},1)="▲"'],
                                  fill=PatternFill("solid", bgColor=STATUS["warning"]), font=Font(bold=True, color="000000")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({rng.split(":")[0]},1)="■"'],
                                  fill=PatternFill("solid", bgColor=STATUS["critical"]), font=Font(bold=True, color="FFFFFF")))


def build_dashboard(wb, weekly_blocks, w_first, w_last):
    ws = wb.create_sheet("Dashboard", 0)
    ws["B1"] = "Manufacturing Lines — throughput & quality"; ws["B1"].font = TITLE_FONT
    ws["B2"] = '="REVO shafts  ·  TABLE  ·  WOOD cues   |   as of "&TEXT(Config!$B$3,"ddd d mmm yyyy")&"   |   Data > Refresh All to update"'
    ws["B2"].font = MUTED

    nxt = kpi_block(ws, 4, "WEEK TO DATE", CFG_WTD)
    nxt = kpi_block(ws, nxt + 1, "MONTH TO DATE", CFG_MTD)
    nxt = kpi_block(ws, nxt + 1, "YEAR TO DATE", CFG_YTD)
    add_flag_formats(ws, f"L6:L{nxt}")
    add_flag_formats(ws, f"M6:M{nxt}")

    ws[f"B{nxt + 1}"] = f"TREND — last {TREND_WEEKS} weeks (Weekly sheet)"; ws[f"B{nxt + 1}"].font = BOLD
    for c in range(2, 14):
        ws.cell(row=nxt + 1, column=c).fill = SECTION_FILL

    wsk = wb["Weekly"]
    cats = Reference(wsk, min_col=1, min_row=w_first, max_row=w_last)
    anchor_row = nxt + 3
    anchors = {"REVO": f"B{anchor_row}", "TABLE": f"G{anchor_row}", "WOOD": f"L{anchor_row}"}
    for line in LINES:
        L = weekly_blocks[line]
        chart = BarChart()
        chart.type = "col"; chart.grouping = "clustered"; chart.gapWidth = 60; chart.overlap = -10
        chart.title = f"{line} — released vs planned"
        chart.y_axis.title = "Units"; chart.x_axis.number_format = "d-mmm"
        chart.y_axis.majorGridlines = None
        chart.y_axis.scaling.min = 0                  # bars anchor at the baseline, always
        chart.legend.position = "b"
        for metric, color in (("Planned", PLAN_COLOR), ("Released", LINE_COLOR[line])):
            col_idx = wsk[f"{L[metric]}{w_first}"].column
            ref = Reference(wsk, min_col=col_idx, min_row=w_first - 1, max_row=w_last)
            chart.add_data(ref, titles_from_data=True)
            s = chart.series[-1]
            s.graphicalProperties = GraphicalProperties(solidFill=color)
            s.graphicalProperties.line.noFill = True
        chart.set_categories(cats)
        chart.width = 8.4; chart.height = 7
        ws.add_chart(chart, anchors[line])

    quality = LineChart()
    quality.title = f"Reject % by line — last {TREND_WEEKS} weeks"
    quality.y_axis.title = "Reject %"; quality.y_axis.number_format = "0%"; quality.y_axis.majorGridlines = None
    quality.y_axis.scaling.min = 0
    quality.x_axis.number_format = "d-mmm"
    quality.legend.position = "b"
    for line in LINES:
        L = weekly_blocks[line]
        col_idx = wsk[f"{L['Reject %']}{w_first}"].column
        ref = Reference(wsk, min_col=col_idx, min_row=w_first - 1, max_row=w_last)
        quality.add_data(ref, titles_from_data=True)
        s = quality.series[-1]
        s.graphicalProperties.line.solidFill = LINE_COLOR[line]
        s.graphicalProperties.line.width = 25000      # EMU ≈ 2px
        s.marker.symbol = "circle"; s.marker.size = 6
        s.marker.graphicalProperties = GraphicalProperties(solidFill=LINE_COLOR[line])
        s.smooth = False
    # series names come from row 4 headers, which read "Reject %" for every line — rename to the line
    for s, line in zip(quality.series, LINES):
        from openpyxl.chart.series import SeriesLabel
        s.tx = SeriesLabel(v=line)
    quality.set_categories(cats)
    quality.width = 23.5; quality.height = 7
    ws.add_chart(quality, f"B{anchor_row + 15}")

    fpy = LineChart()
    fpy.title = f"First pass yield by line — last {TREND_WEEKS} weeks"
    fpy.y_axis.title = "FPY"; fpy.y_axis.number_format = "0%"; fpy.y_axis.majorGridlines = None
    fpy.y_axis.scaling.min = 0; fpy.y_axis.scaling.max = 1
    fpy.x_axis.number_format = "d-mmm"; fpy.legend.position = "b"
    for line in LINES:
        L = weekly_blocks[line]
        col_idx = wsk[f"{L['FPY']}{w_first}"].column
        ref = Reference(wsk, min_col=col_idx, min_row=w_first - 1, max_row=w_last)
        fpy.add_data(ref, titles_from_data=True)
        s = fpy.series[-1]
        s.graphicalProperties.line.solidFill = LINE_COLOR[line]
        s.graphicalProperties.line.width = 25000
        s.marker.symbol = "circle"; s.marker.size = 6
        s.marker.graphicalProperties = GraphicalProperties(solidFill=LINE_COLOR[line])
        s.smooth = False
    for s, line in zip(fpy.series, LINES):
        from openpyxl.chart.series import SeriesLabel
        s.tx = SeriesLabel(v=line)
    fpy.set_categories(cats)
    fpy.width = 23.5; fpy.height = 7
    ws.add_chart(fpy, f"B{anchor_row + 30}")

    # Series names for the bar charts: "Planned"/"Released" headers are right as they are.
    widths(ws, {"A": 2, "B": 10, "C": 9, "D": 10, "E": 9, "F": 9, "G": 9, "H": 11, "I": 14, "J": 9, "K": 9, "L": 13, "M": 11})
    ws.freeze_panes = "B4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def build_readme(wb):
    ws = wb.create_sheet("Read me", 0)
    ws["B1"] = "Manufacturing_Lines — one view of REVO, TABLE and WOOD"; ws["B1"].font = TITLE_FONT
    rows = [
        ("What it answers", "What happened (released vs planned, WTD / MTD / YTD), why (reject Pareto), the risk (FPY and reject % trend) and what to do (flags against the targets on Config)."),
        ("How it gets data", "Power Query. The queries in the repo folder manufacturing_lines/queries read each line's Release Log, Reject Log and Production Plan straight from the line workbooks on OneDrive and land them on Data_Releases, Data_Rejects, Data_Plan and Data_Pareto. Nobody types here."),
        ("Refresh", "Data > Refresh All. Set the three workbook paths on Config first. A missing file fails the refresh on purpose rather than showing a line as zero."),
        ("Install (once)", "Data > Get Data > From Other Sources > Blank Query > Advanced Editor: paste each .m file, name the query exactly as the file. fnLine* queries: Close & Load To > Only Create Connection. Releases, Rejects, Plan, RejectPareto: Close & Load To > Existing worksheet, cell A1 of the matching Data_ sheet (accept the overwrite of the header row)."),
        ("Lines", "REVO = carbon shafts (cart-based release flow). TABLE = pool tables (one row per serialised table). WOOD = Panthera cue butts (one row per serialised cue)."),
        ("Weeks", "Monday to Sunday. Planned counts the weeks whose Monday falls in the period, so a week counts fully in the period it starts in."),
        ("Pareto", "Top ten reject reasons per line over the last 13 weeks. Reason text is what operators typed; once each line fills 'Category (later)' on its QC Reasons sheet, switch the grouping to category."),
        ("Known limits", "REVO's Reject Log was written with SKU and WO swapped by the old macro; the REVO query un-swaps them, verify against the live file. Charts show the last 13 weeks only. No per-SKU view here; that lives in each line's workbook."),
    ]
    for i, (k, v) in enumerate(rows, start=3):
        ws[f"B{i}"] = k; ws[f"B{i}"].font = BOLD
        ws[f"C{i}"] = v; ws[f"C{i}"].alignment = Alignment(wrap_text=True, vertical="top")
    widths(ws, {"A": 2, "B": 18, "C": 120})
    return ws


def add_sample_rows(wb) -> None:
    """Test data only (--sample): proves the formulas and charts against known numbers. Never ship this."""
    from datetime import date, timedelta
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    rel = wb["Data_Releases"]; rej = wb["Data_Rejects"]; plan = wb["Data_Plan"]; par = wb["Data_Pareto"]
    per_week = {"REVO": (240, 6, 4, 10), "TABLE": (8, 1, 0, 1), "WOOD": (9, 2, 0, 1)}   # released, rework, bgrade, reject
    planned = {"REVO": 250, "TABLE": 8, "WOOD": 9}
    for w in range(TREND_WEEKS):
        wk = monday - timedelta(days=7 * (TREND_WEEKS - 1 - w))
        for line, (r_, rw, bg, rj) in per_week.items():
            plan.append([line, wk, f"{line}-SKU", planned[line]])
            rel.append([line, wk + timedelta(days=2), f"{line}-SKU", f"U-{w:03d}", r_ + (w % 3), rw, bg, rj + (1 if w == TREND_WEEKS - 1 else 0), f"SN{w}", "WO1"])
            if rj:
                rej.append([line, wk + timedelta(days=2), f"{line}-SKU", "WO1", f"U-{w:03d}", rj, "Warpage — moved after cut" if w % 2 else "Chatter", f"SN{w}", "WS03"])
    rank = {}
    for line, reasons in {"REVO": [("Tip glue", 40), ("Ferrule crack", 22)], "TABLE": [("Slate chip", 5), ("Rail misfit", 3)],
                          "WOOD": [("Warpage", 7), ("Chatter", 6), ("Pin dot", 1)]}.items():
        for i, (reason, qty) in enumerate(reasons, start=1):
            par.append([line, reason, qty, i, f"{line}|{i}"])


def build(out: Path, sample: bool = False) -> Path:
    wb = Workbook()
    wb.remove(wb.active)
    build_config(wb)
    build_data_sheet(wb, "Data_Releases", REL_COLS, "Loaded by Power Query 'Releases' (load to this sheet, cell A1). One row per Release Log line from each line workbook.")
    build_data_sheet(wb, "Data_Rejects", REJ_COLS, "Loaded by Power Query 'Rejects' (cell A1). One row per Reject Log line.")
    build_data_sheet(wb, "Data_Plan", PLAN_COLS, "Loaded by Power Query 'Plan' (cell A1). Production Plan grids unpivoted to Line x Week Start x SKU.")
    build_data_sheet(wb, "Data_Pareto", PARETO_COLS, "Loaded by Power Query 'RejectPareto' (cell A1). Reject qty by Line x Reason, last 13 weeks, ranked within line. Key = Line|Rank.")
    wsk, blocks, w_first, w_last = build_weekly(wb)
    build_pareto(wb)
    build_dashboard(wb, blocks, w_first, w_last)
    build_readme(wb)
    if sample:
        add_sample_rows(wb)
    order = ["Read me", "Dashboard", "Weekly", "Pareto", "Config", "Data_Releases", "Data_Rejects", "Data_Plan", "Data_Pareto"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 1
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None and cell.font == Font():
                    cell.font = NORMAL
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    target = Path(args[0]) if args else Path("work/Manufacturing_Lines.xlsx")
    print(f"wrote {build(target, sample='--sample' in sys.argv)}")
