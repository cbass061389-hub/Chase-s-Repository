"""Build the JAX inbound launch-allocation workbook from an HIE packing list
and the EMEA Q4 allocation sheet.

Usage:
    python tools/jax_shipment_allocation.py <packing_list.xlsx> <emea_alloc.xlsx> <out.xlsx>

Rule (Jacksonville-bound stock):
  * ASIA / AP allocations and EMEA *Direct* allocations are NOT served from
    this stock, so they are excluded. Only EMEA-via-JAX plus the other
    JAX-fulfilled channels (B2C, Americas, sponsorship / tradeshow / marketing)
    are allocated here.
  * EMEA to ship from JAX = MIN(arriving, EMEA via JAX - already sent).
  * Everything else arriving stays for the JAX-fulfilled regions.
  * Days remaining = launch date - as-of date; ship-by = launch - EMEA transit days.

SKUs are matched on a token-sorted key so "P3 MAP TEAL NW RAD" and
"P3 MAP TEAL RAD NW" line up. Samples on the packing list are listed but not
allocated. All derived numbers are live formulas; yellow cells are inputs.
"""
from __future__ import annotations

import datetime as dt
import re
import sys
import warnings
from collections import OrderedDict

import openpyxl
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.table import Table, TableStyleInfo

warnings.filterwarnings("ignore")

INK, GOLD, BLUE, GREY = "FF1C1610", "FFC8922A", "FF2E75B6", "FF595959"
INPUT = PatternFill("solid", fgColor="FFFFF2CC")
HDR = PatternFill("solid", fgColor="FF44546A")
SEC = PatternFill("solid", fgColor=GOLD)
RED = PatternFill("solid", fgColor="FFFBE3E0")
AMB = PatternFill("solid", fgColor="FFFFF0D4")
GRN = PatternFill("solid", fgColor="FFE2F0E4")
THIN = Side(style="thin", color="FFD9D9D9")
NUM = "#,##0;\\(#,##0\\);\\-"
DATE = "m/d/yyyy"
F = lambda **k: Font(name="Arial", size=k.pop("size", 10), **k)


def key(s) -> str:
    return " ".join(sorted(re.sub(r"\s+", " ", str(s or "").strip().upper()).split()))


def item_type(item: str, po, remark: str) -> str:
    po = str(po or "")
    if "sample" in (remark or "").lower() or po.startswith("TCY") or "SAMPLE" in item.upper() or "FOREARM" in item.upper():
        return "Sample"
    if item.upper().startswith("S "):
        return "Shaft"
    if item.split(" ")[0].upper() in ("BCP", "BCC", "CP", "BC8"):
        return "Cue"
    return "Other"


def read_packing_list(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    hdr = {str(c.value).strip(): c.column for c in ws[10] if c.value}
    meta = {}
    for r in ws.iter_rows(min_row=1, max_row=9, values_only=True):
        for v in r:
            if isinstance(v, str):
                for lab in ("DELIVERY DATE:", "REF.NO.:", "VESSEL/VOYAGE:", "CUSTOMER:", "FROM:"):
                    if v.strip().upper().startswith(lab):
                        meta[lab.rstrip(":")] = v.split(":", 1)[1].strip()
    rows = []
    for i, r in enumerate(ws.iter_rows(min_row=11, values_only=True), start=11):
        g = lambda name: r[hdr[name] - 1] if name in hdr else None
        item, qty = g("ITEM"), g("QTY")
        if item is None or not isinstance(qty, (int, float)):
            if isinstance(item, str) and item.strip().upper().startswith("TOTAL"):
                break
            continue
        item = re.sub(r"\s+", " ", str(item).strip())
        remark = str(g("DESCRIPTION") or "").strip()
        rows.append(dict(line=i, po=g("PO_NUMBER"), item=item, qty=qty, remark=remark, ctns=g("CTNS_TOTAL") or 0,
                         net=g("NET_WEIGHT_TOTAL") or 0, gross=g("GROSS_WEIGHT_TOTAL") or 0, cuft=g("CUFT_TOTAL") or 0,
                         type=item_type(item, g("PO_NUMBER"), remark), key=key(item)))
    m = re.search(r"([A-Za-z]{3})\.?\s*(\d{1,2})\.?\s*(\d{4})", meta.get("DELIVERY DATE", ""))
    delivery = dt.datetime.strptime(" ".join(m.groups()), "%b %d %Y").date() if m else None
    return rows, meta, delivery


def read_emea(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    hdr = [str(c.value).strip() for c in ws[1]]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(hdr, r))
        if not d.get("Item"):
            continue
        item = re.sub(r"\s+", " ", str(d["Item"]).strip())
        out.append(dict(project=d.get("Project"), item=item, launch=d.get("Launch Date"), emea=d.get("EMEA") or 0,
                        direct=d.get("Direct") or 0, jax=d.get("Jax") or 0, type=item_type(item, "", ""), key=key(item)))
    return out


def hdr_row(ws, row, headers, widths=None, fill=HDR, height=32):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row, c, h)
        cell.font = F(bold=True, color="FFFFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if widths:
            ws.column_dimensions[L(c)].width = widths[c - 1]
    ws.row_dimensions[row].height = height


def title(ws, text, sub, cols):
    ws["A1"] = text
    ws["A1"].font = F(size=14, bold=True, color="FFFFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=INK)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
    ws["A2"] = sub
    ws["A2"].font = F(size=9, color=GREY)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=cols)
    ws.row_dimensions[2].height = 30


def add_table(ws, name, ref):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)


def build(pl_path, emea_path, out):
    lines, meta, delivery = read_packing_list(pl_path)
    emea = read_emea(emea_path)
    ref = meta.get("REF.NO", "").strip() or "shipment"
    wb = openpyxl.Workbook()

    # ---------------- Rules & Inputs -------------------------------------------
    ru = wb.active
    ru.title = "Rules & Inputs"
    title(ru, f"Rules & Inputs — JAX inbound allocation, packing list {ref}",
          "Yellow cells are the only inputs on this tab. Every number on the Allocation tab recalculates from them.", 3)
    ru.column_dimensions["A"].width = 34; ru.column_dimensions["B"].width = 16; ru.column_dimensions["C"].width = 95
    inputs = [
        ("Delivery date at JAX", delivery, DATE, f"From the packing list header (DELIVERY DATE). Carrier: {meta.get('VESSEL/VOYAGE', '')}."),
        ("As-of date", "=TODAY()", DATE, "Days-to-launch and days-left use this. Overwrite with a fixed date to freeze the view."),
        ("EMEA transit lead time (days)", 7, "0", "ASSUMPTION — days needed to get stock from JAX to the EMEA warehouse before launch. Ship-by date = launch date − this."),
        ("Urgent threshold (days)", 7, "0", "Days-left at or below this flags URGENT (amber). Negative = overdue (red)."),
    ]
    hdr_row(ru, 4, ["Setting", "Value", "Note"], height=20)
    for i, (lab, val, nf, note) in enumerate(inputs, start=5):
        ru.cell(i, 1, lab).font = F()
        c = ru.cell(i, 2, val); c.font = F(color="FF0000FF"); c.fill = INPUT; c.number_format = nf
        ru.cell(i, 3, note).font = F(size=9, color=GREY)
        ru.cell(i, 3).alignment = Alignment(wrap_text=True, vertical="top")
    DELIV, ASOF, TRANSIT, URG = "'Rules & Inputs'!$B$5", "'Rules & Inputs'!$B$6", "'Rules & Inputs'!$B$7", "'Rules & Inputs'!$B$8"
    rules = [
        ("Scope", f"Cues on HIE packing list {ref} heading to Jacksonville. Samples and shafts on the list are shown on the Inbound tab but not allocated."),
        ("Excluded", "ASIA / AP allocations and EMEA DIRECT allocations are not served from Jacksonville stock, so they are left out. Only EMEA-via-JAX and the other JAX-fulfilled channels are allocated."),
        ("EMEA to ship from JAX", "MIN(qty arriving, EMEA via JAX − EMEA already sent). Shortfall = EMEA via JAX − already sent − arriving (0 if covered)."),
        ("Remaining for JAX regions", "Qty arriving − EMEA to ship. Optional yellow columns on the Allocation tab (B2C, Americas, sponsorship/tradeshow/marketing) come from the NetSuite New Product Allocations export; when filled they split this remainder and show the surplus / shortfall."),
        ("Days remaining", "Days to launch = launch date − as-of date. Ship EMEA by = launch date − transit lead time. Days left = ship-by − as-of date."),
        ("Matching", "Packing-list items and EMEA items are matched on a token-sorted key (Match Key column), so 'P3 MAP TEAL NW RAD' = 'P3 MAP TEAL RAD NW'. Extra spaces are ignored."),
        ("Sources", f"Packing list {ref} ({meta.get('CUSTOMER', '')}, from {meta.get('FROM', '')}) and EMEA_AllocationQ4.xlsx (Project, Item, Launch Date, EMEA, Direct, Jax)."),
        ("Not on file", "A cue with no row in the EMEA file has no Q4 EMEA launch allocation: all units stay for the JAX-fulfilled regions. Numbered '888' editions are tracked separately (China 888 overflow)."),
    ]
    hdr_row(ru, 11, ["Rule", "", "Detail"], height=20)
    ru.merge_cells("A11:B11")
    for i, (a, b) in enumerate(rules, start=12):
        ru.cell(i, 1, a).font = F(bold=True)
        ru.merge_cells(start_row=i, start_column=1, end_row=i, end_column=2)
        c = ru.cell(i, 3, b); c.font = F(); c.alignment = Alignment(wrap_text=True, vertical="top")
        ru.row_dimensions[i].height = 30

    # ---------------- Inbound lines ---------------------------------------------------
    ib = wb.create_sheet(f"Inbound {ref}")
    title(ib, f"Inbound {ref} — packing list lines as shipped",
          "Source values from the HIE packing list (grey). Type and Match Key are derived. Allocation tab sums Qty by Match Key.", 11)
    IH = ["PO", "Item", "Qty", "Remark", "Cartons", "Net Wt", "Gross Wt", "CUFT", "Type", "PL Line", "Match Key"]
    hdr_row(ib, 4, IH, [14, 36, 8, 24, 9, 9, 9, 8, 8, 8, 36])
    for i, ln in enumerate(lines, start=5):
        vals = [ln["po"], ln["item"], ln["qty"], ln["remark"], ln["ctns"], ln["net"], ln["gross"], ln["cuft"], ln["type"], ln["line"], ln["key"]]
        for c, v in enumerate(vals, start=1):
            cell = ib.cell(i, c, v); cell.font = F(color=GREY if c <= 8 else None)
            if c in (3, 5, 10): cell.number_format = NUM
            if c in (6, 7, 8): cell.number_format = "#,##0.00"
    last_ib = 4 + len(lines)
    add_table(ib, "tblInbound", f"A4:K{last_ib}")
    ib.cell(last_ib + 2, 2, "Total units (check vs packing list TOTAL)").font = F(bold=True)
    c = ib.cell(last_ib + 2, 3, f"=SUM(C5:C{last_ib})"); c.font = F(bold=True); c.number_format = NUM
    ib.cell(last_ib + 3, 2, "Cue units").font = F()
    c = ib.cell(last_ib + 3, 3, f'=SUMIFS(C5:C{last_ib},I5:I{last_ib},"Cue")'); c.number_format = NUM
    ib.freeze_panes = "C5"
    IBQ, IBK, IBC, IBT = (f"'Inbound {ref}'!${x}$5:${x}${last_ib}" for x in ("C", "K", "E", "I"))

    # ---------------- EMEA allocation ---------------------------------------------------
    em = wb.create_sheet("EMEA Allocation Q4")
    title(em, "EMEA Allocation Q4 — launch allocations by item",
          "Grey = EMEA_AllocationQ4.xlsx as received. Jax = EMEA units fulfilled from Jacksonville; Direct = shipped direct (excluded from JAX stock). Blue = derived.", 11)
    EH = ["Project", "Item", "Launch Date", "EMEA Total", "EMEA Direct", "EMEA via Jax", "Type", "Match Key", f"Arriving on {ref}", "Days to Launch", "Status"]
    hdr_row(em, 4, EH, [46, 32, 11, 9, 9, 9, 8, 32, 10, 9, 34])
    for i, e in enumerate(emea, start=5):
        vals = [e["project"], e["item"], e["launch"], e["emea"], e["direct"], e["jax"], e["type"], e["key"]]
        for c, v in enumerate(vals, start=1):
            cell = em.cell(i, c, v); cell.font = F(color=GREY if c <= 6 else None)
            if c == 3: cell.number_format = DATE
            if c in (4, 5, 6): cell.number_format = NUM
        em.cell(i, 9, f"=SUMIFS({IBQ},{IBK},$H{i})").number_format = NUM
        em.cell(i, 10, f'=IF(C{i}="","",C{i}-{ASOF})').number_format = "0;[Red]-0"
        em.cell(i, 11, f'=IF(I{i}>0,"On {ref}: "&TEXT(I{i},"#,##0")&" arriving",IF(F{i}=0,"Direct-only / no JAX share","NOT on this shipment — "&F{i}&" via JAX still needed"))')
        for c in (9, 10, 11): em.cell(i, c).font = F(color="FF1F4E79")
    last_em = 4 + len(emea)
    add_table(em, "tblEMEA", f"A4:K{last_em}")
    em.freeze_panes = "C5"
    em.conditional_formatting.add(f"K5:K{last_em}", FormulaRule(formula=[f'LEFT(K5,3)="NOT"'], fill=AMB))
    em.conditional_formatting.add(f"J5:J{last_em}", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="FFB3261E", bold=True)))
    EM = lambda col: f"'EMEA Allocation Q4'!${col}$5:${col}${last_em}"

    # ---------------- Allocation (main) -------------------------------------------------------
    al = wb.create_sheet("Allocation", 0)
    cues = OrderedDict()
    for ln in lines:
        if ln["type"] != "Cue":
            continue
        a = cues.setdefault(ln["key"], dict(item=ln["item"], pos=[], lines=0, remarks=[]))
        a["lines"] += 1
        if str(ln["po"]) not in a["pos"]: a["pos"].append(str(ln["po"]))
        if ln["remark"] and ln["remark"] not in a["remarks"]: a["remarks"].append(ln["remark"])
    launch_of = {e["key"]: e["launch"] for e in emea}
    order = sorted(cues.items(), key=lambda kv: (launch_of.get(kv[0]) is None, launch_of.get(kv[0]) or dt.datetime.max, kv[1]["item"]))

    AH = ["SKU (as packed)", "Project / Launch", "PO(s)", "PL Lines", "Cartons", "Qty Arriving", "Launch Date", "Days to Launch (as-of)",
          "Days Delivery → Launch", "EMEA Total", "EMEA Direct (excluded)", "EMEA via JAX", "EMEA Already Sent (input)", "EMEA to Ship from JAX",
          "EMEA Shortfall", "Ship EMEA by", "Days Left to Ship EMEA", "B2C Alloc (input)", "Americas Alloc (input)", "Sponsor / Tradeshow / Mktg (input)",
          "Remaining for JAX Regions", "Surplus / (Short) vs JAX Allocs", "Status / Action", "Packing-list Remarks", "Match Key"]
    AW = [34, 40, 16, 6, 7, 9, 11, 9, 9, 8, 9, 8, 10, 10, 9, 11, 9, 9, 9, 11, 10, 11, 70, 26, 34]
    title(al, f"JAX inbound {ref} — launch allocation (Jacksonville-fulfilled regions only)",
          f"Delivered to JAX {delivery:%m/%d/%Y} · HIE (Hamson) · {meta.get('VESSEL/VOYAGE', '')} · Cues only. "
          "EMEA to ship from JAX = MIN(arriving, EMEA via JAX − already sent). ASIA/AP and EMEA Direct are excluded (not served from JAX). "
          "Everything else stays for the JAX-fulfilled regions. Yellow = input. Dates and days use Rules & Inputs.", len(AH))
    # KPIs
    first, last = 9, 9 + len(order) - 1
    col = lambda c: f"${L(c)}${first}:${L(c)}${last}"
    kpis = [
        ("Cue units arriving", f"=SUM({col(6)})", NUM),
        ("Cue SKUs", f"=COUNTA({col(1)})", NUM),
        ("SKUs with EMEA via-JAX allocation", f'=COUNTIF({col(12)},">0")', NUM),
        ("EMEA units to ship from JAX", f"=SUM({col(14)})", NUM),
        ("EMEA shortfall units", f"=SUM({col(15)})", NUM),
        ("Units remaining for JAX regions", f"=SUM({col(21)})", NUM),
        ("SKUs overdue / urgent (EMEA)", f'=COUNTIFS({col(17)},"<="&{URG},{col(14)},">0")', NUM),
        ("EMEA launch items NOT on this shipment", f"=COUNTIF({EM('K')},\"NOT*\")", NUM),
    ]
    for i, (lab, f, nf) in enumerate(kpis):
        c0 = 1 + i * 3
        a = al.cell(4, c0, lab); a.font = F(size=9, color=GREY); a.alignment = Alignment(wrap_text=True, vertical="top")
        al.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + 2)
        v = al.cell(5, c0, f); v.font = F(size=16, bold=True); v.number_format = nf
        al.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c0 + 2)
    al.row_dimensions[4].height = 26; al.row_dimensions[5].height = 24
    al.cell(7, 1, "BY SKU — sorted by launch date (non-launch cues at the bottom)").font = F(bold=True, color="FFFFFFFF")
    al.cell(7, 1).fill = SEC
    al.merge_cells(start_row=7, start_column=1, end_row=7, end_column=len(AH))
    hdr_row(al, 8, AH, AW, height=42)
    for r, (k, a) in enumerate(order, start=first):
        al.cell(r, 1, a["item"]).font = F(bold=True)
        al.cell(r, 2, f'=IFERROR(INDEX({EM("A")},MATCH($Y{r},{EM("H")},0)),"")')
        al.cell(r, 3, ", ".join(a["pos"]))
        al.cell(r, 4, a["lines"])
        al.cell(r, 5, f"=SUMIFS({IBC},{IBK},$Y{r})")
        al.cell(r, 6, f"=SUMIFS({IBQ},{IBK},$Y{r})")
        al.cell(r, 7, f'=IFERROR(INDEX({EM("C")},MATCH($Y{r},{EM("H")},0)),"")')
        al.cell(r, 8, f'=IF(G{r}="","",G{r}-{ASOF})')
        al.cell(r, 9, f'=IF(G{r}="","",G{r}-{DELIV})')
        al.cell(r, 10, f'=SUMIFS({EM("D")},{EM("H")},$Y{r})')
        al.cell(r, 11, f'=SUMIFS({EM("E")},{EM("H")},$Y{r})')
        al.cell(r, 12, f'=SUMIFS({EM("F")},{EM("H")},$Y{r})')
        al.cell(r, 13).fill = INPUT
        al.cell(r, 14, f"=IF(L{r}=0,0,MIN(F{r},MAX(0,L{r}-M{r})))")
        al.cell(r, 15, f"=IF(L{r}=0,0,MAX(0,L{r}-M{r}-F{r}))")
        al.cell(r, 16, f'=IF(OR(G{r}="",L{r}=0),"",G{r}-{TRANSIT})')
        al.cell(r, 17, f'=IF(P{r}="","",P{r}-{ASOF})')
        for c in (18, 19, 20): al.cell(r, c).fill = INPUT
        al.cell(r, 21, f"=F{r}-N{r}")
        al.cell(r, 22, f'=IF(R{r}+S{r}+T{r}=0,"",U{r}-(R{r}+S{r}+T{r}))')
        al.cell(r, 23, (f'=IF(G{r}="","No Q4 EMEA launch allocation on file — all "&TEXT(F{r},"#,##0")&" units stay for JAX regions",'
                        f'IF(L{r}=0,"EMEA served Direct only — all "&TEXT(F{r},"#,##0")&" units stay for JAX regions",'
                        f'IF(O{r}>0,"SHORT "&TEXT(O{r},"#,##0")&" for EMEA — ship all "&TEXT(N{r},"#,##0")&" arriving now; cover balance from on hand / next shipment",'
                        f'IF(N{r}=0,"EMEA via-JAX allocation already sent — all units stay for JAX regions",'
                        f'IF(Q{r}<0,"OVERDUE — ship "&TEXT(N{r},"#,##0")&" to EMEA immediately (launch "&TEXT(G{r},"m/d")&", "&TEXT(-Q{r},"0")&" days past ship-by)",'
                        f'IF(Q{r}<={URG},"URGENT — ship "&TEXT(N{r},"#,##0")&" to EMEA within "&TEXT(Q{r},"0")&" days (launch "&TEXT(G{r},"m/d")&")",'
                        f'"Ship "&TEXT(N{r},"#,##0")&" to EMEA by "&TEXT(P{r},"m/d")&" ("&TEXT(Q{r},"0")&" days left); "&TEXT(U{r},"#,##0")&" stay for JAX regions"))))))'))
        al.cell(r, 24, " | ".join(a["remarks"]))
        al.cell(r, 25, k)
        for c in range(1, len(AH) + 1):
            cell = al.cell(r, c)
            if c != 1: cell.font = F(color="FF0000FF" if c in (13, 18, 19, 20) else None)
            if c in (4, 5, 6, 10, 11, 12, 13, 14, 15, 18, 19, 20, 21): cell.number_format = NUM
            if c == 22: cell.number_format = "#,##0;[Red]\\(#,##0\\);\\-"
            if c in (7, 16): cell.number_format = DATE
            if c in (8, 9, 17): cell.number_format = "0;[Red]-0"
            if c in (23, 24): cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = Border(bottom=THIN)
        al.row_dimensions[r].height = 28
    tot = last + 1
    al.cell(tot, 1, "TOTAL").font = F(bold=True)
    for c in (4, 5, 6, 10, 11, 12, 13, 14, 15, 18, 19, 20, 21):
        x = al.cell(tot, c, f"=SUM({L(c)}{first}:{L(c)}{last})"); x.font = F(bold=True); x.number_format = NUM
    al.freeze_panes = "B9"
    rng = lambda c: f"{L(c)}{first}:{L(c)}{last}"
    al.conditional_formatting.add(rng(17), CellIsRule(operator="lessThan", formula=["0"], fill=RED, font=Font(color="FFB3261E", bold=True)))
    al.conditional_formatting.add(rng(17), FormulaRule(formula=[f'AND(ISNUMBER(Q{first}),Q{first}>=0,Q{first}<={URG})'], fill=AMB))
    al.conditional_formatting.add(rng(15), CellIsRule(operator="greaterThan", formula=["0"], fill=RED, font=Font(color="FFB3261E", bold=True)))
    al.conditional_formatting.add(rng(14), CellIsRule(operator="greaterThan", formula=["0"], fill=GRN))
    al.conditional_formatting.add(rng(23), FormulaRule(formula=[f'OR(LEFT(W{first},5)="SHORT",LEFT(W{first},7)="OVERDUE")'], font=Font(color="FFB3261E", bold=True)))
    al.conditional_formatting.add(rng(23), FormulaRule(formula=[f'LEFT(W{first},6)="URGENT"'], font=Font(color="FFB26A00", bold=True)))
    al.conditional_formatting.add(rng(8), CellIsRule(operator="lessThan", formula=["0"], font=Font(color="FFB3261E")))
    al.cell(tot + 2, 1, "Legend: yellow = input · green = EMEA units to pull for transshipment · red = shortfall or overdue · amber = urgent. "
            "EMEA items not on this shipment: see EMEA Allocation Q4 tab, Status column.").font = F(size=9, color=GREY)
    al.merge_cells(start_row=tot + 2, start_column=1, end_row=tot + 2, end_column=len(AH))

    # non-cue items note
    others = [ln for ln in lines if ln["type"] != "Cue"]
    al.cell(tot + 3, 1, f"Not allocated (not cues): {len(others)} packing-list lines, {sum(l['qty'] for l in others):,} units — "
            + "; ".join(sorted({f"{l['type']}: {l['item']}" for l in others}))[:900]).font = F(size=9, color=GREY)
    al.cell(tot + 3, 1).alignment = Alignment(wrap_text=True, vertical="top")
    al.merge_cells(start_row=tot + 3, start_column=1, end_row=tot + 3, end_column=len(AH))
    al.row_dimensions[tot + 3].height = 40

    wb.save(out)
    print(f"wrote {out}: {len(order)} cue SKUs, {len(lines)} packing-list lines, {len(emea)} EMEA rows")


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2], sys.argv[3])
