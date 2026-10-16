"""Build the JAX inbound launch-allocation workbook.

Usage:
    python tools/jax_shipment_allocation.py <packing_list.xlsx> <emea_alloc.xlsx> <out.xlsx> [npa_export.csv|xlsx]

Inputs
  packing list   HIE packing list (one sheet, header row 10: PO_NUMBER, ITEM, QTY, ...)
  emea_alloc     EMEA_AllocationQ4.xlsx (Project, Item, Launch Date, EMEA, Direct, Jax)
  npa_export     optional NetSuite "New Product Allocations" export (Internal ID, Name,
                 Description, Base Price, Targeted Launch Date, B2C, Sponsorships, EMEA,
                 Americas, ASIA, Tradeshows, Marketing). If omitted the NPA Import tab is
                 left as a paste area and the channel columns fill when it is pasted.

Rule: stock is Jacksonville-bound, so only JAX-fulfilled channels are allocated:
EMEA via JAX (EMEA total minus Direct), B2C, Americas, Sponsorship, Tradeshow,
Marketing. EMEA Direct and ASIA/AP are excluded. Left for stock = arriving minus
the sum of those allocations; negative = short.
"""
from __future__ import annotations

import csv
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

INK, GOLD, GREY, BLUE = "FF1C1610", "FFC8922A", "FF595959", "FF1F4E79"
INPUT = PatternFill("solid", fgColor="FFFFF2CC")
HDR = PatternFill("solid", fgColor="FF44546A")
SEC = PatternFill("solid", fgColor=GOLD)
RED = PatternFill("solid", fgColor="FFFBE3E0")
AMB = PatternFill("solid", fgColor="FFFFF0D4")
GRN = PatternFill("solid", fgColor="FFE2F0E4")
THIN = Side(style="thin", color="FFD9D9D9")
NUM = "#,##0;[Red]\\(#,##0\\);\\-"
DATE = "m/d/yyyy"
NPA_COLS = ["Internal ID", "Name", "Description", "Base Price", "Targeted Launch Date", "B2C", "Sponsorships", "EMEA", "Americas", "ASIA", "Tradeshows", "Marketing"]
NPA_ROWS = 400  # paste area size


def F(**k):
    return Font(name="Arial", size=k.pop("size", 10), **k)


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
                        meta[lab.rstrip(":.")] = v.split(":", 1)[-1].strip()
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
                         type=item_type(item, g("PO_NUMBER"), remark), key=key(item)))
    m = re.search(r"([A-Za-z]{3})\.?\s*(\d{1,2})\.?\s*(\d{4})", meta.get("DELIVERY DATE", ""))
    delivery = dt.datetime.strptime(" ".join(m.groups()), "%b %d %Y").date() if m else dt.date.today()
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
                        direct=d.get("Direct") or 0, jax=d.get("Jax") or 0, key=key(item)))
    return out


def read_npa(path):
    if not path:
        return []
    if path.lower().endswith(".csv"):
        with open(path, newline="", encoding="utf-8-sig") as fh:
            rows = list(csv.reader(fh))
    else:
        ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
        rows = [[c for c in r] for r in ws.iter_rows(values_only=True)]
    hdr = [str(h or "").strip() for h in rows[0]]
    out = []
    for r in rows[1:]:
        d = dict(zip(hdr, r))
        if not d.get("Name"):
            continue
        row = []
        for c in NPA_COLS:
            v = d.get(c)
            if isinstance(v, str):
                t = v.strip().replace(",", "")
                if re.fullmatch(r"-?\d+(\.\d+)?", t):
                    v = float(t) if "." in t else int(t)
                elif c == "Targeted Launch Date":
                    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
                        try:
                            v = dt.datetime.strptime(t, fmt); break
                        except ValueError:
                            pass
            row.append(v)
        out.append(row)
    return out


def hdr_row(ws, row, headers, widths=None, height=30):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row, c, h)
        cell.font = F(bold=True, color="FFFFFFFF")
        cell.fill = HDR
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
    ws.row_dimensions[2].height = 28


def section(ws, row, text, cols):
    c = ws.cell(row, 1, text)
    c.font = F(bold=True, color="FFFFFFFF")
    c.fill = SEC
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)


def add_table(ws, name, ref):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)


def build(pl_path, emea_path, out, npa_path=None):
    lines, meta, delivery = read_packing_list(pl_path)
    emea = read_emea(emea_path)
    npa = read_npa(npa_path)
    ref = (meta.get("REF.NO") or "shipment").strip()
    wb = openpyxl.Workbook()

    # ============ NPA Import (paste area) =================================================
    np_ = wb.create_sheet("NPA Import")
    title(np_, "NPA Import — NetSuite New Product Allocations export",
          "Paste the New Product Allocations saved-search export here from row 4 (keep these headers in this order). "
          "The Allocation tab reads B2C, Sponsorships, Americas, Tradeshows and Marketing by SKU from this tab. EMEA and ASIA are not used from here: "
          "EMEA comes from the EMEA Q4 tab (Direct excluded) and ASIA is not served from Jacksonville.", 12)
    hdr_row(np_, 3, NPA_COLS, [10, 34, 44, 9, 12, 8, 11, 8, 9, 8, 10, 9], height=24)
    np_first, np_last = 4, 3 + max(NPA_ROWS, len(npa) + 50)
    for r in range(np_first, np_last + 1):
        for c in range(1, 13):
            cell = np_.cell(r, c)
            cell.fill = INPUT
            cell.font = F(color="FF0000FF")
            if c == 5: cell.number_format = DATE
            if c >= 6: cell.number_format = "#,##0;\\(#,##0\\);"
    for i, row in enumerate(npa, start=np_first):
        for c, v in enumerate(row, start=1):
            np_.cell(i, c, v)
    np_.freeze_panes = "C4"
    NPH = f"'NPA Import'!$A$3:$L$3"
    NPD = f"'NPA Import'!$A${np_first}:$L${np_last}"
    NPN = f"'NPA Import'!$B${np_first}:$B${np_last}"

    NPC = {c: L(i + 1) for i, c in enumerate(NPA_COLS)}  # fixed column letters, header order is part of the contract

    def npa_sum(colname, sku_cell):
        col = NPC[colname]
        return f"SUMIFS('NPA Import'!${col}${np_first}:${col}${np_last},{NPN},\"*\"&{sku_cell})"

    def npa_launch(sku_cell):
        col = NPC["Targeted Launch Date"]
        return f"IFERROR(INDEX('NPA Import'!${col}${np_first}:${col}${np_last},MATCH(\"*\"&{sku_cell},{NPN},0)),\"\")"

    # ============ Inbound lines ===========================================================
    ib = wb.create_sheet(f"Inbound {ref}")
    title(ib, f"Inbound {ref} — packing list lines as shipped",
          "Values from the HIE packing list. Type and Match Key are derived; the Allocation tab sums Qty by Match Key.", 8)
    IH = ["PO", "Item", "Qty", "Remark", "Cartons", "Type", "PL Line", "Match Key"]
    hdr_row(ib, 4, IH, [14, 36, 8, 26, 9, 8, 8, 36], height=22)
    for i, ln in enumerate(lines, start=5):
        for c, v in enumerate([ln["po"], ln["item"], ln["qty"], ln["remark"], ln["ctns"], ln["type"], ln["line"], ln["key"]], start=1):
            cell = ib.cell(i, c, v); cell.font = F(color=GREY if c <= 5 else None)
            if c in (3, 5): cell.number_format = "#,##0;\\(#,##0\\);\\-"
    last_ib = 4 + len(lines)
    add_table(ib, "tblInbound", f"A4:H{last_ib}")
    ib.cell(last_ib + 2, 2, "Total units (ties to packing list TOTAL)").font = F(bold=True)
    c = ib.cell(last_ib + 2, 3, f"=SUM(C5:C{last_ib})"); c.font = F(bold=True); c.number_format = "#,##0"
    ib.cell(last_ib + 3, 2, "Cue units").font = F()
    c = ib.cell(last_ib + 3, 3, f'=SUMIFS(C5:C{last_ib},F5:F{last_ib},"Cue")'); c.number_format = "#,##0"
    ib.freeze_panes = "C5"
    IBQ, IBK = f"'Inbound {ref}'!$C$5:$C${last_ib}", f"'Inbound {ref}'!$H$5:$H${last_ib}"

    # ============ EMEA Q4 ==================================================================
    em = wb.create_sheet("EMEA Q4")
    title(em, "EMEA Q4 launch allocations — and which ones this shipment covers",
          "Grey = EMEA_AllocationQ4.xlsx as received. Via JAX = EMEA units fulfilled from Jacksonville (Direct is shipped direct and excluded). Blue = derived.", 10)
    EH = ["Launch", "SKU", "Launch Date", "Days to Launch", "EMEA Total", "EMEA Direct (excluded)", "EMEA via JAX", f"Arriving on {ref}", "Status", "Match Key"]
    hdr_row(em, 4, EH, [46, 32, 11, 9, 9, 10, 9, 10, 40, 32], height=30)
    emea_sorted = sorted(emea, key=lambda e: (e["launch"] or dt.datetime.max, e["item"]))
    for i, e in enumerate(emea_sorted, start=5):
        for c, v in enumerate([e["project"], e["item"], e["launch"], None, e["emea"], e["direct"], e["jax"], None, None, e["key"]], start=1):
            if v is not None:
                cell = em.cell(i, c, v); cell.font = F(color=GREY if c in (1, 2, 3, 5, 6, 7) else None)
        em.cell(i, 3).number_format = DATE
        for c in (5, 6, 7): em.cell(i, c).number_format = "#,##0;\\(#,##0\\);\\-"
        em.cell(i, 4, f'=IF(C{i}="","",C{i}-ASOF)').number_format = "0;[Red]-0"
        em.cell(i, 8, f"=SUMIFS({IBQ},{IBK},$J{i})").number_format = "#,##0;\\(#,##0\\);\\-"
        em.cell(i, 9, f'=IF(G{i}=0,"Direct only — nothing via JAX",IF(H{i}>=G{i},"Covered by this shipment",IF(H{i}>0,"PARTIAL — "&TEXT(G{i}-H{i},"#,##0")&" more needed via JAX","NOT on this shipment — "&TEXT(G{i},"#,##0")&" needed via JAX")))')
        for c in (4, 8, 9): em.cell(i, c).font = F(color=BLUE)
    last_em = 4 + len(emea)
    add_table(em, "tblEMEA", f"A4:J{last_em}")
    em.freeze_panes = "C5"
    em.conditional_formatting.add(f"I5:I{last_em}", FormulaRule(formula=['LEFT(I5,3)="NOT"'], fill=RED))
    em.conditional_formatting.add(f"I5:I{last_em}", FormulaRule(formula=['LEFT(I5,7)="PARTIAL"'], fill=AMB))
    em.conditional_formatting.add(f"I5:I{last_em}", FormulaRule(formula=['LEFT(I5,7)="Covered"'], fill=GRN))
    EM = lambda col: f"'EMEA Q4'!${col}$5:${col}${last_em}"

    # ============ Allocation (main) ========================================================
    al = wb.create_sheet("Allocation", 0)
    del wb["Sheet"]
    cues = OrderedDict()
    for ln in lines:
        if ln["type"] != "Cue":
            continue
        a = cues.setdefault(ln["key"], dict(item=ln["item"], remarks=[]))
        if ln["remark"] and ln["remark"] not in a["remarks"]:
            a["remarks"].append(ln["remark"])
    launch_of = {e["key"]: e["launch"] for e in emea}
    launch_items = sorted([kv for kv in cues.items() if kv[0] in launch_of], key=lambda kv: (launch_of[kv[0]] or dt.datetime.max, kv[1]["item"]))
    other_items = sorted([kv for kv in cues.items() if kv[0] not in launch_of], key=lambda kv: kv[1]["item"])

    AH = ["Launch", "SKU", "Launch Date", "Days to Launch", "Arriving", "EMEA via JAX", "B2C", "Americas", "Sponsorship", "Tradeshow", "Marketing",
          "Total to Allocate", "Left for Stock / (Short)", "Action"]
    AW = [40, 32, 11, 9, 9, 9, 8, 9, 10, 9, 9, 10, 11, 64]
    NCOL = len(AH)
    title(al, f"Packing list {ref} → Jacksonville: launch allocation by SKU",
          f"Delivered to JAX {delivery:%m/%d/%Y} · HIE (Hamson) · {meta.get('VESSEL/VOYAGE', '')}. "
          "Jacksonville-fulfilled channels only: EMEA via JAX (EMEA less Direct), B2C, Americas, Sponsorship, Tradeshow, Marketing. "
          "EMEA Direct and ASIA/AP are excluded. Left for stock = arriving − total to allocate; a negative number in red is a shortfall.", NCOL)

    # settings (top right)
    al.cell(4, NCOL, "As-of date").font = F(size=9, color=GREY)
    s1 = al.cell(4, NCOL + 1, "=TODAY()"); s1.number_format = DATE
    al.cell(5, NCOL, "Urgent if days to launch ≤").font = F(size=9, color=GREY)
    s2 = al.cell(5, NCOL + 1, 7)
    for s in (s1, s2): s.fill = INPUT; s.font = F(color="FF0000FF")
    al.column_dimensions[L(NCOL + 1)].width = 11
    wb.defined_names["ASOF"] = openpyxl.workbook.defined_name.DefinedName("ASOF", attr_text=f"Allocation!${L(NCOL + 1)}$4")
    wb.defined_names["URGENT"] = openpyxl.workbook.defined_name.DefinedName("URGENT", attr_text=f"Allocation!${L(NCOL + 1)}$5")

    # totals band (rows 4-5)
    first = 9
    n_launch, n_other = len(launch_items), len(other_items)
    sec2 = first + n_launch + 1          # section header row for other cues
    first2 = sec2 + 1
    last2 = first2 + n_other - 1
    last_all = last2
    def both(col):  # range covering both sections
        return f"${col}${first}:${col}${last_all}"
    band = [("Cue units arriving", "E"), ("EMEA via JAX", "F"), ("B2C", "G"), ("Americas", "H"), ("Sponsorship", "I"), ("Tradeshow", "J"), ("Marketing", "K"), ("Total to allocate", "L"), ("Left for stock / (short)", "M")]
    for i, (lab, col) in enumerate(band, start=1):
        a = al.cell(4, i, lab); a.font = F(size=9, color=GREY); a.alignment = Alignment(wrap_text=True, vertical="bottom")
        v = al.cell(5, i, f"=SUM({both(col)})" if col != "M" else f"=SUMIF({both(col)},\"<0\")"); v.font = F(size=14, bold=True)
        v.number_format = NUM if col != "M" else "[Red]\\(#,##0\\);[Red]\\(#,##0\\);\\-"
    al.cell(4, 9).value = "Shortfall units (sum of negatives)"
    al.row_dimensions[4].height = 26; al.row_dimensions[5].height = 22
    al.cell(6, 1, "Channel columns B2C … Marketing fill from the NPA Import tab once the NetSuite New Product Allocations export is pasted there. Blank = not on file yet.").font = F(size=9, color=GREY)
    al.merge_cells(start_row=6, start_column=1, end_row=6, end_column=NCOL)

    section(al, 7, f"Q4 LAUNCH ITEMS ON THIS SHIPMENT ({n_launch} SKUs) — sorted by launch date", NCOL)
    hdr_row(al, 8, AH, AW, height=30)

    def write_row(r, k, a, is_launch):
        al.cell(r, 1, f'=IFERROR(INDEX({EM("A")},MATCH("{k}",{EM("J")},0)),"")')
        al.cell(r, 2, a["item"]).font = F(bold=True)
        if is_launch:
            al.cell(r, 3, f'=IFERROR(INDEX({EM("C")},MATCH("{k}",{EM("J")},0)),"")')
        else:
            al.cell(r, 3, f"={npa_launch(f'$B{r}')}")
        al.cell(r, 4, f'=IF(C{r}="","",C{r}-ASOF)')
        al.cell(r, 5, f'=SUMIFS({IBQ},{IBK},"{k}")')
        al.cell(r, 6, f'=SUMIFS({EM("G")},{EM("J")},"{k}")')
        for c, nm in zip(range(7, 12), ["B2C", "Americas", "Sponsorships", "Tradeshows", "Marketing"]):
            al.cell(r, c, f"={npa_sum(nm, f'$B{r}')}")
        al.cell(r, 12, f"=SUM(F{r}:K{r})")
        al.cell(r, 13, f"=E{r}-L{r}")
        al.cell(r, 14, (f'=IF(L{r}=0,"No JAX-channel allocation on file — all "&TEXT(E{r},"#,##0")&" to stock",'
                        f'IF(M{r}<0,"SHORT "&TEXT(-M{r},"#,##0")&" — fill EMEA first, then channels; balance from on hand / next shipment",'
                        f'IF(D{r}="","Allocate per channels; no launch date on file",'
                        f'IF(D{r}<0,"Launch passed — distribute "&TEXT(L{r},"#,##0")&" now; "&TEXT(M{r},"#,##0")&" to stock",'
                        f'IF(D{r}<=URGENT,"URGENT — distribute "&TEXT(L{r},"#,##0")&" within "&TEXT(D{r},"0")&" days; "&TEXT(M{r},"#,##0")&" to stock",'
                        f'"Distribute "&TEXT(L{r},"#,##0")&" before "&TEXT(C{r},"m/d")&" ("&TEXT(D{r},"0")&" days); "&TEXT(M{r},"#,##0")&" to stock")))))'))
        for c in range(1, NCOL + 1):
            cell = al.cell(r, c)
            if c != 2: cell.font = F()
            if c == 3: cell.number_format = DATE
            if c == 4: cell.number_format = "0;[Red]-0"
            if 5 <= c <= 13: cell.number_format = NUM
            if c == 14: cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.border = Border(bottom=THIN)
        al.row_dimensions[r].height = 26

    for i, (k, a) in enumerate(launch_items):
        write_row(first + i, k, a, True)
    section(al, sec2, f"OTHER CUES ON THIS SHIPMENT ({n_other} SKUs) — no Q4 EMEA launch allocation on file; channel columns fill from NPA Import if present", NCOL)
    for i, (k, a) in enumerate(other_items):
        write_row(first2 + i, k, a, False)
    tot = last_all + 1
    al.cell(tot, 2, "TOTAL — all cues").font = F(bold=True)
    for c in range(5, 14):
        x = al.cell(tot, c, f"=SUM({L(c)}{first}:{L(c)}{last_all})"); x.font = F(bold=True); x.number_format = NUM
    al.freeze_panes = "C9"
    rng = lambda c: f"{L(c)}{first}:{L(c)}{last_all}"
    al.conditional_formatting.add(rng(13), CellIsRule(operator="lessThan", formula=["0"], fill=RED, font=Font(color="FFB3261E", bold=True)))
    al.conditional_formatting.add(rng(4), CellIsRule(operator="lessThan", formula=["0"], fill=RED, font=Font(color="FFB3261E", bold=True)))
    al.conditional_formatting.add(rng(4), FormulaRule(formula=[f"AND(ISNUMBER(D{first}),D{first}>=0,D{first}<=URGENT)"], fill=AMB))
    al.conditional_formatting.add(rng(6), CellIsRule(operator="greaterThan", formula=["0"], fill=GRN))
    for c in range(7, 12):
        al.conditional_formatting.add(rng(c), CellIsRule(operator="greaterThan", formula=["0"], fill=GRN))
    al.conditional_formatting.add(rng(14), FormulaRule(formula=[f'LEFT(N{first},5)="SHORT"'], font=Font(color="FFB3261E", bold=True)))
    al.conditional_formatting.add(rng(14), FormulaRule(formula=[f'OR(LEFT(N{first},6)="URGENT",LEFT(N{first},6)="Launch")'], font=Font(color="FFB26A00", bold=True)))
    others = [ln for ln in lines if ln["type"] != "Cue"]
    al.cell(tot + 2, 1, f"Not allocated: {len(others)} non-cue lines ({sum(l['qty'] for l in others):,} units — samples and one shaft), listed on the Inbound {ref} tab. "
            f"EMEA Q4 launch items that are NOT on this shipment are listed on the EMEA Q4 tab (red status).").font = F(size=9, color=GREY)
    al.merge_cells(start_row=tot + 2, start_column=1, end_row=tot + 2, end_column=NCOL)

    wb._sheets = [wb["Allocation"], wb["EMEA Q4"], wb["NPA Import"], wb[f"Inbound {ref}"]]
    wb.save(out)
    print(f"wrote {out}: {n_launch} launch SKUs + {n_other} other cues, {len(lines)} lines, {len(emea)} EMEA rows, {len(npa)} NPA rows")


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4] if len(sys.argv) > 4 else None)
