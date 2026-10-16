"""Build the JAX inbound launch-allocation workbook.

Usage:
    python tools/jax_shipment_allocation.py <packing_list.xlsx> <emea_alloc.xlsx> <launch_plan.csv|xlsx> <out.xlsx> [plan_asks.csv] [ask_column]

Inputs
  packing list   HIE packing list (one sheet, header row 10: PO_NUMBER, ITEM, QTY, ...)
  emea_alloc     EMEA_AllocationQ4.xlsx (Project, Item, Launch Date, EMEA, Direct, Jax)
  launch plan    2026 Launch Plan Allocation export (Project, Item, Launch Date, B2C, EMEA,
                 NA, AP, Trade show, Sponsorship, Marketing, ...). Extra empty columns are ignored.
  plan asks      optional CSV exported from the Shipment Request workbook's 'JAX Shipments by Date'
                 tab (sku, block, alloc, recd, jax_need, rule, flag, ask_<date>..., finished, completes,
                 note, instr, held). Gives 'already received', the remaining JAX-bound need and what
                 was actually requested for this sailing, so shortfalls are measured against the
                 request that passed the 85% gate, not against the full launch allowance.

Rule (stock is Jacksonville-bound):
  Allocated here  = EMEA via JAX (EMEA less Direct, from the EMEA file) + B2C + NA + Trade show
                    + Sponsorship + Marketing (from the launch plan).
  Excluded        = AP and EMEA Direct (not served from Jacksonville).
  Left for stock  = units arriving - total allocated; negative = shortfall.

SKUs match on a token-sorted key (so "MAP TEAL NW RAD" = "MAP TEAL RAD NW"); a launch-plan
suffix after " - " (e.g. "- BOX", "- B2C / AP") is ignored for matching and duplicate plan rows
for one SKU are summed. Samples on the packing list are listed but not allocated.
All derived numbers are live formulas; yellow cells are inputs.
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
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

warnings.filterwarnings("ignore")

INK, GOLD, GREY, BLUE = "FF1C1610", "FFC8922A", "FF595959", "FF1F4E79"
INPUT = PatternFill("solid", fgColor="FFFFF2CC")
HDR = PatternFill("solid", fgColor="FF44546A")
HDR_EX = PatternFill("solid", fgColor="FF8C8C8C")
SEC = PatternFill("solid", fgColor=GOLD)
RED = PatternFill("solid", fgColor="FFFBE3E0")
AMB = PatternFill("solid", fgColor="FFFFF0D4")
GRN = PatternFill("solid", fgColor="FFE2F0E4")
THIN = Side(style="thin", color="FFD9D9D9")
NUM = "#,##0;[Red]\\(#,##0\\);\\-"
DATE = "m/d/yyyy"
PLAN_COLS = ["Project", "Item", "Exclusive", "On Current Launch", "Launch Date", "Butt", "Shaft", "MSRP", "Dealer ASP",
             "B2C", "EMEA", "NA", "AP", "Trade show", "Sponsorship", "Marketing", "B2B TLT Units", "B2C TLT Units", "Total Units/ordering", "EST REV"]
CHANNELS = [("B2C", "B2C"), ("NA", "NA (Americas)"), ("Trade show", "Trade show"), ("Sponsorship", "Sponsorship"), ("Marketing", "Marketing")]


def F(**k):
    return Font(name="Arial", size=k.pop("size", 10), **k)


def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip())


def key(s) -> str:
    base = re.sub(r"\s+-\s+.*$", "", norm(s))  # drop " - BOX", " - B2C / AP" program suffixes
    return " ".join(sorted(base.upper().split()))


def num(v):
    if isinstance(v, (int, float)):
        return v
    t = str(v or "").strip().replace(",", "").replace("$", "")
    if t in ("", "-"):
        return None
    try:
        return float(t) if "." in t else int(t)
    except ValueError:
        return v


def date(v):
    if isinstance(v, (dt.date, dt.datetime)):
        return v
    t = str(v or "").strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(t, fmt)
        except ValueError:
            pass
    return None


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
        item = norm(item)
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
        item = norm(d["Item"])
        out.append(dict(project=d.get("Project"), item=item, launch=date(d.get("Launch Date")), emea=num(d.get("EMEA")) or 0,
                        direct=num(d.get("Direct")) or 0, jax=num(d.get("Jax")) or 0, key=key(item)))
    return out


def read_plan(path):
    if path.lower().endswith(".csv"):
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as fh:
            raw = [r[:40] for r in csv.reader(fh)]
    else:
        ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
        raw = [list(r[:40]) for r in ws.iter_rows(values_only=True)]
    hdr = [norm(h).replace('"', "") for h in raw[0]]
    out = []
    for r in raw[1:]:
        d = dict(zip(hdr, r))
        item = norm(d.get("Item"))
        if not item:
            continue
        row = {c: d.get(c) for c in PLAN_COLS}
        row["Item"] = item
        row["Launch Date"] = date(row["Launch Date"])
        for c in PLAN_COLS[5:]:
            row[c] = num(row[c])
        row["key"] = key(item)
        out.append(row)
    return out


def read_asks(path, ask_col=None):
    if not path:
        return []
    with open(path, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    ask_col = ask_col or [c for c in rows[0] if c.startswith("ask_")][-1]
    for r in rows:
        r["sku"] = norm(r["sku"]); r["key"] = key(r["sku"])
        for c, v in list(r.items()):
            if c not in ("sku", "key", "block", "rule", "flag", "note", "instr", "launch"):
                r[c] = num(v) or 0
        if r.get("block") == "CORE":
            # CORE rows carry a 3-window replenishment need; what is owed from THIS sailing is its ask
            r["note"] = f"3-window replenishment need {r['jax_need']:,.0f}; " + (r.get("note") or "")
            r["jax_need"] = r[ask_col]
            r["rule"] = r["rule"] or "n/a (CORE)"
            r["flag"] = r["flag"] or "n/a"
    launch_keys = {r["key"] for r in rows if r.get("block") == "Launch"}
    dupes = [r for r in rows if r.get("block") == "CORE" and r["key"] in launch_keys]
    for r in dupes:
        print(f"  plan asks: dropping CORE duplicate of launch SKU {r['sku']} (ask {r.get('ask_1016', '')})")
    return [r for r in rows if r not in dupes]


def hdr_row(ws, row, headers, widths=None, height=30, fills=None):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row, c, h)
        cell.font = F(bold=True, color="FFFFFFFF")
        cell.fill = (fills or {}).get(c, HDR)
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


def section(ws, row, text, cols):
    c = ws.cell(row, 1, text)
    c.font = F(bold=True, color="FFFFFFFF")
    c.fill = SEC
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)


def add_table(ws, name, ref):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)


def build(pl_path, emea_path, plan_path, out, asks_path=None, ask_col=None):
    lines, meta, delivery = read_packing_list(pl_path)
    emea = read_emea(emea_path)
    plan = read_plan(plan_path)
    asks = read_asks(asks_path, ask_col)
    if asks and not ask_col:
        ask_col = [c for c in asks[0] if c.startswith("ask_")][-1]
    ref = (meta.get("REF.NO") or "shipment").strip()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ============ Launch Plan (source) =====================================================
    lp = wb.create_sheet("Launch Plan")
    title(lp, "Launch Plan — 2026 Launch Plan Allocation (as received)",
          "Grey = the export as received. Match Key is derived (token-sorted SKU, ' - BOX' / ' - B2C / AP' suffix dropped). "
          "The Allocation tab sums B2C, NA, Trade show, Sponsorship and Marketing by Match Key, so duplicate rows for one SKU add up. "
          "EMEA here is used only as a cross-check; the EMEA Direct / via-JAX split comes from the EMEA Q4 tab. AP is excluded.", 21)
    LH = PLAN_COLS + ["Match Key"]
    hdr_row(lp, 4, LH, [36, 34, 8, 8, 11, 9, 9, 10, 10, 7, 7, 7, 7, 8, 9, 8, 8, 8, 10, 10, 34], height=30)
    for i, p in enumerate(plan, start=5):
        for c, col in enumerate(LH, start=1):
            v = p.get(col) if col != "Match Key" else p["key"]
            if v is None:
                continue
            cell = lp.cell(i, c, v); cell.font = F(color=GREY if col != "Match Key" else None)
            if col == "Launch Date": cell.number_format = DATE
            elif col in ("Butt", "Shaft", "MSRP", "Dealer ASP", "EST REV"): cell.number_format = "$#,##0"
            elif c >= 10: cell.number_format = "#,##0;\\(#,##0\\);\\-"
    last_lp = 4 + len(plan)
    add_table(lp, "tblLaunchPlan", f"A4:U{last_lp}")
    lp.freeze_panes = "C5"
    PK = f"'Launch Plan'!$U$5:$U${last_lp}"
    PC = {col: f"'Launch Plan'!${L(i + 1)}$5:${L(i + 1)}${last_lp}" for i, col in enumerate(LH)}

    # ============ Inbound lines ===========================================================
    ib = wb.create_sheet(f"Inbound {ref}")
    title(ib, f"Inbound {ref} — packing list lines as shipped",
          "Values from the HIE packing list. Type and Match Key are derived; the Allocation tab sums Qty by Match Key.", 8)
    hdr_row(ib, 4, ["PO", "Item", "Qty", "Remark", "Cartons", "Type", "PL Line", "Match Key"], [14, 36, 8, 26, 9, 8, 8, 36], height=22)
    for i, ln in enumerate(lines, start=5):
        for c, v in enumerate([ln["po"], ln["item"], ln["qty"], ln["remark"], ln["ctns"], ln["type"], ln["line"], ln["key"]], start=1):
            cell = ib.cell(i, c, v); cell.font = F(color=GREY if c <= 5 else None)
            if c in (3, 5): cell.number_format = "#,##0;\\(#,##0\\);\\-"
    last_ib = 4 + len(lines)
    add_table(ib, "tblInbound", f"A4:H{last_ib}")
    ib.cell(last_ib + 2, 2, "Total units (ties to packing list TOTAL)").font = F(bold=True)
    c = ib.cell(last_ib + 2, 3, f"=SUM(C5:C{last_ib})"); c.font = F(bold=True); c.number_format = "#,##0"
    ib.cell(last_ib + 3, 2, "Cue units").font = F()
    ib.cell(last_ib + 3, 3, f'=SUMIFS(C5:C{last_ib},F5:F{last_ib},"Cue")').number_format = "#,##0"
    ib.freeze_panes = "C5"
    IBQ, IBK = f"'Inbound {ref}'!$C$5:$C${last_ib}", f"'Inbound {ref}'!$H$5:$H${last_ib}"

    # ============ EMEA Q4 ==================================================================
    em = wb.create_sheet("EMEA Q4")
    title(em, "EMEA Q4 launch allocations — Direct vs via Jacksonville, and what this shipment covers",
          "Grey = EMEA_AllocationQ4.xlsx as received. Via JAX = EMEA units fulfilled from Jacksonville; Direct ships direct and is excluded. Blue = derived.", 10)
    hdr_row(em, 4, ["Launch", "SKU", "Launch Date", "Days to Launch", "EMEA Total", "EMEA Direct (excluded)", "EMEA via JAX", f"Arriving on {ref}", "Status", "Match Key"],
            [46, 32, 11, 9, 9, 10, 9, 10, 40, 32], height=30)
    emea_sorted = sorted(emea, key=lambda e: (e["launch"] or dt.datetime.max, e["item"]))
    for i, e in enumerate(emea_sorted, start=5):
        for c, v in enumerate([e["project"], e["item"], e["launch"], None, e["emea"], e["direct"], e["jax"], None, None, e["key"]], start=1):
            if v is not None:
                cell = em.cell(i, c, v); cell.font = F(color=GREY if c <= 7 else None)
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

    # ============ Plan Asks (from the Shipment Request workbook) =============================
    AK = None
    if asks:
        pa = wb.create_sheet("Plan Asks")
        title(pa, "Plan Asks — what the shipment plan requested from HIE (JAX Shipments by Date tab)",
              "Values exported from Shipment_Request_HIE - Updated.xlsm, 'JAX Shipments by Date'. Launch block: allocation, already received/invoiced in JAX, "
              "remaining JAX-bound need (net of receipts and the AP direct share), the 85% rule (MEETS/MOVE on projected completion by launch), and the units "
              "requested per sailing. CORE block: replenishment asks. CORE rows that duplicate a launch SKU were dropped. Match Key is derived.", 16)
            
        acols = ["sku", "block", "launch", "alloc", "recd", "jax_need", "rule", "flag"] + [c for c in asks[0] if c.startswith("ask_")] + ["finished", "completes", "held", "note", "instr"]
        alabels = {"sku": "SKU", "block": "Block", "launch": "Launch Date", "alloc": "Allocation", "recd": "Already Recd / Invoiced", "jax_need": "JAX-Bound Need (remaining; CORE = this sailing)",
                   "rule": "85% Rule", "flag": "Channel Gate Flag", "finished": "Finished at HIE (when asked)", "completes": "Completes by Cutoff", "held": "Held by Gate", "note": "Note", "instr": "Instruction"}
        hdrs = [alabels.get(c, "Ask " + c[4:6] + "/" + c[6:8]) for c in acols] + ["Match Key"]
        hdr_row(pa, 4, hdrs, [34, 8, 11, 9, 9, 10, 8, 12] + [9] * sum(c.startswith("ask_") for c in acols) + [10, 9, 8, 40, 26, 34], height=32)
        for i, r in enumerate(asks, start=5):
            for c, col in enumerate(acols, start=1):
                v = r.get(col)
                if col == "launch" and v not in (None, ""):
                    try: v = dt.datetime(1899, 12, 30) + dt.timedelta(days=int(float(v)))
                    except (TypeError, ValueError): pass
                if v in (None, ""): continue
                cell = pa.cell(i, c, v); cell.font = F(color=GREY)
                if col == "launch": cell.number_format = DATE
                elif isinstance(v, (int, float)): cell.number_format = "#,##0;\\(#,##0\\);\\-"
            pa.cell(i, len(acols) + 1, r["key"]).font = F()
        last_pa = 4 + len(asks)
        add_table(pa, "tblPlanAsks", f"A4:{L(len(acols) + 1)}{last_pa}")
        pa.freeze_panes = "C5"
        acol = {c: f"'Plan Asks'!${L(i + 1)}$5:${L(i + 1)}${last_pa}" for i, c in enumerate(acols)}
        AK = f"'Plan Asks'!${L(len(acols) + 1)}$5:${L(len(acols) + 1)}${last_pa}"

    # ============ Allocation (main) ========================================================
    al = wb.create_sheet("Allocation", 0)
    cues = OrderedDict()
    for ln in lines:
        if ln["type"] != "Cue":
            continue
        cues.setdefault(ln["key"], dict(item=ln["item"]))
    plan_keys = {p["key"] for p in plan}
    emea_keys = {e["key"] for e in emea}
    launch_dt = {}
    for p in plan:
        if p["Launch Date"] and p["key"] not in launch_dt:
            launch_dt[p["key"]] = p["Launch Date"]
    for e in emea:
        if e["launch"]:
            launch_dt.setdefault(e["key"], e["launch"])
    with_alloc = sorted([kv for kv in cues.items() if kv[0] in plan_keys or kv[0] in emea_keys],
                        key=lambda kv: (launch_dt.get(kv[0]) or dt.datetime.max, kv[1]["item"]))
    no_alloc = sorted([kv for kv in cues.items() if kv[0] not in plan_keys and kv[0] not in emea_keys], key=lambda kv: kv[1]["item"])

    CH = ["EMEA via JAX", "B2C", "NA (Americas)", "Trade show", "Sponsorship", "Marketing"]
    ask_lab = f"Requested for this sailing ({ask_col[4:6]}/{ask_col[6:8]})" if ask_col else "Requested for this sailing"
    AH = (["Launch", "SKU", "Launch Date", "Days to Launch", "Arriving", "Already in JAX (recd / invoiced)", "Remaining JAX Need (plan)", ask_lab, "Arriving vs Request", "85% Rule (plan)"]
          + CH + ["Total JAX Allowance"]
          + CH + ["To Stock", "Still Owed after this shipment", "Action", "AP (excluded)", "EMEA Direct (excluded)", "Pool", "Factor", "Residual"])
    AW = [38, 32, 11, 8, 9, 10, 10, 11, 9, 9] + [9, 8, 9, 8, 9, 9] + [9] + [9, 8, 9, 8, 9, 9] + [8, 10, 66, 8, 9, 7, 7, 7]
    NCOL = len(AH)
    cA, cRECD, cNEED, cREQ, cVS, cRULE = 5, 6, 7, 8, 9, 10
    cE = 11                                             # first plan allowance (EMEA via JAX)
    cP = {n: cE + i for i, n in enumerate(CH)}
    cTOT = cE + 6
    cN0 = cTOT + 1
    cS = {n: cN0 + i for i, n in enumerate(CH)}
    cSTK, cOWE, cACT = cN0 + 6, cN0 + 7, cN0 + 8
    cAP, cDIR, cPOOL, cFAC, cRES = cACT + 1, cACT + 2, cACT + 3, cACT + 4, cACT + 5
    col = lambda c, r: f"{L(c)}{r}"

    title(al, f"Packing list {ref} → Jacksonville: what arrived vs what was requested, and how to allocate it",
          f"Delivered to JAX {delivery:%m/%d/%Y} · HIE (Hamson) · {meta.get('VESSEL/VOYAGE', '')}. "
          "'Remaining JAX Need' and 'Requested' come from the shipment plan (Plan Asks tab): allocation less what Jacksonville already received, less the AP direct share, "
          "requested only where the launch MEETS the 85% rule. Shortfalls are measured against that request. Allowances (B2C, NA, Trade show, Sponsorship, Marketing from the Launch Plan; "
          "EMEA via JAX from the EMEA Q4 file) drive the channel split: EMEA first, the rest pro-rata. AP and EMEA Direct are excluded.", NCOL)
    al.cell(4, cACT, "As-of date").font = F(size=9, color=GREY)
    s1 = al.cell(4, cACT + 1, "=TODAY()"); s1.number_format = DATE
    al.cell(5, cACT, "Urgent if days to launch ≤").font = F(size=9, color=GREY)
    s2 = al.cell(5, cACT + 1, 7)
    for s_ in (s1, s2): s_.fill = INPUT; s_.font = F(color="FF0000FF")
    wb.defined_names["ASOF"] = DefinedName("ASOF", attr_text=f"Allocation!${L(cACT + 1)}$4")
    wb.defined_names["URGENT"] = DefinedName("URGENT", attr_text=f"Allocation!${L(cACT + 1)}$5")

    first = 9
    sec2 = first + len(with_alloc) + 1
    first2 = sec2 + 1
    last_all = first2 + len(no_alloc) - 1
    both = lambda c: f"${L(c)}${first}:${L(c)}${last_all}"
    band = [("Cue units arriving", f"=SUM({both(cA)})", NUM), ("Requested for this sailing", f"=SUM({both(cREQ)})", NUM),
            ("Under request (HIE shipped less than asked)", f'=-SUMIF({both(cVS)},"<0")', "[Red]#,##0;[Red]\\(#,##0\\);\\-"),
            ("Over request / not requested", f'=SUMIF({both(cVS)},">0")', NUM),
            ("Already in JAX before this shipment", f"=SUM({both(cRECD)})", NUM), ("Still owed after this shipment", f"=SUM({both(cOWE)})", NUM)] + \
           [(f"→ {n}", f"=SUM({both(cS[n])})", NUM) for n in CH] + [("→ To stock", f"=SUM({both(cSTK)})", NUM)]
    for i, (lab, f, nf) in enumerate(band, start=1):
        a = al.cell(4, i, lab); a.font = F(size=9, color=GREY); a.alignment = Alignment(wrap_text=True, vertical="bottom")
        v = al.cell(5, i, f); v.font = F(size=14, bold=True); v.number_format = nf
    al.row_dimensions[4].height = 30; al.row_dimensions[5].height = 22
    section(al, 6, f"LAUNCH ITEMS ON THIS SHIPMENT ({len(with_alloc)} SKUs) — sorted by launch date · then other cues", NCOL)
    groups = [(1, cA, "SHIPMENT"), (cRECD, cRULE, "SHIPMENT PLAN (Plan Asks tab)"), (cE, cTOT, "LAUNCH PLAN ALLOWANCES — Jacksonville channels"), (cN0, cOWE, "ALLOCATE THIS SHIPMENT"), (cACT, cACT, ""), (cAP, cRES, "Excluded / workings")]
    for a, b_, txt in groups:
        c = al.cell(7, a, txt or None); c.font = F(bold=True, size=9, color="FFFFFFFF"); c.alignment = Alignment(horizontal="center")
        c.fill = HDR_EX if a >= cAP else PatternFill("solid", fgColor="FF2E75B6" if a == cN0 else ("FF7A6A1E" if a == cRECD else "FF44546A"))
        if b_ > a: al.merge_cells(start_row=7, start_column=a, end_row=7, end_column=b_)
    fills = {**{c: HDR_EX for c in range(cAP, cRES + 1)}, **{c: PatternFill("solid", fgColor="FF2E75B6") for c in range(cN0, cOWE + 1)}, **{c: PatternFill("solid", fgColor="FF7A6A1E") for c in range(cRECD, cRULE + 1)}}
    hdr_row(al, 8, AH, AW, height=40, fills=fills)

    def write_row(r, k, a):
        E, TOT = col(cA, r), col(cTOT, r)
        RECD, NEED, REQ, VS, RULE = col(cRECD, r), col(cNEED, r), col(cREQ, r), col(cVS, r), col(cRULE, r)
        POOL, FAC, RES = col(cPOOL, r), col(cFAC, r), col(cRES, r)
        plan_rng = f"{col(cP['B2C'], r)}:{col(cP['Marketing'], r)}"
        al.cell(r, 1, f'=IFERROR(INDEX({PC["Project"]},MATCH("{k}",{PK},0)),IFERROR(INDEX({EM("A")},MATCH("{k}",{EM("J")},0)),""))')
        al.cell(r, 2, a["item"]).font = F(bold=True)
        al.cell(r, 3, f'=IFERROR(INDEX({PC["Launch Date"]},MATCH("{k}",{PK},0)),IFERROR(INDEX({EM("C")},MATCH("{k}",{EM("J")},0)),""))')
        al.cell(r, 4, f'=IF(C{r}="","",C{r}-ASOF)')
        al.cell(r, cA, f'=SUMIFS({IBQ},{IBK},"{k}")')
        if AK:
            has = f'COUNTIF({AK},"{k}")>0'
            al.cell(r, cRECD, f'=IF({has},SUMIFS({acol["recd"]},{AK},"{k}"),"")')
            al.cell(r, cNEED, f'=IF({has},SUMIFS({acol["jax_need"]},{AK},"{k}"),"")')
            al.cell(r, cREQ, f'=IF({has},SUMIFS({acol[ask_col]},{AK},"{k}"),"")')
            al.cell(r, cVS, f'=IF({REQ}="","",{E}-{REQ})')
            al.cell(r, cRULE, f'=IFERROR(INDEX({acol["rule"]},MATCH("{k}",{AK},0)),"")')
        else:
            for c in (cRECD, cNEED, cREQ, cVS, cRULE): al.cell(r, c, "")
        al.cell(r, cP["EMEA via JAX"], f'=SUMIFS({EM("G")},{EM("J")},"{k}")')
        for n, src in [("B2C", "B2C"), ("NA (Americas)", "NA"), ("Trade show", "Trade show"), ("Sponsorship", "Sponsorship"), ("Marketing", "Marketing")]:
            al.cell(r, cP[n], f'=SUMIFS({PC[src]},{PK},"{k}")')
        al.cell(r, cTOT, f"=SUM({col(cE, r)}:{col(cP['Marketing'], r)})")
        al.cell(r, cPOOL, f"={E}-{col(cS['EMEA via JAX'], r)}")
        al.cell(r, cFAC, f"=IF(SUM({plan_rng})=0,0,MIN(1,{POOL}/SUM({plan_rng})))")
        al.cell(r, cRES, f"=IF(AND({FAC}<1,SUM({plan_rng})>0),{POOL}-SUMPRODUCT(INT({plan_rng}*{FAC})),0)")
        al.cell(r, cS["EMEA via JAX"], f"=MIN({col(cP['EMEA via JAX'], r)},{E})")
        for pos, n in enumerate(["B2C", "NA (Americas)", "Trade show", "Sponsorship", "Marketing"], start=1):
            P = col(cP[n], r)
            al.cell(r, cS[n], f"=INT({P}*{FAC})+IF(AND({FAC}<1,MATCH(MAX({plan_rng}),{plan_rng},0)={pos}),{RES},0)")
        now_rng = f"{col(cS['EMEA via JAX'], r)}:{col(cS['Marketing'], r)}"
        al.cell(r, cSTK, f"={E}-SUM({now_rng})")
        al.cell(r, cOWE, f'=IF({NEED}<>"",MAX(0,{NEED}-{E}),MAX(0,{TOT}-{E}))')
        al.cell(r, cAP, f'=SUMIFS({PC["AP"]},{PK},"{k}")')
        al.cell(r, cDIR, f'=SUMIFS({EM("F")},{EM("J")},"{k}")')
        D, C, STK, OWE = f"D{r}", f"C{r}", col(cSTK, r), col(cOWE, r)
        urg = f'IF({D}="","",IF({D}<0,"LAUNCH PASSED — ",IF({D}<=URGENT,"URGENT ("&TEXT({D},"0")&" days) — ","")))'
        noplan = (f'IF({TOT}=0,"No Jacksonville allocation on file — all "&TEXT({E},"#,##0")&" to stock",'
                  f'IF({OWE}>0,"No plan request on file — distribute all "&TEXT({E},"#,##0")&" as shown; "&TEXT({OWE},"#,##0")&" of the allowance not yet covered",'
                  f'"Plan covered — distribute "&TEXT({TOT},"#,##0")&" per allowances; "&TEXT({STK},"#,##0")&" to stock"))')
        withplan = (f'IF(AND({REQ}=0,{RULE}="MOVE"),"NOT REQUESTED — launch held by the 85% gate (MOVE); "&TEXT({E},"#,##0")&" arrived anyway — hold in stock until the gate clears",'
                    f'IF({REQ}=0,"Not requested on this sailing — "&TEXT({E},"#,##0")&" arrived; distribute as shown or hold",'
                    f'IF({VS}<0,"SHORT "&TEXT(-{VS},"#,##0")&" vs the "&TEXT({REQ},"#,##0")&" requested — HIE shipped less than asked; distribute the "&TEXT({E},"#,##0")&" as shown, chase HIE for the balance",'
                    f'IF({VS}>0,"Over request by "&TEXT({VS},"#,##0")&" — distribute as shown; extra to stock",'
                    f'"Matches request ("&TEXT({REQ},"#,##0")&") — distribute as shown"))))')
        al.cell(r, cACT, f'={urg}&IF({REQ}="",{noplan},{withplan})')
        for c in range(1, NCOL + 1):
            cell = al.cell(r, c)
            if c != 2: cell.font = F(color=GREY if c >= cAP else None)
            if c == 3: cell.number_format = DATE
            if c == 4: cell.number_format = "0;[Red]-0"
            if c == cVS: cell.number_format = "+#,##0;[Red]-#,##0;0"
            elif c == cFAC: cell.number_format = "0.00"
            elif c >= 5 and c not in (cACT, cRULE): cell.number_format = NUM
            if c == cACT: cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.border = Border(bottom=THIN)
        al.row_dimensions[r].height = 28

    for i, (k, a) in enumerate(with_alloc):
        write_row(first + i, k, a)
    section(al, sec2, f"OTHER CUES ON THIS SHIPMENT ({len(no_alloc)} SKUs) — no launch allowance; replenishment (CORE) asks shown where the plan had one", NCOL)
    for i, (k, a) in enumerate(no_alloc):
        write_row(first2 + i, k, a)
    tot = last_all + 1
    al.cell(tot, 2, "TOTAL — all cues").font = F(bold=True)
    for c in [cA, cRECD, cNEED, cREQ, cVS] + list(range(cE, cTOT + 1)) + list(range(cN0, cOWE + 1)) + [cAP, cDIR]:
        x = al.cell(tot, c, f"=SUM({L(c)}{first}:{L(c)}{last_all})"); x.font = F(bold=True, color=GREY if c >= cAP else None); x.number_format = NUM if c != cVS else "+#,##0;[Red]-#,##0;0"
    al.freeze_panes = "C9"
    rng = lambda c: f"{L(c)}{first}:{L(c)}{last_all}"
    al.conditional_formatting.add(rng(4), CellIsRule(operator="lessThan", formula=["0"], fill=RED, font=Font(color="FFB3261E", bold=True)))
    al.conditional_formatting.add(rng(4), FormulaRule(formula=[f"AND(ISNUMBER(D{first}),D{first}>=0,D{first}<=URGENT)"], fill=AMB))
    al.conditional_formatting.add(rng(cVS), CellIsRule(operator="lessThan", formula=["0"], fill=RED, font=Font(color="FFB3261E", bold=True)))
    al.conditional_formatting.add(rng(cVS), FormulaRule(formula=[f'AND(ISNUMBER({L(cVS)}{first}),{L(cVS)}{first}=0)'], fill=GRN))
    al.conditional_formatting.add(rng(cRULE), FormulaRule(formula=[f'{L(cRULE)}{first}="MOVE"'], fill=RED, font=Font(color="FFB3261E", bold=True)))
    for c in range(cN0, cN0 + 6):
        al.conditional_formatting.add(rng(c), CellIsRule(operator="greaterThan", formula=["0"], fill=GRN, font=Font(bold=True)))
    al.conditional_formatting.add(rng(cOWE), CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="FFB3261E")))
    al.conditional_formatting.add(rng(cACT), FormulaRule(formula=[f'OR(LEFT({L(cACT)}{first},6)="URGENT",LEFT({L(cACT)}{first},6)="LAUNCH")'], font=Font(color="FFB26A00", bold=True)))
    al.conditional_formatting.add(rng(cACT), FormulaRule(formula=[f'ISNUMBER(SEARCH("SHORT ",{L(cACT)}{first}))'], font=Font(color="FFB3261E", bold=True)))
    al.column_dimensions.group(L(cPOOL), L(cRES), hidden=True)
    others = [ln for ln in lines if ln["type"] != "Cue"]
    al.cell(tot + 2, 1, f"Not allocated: {len(others)} non-cue lines ({sum(l['qty'] for l in others):,} units — samples and one shaft) on the Inbound {ref} tab. "
            "EMEA launch items NOT on this shipment: EMEA Q4 tab, red status. 'Still owed' = remaining JAX need from the plan less this shipment (falls back to allowance less arriving when the plan has no row). "
            "Grey columns at the right are excluded from the allocation (AP, EMEA Direct); Pool / Factor / Residual are the pro-rata workings (grouped, click + to expand).").font = F(size=9, color=GREY)
    al.merge_cells(start_row=tot + 2, start_column=1, end_row=tot + 2, end_column=NCOL)

    wb._sheets = [wb["Allocation"]] + ([wb["Plan Asks"]] if asks else []) + [wb["EMEA Q4"], wb["Launch Plan"], wb[f"Inbound {ref}"]]
    wb.save(out)
    print(f"wrote {out}: {len(with_alloc)} launch SKUs + {len(no_alloc)} other cues, {len(lines)} lines, {len(emea)} EMEA rows, {len(plan)} plan rows")


if __name__ == "__main__":
    build(*sys.argv[1:7])
