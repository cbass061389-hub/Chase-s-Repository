"""Regenerate the REVO seed documents for FLOOR-MKR from Revo_Production - New.xlsm.

Reads: REVO Floor (rows 9+: WO#, SKU, CART, Qty, Forecast, DATE, op ticks, Released To Date,
Remaining; group labels in column A), Plan Config (op order, cure flags), Quality_Codes,
Holds, Production Plan (row 2 = Friday week-end dates -> Monday week-start), Release Log and
Reject Log. Writes one JSON per document under <out_dir>/<collection>/<doc_id>.json and
replaces the REVO entry in an existing config/lines.json. Write them with the ArtifactData
tool in batches of 50 (config/lines needs if_version).

Usage: python floor_mkr/tools/seed_revo.py <revo.xlsm> <existing config lines.json> <out_dir> [since YYYY-MM-DD]
"""
import json, os, re, sys
from datetime import datetime, timedelta

import openpyxl


def seg(s): return re.sub(r"[^A-Za-z0-9_\-.~:@+]", "_", str(s))[:180]


def main(xlsm, config_json, out, since="2026-08-10"):
    now = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    def w(coll, did, data):
        os.makedirs(f"{out}/{coll}", exist_ok=True)
        json.dump(data, open(f"{out}/{coll}/{did}.json", "w"), separators=(",", ":"))
    wb = openpyxl.load_workbook(xlsm, keep_vba=True, data_only=True)
    fl = wb["REVO Floor"]
    ops = [str(fl.cell(row=9, column=c).value) for c in range(8, 23)]
    qc = wb["Quality_Codes"]
    col = lambda c: [qc.cell(row=r, column=c).value for r in range(5, 40) if qc.cell(row=r, column=c).value]
    cfg = json.load(open(config_json))
    cure = {"100", "40"}
    cfg["lines"]["REVO"] = {"name": "REVO carbon shafts", "unitNoun": "cart", "color": "var(--line-REVO)", "partialRelease": True,
        "stations": [{"id": o, "name": ("Incoming inspection gate" if o == "IC" else "Assembly (parallel, off main line)" if o == "ASSYM" else "Final QC" if o == "FQC" else f"Op {o}") + (" (cure before next op)" if o in cure else ""), "zone": "", "restDays": 1 if o in cure else 0} for o in ops],
        "quality": {"defectTypes": col(1), "locations": col(2), "rootCauses": col(3), "correctiveActions": col(5)}}
    cfg["updatedAt"] = now
    w("config", "lines", cfg)
    for t in cfg["lines"]["REVO"]["quality"]["defectTypes"]:
        w("reasons", seg("REVO_" + re.sub(r"\s+", "_", t.lower())), {"line": "REVO", "text": t, "timesUsed": 0, "firstSeen": "", "lastUsed": "", "active": True, "category": "Defect type", "seed": True})
    group, order = None, 0
    for r in range(10, fl.max_row + 1):
        a = fl.cell(row=r, column=1).value
        if isinstance(a, str) and a.strip(): group = a.strip()
        cart, sku, qty = fl.cell(row=r, column=4).value, fl.cell(row=r, column=3).value, fl.cell(row=r, column=5).value
        wo = fl.cell(row=r, column=2).value or ""
        if cart in (None, "") or not sku or qty is None or str(wo).startswith("CART LOOKUP"): continue
        order += 1
        d = fl.cell(row=r, column=7).value; d = d.date().isoformat() if isinstance(d, datetime) else ""
        rel = fl.cell(row=r, column=25).value or 0; rem = fl.cell(row=r, column=26).value; rem = qty if rem is None else rem
        w("units", f"REVO_{seg(cart)}", {"line": "REVO", "unitId": str(cart), "wo": wo, "sku": sku, "serial": "", "qty": qty, "batch": f"{group} · due {datetime.fromisoformat(d).strftime('%d %b')}" if d else group,
            "target": d, "stations": {o: True for i, o in enumerate(ops) if fl.cell(row=r, column=8 + i).value}, "stationDates": {},
            "status": "RELEASED" if rem == 0 and rel > 0 else "", "statusDate": d if rem == 0 and rel > 0 else "", "qtyReleased": rel, "qtyOut": max(0, qty - rel - rem), "order": order, "createdAt": now, "seed": True})
    for i, row in enumerate([r for r in wb["Holds"].iter_rows(min_row=2, values_only=True) if r[0] is not None], 1):
        cart, op, reason, d, released = row[:5]; d = d.date().isoformat() if isinstance(d, datetime) else str(d)
        w("holds", f"REVO_hold_{i:02d}", {"line": "REVO", "unitDoc": f"REVO_{seg(cart)}", "unitId": str(cart), "station": str(op), "reason": (reason or "").strip(), "dateHeld": d, "released": bool(released), "releasedAt": d if released else "", "by": None, "ts": d + "T08:00:00Z", "seed": True})
    pp = wb["Production Plan"]
    dates = {c: pp.cell(row=2, column=c).value for c in range(2, pp.max_column + 1) if isinstance(pp.cell(row=2, column=c).value, datetime)}
    for r in range(5, pp.max_row + 1):
        sku = pp.cell(row=r, column=1).value
        if not sku: continue
        for c, d in dates.items():
            q = pp.cell(row=r, column=c).value
            if q:
                mon = (d - timedelta(days=4)).date().isoformat()
                w("plan", f"REVO_{mon}_{seg(sku)}", {"line": "REVO", "weekStart": mon, "sku": sku, "qty": q, "updatedAt": now})
    rej = {}
    for row in wb["Reject Log"].iter_rows(min_row=2, values_only=True):
        if isinstance(row[0], datetime): rej.setdefault((row[0].date().isoformat(), str(row[3])), []).append(row[5] or "")
    n = 0
    for row in wb["Release Log"].iter_rows(min_row=4, values_only=True):
        if not isinstance(row[0], datetime) or row[0].date().isoformat() < since: continue
        d, sku, cart, qrel, wo, qrw, qbg, qrj = row[0].date().isoformat(), row[1], str(row[2]), row[3] or 0, row[4] or "", row[5] or 0, row[6] or 0, row[7] or 0
        n += 1; reason = "; ".join(rej.pop((d, cart), []))
        out_ = "RELEASED" if qrel > 0 else ("REJECTED" if qrj >= max(qrw, qbg) else ("REWORK" if qrw >= qbg else "B-GRADE"))
        w("events", f"REVO_ev_{n:04d}", {"ts": d + "T12:00:00Z", "date": d, "line": "REVO", "unitDoc": f"REVO_{seg(cart)}", "unitId": cart, "wo": wo, "sku": sku, "batch": "", "serial": "", "outcome": out_, "station": "FQC", "reason": reason, "notes": "imported from Release Log" + (" + Reject Log" if reason else ""), "zone": "",
            "qtyPass": qrel + qrw + qbg + qrj, "qtyReleased": qrel, "qtyRework": qrw, "qtyBGrade": qbg, "qtyReject": qrj, "qtyHold": 0, "details": [], "by": None, "seed": True})
    for (d, cart), descs in rej.items():
        if d < since: continue
        n += 1
        w("events", f"REVO_ev_{n:04d}", {"ts": d + "T12:00:00Z", "date": d, "line": "REVO", "unitDoc": f"REVO_{seg(cart)}", "unitId": cart, "wo": "", "sku": "", "batch": "", "serial": "", "outcome": "REJECTED", "station": "FQC", "reason": "; ".join(descs), "notes": "imported from Reject Log (no matching release row)", "zone": "",
            "qtyPass": len(descs), "qtyReleased": 0, "qtyRework": 0, "qtyBGrade": 0, "qtyReject": len(descs), "qtyHold": 0, "details": [], "by": None, "seed": True})


if __name__ == "__main__":
    main(*sys.argv[1:])
