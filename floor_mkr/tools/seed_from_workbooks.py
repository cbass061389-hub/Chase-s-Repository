"""Regenerate FLOOR-MKR seed documents (one JSON per document) under a target folder.

Sources: the TABLE Floor rows extracted from Table_Production.xlsm (see the extraction
in the session notes), the WOOD batch plan (wood_production/tools/build_wood_production.py),
and the REVO op list from REVO_Ops. Write them with the ArtifactData tool as batches of 50.

Usage: python floor_mkr/tools/seed_from_workbooks.py <table_units.json> <out_dir>
"""
import json, math, os, re, sys
from datetime import date, timedelta

def seg(s): return re.sub(r"[^A-Za-z0-9_\-.~:@+]", "_", s)[:180]

def main(table_units_path, out):
    now = "2026-11-08T22:30:00Z"
    def w(coll, did, data):
        os.makedirs(f"{out}/{coll}", exist_ok=True)
        json.dump(data, open(f"{out}/{coll}/{did}.json", "w"), separators=(",", ":"))
    revo_ops = ["IC", "10", "20", "30", "100", "40", "110", "50", "60", "70", "120", "130", "160"]
    w("config", "lines", {"lines": {
        "REVO": {"name": "REVO carbon shafts", "unitNoun": "cart", "color": "var(--line-REVO)",
                 "stations": [{"id": o, "name": "Op " + o, "zone": "", "restDays": 0} for o in revo_ops]},
        "TABLE": {"name": "Pool tables", "unitNoun": "table", "color": "var(--line-TABLE)", "stations": [
            {"id": "PreAssym", "name": "Pre-assembly", "zone": "Pre-Assembly", "restDays": 0},
            {"id": "ST01", "name": "Kitting, Base & Frame", "zone": "Pre-Build", "restDays": 0},
            {"id": "ST02", "name": "Leveling, Ball Return & Legs", "zone": "OP1", "restDays": 0},
            {"id": "ST03", "name": "Slate Install & Cloth", "zone": "OP2", "restDays": 0},
            {"id": "ST04", "name": "Rails Prep & Install", "zone": "OP2", "restDays": 0},
            {"id": "ST05", "name": "QC, Serial, ERP & Move to Warehouse", "zone": "Finishing", "restDays": 0}]},
        "WOOD": {"name": "Panthera cue butts", "unitNoun": "cue", "color": "var(--line-WOOD)", "stations": [
            {"id": "WS01", "name": "Receive, inspect & cut blanks", "zone": "Pre-Build", "restDays": 0},
            {"id": "WS02", "name": "Rough turn forearm & handle to 1.300 straight", "zone": "Turning", "restDays": 14},
            {"id": "WS03", "name": "Rough taper & forearm coring", "zone": "Turning", "restDays": 14},
            {"id": "WS04", "name": "Cut hardware; A-joint, weight cartridge, butt cap", "zone": "Machining", "restDays": 0},
            {"id": "WS05", "name": "Joint pin & preform glue-up", "zone": "Assembly", "restDays": 1},
            {"id": "WS06", "name": "Final taper 1.350 to 1.250", "zone": "Turning", "restDays": 0},
            {"id": "WS07", "name": "Independent inspection, robot test, clear-coat prep", "zone": "Finishing", "restDays": 0}]}},
        "updatedAt": now})
    w("settings", "targets", {"attainment": 0.95, "fpy": 0.9, "reject": 0.05, "rework": 0.1, "updatedAt": now})
    for r in ["Warpage / movement after cut", "Moisture content out of range", "Runout — not straight", "Dimension out of tolerance",
              "Chatter / tool marks", "Grain defect / check / crack", "Joint fit (A-joint)", "Glue line / glue-up failure",
              "Pin dot / surface blemish", "Weight cartridge hole off", "Operator error / wrong setup sheet", "Material shortage (hardware)"]:
        w("reasons", seg("WOOD_" + re.sub(r"\s+", "_", r.lower())), {"line": "WOOD", "text": r, "timesUsed": 0, "firstSeen": "", "lastUsed": "", "active": True, "category": "", "seed": True})
    batches = [("B1 — Prototype — Broomstick 1", "PAN-BRM-001", 1, 0.0, date(2026, 12, 14), date(2026, 12, 14), "2026-12-18", "WO-PAN-001"),
               ("B2 — 25 identical broomsticks", "PAN-BRM-001", 25, 0.10, date(2027, 1, 4), date(2027, 1, 18), "2027-01-20", "WO-PAN-002"),
               ("B3 — 50 sellable cues (STR3)", "PAN-SELL-001", 50, 0.10, date(2027, 3, 29), date(2027, 6, 7), "2027-06-11", "WO-PAN-003")]
    n = 0; plan = {}
    for label, sku, good, allow, start, end, target, wo in batches:
        starts = math.ceil(round(good * (1 + allow), 6)); weeks = (end - start).days // 7 + 1; base = starts // weeks; rem = starts - base * weeks
        for i in range(weeks):
            wk = (start + timedelta(days=7 * i)).isoformat(); q = base + (rem if i == weeks - 1 else 0)
            if q: plan[(wk, sku)] = plan.get((wk, sku), 0) + q
        for _ in range(starts):
            n += 1; uid = f"C-{n:04d}"
            w("units", f"WOOD_{uid}", {"line": "WOOD", "unitId": uid, "wo": wo, "sku": sku, "serial": "", "qty": 1, "batch": label, "target": target,
                                        "stations": {}, "stationDates": {}, "status": "", "statusDate": "", "order": n, "createdAt": now, "seed": True})
    for (wk, sku), q in plan.items():
        w("plan", f"WOOD_{wk}_{seg(sku)}", {"line": "WOOD", "weekStart": wk, "sku": sku, "qty": q, "updatedAt": now})
    for i, u in enumerate(json.load(open(table_units_path)), 1):
        w("units", f"TABLE_{u['unitId']}", {"line": "TABLE", "unitId": u["unitId"], "wo": u["wo"], "sku": u["sku"], "serial": u["serial"], "qty": u["qty"],
                                            "batch": u["batch"], "target": u["target"] or "", "stations": {k: v for k, v in u["stations"].items() if v},
                                            "stationDates": {}, "status": "", "statusDate": "", "order": i, "createdAt": now, "seed": True})

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
