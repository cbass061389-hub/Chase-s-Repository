"""Turn the weekly OOL long table into capacity, slippage, reject and revenue-timing metrics.

Line identity: HIE splits each PO+item into delivery lines keyed by PI date. Lines are tracked week to week
within (PO, item): exact PI-date match first, then the remaining lines are paired by order qty and nearest
PI date (that is a reschedule). Each tracked line keeps its FIRST-SEEN PI / scheduled-finish date (the original
commitment) and its latest state.

Definitions (all units):
  good      = shipped + packing + warehouse
  rejects   = bgrade + defect
  completed = good + rejects                     (units through production)
  weekly output = increase in completed per tracked line between consecutive reports
  finish date   = actual_finish if HIE filled it, else the report date where online first hit 0 with good>0
  slip days     = finish date - ORIGINAL PI date (HIE's committed date, first time the line appeared)
  overdue at a date = units not yet good on lines whose original scheduled finish is before that date
"""
import csv, os, json
from collections import defaultdict
from datetime import date, datetime, timedelta

D = os.path.dirname(os.path.abspath(__file__))
TODAY = date(2026, 11, 6)


def d(s):
    return datetime.strptime(s, '%Y-%m-%d').date() if s else None


def f(s):
    try:
        return float(s)
    except (TypeError, ValueError):
        return 0.0


rows = list(csv.DictReader(open(os.path.join(D, 'ool_lines.csv'), encoding='utf-8')))
for r in rows:
    r['rd'] = d(r['report_date'])
    for k in ('pi_date', 'sched_finish', 'actual_finish'):
        r[k] = d(r[k])
    for k in ('order_qty', 'prod_qty', 'shipped', 'packing', 'warehouse', 'bgrade', 'defect', 'makeup', 'online'):
        r[k] = f(r[k])
    r['good'] = r['shipped'] + r['packing'] + r['warehouse']
    r['rej'] = r['bgrade'] + r['defect']
    r['done'] = r['good'] + r['rej']

reports = sorted({r['rd'] for r in rows})
by_rep = defaultdict(lambda: defaultdict(list))
for r in rows:
    by_rep[r['rd']][(r['po'], r['item'])].append(r)

# ---------------- line tracking
lines = {}           # lid -> dict
cur = {}             # (po,item) -> list of (lid,row) in previous report
nid = 0
obs = []             # (lid, rd, row)
for rd in reports:
    nxt = {}
    for key, rs in by_rep[rd].items():
        prev = list(cur.get(key, []))
        assigned = []
        rest = []
        for r in rs:
            hit = next((p for p in prev if p[1]['pi_date'] == r['pi_date'] and p[1]['order_qty'] == r['order_qty']), None) \
                or next((p for p in prev if p[1]['pi_date'] == r['pi_date']), None)
            if hit:
                prev.remove(hit); assigned.append((hit[0], r))
            else:
                rest.append(r)
        for r in rest:
            # reschedule = a still-open line of the same qty whose PI date moved; never re-use a completed line
            cands = [p for p in prev if p[1]['order_qty'] == r['order_qty'] and p[1]['done'] < p[1]['order_qty']
                     and r['done'] >= p[1]['done'] - 0.5]
            if cands and r['pi_date']:
                hit = min(cands, key=lambda p: abs(((p[1]['pi_date'] or r['pi_date']) - r['pi_date']).days))
                prev.remove(hit); assigned.append((hit[0], r))
            else:
                nid += 1
                lines[nid] = {'po': key[0], 'item': key[1], 'sheet': r['sheet'], 'first_rd': rd,
                              'orig_pi': r['pi_date'], 'orig_sched': r['sched_finish'], 'first_done': r['done'],
                              'first_online': r['online']}
                assigned.append((nid, r))
        nxt[key] = assigned
        for lid, r in assigned:
            L = lines[lid]
            L.update({'last_rd': rd, 'order_qty': r['order_qty'], 'pi': r['pi_date'], 'sched': r['sched_finish'],
                      'good': r['good'], 'bgrade': r['bgrade'], 'defect': r['defect'], 'makeup': r['makeup'],
                      'online': r['online'], 'done': r['done'], 'sheet': r['sheet']})
            complete = r['order_qty'] > 0 and r['done'] >= r['order_qty'] - 0.5
            if complete and 'reached_rd' not in L:
                L['reached_rd'] = rd
                # HIE's 'Actually Finished Date' is a forecast on open lines; trust it only once the line is complete
                af = r['actual_finish']
                L['actual_finish'] = af if (af and af <= rd and (L.get('prev_rd') is None or af > L['prev_rd'] - timedelta(days=60))) else None
            L['prev_rd'] = rd
            if not L.get('orig_pi') and r['pi_date']:
                L['orig_pi'] = r['pi_date']; L['pi_basis'] = 'PI (first filled)'
            if not L.get('orig_pi') and not L.get('sched_fallback') and r['sched_finish']:
                L['sched_fallback'] = r['sched_finish'] + timedelta(days=14)
            n_resched = L.setdefault('resched', 0)
            if L.get('_lastpi') and r['pi_date'] and r['pi_date'] != L['_lastpi']:
                L['resched'] = n_resched + 1
            L['_lastpi'] = r['pi_date']
            obs.append((lid, rd, r))
    cur = nxt

first_rep, last_rep = reports[0], reports[-1]
_pf = os.path.join(D, '..', 'pi', 'pi_first_dates.json')
PI_FIRST = json.load(open(_pf)) if os.path.exists(_pf) else {}
_grp = defaultdict(list)
for lid, L in lines.items():
    _grp[(L['po'], ' '.join(L['item'].upper().split()))].append(lid)
for key, lids in _grp.items():
    pf = PI_FIRST.get(key[0] + '|' + key[1])
    if not pf:
        continue
    lids.sort(key=lambda i: (lines[i].get('orig_pi') or lines[i].get('sched_fallback') or date(2100, 1, 1), lines[i]['first_rd']))
    for rank, i in enumerate(lids):
        if rank < len(pf):
            lines[i]['ool_pi'] = lines[i].get('orig_pi')
            lines[i]['orig_pi'] = datetime.strptime(pf[rank], '%Y-%m-%d').date()
            lines[i]['pi_basis'] = 'HIE PI email (first quote)'
for lid, L in lines.items():
    if L.get('pi_basis') == 'HIE PI email (first quote)':
        pass
    elif L.get('orig_pi'):
        L.setdefault('pi_basis', 'PI date')
    elif L.get('sched_fallback'):
        L['orig_pi'] = L['sched_fallback']; L['pi_basis'] = 'Sched finish + 14d'
    else:
        L['pi_basis'] = 'none'
    L['closed'] = L['last_rd'] < last_rep           # dropped off the list = fully shipped/closed
    fin = L.get('actual_finish') or L.get('reached_rd')
    L['dropped'] = False
    if not fin and L['closed']:
        if L['done'] >= 0.95 * L['order_qty'] and L['order_qty'] > 0:
            nxt = next((x for x in reports if x > L['last_rd']), None)   # dropped off complete: finished by next report
            fin = nxt or L['last_rd']
        else:
            L['dropped'] = True   # vanished before being produced = cancelled / re-split by HIE; no slip measured
    L['finish'] = fin
    # usable for slippage only if we saw the line before it finished
    L['observed_open'] = L['first_online'] > 0 or L['first_done'] < L['order_qty']
    L['slip'] = (fin - L['orig_pi']).days if (fin and L['orig_pi'] and L['observed_open']) else None
    L['open_late'] = (not fin) and L['orig_pi'] and L['orig_pi'] < TODAY
    L['days_overdue_now'] = (TODAY - L['orig_pi']).days if L['open_late'] else None

# ---------------- weekly output (completed-unit deltas)
last_done = {}
week_out = defaultdict(float); week_good = defaultdict(float); week_rej = defaultdict(float)
sku_week = defaultdict(float)
prev_rep = {}
for i, rd in enumerate(reports):
    span = (rd - reports[i - 1]).days if i else 7
    prev_rep[rd] = span
for lid, rd, r in obs:
    p = last_done.get(lid)
    L = lines[lid]
    if p is None:
        base = r['done'] if rd == first_rep or not L['observed_open'] else 0.0
        pg, pr = (r['good'], r['rej']) if base else (0.0, 0.0)
        last_done[lid] = (r['done'], r['good'], r['rej']) if base else (0.0, 0.0, 0.0)
        if base:
            continue
        p = (0.0, 0.0, 0.0)
    dd, dg, dr = r['done'] - p[0], r['good'] - p[1], r['rej'] - p[2]
    if dd > 0:
        week_out[rd] += dd; sku_week[(L['item'], rd)] += dd
    if dg > 0: week_good[rd] += dg
    if dr > 0: week_rej[rd] += dr
    last_done[lid] = (r['done'], r['good'], r['rej'])

# ---------------- weekly plan vs actual (HIE's own schedule grid, as published the report before)
plan_by_week = defaultdict(float)
plan_rows = list(csv.DictReader(open(os.path.join(D, 'ool_plan.csv'), encoding='utf-8')))
pp = defaultdict(lambda: defaultdict(float))
for p in plan_rows:
    pp[d(p['report_date'])][d(p['week'])] += f(p['qty'])
for i, rd in enumerate(reports[1:], 1):
    prv = reports[i - 1]
    plan_by_week[rd] = sum(q for wk, q in pp[prv].items() if prv < wk <= rd)

# ---------------- overdue snapshot each report (units past ORIGINAL commitment, not yet good)
overdue_units = defaultdict(float); overdue_lines = defaultdict(int)
overdue_sku = defaultdict(lambda: defaultdict(float))
for lid, rd, r in obs:
    L = lines[lid]
    if L['orig_pi'] and L['orig_pi'] < rd and L['observed_open']:
        rem = max(0.0, r['order_qty'] - r['good'] - r['rej'])
        if rem > 0:
            overdue_units[rd] += rem; overdue_lines[rd] += 1
            overdue_sku[rd][L['item']] += rem

json.dump({'reports': [str(x) for x in reports]}, open(os.path.join(D, 'reports.json'), 'w'))

# ---------------- export
with open(os.path.join(D, 'lines_tracked.csv'), 'w', newline='', encoding='utf-8') as fo:
    cols = ['lid', 'po', 'item', 'sheet', 'first_rd', 'last_rd', 'order_qty', 'orig_pi', 'orig_sched', 'pi', 'sched',
            'finish', 'slip', 'resched', 'good', 'bgrade', 'defect', 'makeup', 'online', 'closed', 'observed_open',
            'days_overdue_now', 'dropped', 'pi_basis', 'ool_pi']
    w = csv.writer(fo); w.writerow(cols)
    for lid, L in lines.items():
        w.writerow([lid] + [L.get(c) for c in cols[1:]])
with open(os.path.join(D, 'weekly.csv'), 'w', newline='') as fo:
    w = csv.writer(fo); w.writerow(['report_date', 'days_since_prior', 'completed', 'good', 'rejects', 'hie_plan',
                                   'overdue_units', 'overdue_lines'])
    for rd in reports:
        w.writerow([rd, prev_rep[rd], week_out[rd], week_good[rd], week_rej[rd], plan_by_week.get(rd, ''),
                    overdue_units[rd], overdue_lines[rd]])
with open(os.path.join(D, 'overdue_sku.csv'), 'w', newline='') as fo:
    w = csv.writer(fo); w.writerow(['report_date', 'item', 'overdue_units'])
    for rd in reports:
        for it, q in overdue_sku[rd].items():
            w.writerow([rd, it, q])
with open(os.path.join(D, 'sku_week_output.csv'), 'w', newline='') as fo:
    w = csv.writer(fo); w.writerow(['item', 'report_date', 'completed'])
    for (it, rd), q in sku_week.items():
        w.writerow([it, rd, q])
s = [L['slip'] for L in lines.values() if L['slip'] is not None]
print(len(reports), 'reports', len(lines), 'tracked lines', len(s), 'with slip; median slip',
      sorted(s)[len(s) // 2] if s else None)
