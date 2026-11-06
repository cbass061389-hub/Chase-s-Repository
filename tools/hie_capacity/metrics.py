"""Aggregate tracked lines + weekly deltas into the JSON the HTML report renders."""
import csv, json, os, statistics, re
from collections import defaultdict
from datetime import date, datetime, timedelta

D = os.path.dirname(os.path.abspath(__file__))
TODAY = date(2026, 11, 6)
GRACE = 7  # days past original commitment before a unit counts as late


def d(s):
    return datetime.strptime(s, '%Y-%m-%d').date() if s and s != 'None' else None


def f(s):
    try:
        return float(s)
    except (TypeError, ValueError):
        return 0.0


def mkey(x):
    return f'{x.year}-{x.month:02d}'


# ---------------- prices
price, pbasis, cat = {}, {}, {}
for r in csv.DictReader(open(os.path.join(D, 'prices.csv'))):
    k = re.sub(r'\s+', ' ', r['sku'].strip().upper())
    if r['sell_price']:
        price[k] = float(r['sell_price']); pbasis[k] = r['price_basis']
    cat[k] = r['category']
FALLBACK = statistics.median(price.values())


_fam = {}


def p_of(item):
    k = re.sub(r'\s+', ' ', item.strip().upper())
    if k in price:
        return price[k], pbasis[k]
    if k in _fam:
        return _fam[k]
    toks = k.replace('-', ' ').split()
    for n in range(len(toks) - 1, 1, -1):          # longest shared prefix with priced SKUs (>=2 tokens)
        pre = ' '.join(toks[:n])
        hits = [v for s_, v in price.items() if s_.replace('-', ' ').startswith(pre + ' ') or s_.replace('-', ' ') == pre]
        if len(hits) >= 2 or (hits and n >= 4):
            _fam[k] = (statistics.median(hits), f'Family median ({pre})')
            return _fam[k]
    _fam[k] = (FALLBACK, 'Unmatched SKU - median price')
    return _fam[k]


def line_of(item, sheet):
    k = re.sub(r'\s+', ' ', item.strip().upper())
    c = cat.get(k, '')
    if c:
        return c.split(' : ')[-1]
    return {'P2 (UNITBC)': 'Cues (P2)', 'shaft': 'Shafts', 'Carom': 'Carom', 'Radial RD': 'Radial', 'AIR2 & EXTENDER':
            'Jump/Ext', 'ACC': 'Accessories', 'cues': 'Cues', 'Partial': 'Partial', 'PO': 'PO'}.get(sheet, sheet)


L = list(csv.DictReader(open(os.path.join(D, 'lines_tracked.csv'), encoding='utf-8')))
for x in L:
    for k in ('first_rd', 'last_rd', 'orig_pi', 'orig_sched', 'pi', 'sched', 'finish'):
        x[k] = d(x[k])
    for k in ('order_qty', 'good', 'bgrade', 'defect', 'makeup', 'online'):
        x[k] = f(x[k])
    x['slip'] = int(float(x['slip'])) if x['slip'] not in ('', 'None') else None
    x['resched'] = int(float(x['resched'] or 0))
    x['observed_open'] = x['observed_open'] == 'True'
    x['price'], x['pbasis'] = p_of(x['item'])
    x['pline'] = line_of(x['item'], x['sheet'])

W = list(csv.DictReader(open(os.path.join(D, 'weekly.csv'))))
for w in W:
    w['rd'] = d(w['report_date'])
    for k in ('days_since_prior', 'completed', 'good', 'rejects', 'overdue_units', 'overdue_lines'):
        w[k] = f(w[k])
reports = [w['rd'] for w in W]
first, last = reports[0], reports[-1]

# report gaps (missing weekly files)
gaps = [(str(a), str(b), (b - a).days) for a, b in zip(reports, reports[1:]) if (b - a).days > 10]

# ---------------- monthly capacity (spread each interval's output evenly over its days)
mon_done, mon_good, mon_rej, mon_days = defaultdict(float), defaultdict(float), defaultdict(float), defaultdict(int)
for w in W[1:]:
    days = int(w['days_since_prior'])
    for i in range(days):
        day = w['rd'] - timedelta(days=i)
        mon_done[mkey(day)] += w['completed'] / days
        mon_good[mkey(day)] += w['good'] / days
        mon_rej[mkey(day)] += w['rejects'] / days
        mon_days[mkey(day)] += 1
months = [m for m in sorted(mon_done) if mon_days[m] >= 14]
cap_month = [{'m': m, 'done': round(mon_done[m]), 'good': round(mon_good[m]), 'rej': round(mon_rej[m]),
              'perweek': round(mon_done[m] / (mon_days[m] / 7)) if mon_days[m] else 0,
              'coverage': mon_days[m]} for m in months]

# ---------------- slippage / on-time (lines we saw open, now finished)
fin = [x for x in L if x['slip'] is not None and x['finish'] and x['finish'] <= TODAY]
open_late = [x for x in L if not x['finish'] and x['orig_pi'] and x['orig_pi'] < TODAY and x['online'] + 0 >= 0
             and x['observed_open'] and x['last_rd'] == last and x['order_qty'] - x['good'] - x['bgrade'] - x['defect'] > 0]


def wq(xs, key):
    tot = sum(x['order_qty'] for x in xs)
    return sum(key(x) * x['order_qty'] for x in xs) / tot if tot else 0


slips = sorted(x['slip'] for x in fin)
ontime = [x for x in fin if x['slip'] <= GRACE]
hist_bins = [(-10**6, -30, 'Early >30d'), (-30, 0, 'Early ≤30d'), (0, 8, 'On time (≤7d)'), (8, 31, '8–30d late'),
             (31, 61, '31–60d late'), (61, 91, '61–90d late'), (91, 181, '91–180d late'), (181, 10**6, '>180d late')]
hist = []
for lo, hi, lab in hist_bins:
    xs = [x for x in fin if lo <= x['slip'] < hi]
    hist.append({'b': lab, 'lines': len(xs), 'units': round(sum(x['order_qty'] for x in xs))})

# ---------------- revenue timing by ORIGINAL commitment month
rev = defaultdict(lambda: {'due_units': 0, 'due_rev': 0, 'late_units': 0, 'late_rev': 0, 'late30': 0, 'late90': 0,
                           'late_more': 0, 'still_open': 0, 'rev_days': 0})
for x in fin + open_late:
    if not x['orig_pi'] or x['orig_pi'] < first:
        continue
    m = mkey(x['orig_pi'])
    v = x['order_qty'] * x['price']
    r = rev[m]
    r['due_units'] += x['order_qty']; r['due_rev'] += v
    late = (x['slip'] if x['slip'] is not None else (TODAY - x['orig_pi']).days)
    if late > GRACE:
        r['late_units'] += x['order_qty']; r['late_rev'] += v; r['rev_days'] += v * late
        if not x['finish']:
            r['still_open'] += v
        elif late <= 30:
            r['late30'] += v
        elif late <= 90:
            r['late90'] += v
        else:
            r['late_more'] += v
rev_month = [{'m': m, **{k: round(v) for k, v in rev[m].items()},
              'pct_late': round(100 * rev[m]['late_rev'] / rev[m]['due_rev'], 1) if rev[m]['due_rev'] else 0,
              'avg_days_late': round(rev[m]['rev_days'] / rev[m]['late_rev']) if rev[m]['late_rev'] else 0}
             for m in sorted(rev) if m <= mkey(TODAY)]

# ---------------- overdue backlog over time ($ at each report)
od_val = defaultdict(float)
for r in csv.DictReader(open(os.path.join(D, 'overdue_sku.csv'), encoding='utf-8')):
    od_val[r['report_date']] += f(r['overdue_units']) * p_of(r['item'])[0]
backlog = [{'d': str(w['rd']), 'units': round(w['overdue_units']), 'rev': round(od_val[str(w['rd'])]),
            'lines': int(w['overdue_lines'])} for w in W]

# ---------------- rejects by month
rej_month = [{'m': c['m'], 'rej': c['rej'], 'done': c['done'],
              'rate': round(100 * c['rej'] / c['done'], 2) if c['done'] else 0} for c in cap_month]

# ---------------- SKU table
sku = defaultdict(list)
for x in L:
    sku[x['item']].append(x)
sku_rows = []
for it, xs in sku.items():
    f_ = [x for x in xs if x['slip'] is not None and x['finish'] and x['finish'] <= TODAY]
    ol = [x for x in open_late if x['item'] == it]
    rej_b = sum(x['bgrade'] for x in xs); rej_d = sum(x['defect'] for x in xs)
    made = sum(x['good'] for x in xs) + rej_b + rej_d
    od_units = sum(x['order_qty'] - x['good'] - x['bgrade'] - x['defect'] for x in ol)
    p, pb = p_of(it)
    late_rev = sum(x['order_qty'] * p for x in f_ if x['slip'] > GRACE) + od_units * p
    sku_rows.append({
        'sku': it, 'line': xs[0]['pline'], 'pos': len({x['po'] for x in xs}), 'lines': len(xs),
        'ordered': round(sum(x['order_qty'] for x in xs)), 'made': round(made),
        'avg_slip': round(statistics.mean([x['slip'] for x in f_]), 1) if f_ else None,
        'med_slip': statistics.median([x['slip'] for x in f_]) if f_ else None,
        'max_slip': max([x['slip'] for x in f_]) if f_ else None,
        'ontime': round(100 * sum(1 for x in f_ if x['slip'] <= GRACE) / len(f_)) if f_ else None,
        'resched': sum(x['resched'] for x in xs),
        'bgrade': round(rej_b), 'defect': round(rej_d), 'makeup': round(sum(x['makeup'] for x in xs)),
        'rej_rate': round(100 * (rej_b + rej_d) / made, 2) if made else None,
        'od_units': round(od_units), 'od_rev': round(od_units * p),
        'max_od': max([(TODAY - x['orig_pi']).days for x in ol]) if ol else None,
        'price': round(p, 2), 'pbasis': pb, 'late_rev': round(late_rev)})
sku_rows.sort(key=lambda r: -r['late_rev'])

# product-line rollup
pl = defaultdict(lambda: defaultdict(float))
for r in sku_rows:
    a = pl[r['line']]
    a['ordered'] += r['ordered']; a['made'] += r['made']; a['rej'] += r['bgrade'] + r['defect']
    a['od_units'] += r['od_units']; a['od_rev'] += r['od_rev']; a['late_rev'] += r['late_rev']; a['skus'] += 1
for it, xs in sku.items():
    for x in xs:
        if x['slip'] is not None and x['finish'] and x['finish'] <= TODAY:
            a = pl[x['pline']]; a['slip_w'] += x['slip'] * x['order_qty']; a['slip_q'] += x['order_qty']
            a['ontime_q'] += x['order_qty'] if x['slip'] <= GRACE else 0
pl_rows = sorted([{'line': k, 'skus': int(v['skus']), 'ordered': round(v['ordered']), 'made': round(v['made']),
                   'avg_slip': round(v['slip_w'] / v['slip_q'], 1) if v['slip_q'] else None,
                   'ontime': round(100 * v['ontime_q'] / v['slip_q']) if v['slip_q'] else None,
                   'rej_rate': round(100 * v['rej'] / v['made'], 2) if v['made'] else None,
                   'od_units': round(v['od_units']), 'od_rev': round(v['od_rev']), 'late_rev': round(v['late_rev'])}
                  for k, v in pl.items()], key=lambda r: -r['late_rev'])

# open overdue PO lines now
od_lines = sorted([{'po': x['po'], 'sku': x['item'], 'line': x['pline'], 'qty_open': round(x['order_qty'] - x['good'] - x['bgrade'] - x['defect']),
                    'orig': str(x['orig_pi']), 'current': str(x['pi']) if x['pi'] else '', 'sched': str(x['sched']) if x['sched'] else '', 'days': (TODAY - x['orig_pi']).days,
                    'resched': x['resched'], 'rev': round((x['order_qty'] - x['good'] - x['bgrade'] - x['defect']) * x['price'])}
                   for x in open_late], key=lambda r: -r['rev'])

# ---------------- KPIs
last13 = [c for c in cap_month if c['coverage'] >= 28][-3:]  # last 3 full months
tot_due = sum(r['due_rev'] for r in rev_month); tot_late = sum(r['late_rev'] for r in rev_month)
kpi = {
    'first': str(first), 'last': str(last), 'n_reports': len(reports), 'n_lines': len(L), 'n_skus': len(sku),
    'n_pos': len({x['po'] for x in L}), 'gaps': gaps,
    'units_done': round(sum(c['done'] for c in cap_month)),
    'avg_week_recent': round(statistics.mean([c['perweek'] for c in last13])) if last13 else 0,
    'avg_week_all': round(sum(c['done'] for c in cap_month) / (sum(c['coverage'] for c in cap_month) / 7)),
    'med_slip': statistics.median(slips) if slips else None,
    'wavg_slip': round(wq(fin, lambda x: x['slip']), 1),
    'p75_slip': slips[int(len(slips) * .75)] if slips else None,
    'ontime_pct': round(100 * len(ontime) / len(fin)) if fin else None,
    'ontime_units_pct': round(100 * sum(x['order_qty'] for x in ontime) / sum(x['order_qty'] for x in fin)) if fin else None,
    'n_fin': len(fin),
    'od_units_now': round(sum(r['qty_open'] for r in od_lines)), 'od_rev_now': round(sum(r['rev'] for r in od_lines)),
    'od_lines_now': len(od_lines),
    'rej_units': round(sum(r['bgrade'] + r['defect'] for r in sku_rows)),
    'rej_rate': round(100 * sum(r['bgrade'] + r['defect'] for r in sku_rows) / max(1, sum(r['made'] for r in sku_rows)), 2),
    'late_rev_total': round(tot_late), 'due_rev_total': round(tot_due),
    'late_rev_pct': round(100 * tot_late / tot_due, 1) if tot_due else None,
    'late_rev_month_avg': round(tot_late / max(1, len([r for r in rev_month if r['due_rev']]))),
    'priced_skus': sum(1 for r in sku_rows if r['pbasis'] == 'Open SO rate'),
    'cost_skus': sum(1 for r in sku_rows if r['pbasis'].startswith('Cost')),
    'unpriced_skus': sum(1 for r in sku_rows if r['pbasis'].startswith('Unmatched')),
    'family_skus': sum(1 for r in sku_rows if r['pbasis'].startswith('Family')),
    'fallback_price': round(FALLBACK, 2), 'grace': GRACE, 'today': str(TODAY),
}

out = {'kpi': kpi, 'cap_month': cap_month, 'rev_month': rev_month, 'backlog': backlog, 'hist': hist,
       'rej_month': rej_month, 'sku': sku_rows, 'pl': pl_rows, 'od_lines': od_lines}
json.dump(out, open(os.path.join(D, 'metrics.json'), 'w'), default=str)
print(json.dumps(kpi, indent=1, default=str))
