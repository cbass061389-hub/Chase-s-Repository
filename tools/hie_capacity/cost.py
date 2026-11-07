"""HIE unit-price history by SKU and PO -> year-over-year cost change.

Inputs
  wh_raw/*.txt          weekly 'Predator warehouse inventory list' text (sheet 'regular inventory':
                        ITEM# | PO# | QTY | Unit Price | Amount)
  ../cap/ool_lines.csv  open order list history (PO first-seen dates, order qty per PO/item)
Outputs
  po_prices.csv   one row per (SKU, PO): unit price, PO date, year, basis
  cost.json       aggregates for the HTML section
Method
  PO date = first report the PO appeared on HIE's open order list (POs that first appear after the first report).
            POs already open in Jun 2023 and older: placed by PO number, interpolating/extrapolating the
            PO-number -> date line fitted on the dated POs. PO 7666 (HIE overstock bucket) is excluded.
  Year price per SKU = quantity-weighted unit price across that year's POs.
  YoY % per SKU = price(year) / price(prior year) - 1 (only SKUs bought in both years).
  Index = chain-linked, weighted by the later year's units ordered.
  $ impact of a year's increase = (price_y - price_{y-1}) x units ordered in year y.
"""
import csv, glob, json, os, re, statistics
from collections import defaultdict
from datetime import date, datetime, timedelta

D = os.path.dirname(os.path.abspath(__file__))
CAP = os.path.join(D, '..', 'cap')


def norm(s):
    return re.sub(r'\s+', ' ', (s or '').strip().upper())


def f(s):
    try:
        return float(str(s).replace(',', ''))
    except ValueError:
        return None


# ---------------- warehouse unit prices
obs = []   # (report_date, sku, po, qty, price)
for fp in sorted(glob.glob(os.path.join(D, 'wh_raw', '*.txt'))):
    m = re.search(r'(\d{4}-\d{2}-\d{2})', os.path.basename(fp))
    rd = m.group(1) if m else ''
    t = open(fp, encoding='utf-8', errors='replace').read()
    if t.lstrip().startswith('{') or t.lstrip().startswith('"'):
        try:
            j = json.loads(t); t = j if isinstance(j, str) else (j.get('text') or json.dumps(j))
        except Exception:
            pass
    blocks = re.split(r'^=== Sheet: (.*?) ===\s*$', t, flags=re.M)
    for i in range(1, len(blocks), 2):
        rows = blocks[i + 1].split('\n')
        hi = next((k for k, r in enumerate(rows[:6]) if 'ITEM' in r.upper() and 'PRICE' in r.upper()), None)
        if hi is None:
            continue
        hdr = [h.strip().upper() for h in rows[hi].split('\t')]
        ci = {k: next((n for n, h in enumerate(hdr) if k in h), None) for k in ('ITEM', 'PO', 'QTY', 'PRICE')}
        if None in ci.values():
            continue
        for r in rows[hi + 1:]:
            c = r.split('\t')
            if len(c) <= max(ci.values()):
                continue
            sku, po = norm(c[ci['ITEM']]), c[ci['PO']].strip()
            q, p = f(c[ci['QTY']]), f(c[ci['PRICE']])
            if not sku or sku.startswith('TOTAL') or not re.match(r'^\d{3,6}$', po) or not p or p <= 0:
                continue
            obs.append((rd, sku, po, q or 0, round(p, 2)))

# ---------------- PO dates
ool = list(csv.DictReader(open(os.path.join(CAP, 'ool_lines.csv'), encoding='utf-8')))
first_rep = min(r['report_date'] for r in ool)
first_seen, order_qty = {}, defaultdict(float)
seen_line = set()
for r in ool:
    po = r['po']
    first_seen.setdefault(po, r['report_date'])
    key = (po, norm(r['item']), r['pi_date'], r['order_qty'])
    if key not in seen_line:          # count each delivery line once
        seen_line.add(key)
        order_qty[(po, norm(r['item']))] += f(r['order_qty']) or 0
# Anchors: HIE quote date of each PI-quoted PO (earliest quote) - real dates, unlike open-order-list sightings
anchors = {}
_pif = os.path.join(D, '..', 'pi', 'pi_prices.csv')
if os.path.exists(_pif):
    for q in csv.DictReader(open(_pif, encoding='utf-8')):
        if q['po'].isdigit() and 9000 < int(q['po']) < 20000 and q['quote_date']:
            d0 = datetime.strptime(q['quote_date'], '%Y-%m-%d').date()
            anchors[int(q['po'])] = min(anchors.get(int(q['po']), d0), d0)
dated = sorted(anchors.items()) if len(anchors) > 20 else sorted(
    (int(p), datetime.strptime(d, '%Y-%m-%d').date()) for p, d in first_seen.items()
    if p.isdigit() and d > first_rep and 9000 < int(p) < 20000)
# monotone smoothing: running max of dates along PO number
mono, last = [], date(2000, 1, 1)
for p, d in dated:
    last = max(last, d); mono.append((p, last))
n = len(mono)
xs = [p for p, _ in mono]; ys = [(d - date(2020, 1, 1)).days for _, d in mono]
mx, my = statistics.mean(xs), statistics.mean(ys)
slope = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / sum((x - mx) ** 2 for x in xs)
intercept = my - slope * mx


def po_date(po):
    if not po.isdigit() or not (9000 < int(po) < 20000):
        return None, 'n/a'
    p = int(po)
    if p in anchors:
        return anchors[p], 'HIE quote date'
    if xs[0] <= p <= xs[-1]:
        for (p0, d0), (p1, d1) in zip(mono, mono[1:]):
            if p0 <= p <= p1:
                fr = (p - p0) / (p1 - p0) if p1 > p0 else 0
                return d0 + timedelta(days=round(fr * (d1 - d0).days)), 'Interpolated from PO number'
    return date(2020, 1, 1) + timedelta(days=round(intercept + slope * p)), 'Extrapolated from PO number'


# ---------------- (sku, po) price table
sp = defaultdict(lambda: {'prices': defaultdict(float), 'first': None, 'last': None, 'wh_qty_max': 0})
for rd, sku, po, q, p in obs:
    if po == '7666':
        continue
    e = sp[(sku, po)]
    e['prices'][p] += max(q, 1)
    e['first'] = min(e['first'] or rd, rd); e['last'] = max(e['last'] or rd, rd)
    e['wh_qty_max'] = max(e['wh_qty_max'], q)
rows = []
for (sku, po), e in sp.items():
    p = max(e['prices'], key=e['prices'].get)            # the price HIE carries for that PO (mode)
    pd_, basis = po_date(po)
    oq = order_qty.get((po, sku)) or e['wh_qty_max']
    rows.append({'sku': sku, 'po': po, 'price': p, 'n_prices': len(e['prices']), 'po_date': str(pd_) if pd_ else '',
                 'year': pd_.year if pd_ else None, 'date_basis': basis, 'units': round(oq),
                 'units_basis': 'PO order qty (open order list)' if order_qty.get((po, sku)) else 'Max qty seen in warehouse',
                 'wh_first': e['first'], 'wh_last': e['last']})
# ---------------- HIE PI quotes override warehouse prices (quote date = PO placement)
pif = os.path.join(D, '..', 'pi', 'pi_prices.csv')
n_pi = 0
if os.path.exists(pif):
    have = {(r['sku'], r['po']): r for r in rows}
    for q in csv.DictReader(open(pif, encoding='utf-8')):
        p = f(q['price'])
        if not p or p <= 0:
            continue
        sku, po = norm(q['sku']), q['po']
        try:
            qd = datetime.strptime(q['quote_date'], '%Y-%m-%d').date()
        except ValueError:
            continue
        units = f(q['units']) or 0
        py = int(q['price_year']) if (q.get('price_year') or '').isdigit() else qd.year
        rec = {'sku': sku, 'po': po, 'price': round(p, 2), 'n_prices': 1, 'po_date': str(qd), 'year': py,
               'year_basis': 'Production year (' + (q.get('price_year_basis') or 'quote date') + ')',
               'date_basis': 'HIE quote date (' + q['source'] + ')', 'units': round(order_qty.get((po, sku)) or units),
               'units_basis': 'PO order qty (open order list)' if order_qty.get((po, sku)) else 'PI qty',
               'wh_first': '', 'wh_last': '', 'price_source': 'HIE PI' if q['source'] == 'PI' else 'HIE pending list'}
        old = have.get((sku, po))
        if old:
            rec['wh_price'] = old['price']; rec['wh_first'] = old['wh_first']; rec['wh_last'] = old['wh_last']
            rows.remove(old)
        rows.append(rec); n_pi += 1
for r in rows:
    r.setdefault('price_source', 'Warehouse list'); r.setdefault('wh_price', r['price'] if r['price_source'] == 'Warehouse list' else '')
rows.sort(key=lambda r: (r['sku'], r['po_date']))
with open(os.path.join(D, 'po_prices.csv'), 'w', newline='', encoding='utf-8') as fo:
    w = csv.DictWriter(fo, sorted({k for r in rows for k in r})); w.writeheader(); w.writerows(rows)

# ---------------- product line lookup (reuse capacity pricing categories)
cat = {}
pf = os.path.join(CAP, 'prices.csv')
if os.path.exists(pf):
    for r in csv.DictReader(open(pf)):
        cat[norm(r['sku'])] = r['category'].split(' : ')[-1] if r['category'] else ''


def pline(sku):
    if sku in cat and cat[sku]:
        return cat[sku]
    toks = sku.split()
    for k in range(len(toks) - 1, 1, -1):
        pre = ' '.join(toks[:k])
        hits = [v for s, v in cat.items() if s.startswith(pre + ' ') and v]
        if hits:
            return max(set(hits), key=hits.count)
    return 'Other'



# ---------------- price year = year of the PO's earliest PI date (HIE prices by PI date), any source
po_pi = {}
_lf = os.path.join(D, '..', 'pi', 'pi_ledger.csv')
if os.path.exists(_lf):
    for r in csv.DictReader(open(_lf, encoding='utf-8')):
        if r['pi_first_min']:
            po_pi[r['po']] = min(po_pi.get(r['po'], r['pi_first_min']), r['pi_first_min'])
_tf = os.path.join(CAP, 'lines_tracked.csv')
if os.path.exists(_tf):
    for r in csv.DictReader(open(_tf, encoding='utf-8')):
        v = r.get('orig_pi') or ''
        if re.match(r'\d{4}-\d{2}-\d{2}$', v) and r['po'] not in po_pi:
            po_pi.setdefault('_ool_' + r['po'], v)
            po_pi['_ool_' + r['po']] = min(po_pi['_ool_' + r['po']], v)
n_pi_year = 0
for r in rows:
    v = po_pi.get(r['po']) or po_pi.get('_ool_' + r['po'])
    if v:
        r['year'] = int(v[:4]); r['year_basis'] = 'PO earliest PI date (' + ('PI' if r['po'] in po_pi else 'open order list') + ')'
        n_pi_year += 1
    else:
        r.setdefault('year_basis', r.get('date_basis', ''))

# ---------------- HIE-announced price increases (price tables with 'previous year final price' in PI threads)
announced = {}
for fp in sorted(glob.glob(os.path.join(D, '..', 'pi', 'pi_lines_*.tsv'))):
    for r in csv.DictReader(open(fp, encoding='utf-8', errors='replace'), delimiter='\t'):
        rem = r.get('remark') or ''
        m0 = re.search(r'(20\d\d) final price \$?([\d,]+\.?\d*)', rem, re.I)
        p1 = f(r.get('unit_price'))
        if not m0 or not p1:
            continue
        p0 = f(m0.group(2))
        if not p0:
            continue
        sku = norm(r['item'])
        announced[sku] = {'sku': sku, 'line': '', 'from_year': int(m0.group(1)), 'to_year': int(m0.group(1)) + 1,
                          'p0': p0, 'p1': round(p1, 2), 'pct': round(100 * (p1 / p0 - 1), 1),
                          'date': (r.get('received_utc') or '')[:10], 'subject': r.get('subject', '')}

# ---------------- year prices per SKU
yr = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0]))   # sku -> year -> [sum p*q, sum q]
for r in rows:
    if r['year'] is None:
        continue
    a = yr[r['sku']][r['year']]
    a[0] += r['price'] * max(r['units'], 1); a[1] += max(r['units'], 1)
years = sorted({y for s in yr.values() for y in s})
sku_year = {s: {y: (v[0] / v[1], v[1]) for y, v in ys_.items()} for s, ys_ in yr.items()}

pairs = []   # yoy per sku
for s, ys_ in sku_year.items():
    yl = sorted(ys_)
    for a, b in zip(yl, yl[1:]):
        if b - a != 1:
            continue
        pa, pb = ys_[a][0], ys_[b][0]
        pairs.append({'sku': s, 'line': pline(s), 'from': a, 'to': b, 'p0': round(pa, 2), 'p1': round(pb, 2),
                      'pct': round(100 * (pb / pa - 1), 2), 'units': round(ys_[b][1]),
                      'impact': round((pb - pa) * ys_[b][1])})

yoy = []
idx = 100.0
pairs = [p for p in pairs if p['to'] <= date.today().year]   # future price years = orders not yet built
pair_years = sorted({p['from'] for p in pairs})
years = [y for y in years if pair_years and y >= pair_years[0]]
index = [{'year': years[0], 'index': 100.0}] if years else []
for a, b in zip(years, years[1:]):
    P = [p for p in pairs if p['from'] == a and p['to'] == b]
    if len(P) < 5:            # too few SKUs bought in both years to call it a trend
        if index and index[-1]['year'] == a and len(index) == 1:
            index = [{'year': b, 'index': 100.0}]
        continue
    wsum = sum(p['units'] for p in P)
    wpct = sum(p['pct'] * p['units'] for p in P) / wsum if wsum else 0
    idx *= 1 + wpct / 100
    index.append({'year': b, 'index': round(idx, 1)})
    yoy.append({'from': a, 'to': b, 'skus': len(P), 'units': wsum, 'wpct': round(wpct, 2),
                'median_pct': round(statistics.median([p['pct'] for p in P]), 2),
                'up': sum(1 for p in P if p['pct'] > 0.5), 'flat': sum(1 for p in P if abs(p['pct']) <= 0.5),
                'down': sum(1 for p in P if p['pct'] < -0.5), 'impact': sum(p['impact'] for p in P)})

by_line = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0, 0.0, 0]))
for p in pairs:
    a = by_line[p['line']][p['to']]
    a[0] += p['pct'] * p['units']; a[1] += p['units']; a[2] += p['impact']; a[3] += 1
line_rows = []
for l, d_ in by_line.items():
    row = {'line': l, 'impact': round(sum(v[2] for v in d_.values()))}
    for y, v in d_.items():
        row[str(y)] = {'pct': round(v[0] / v[1], 2) if v[1] else None, 'skus': v[3], 'impact': round(v[2])}
    line_rows.append(row)
line_rows.sort(key=lambda r: -r['impact'])

# SKU price ladder
sku_rows = []
for s, ys_ in sku_year.items():
    yl = sorted(ys_)
    r = {'sku': s, 'line': pline(s), 'years': {str(y): round(ys_[y][0], 2) for y in yl},
         'units': {str(y): round(ys_[y][1]) for y in yl}}
    if len(yl) >= 2:
        r['first_y'], r['last_y'] = yl[0], yl[-1]
        r['total_pct'] = round(100 * (ys_[yl[-1]][0] / ys_[yl[0]][0] - 1), 2)
        r['cagr'] = round(100 * ((ys_[yl[-1]][0] / ys_[yl[0]][0]) ** (1 / (yl[-1] - yl[0])) - 1), 2)
    r['impact'] = sum(p['impact'] for p in pairs if p['sku'] == s)
    r['n_po'] = sum(1 for x in rows if x['sku'] == s)
    sku_rows.append(r)
sku_rows.sort(key=lambda r: -(r['impact'] or 0))

# same-year price spread (different prices on different POs in one year)
spread = []
for s in yr:
    for y in years:
        ps = sorted({x['price'] for x in rows if x['sku'] == s and x['year'] == y})
        if len(ps) > 1 and ps[0] > 0 and ps[-1] / ps[0] > 1.03:
            spread.append({'sku': s, 'year': y, 'low': ps[0], 'high': ps[-1], 'pct': round(100 * (ps[-1] / ps[0] - 1), 1)})
spread.sort(key=lambda r: -r['pct'])

out = {'meta': {'n_obs': len(obs), 'n_files': len(glob.glob(os.path.join(D, 'wh_raw', '*.txt'))),
                'n_sku_po': len(rows), 'n_skus': len(yr), 'years': years, 'n_pi_prices': n_pi, 'n_pi_year': n_pi_year,
                'price_source_counts': {k: sum(1 for r in rows if r['price_source'] == k) for k in ('HIE PI', 'HIE pending list', 'Warehouse list')},
                'po_slope_days_per_po': round(slope, 3),
                'basis_counts': {b: sum(1 for r in rows if r['date_basis'] == b) for b in {r['date_basis'] for r in rows}}},
       'yoy': yoy, 'index': index, 'announced': sorted(announced.values(), key=lambda a: (-a['pct'], a['sku'])), 'lines': line_rows, 'skus': sku_rows, 'pairs': pairs, 'spread': spread[:200]}
json.dump(out, open(os.path.join(D, 'cost.json'), 'w'), default=str)
print(json.dumps(out['meta'], indent=1)); print(json.dumps(yoy, indent=1)); print(index)
