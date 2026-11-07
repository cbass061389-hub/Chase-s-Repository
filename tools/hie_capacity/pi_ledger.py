"""Merge every parsed HIE PI line into one ledger and derive price-by-year and commitment metrics.

Inputs : pi_lines_*.tsv (one row per PI line per email), cancel_lines.tsv
Outputs: pi_ledger.csv      one row per (PO, item): first quote, latest quote, first PI dates, request dates
         pi_prices.csv      one row per (SKU, PO) with the PO's quoted unit price and quote date (feeds cost.py)
         pi_summary.json    counts, request-vs-PI gap, re-quotes, cancellations
Rules  : 'Pending order list' rows are kept only as a fallback source of price/date for POs without a PI email.
         First quote = earliest received email that priced the line; original PI date = earliest email's PI dates.
"""
import csv, glob, json, os, re, statistics
from collections import defaultdict
from datetime import datetime, date

D = os.path.dirname(os.path.abspath(__file__))


def norm(s):
    return re.sub(r'\s+', ' ', (s or '').strip().upper())


def num(s):
    try:
        return float(str(s).replace(',', '').replace('$', '').strip())
    except ValueError:
        return None


def dt(s):
    s = (s or '').strip()[:10]
    for f in ('%Y-%m-%d', '%m/%d/%Y', '%Y/%m/%d'):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


rows = []
for fp in sorted(glob.glob(os.path.join(D, 'pi_lines_*.tsv'))):
    for r in csv.DictReader(open(fp, encoding='utf-8', errors='replace'), delimiter='\t'):
        po = re.sub(r'\D', '', r.get('po') or '')
        item = norm(r.get('item'))
        if not po or not item or item.startswith('TOTAL'):
            continue
        src = 'pending list' if 'pending order list' in (r.get('subject', '') + r.get('attachment_name', '')).lower() else 'PI'
        rows.append({'recv': dt(r.get('received_utc')), 'po': po, 'item': item, 'qty': num(r.get('qty')),
                     'price': num(r.get('unit_price')), 'pi': dt(r.get('pi_date')), 'req': dt(r.get('request_date')),
                     'start': dt(r.get('start_date')), 'rev': (r.get('is_revision') or '').upper().startswith('Y'),
                     'src': src, 'file': os.path.basename(fp), 'subject': r.get('subject', ''), 'remark': r.get('remark', '')})

# dedupe identical lines repeated across reply emails
seen, uniq = set(), []
for r in sorted(rows, key=lambda x: (x['recv'] or date(2100, 1, 1))):
    k = (r['po'], r['item'], r['qty'], r['price'], r['pi'], r['req'], r['recv'])
    if k not in seen:
        seen.add(k); uniq.append(r)
rows = uniq

by = defaultdict(list)
for r in rows:
    by[(r['po'], r['item'])].append(r)

ledger = []
for (po, item), rs in by.items():
    pis = [r for r in rs if r['src'] == 'PI'] or rs
    pis.sort(key=lambda x: x['recv'] or date(2100, 1, 1))
    first_recv = pis[0]['recv']
    first = [r for r in pis if r['recv'] == first_recv]
    last_recv = pis[-1]['recv']
    last = [r for r in pis if r['recv'] == last_recv]
    p_first = [r['price'] for r in first if r['price']]
    p_last = [r['price'] for r in last if r['price']]
    q_first = sum(r['qty'] or 0 for r in first)
    pi_first = [r['pi'] for r in first if r['pi']]
    req_first = [r['req'] for r in first if r['req']]
    gaps = [(r['pi'] - r['req']).days for r in first if r['pi'] and r['req']]
    ledger.append({
        'po': po, 'item': item, 'source': pis[0]['src'], 'first_quote': str(first_recv) if first_recv else '',
        'last_quote': str(last_recv) if last_recv else '', 'n_quotes': len({r['recv'] for r in pis}),
        'qty_first': q_first, 'price_first': statistics.median(p_first) if p_first else None,
        'price_last': statistics.median(p_last) if p_last else None,
        'pi_first_min': str(min(pi_first)) if pi_first else '', 'pi_first_max': str(max(pi_first)) if pi_first else '',
        'req_first_min': str(min(req_first)) if req_first else '',
        'pi_minus_req_days': round(statistics.mean(gaps)) if gaps else None,
        'pi_last_max': str(max([r['pi'] for r in last if r['pi']], default='')) if last else '',
    })

first_dates = {}
for (po, item), rs in by.items():
    pis = [r for r in rs if r['src'] == 'PI'] or rs
    pis.sort(key=lambda x: x['recv'] or date(2100, 1, 1))
    fr = pis[0]['recv']
    first_dates[po + '|' + item] = sorted(str(r['pi']) for r in pis if r['recv'] == fr and r['pi'])
json.dump(first_dates, open(os.path.join(D, 'pi_first_dates.json'), 'w'))
with open(os.path.join(D, 'pi_ledger.csv'), 'w', newline='', encoding='utf-8') as fo:
    w = csv.DictWriter(fo, list(ledger[0].keys())); w.writeheader(); w.writerows(ledger)
with open(os.path.join(D, 'pi_prices.csv'), 'w', newline='', encoding='utf-8') as fo:
    w = csv.writer(fo); w.writerow(['sku', 'po', 'price', 'quote_date', 'units', 'source'])
    for L in ledger:
        if L['price_first']:
            w.writerow([L['item'], L['po'], L['price_first'], L['first_quote'], L['qty_first'], L['source']])

# re-quotes: price changed between first and latest quote on the same PO line
requotes = [L for L in ledger if L['price_first'] and L['price_last'] and abs(L['price_last'] - L['price_first']) > 0.01]
gaps = [L['pi_minus_req_days'] for L in ledger if L['pi_minus_req_days'] is not None]
pi_moves = [L for L in ledger if L['pi_first_max'] and L['pi_last_max'] and L['pi_last_max'] > L['pi_first_max']]

cancels = list(csv.DictReader(open(os.path.join(D, 'cancel_lines.tsv'), encoding='utf-8', errors='replace'), delimiter='\t')) \
    if os.path.exists(os.path.join(D, 'cancel_lines.tsv')) else []

summary = {
    'pi_rows': len(rows), 'po_lines': len(ledger), 'pos': len({L['po'] for L in ledger}),
    'skus': len({L['item'] for L in ledger}), 'files': sorted({r['file'] for r in rows}),
    'by_year': {y: len({L['po'] for L in ledger if L['first_quote'][:4] == y}) for y in ('2023', '2024', '2025', '2026')},
    'requotes': len(requotes),
    'requote_examples': sorted([{'po': L['po'], 'item': L['item'], 'from': L['price_first'], 'to': L['price_last'],
                                 'pct': round(100 * (L['price_last'] / L['price_first'] - 1), 1)} for L in requotes],
                               key=lambda x: -abs(x['pct']))[:40],
    'pi_vs_request_median_days': statistics.median(gaps) if gaps else None,
    'pi_vs_request_mean_days': round(statistics.mean(gaps), 1) if gaps else None,
    'pi_later_than_request_pct': round(100 * sum(1 for g in gaps if g > 7) / len(gaps)) if gaps else None,
    'pi_redated_later_lines': len(pi_moves),
    'cancel_rows': len(cancels),
}
json.dump(summary, open(os.path.join(D, 'pi_summary.json'), 'w'), indent=1, default=str)
print(json.dumps({k: v for k, v in summary.items() if k not in ('requote_examples', 'files')}, indent=1))
