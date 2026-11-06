"""Parse every HIE weekly open order list text dump into one long line-level table.

Input : cap/ool_raw/*.txt  (raw read_resource text; '=== Sheet: X ===' blocks, tab-separated rows)
Output: cap/ool_lines.csv  one row per report x PO line
        cap/ool_plan.csv   one row per report x PO line x planned week (weekly schedule grid)
        cap/ool_reports.csv one row per report file (date, sheets, rows)
"""
import csv, glob, os, re, json
from datetime import datetime, timedelta

D = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(D, 'ool_raw')
EPOCH = datetime(1899, 12, 30)


def xl(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if 30000 < f < 60000:
        return (EPOCH + timedelta(days=f)).date()
    return None


def num(v):
    v = (v or '').strip().replace(',', '')
    try:
        return float(v)
    except ValueError:
        return None


def load_text(path):
    t = open(path, encoding='utf-8', errors='replace').read()
    s = t.lstrip()
    if s.startswith('{') or s.startswith('"'):
        try:
            j = json.loads(t)
            if isinstance(j, str):
                t = j
            elif isinstance(j, dict):
                t = j.get('text') or j.get('content') or json.dumps(j)
        except Exception:
            pass
    return t


def report_date(path, text):
    m = re.search(r'UPDATED:\s*\t?\s*([\d.]+)', text)
    if m and xl(m.group(1)):
        return xl(m.group(1))
    m = re.search(r'(\d{4})-(\d{2})-(\d{2})', os.path.basename(path))
    return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).date() if m else None


FIELDS = ['po', 'item', 'order_date', 'order_qty', 'prod_qty', 'pi_date', 'sched_finish', 'actual_finish',
          'pg_month', 'pg_qty', 'difference', 'shipped', 'packing', 'warehouse', 'accum', 'bgrade', 'defect',
          'makeup', 'online', 'tba']
DATEF = {'order_date', 'pi_date', 'sched_finish', 'actual_finish'}
NUMF = {'order_qty', 'prod_qty', 'difference', 'shipped', 'packing', 'warehouse', 'accum', 'bgrade', 'defect',
        'makeup', 'online', 'pg_qty'}


def parse(path):
    text = load_text(path)
    rdate = report_date(path, text)
    lines, plans, sheets = [], [], []
    blocks = re.split(r'^=== Sheet: (.*?) ===\s*$', text, flags=re.M)
    for i in range(1, len(blocks), 2):
        sheet, body = blocks[i].strip(), blocks[i + 1]
        rows = body.split('\n')
        hi = next((k for k, r in enumerate(rows[:12]) if r.split('\t')[:2] == ['PO#', 'Item#']), None)
        if hi is None:
            sheets.append((sheet, 0, 'no-header'))
            continue
        hdr = rows[hi].split('\t')
        sub = rows[hi + 1].split('\t') if hi + 1 < len(rows) else []
        # weekly plan columns: serial dates in the sub-header from col 20 on
        wk = {c: xl(v) for c, v in enumerate(sub) if c >= 20 and xl(v)}
        rem_c = next((c for c, h in enumerate(hdr) if '異常' in h or '備注' in h), None)
        log_c = next((c for c, h in enumerate(hdr) if '異動' in h), None)
        n = 0
        for r in rows[hi + 2:]:
            c = r.split('\t')
            if len(c) < 12:
                continue
            po, item = c[0].strip(), c[1].strip()
            if not re.match(r'^\d{3,6}', po) or not item or re.search(r'total|合計|合计', item, re.I):
                continue
            rec = {'report_date': rdate, 'sheet': sheet, 'po': po, 'item': re.sub(r'\s+', ' ', item)}
            for k, f in enumerate(FIELDS[2:], start=2):
                v = c[k] if k < len(c) else ''
                rec[f] = xl(v) if f in DATEF else (num(v) if f in NUMF else v.strip())
            rec['remark'] = c[rem_c].strip() if rem_c is not None and rem_c < len(c) else ''
            rec['changelog'] = c[log_c].strip() if log_c is not None and log_c < len(c) else ''
            planned = {wk[k]: num(c[k]) for k in wk if k < len(c) and num(c[k])}
            rec['plan_total'] = sum(planned.values()) if planned else 0
            lines.append(rec)
            for wd, q in planned.items():
                plans.append({'report_date': rdate, 'po': po, 'item': rec['item'], 'pi_date': rec['pi_date'],
                              'week': wd, 'qty': q})
            n += 1
        sheets.append((sheet, n, 'ok'))
    return rdate, lines, plans, sheets


def main():
    files = sorted(glob.glob(os.path.join(RAW, '*.txt')))
    by_date = {}
    for f in files:
        rd, L, P, S = parse(f)
        if rd is None or not L:
            print('SKIP', os.path.basename(f), rd, len(L))
            continue
        if rd in by_date and len(by_date[rd][1]) >= len(L):
            continue  # keep the fuller copy of a duplicate week
        by_date[rd] = (f, L, P, S)
    cols = ['report_date', 'sheet'] + FIELDS + ['remark', 'changelog', 'plan_total']
    with open(os.path.join(D, 'ool_lines.csv'), 'w', newline='', encoding='utf-8') as fo, \
         open(os.path.join(D, 'ool_plan.csv'), 'w', newline='', encoding='utf-8') as fp, \
         open(os.path.join(D, 'ool_reports.csv'), 'w', newline='', encoding='utf-8') as fr:
        w = csv.DictWriter(fo, cols); w.writeheader()
        wp = csv.DictWriter(fp, ['report_date', 'po', 'item', 'pi_date', 'week', 'qty']); wp.writeheader()
        wr = csv.writer(fr); wr.writerow(['report_date', 'file', 'lines', 'sheets'])
        for rd in sorted(by_date):
            f, L, P, S = by_date[rd]
            for r in L: w.writerow(r)
            for p in P: wp.writerow(p)
            wr.writerow([rd, os.path.basename(f), len(L), ';'.join(f'{s}:{n}' for s, n, _ in S)])
    print(len(by_date), 'reports;', sum(len(v[1]) for v in by_date.values()), 'line rows')


if __name__ == '__main__':
    main()
