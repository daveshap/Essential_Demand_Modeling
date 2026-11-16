"""Extract BLS Employment Projections Table 1.2 (2025-35) to CSV.

Usage: py -3.13 -I extract_table_1_2.py <occupation.xlsx> <out.csv>

Keeps every row (Summary and Line item) with its indent level, so the
hierarchy is preserved. Only Line item rows are ever summed downstream.
"""
import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_read import load, rows  # noqa: E402


def num(v):
    if v is None:
        return ''
    v = str(v).strip()
    try:
        return round(float(v), 4)
    except ValueError:
        return ''  # BLS uses a dash for unavailable


def main(src, out):
    z, sheets, ss = load(src)
    target = dict(sheets)['Table 1.2']
    header_seen = False
    recs = []
    for r, vals in rows(z, target, ss):
        if not header_seen:
            if vals and vals[0] == '2025 National Employment Matrix title':
                header_seen = True
            continue
        if len(vals) < 13 or vals[2] not in ('Summary', 'Line item'):
            continue
        raw = vals[0] or ''
        level = (len(raw) - len(raw.lstrip(' '))) // 2
        recs.append({
            'soc_code': vals[1].strip(),
            'title': raw.strip(),
            'occupation_type': vals[2],
            'level': level,
            'emp_2025_thousands': num(vals[3]),
            'emp_2035_thousands': num(vals[4]),
            'pct_self_employed_2025': num(vals[9]),
            'median_wage_2025': num(vals[11]),
            'typical_entry_education': (vals[12] or '').strip() if vals[12] and len(vals[12].strip()) > 1 else '',
        })
    # attach major group to each row
    major = None
    for rec in recs:
        if rec['soc_code'].endswith('-0000') and rec['soc_code'] != '00-0000':
            major = rec['soc_code']
        rec['major_group_code'] = major or ''
    fields = ['soc_code', 'title', 'occupation_type', 'level', 'major_group_code',
              'emp_2025_thousands', 'emp_2035_thousands', 'pct_self_employed_2025',
              'median_wage_2025', 'typical_entry_education']
    with open(out, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for rec in recs:
            w.writerow({k: rec[k] for k in fields})
    li = [x for x in recs if x['occupation_type'] == 'Line item']
    print('rows', len(recs), 'line items', len(li),
          'line-item employment (thousands)', round(sum(x['emp_2025_thousands'] for x in li), 1))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
