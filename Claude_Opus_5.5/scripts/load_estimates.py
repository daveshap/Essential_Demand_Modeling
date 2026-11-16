"""Load the BLS baseline and the per-occupation estimates, and validate them.

Each estimates/*.txt line is: code|low|median|high|tags|rationale (percent).
Validation: every BLS line item scored exactly once, no unknown codes,
0 <= low <= median <= high <= 100, known tags, non-empty rationale.
"""
import csv
import glob
import os

ALLOWED_TAGS = {'Presence', 'Provenance', 'Accountability', 'Affinity'}


def load_bls(path):
    rows = list(csv.DictReader(open(path, encoding='utf-8')))
    for r in rows:
        r['level'] = int(r['level'])
        r['emp_2025_thousands'] = float(r['emp_2025_thousands'])
        r['emp_2035_thousands'] = float(r['emp_2035_thousands'])
        r['pct_self_employed_2025'] = float(r['pct_self_employed_2025']) if r['pct_self_employed_2025'] else None
        r['median_wage_2025'] = float(r['median_wage_2025']) if r['median_wage_2025'] else None
    return rows


def load_estimates(est_dir, bls_rows):
    line_items = {r['soc_code']: r for r in bls_rows if r['occupation_type'] == 'Line item'}
    est, errors = {}, []
    for path in sorted(glob.glob(os.path.join(est_dir, '*.txt'))):
        for n, line in enumerate(open(path, encoding='utf-8'), 1):
            line = line.rstrip('\n')
            if not line.strip() or line.startswith('#'):
                continue
            parts = line.split('|')
            where = f'{os.path.basename(path)}:{n}'
            if len(parts) != 6:
                errors.append(f'{where}: expected 6 fields, got {len(parts)}')
                continue
            code, lo, md, hi, tags, why = [p.strip() for p in parts]
            if code not in line_items:
                errors.append(f'{where}: {code} is not a BLS line item')
                continue
            if code in est:
                errors.append(f'{where}: {code} scored twice')
                continue
            try:
                lo, md, hi = float(lo), float(md), float(hi)
            except ValueError:
                errors.append(f'{where}: non-numeric share')
                continue
            if not (0 <= lo <= md <= hi <= 100):
                errors.append(f'{where}: {code} shares out of order or range ({lo}, {md}, {hi})')
            tag_list = [] if tags == 'None' else [t.strip() for t in tags.split(';')]
            bad = [t for t in tag_list if t not in ALLOWED_TAGS]
            if bad:
                errors.append(f'{where}: {code} unknown tags {bad}')
            if md > 0 and not tag_list:
                errors.append(f'{where}: {code} has a nonzero median but no tag')
            if md == 0 and hi == 0 and tag_list:
                errors.append(f'{where}: {code} is all zero but tagged')
            if len(why) < 20:
                errors.append(f'{where}: {code} rationale too short')
            est[code] = {'low': lo, 'median': md, 'high': hi, 'tags': tag_list, 'rationale': why,
                         'source_file': os.path.basename(path)}
    missing = [c for c in line_items if c not in est]
    for c in missing:
        errors.append(f'missing estimate for {c} {line_items[c]["title"]}')
    return est, errors
