"""Validate estimates and print totals: py -3.13 -I check_estimates.py <repo_subdir>"""
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from load_estimates import load_bls, load_estimates  # noqa: E402

base = sys.argv[1]
bls = load_bls(os.path.join(base, 'source_data', 'BLS_EP_Table_1.2_2025-35.csv'))
est, errors = load_estimates(os.path.join(base, 'estimates'), bls)
print('errors:', len(errors))
for e in errors[:50]:
    print('  ', e)
items = [r for r in bls if r['occupation_type'] == 'Line item']
majors = {r['soc_code']: r['title'] for r in bls if r['occupation_type'] == 'Summary' and r['level'] == 1}
tot = sum(r['emp_2025_thousands'] for r in items)
agg = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])
for r in items:
    e = est.get(r['soc_code'])
    if not e:
        continue
    g = agg[r['major_group_code']]
    g[0] += r['emp_2025_thousands']
    for i, k in enumerate(('low', 'median', 'high'), 1):
        g[i] += r['emp_2025_thousands'] * e[k] / 100
print(f"\n{'group':<8}{'title':<58}{'emp K':>10}{'low%':>7}{'med%':>7}{'high%':>7}{'med K':>9}")
for code in sorted(agg):
    emp, lo, md, hi = agg[code]
    print(f"{code:<8}{majors[code][:56]:<58}{emp:>10.1f}{100*lo/emp:>7.2f}{100*md/emp:>7.2f}{100*hi/emp:>7.2f}{md:>9.1f}")
L = sum(v[1] for v in agg.values()); M = sum(v[2] for v in agg.values()); H = sum(v[3] for v in agg.values())
print(f"\nTOTAL employment {tot:.1f}K; essential low {L:.1f}K ({100*L/tot:.2f}%), median {M:.1f}K ({100*M/tot:.2f}%), high {H:.1f}K ({100*H/tot:.2f}%)")
zero = sum(1 for r in items if est.get(r['soc_code'], {}).get('median', 1) == 0)
zero_emp = sum(r['emp_2025_thousands'] for r in items if est.get(r['soc_code'], {}).get('median', 1) == 0)
print(f"occupations with median 0%: {zero} ({zero_emp:.1f}K jobs, {100*zero_emp/tot:.1f}% of employment)")
top = sorted(items, key=lambda r: -r['emp_2025_thousands'] * est.get(r['soc_code'], {'median': 0})['median'])[:25]
print('\nTop 25 contributors to the median total:')
for r in top:
    e = est[r['soc_code']]
    print(f"  {r['soc_code']} {r['title'][:55]:<56}{r['emp_2025_thousands']:>8.1f}K x {e['median']:>5}% = {r['emp_2025_thousands']*e['median']/100:>7.1f}K")
