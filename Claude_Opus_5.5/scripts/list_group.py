"""Print the detailed occupations of one or more SOC major groups: code | title | 2025 emp (K) | % self-employed | median wage."""
import csv, sys
rows = list(csv.DictReader(open(sys.argv[1], encoding='utf-8')))
groups = set(sys.argv[2:])
for r in rows:
    if r['occupation_type'] == 'Line item' and r['major_group_code'][:2] in groups:
        print(f"{r['soc_code']} | {r['title']} | {r['emp_2025_thousands']} | se={r['pct_self_employed_2025']} | ${r['median_wage_2025']}")
