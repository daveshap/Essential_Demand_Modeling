import json
from pathlib import Path

root = Path(__file__).resolve().parent.parent
lines = (root / 'scores_services_production.txt').read_text(encoding='utf-8').splitlines()
records = []
for line in lines:
    if not line.strip():
        continue
    code, low, median, high, explanation = line.split('|', 4)
    records.append(dict(code=code, low=float(low), median=float(median), high=float(high), explanation=explanation))
expected = {r['code'] for r in json.loads((root / 'bls/occupations.json').read_text(encoding='utf-8')) if r['code'][:2] in ['29','31','33','35','37','39','51']}
actual = {r['code'] for r in records}
assert len(records) == len(actual), 'Duplicate score codes'
assert actual == expected, {'missing':sorted(expected-actual),'extra':sorted(actual-expected)}
assert all(0 <= r['low'] <= r['median'] <= r['high'] <= 100 and len(r['explanation']) > 30 for r in records)
(root/'scores_services_production.json').write_text(json.dumps(records, indent=2, ensure_ascii=False)+'\n', encoding='utf-8')
print(f'Validated {len(records)} explicit occupation scores; {sum(r["median"] == 0 for r in records)} zero medians.')
