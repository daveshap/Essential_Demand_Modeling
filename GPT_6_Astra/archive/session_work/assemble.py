import json,csv
from pathlib import Path
from collections import Counter
p=Path(__file__).resolve().parent
src=json.loads((p/'bls/supplied_occupations.json').read_text(encoding='utf-8'))
parts=['professional','education_media','services_production','operations']
scores=[]
for part in parts:
    scores.extend(json.loads((p/f'scores_{part}.json').read_text(encoding='utf-8-sig')))
assert len(scores)==len(src)==831,(len(scores),len(src))
assert len({r['code'] for r in scores})==831
assert {r['code'] for r in scores}=={r['code'] for r in src}
bycode={r['code']:r for r in scores}
for r in scores:
    assert all(isinstance(r[k],(int,float)) for k in ['low','median','high'])
    assert 0<=r['low']<=r['median']<=r['high']<=100,r
    assert len(r['explanation'].strip())>40,r
merged=[dict(r,**{k:v for k,v in bycode[r['code']].items() if k!='code'}) for r in src]
(p/'combined.json').write_text(json.dumps(merged,ensure_ascii=False,indent=2),encoding='utf-8')
jobs=sum(r['employment_2025'] for r in merged)
eq={k:sum(r['employment_2025']*r[k]/100 for r in merged) for k in ['low','median','high']}
summary=dict(rows=len(merged),jobs=jobs,essential_equivalents=eq,essential_shares={k:v/jobs for k,v in eq.items()},
             zero_median=sum(r['median']==0 for r in merged),zero_all=sum(r['high']==0 for r in merged))
(p/'independent_summary.json').write_text(json.dumps(summary,indent=2),encoding='utf-8')
print(json.dumps(summary,indent=2))
print('Largest median contributions:')
for r in sorted(merged,key=lambda x:x['employment_2025']*x['median'],reverse=True)[:18]:
    print(r['code'],r['title'],r['low'],r['median'],r['high'],round(r['employment_2025']*r['median']/100))
