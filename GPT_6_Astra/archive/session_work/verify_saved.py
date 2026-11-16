import json,math,hashlib
from pathlib import Path
import openpyxl
p=Path(__file__).resolve().parent
target=Path((p/'output_path.txt').read_text(encoding='utf-8'))
expected=json.loads((p/'combined.json').read_text(encoding='utf-8'))
controls=json.loads((p/'independent_summary.json').read_text(encoding='utf-8'))
w=openpyxl.load_workbook(target,data_only=False)
v=openpyxl.load_workbook(target,data_only=True)
assert w.sheetnames==['Summary','Occupations','BLS source','Method and sources']
o=v['Occupations']; f=w['Occupations']
assert o.max_row==836
for ix,r in enumerate(expected,6):
    cells=[o.cell(ix,c).value for c in range(1,17)]
    assert cells[0]==r['code'] and cells[1]==r['title']
    assert cells[2]==r['employment_2025']
    assert cells[9]==r['explanation'] and cells[10]==r['major_title']
    assert cells[14]==r['source_row']
    for j,k in enumerate(['low','median','high']):
        assert abs(cells[3+j]-r[k]/100)<1e-12
        assert abs(cells[11+j]-r['employment_2025']*r[k]/100)<1e-7
    assert all(abs(cells[a]+cells[b]-1)<1e-12 for a,b in [(3,8),(4,7),(5,6)])
    assert f.cell(ix,7).value==f'=1-F{ix}'
    assert f.cell(ix,8).value==f'=1-E{ix}'
    assert f.cell(ix,9).value==f'=1-D{ix}'
    assert f.cell(ix,12).value==f'=C{ix}*D{ix}'
    assert f.cell(ix,13).value==f'=C{ix}*E{ix}'
    assert f.cell(ix,14).value==f'=C{ix}*F{ix}'
for j,k in enumerate(['low','median','high'],2):
    assert abs(v['Summary'].cell(8,j).value-controls['essential_shares'][k])<1e-12
    assert abs(v['Summary'].cell(10,j).value-controls['essential_equivalents'][k])<1e-7
assert v['Summary']['B15'].value==170267200
assert v['Summary']['B16'].value==170280800
assert v['Summary']['B17'].value==13600
assert v['Summary']['B47'].value==controls['zero_median']
assert v['Summary']['B48'].value==controls['zero_all']
for row in range(23,45):
    title=v['Summary'].cell(row,1).value
    members=[r for r in expected if r['major_title']==title]
    jobs=sum(r['employment_2025'] for r in members)
    assert v['Summary'].cell(row,2).value==jobs
    for col,key in [(3,'low'),(4,'median'),(5,'high')]:
        control=sum(r['employment_2025']*r[key]/100 for r in members)/jobs
        assert abs(v['Summary'].cell(row,col).value-control)<1e-12
raw=json.loads((p/'bls/supplied_source_rows.json').read_text(encoding='utf8'))
for ix,r in enumerate(raw,5):
    assert [v['BLS source'].cell(ix,c).value for c in range(1,6)]==[r['code'],r['title'],r['type'],r['employment_thousands_2025'],r['source_row']]
errors=[]
for s in v:
    for row in s:
        errors.extend((s.title,c.coordinate,c.value) for c in row if c.data_type=='e')
assert not errors,errors
assert f.freeze_panes=='C6'
assert f.tables['OccupationEstimates'].ref=='A5:P836'
assert f.tables['OccupationEstimates'].autoFilter is not None
assert 'yyyy-mm-dd' in w['Method and sources']['B8'].number_format
assert all(s.sheet_view.showGridLines is False for s in w)
report={'output':str(target),'sha256':hashlib.sha256(target.read_bytes()).hexdigest(),'bytes':target.stat().st_size,
        'occupation_rows_verified':len(expected),'raw_source_rows_verified':len(raw),'formula_errors':errors,
        'checks':'Every score, title, code, employment value, explanation, complement and weighted formula; all group and headline totals; panes, filters, gridlines and timestamp formatting.'}
(p/'saved_workbook_verification.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
print(json.dumps(report,indent=2))
