import json, hashlib, re, shutil
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import openpyxl

root = Path(__file__).resolve().parent
src = Path(r'C:\Users\daves\Downloads\occupation.xlsx')
w = openpyxl.load_workbook(src, read_only=True, data_only=True)
s = w['Table 1.2']
prior = {r['code']: r for r in json.loads((root/'bls/occupations.json').read_text(encoding='utf-8-sig'))}
rows, detailed, major = [], [], None
for ix, row in enumerate(s.values, 1):
    title, code, kind, emp = row[:4]
    if not isinstance(code, str) or not re.fullmatch(r'\d\d-\d{4}', code):
        continue
    item = dict(source_row=ix, title=title.strip(), code=code, type=kind, employment_thousands_2025=emp)
    rows.append(item)
    if code.endswith('-0000') and code != '00-0000':
        major = item
    if kind == 'Line item':
        assert code in prior, code
        assert round(emp*1000) == prior[code]['employment_2025'], code
        assert title.strip() == prior[code]['title'], (code, title, prior[code]['title'])
        link = re.search(r'href="([^"]+)"', row[16] or '')
        detailed.append(dict(code=code,title=title.strip(),major_code=major['code'],major_title=major['title'],
                             employment_2025=round(emp*1000),employment_thousands_2025=emp,
                             source_row=ix,ooh_url=link.group(1) if link else '',self_employed_percent=row[9]))
assert set(prior) == {r['code'] for r in detailed}
meta = dict(source_path=str(src),source_sheet='Table 1.2',sha256=hashlib.sha256(src.read_bytes()).hexdigest(),
            processed_at=datetime.now(ZoneInfo('America/New_York')).isoformat(),
            source_url='https://www.bls.gov/emp/ind-occ-matrix/occupation.xlsx',
            official_html='https://www.bls.gov/emp/tables/occupational-projections-and-characteristics.htm',
            base_year=2025,detailed_count=len(detailed),detailed_employment=sum(r['employment_2025'] for r in detailed),
            published_total=round(rows[0]['employment_thousands_2025']*1000),
            comparison='All detailed codes, titles and 2025 employment match the freshly retrieved official BLS HTML table exactly.')
(root/'bls/supplied_occupations.json').write_text(json.dumps(detailed,ensure_ascii=False,indent=2),encoding='utf-8')
(root/'bls/supplied_source_rows.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
(root/'bls/supplied_metadata.json').write_text(json.dumps(meta,indent=2),encoding='utf-8')
shutil.copy2(src,root/'bls/occupation_user_supplied.xlsx')
print(json.dumps(meta,indent=2))
