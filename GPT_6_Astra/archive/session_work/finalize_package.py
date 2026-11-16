import os,csv,json,hashlib,shutil
from pathlib import Path
work=Path(__file__).resolve().parent
package=work/'package/GPT_6_Astra'
stem='Essential_vs_Derived_Demand_by_Occupation_GPT_6_Astra_2026-11-16_0834'
for name in ['attribution_verification.json']:
    shutil.copy2(work/name,package/'verification'/name)
for name in ['verify_attribution.py','correct_attribution.mjs','package_run.py','finalize_package.py']:
    shutil.copy2(work/name,package/'archive/session_work'/name)
for f in work.glob('corrected_*.png'):
    shutil.copy2(f,package/'verification/previews'/f.name)
# Include the original root delivery in the manifest; deployment moves it to this archive path.
oldroot=Path(r'C:\Essential_Demand_Modeling\Essential_Demand_ChatGPT_6_Pro_2026-11-16_083431_EST.xlsx')
rootarchive=package/'archive/root_delivery'/oldroot.name
rootarchive.parent.mkdir(exist_ok=True)
shutil.copy2(oldroot,rootarchive)

# The tested portable rebuild must reproduce every cell value/formula.
import openpyxl
differences=[]
for cached in [False,True]:
    original=openpyxl.load_workbook(package/(stem+'.xlsx'),data_only=cached)
    rebuilt=openpyxl.load_workbook(package/'rebuild'/(stem+'.xlsx'),data_only=cached)
    for a,b in zip(original,rebuilt):
        for row in a:
            for c in row:
                if c.value!=b[c.coordinate].value:
                    differences.append((cached,a.title,c.coordinate,c.value,b[c.coordinate].value))
assert not differences,differences[:10]
(package/'verification/rebuild_comparison.json').write_text(json.dumps({'all_cell_values_and_formulas_match':True,'compared':'Package root workbook versus portable script rebuild','differences':differences},indent=2),encoding='utf8')

files=[]
def walk(directory):
    for f in sorted(directory.iterdir()):
        if f.is_symlink() or (hasattr(f,'is_junction') and f.is_junction()): continue
        if f.is_dir(): walk(f)
        elif f.name!='manifest_sha256.csv': files.append(f)
walk(package)
with (package/'manifest_sha256.csv').open('w',newline='',encoding='utf-8') as stream:
    writer=csv.writer(stream);writer.writerow(['relative_path','bytes','sha256'])
    for f in files:writer.writerow([f.relative_to(package).as_posix(),f.stat().st_size,hashlib.sha256(f.read_bytes()).hexdigest()])
print(json.dumps({'files':len(files),'bytes':sum(f.stat().st_size for f in files),'package':str(package)},indent=2))
