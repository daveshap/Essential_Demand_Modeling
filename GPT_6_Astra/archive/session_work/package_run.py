import csv,json,shutil,re,hashlib
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

work=Path(__file__).resolve().parent
package=work/'package/GPT_6_Astra'
stem='Essential_vs_Derived_Demand_by_Occupation_GPT_6_Astra_2026-11-16_0834'
for d in ['estimates','scripts','source_data/framework','verification/previews','archive/session_work/bls','archive/superseded_workbooks']:
    (package/d).mkdir(parents=True,exist_ok=True)
def copy(src,dest):
    dest.parent.mkdir(parents=True,exist_ok=True)
    shutil.copy2(src,dest)

# Preserve the complete session files, including intermediate drafts and checks.
# Deliberately skip directories here so the runtime junction and package cannot recurse.
for f in work.iterdir():
    if f.is_file(): copy(f,package/'archive/session_work'/f.name)
for f in (work/'bls').iterdir():
    if f.is_file(): copy(f,package/'archive/session_work/bls'/f.name)

for name in ['scores_professional.json','scores_education_media.json','scores_services_production.json','scores_operations.json','combined.json','scoring_brief.md','review_notes.json']:
    copy(work/name,package/'estimates'/name)
for f in (work/'bls').iterdir():
    if f.is_file() and f.suffix!='.py': copy(f,package/'source_data'/f.name)
copy(work/'bls/occupation_user_supplied.xlsx',package/'source_data/BLS_EP_occupation.xlsx')

framework=Path(r'C:\PaidToBeHuman')
for rel in ['README.md','source/Paid_to_Be_Human_Conversation_Export.md',
            'reference/background/Derived vs Essential Demand v01 (text).md',
            'reference/transcripts/derived vs essential demand.md']:
    copy(framework/rel,package/'source_data/framework'/rel)
for number in ['02','04','06','07','08']:
    files=list((framework/'course').glob(f'Lesson {number} - *.docx'))
    assert len(files)==1
    copy(files[0],package/'source_data/framework/course'/files[0].name)
for f in work.glob('framework_Lesson*.txt'):
    copy(f,package/'source_data/framework/extracted'/f.name)
copy(work/'framework_review.md',package/'verification/framework_review.md')
for name in ['independent_summary.json','calculation_checks.json','formula_error_scan.ndjson','summary_inspection.ndjson','saved_workbook_verification.json','attribution_verification.json','final_export_inspection.ndjson']:
    copy(work/name,package/'verification'/name)
for f in work.glob('*.png'): copy(f,package/'verification/previews'/f.name)

out=work.parent/'outputs'
copy(out/(stem+'.xlsx'),package/(stem+'.xlsx'))
for name in ['Essential_Demand_ChatGPT_6_Pro_2026-11-16_083431_EST.xlsx','Essential_Demand_Astra_2026-11-16_083431_EST.xlsx']:
    if (out/name).exists(): copy(out/name,package/'archive/superseded_workbooks'/name)
for f in out.glob('*.inspect.ndjson'):copy(f,package/'verification'/f.name)

rows=json.loads((work/'combined.json').read_text(encoding='utf8'))
headers=['soc_code','occupation','major_group','employment_2025','essential_low_pct','essential_median_pct','essential_high_pct','derived_low_pct','derived_median_pct','derived_high_pct','essential_equivalents_low','essential_equivalents_median','essential_equivalents_high','explanation','source_workbook_row','bls_description_url']
with (package/(stem+'.csv')).open('w',newline='',encoding='utf-8-sig') as f:
    w=csv.writer(f);w.writerow(headers)
    for r in rows:
        w.writerow([r['code'],r['title'],r['major_title'],r['employment_2025'],r['low'],r['median'],r['high'],100-r['high'],100-r['median'],100-r['low'],*[round(r['employment_2025']*r[k]/100,6) for k in ['low','median','high']],r['explanation'],r['source_row'],r['ooh_url']])
source=json.loads((work/'bls/supplied_source_rows.json').read_text(encoding='utf8'))
with (package/'source_data/BLS_EP_Table_1.2_2025-35.csv').open('w',newline='',encoding='utf-8-sig') as f:
    w=csv.DictWriter(f,fieldnames=['code','title','type','employment_thousands_2025','source_row']);w.writeheader()
    for r in source:w.writerow(r)

run=json.loads((work/'run_metadata.json').read_text())
run.update(filename=stem+'.xlsx',model='GPT-6 Astra',folder='GPT_6_Astra',
           packaged_eastern=datetime.now(ZoneInfo('America/New_York')).isoformat(),
           original_attribution='The initial delivered filename incorrectly used ChatGPT 6 Pro. Corrected to GPT-6 Astra after the user identified the error.',
           scores_changed_during_packaging=False)
(package/'run_metadata.json').write_text(json.dumps(run,indent=2),encoding='utf8')

# Adapt the original authoring script to the self-contained package layout.
builder=(work/'build_workbook.mjs').read_text(encoding='utf8')
builder=builder.replace("const read=async name=>JSON.parse((await fs.readFile(path.join(here,name),'utf8')).replace(/^\\uFEFF/,''));", "const packageRoot=path.resolve(here,'..');\nconst support=path.join(packageRoot,'rebuild','verification');\nawait fs.mkdir(support,{recursive:true});\nconst paths={'combined.json':'estimates/combined.json','bls/supplied_source_rows.json':'source_data/supplied_source_rows.json','bls/supplied_metadata.json':'source_data/supplied_metadata.json','independent_summary.json':'verification/independent_summary.json'};\nconst read=async name=>JSON.parse((await fs.readFile(path.join(packageRoot,paths[name]||name),'utf8')).replace(/^\\uFEFF/,''));")
assert 'const packageRoot=' in builder
builder=builder.replace("const runfile=path.join(here,'run_metadata.json');","const runfile=path.join(packageRoot,'run_metadata.json');")
builder=builder.replace("const outdir=path.resolve(here,'../outputs');","const outdir=path.join(packageRoot,'rebuild');")
builder=builder.replace('Essential_Demand_ChatGPT_6_Pro_','Essential_vs_Derived_Demand_by_Occupation_GPT_6_Astra_')
builder=builder.replace("['ChatGPT 6 Pro (requested run label). Created '","['GPT-6 Astra. Created '")
builder=builder.replace('ChatGPT 6 Pro (requested run label; execution provenance in Method and sources)','Prepared by GPT-6 Astra in Codex')
builder=builder.replace("['Run label','ChatGPT 6 Pro — requested comparison label.']","['Model','GPT-6 Astra']")
builder=builder.replace('Prepared in Codex by a GPT-6-based assistant, with three same-model subagents scoring disjoint occupation sets and a review pass. This session does not independently verify a Pro product/tier designation.','Prepared by GPT-6 Astra in Codex, with three same-model subagents scoring disjoint occupation sets and a review pass.')
builder=builder.replace("path.join(here,'formula_error_scan.ndjson')","path.join(support,'formula_error_scan.ndjson')")
builder=builder.replace("path.join(here,'summary_inspection.ndjson')","path.join(support,'summary_inspection.ndjson')")
builder=builder.replace("path.join(here,'calculation_checks.json')","path.join(support,'calculation_checks.json')")
builder=builder.replace("path.join(here,file)","path.join(support,file)")
builder=builder.replace("path.join(here,'output_path.txt')","path.join(support,'output_path.txt')")
assert 'ChatGPT 6 Pro' not in builder and 'GPT-6-based' not in builder
(package/'scripts/build_workbook.mjs').write_text(builder,encoding='utf8')

revisions=[
 ['2026-11-16 before final export','49-9092','essential_high_pct','1','0','Removed speculative demonstration/heritage allowance from commercial diving. Only the high estimate changed; other scores unchanged.'],
 ['2026-11-16 before final export','25-9044','explanation','Included research duties','Instructional-support duties','Corrected the occupational scope to materials, discussions, examinations and grading. Scores unchanged.'],
 ['2026-11-16 before final export','25-1193','explanation','Sounded like direct fitness coaching','Postsecondary recreation and fitness studies','Clarified professional education, physiology, program design and facilities management. Scores unchanged.'],
 ['2026-11-16 before final export','39-9011','explanation','Generic human attachment rationale','Explicit sustained-care versus preschool instruction assumption','Explained differing childcare and preschool estimates. Scores unchanged.'],
 ['2026-11-16 before final export','Employment formulas','conversion','Raw thousands multiplied by 1000','ROUND(source thousands*1000,0)','Normalized binary floating-point artifacts in four converted job counts. Source values and scores unchanged.'],
 ['2026-11-16 after user correction','Workbook attribution','model','ChatGPT 6 Pro requested label / Codex GPT-6 provenance','GPT-6 Astra','Corrected misleading author label and filenames. Preserved original timestamp and all estimates.'],
 ['2026-11-16 packaging','Delivery organization','folder','Loose workbook in repository root','GPT_6_Astra','Saved source data, estimates, scripts, methods, checks, revision history and original session files. Claude folder untouched.']
]
with (package/'revision_log.csv').open('w',newline='',encoding='utf-8-sig') as f:
    w=csv.writer(f);w.writerow(['when','scope','field','previous','current','reason']);w.writerows(revisions)
print(json.dumps({'package':str(package),'workbook':stem+'.xlsx','rows':len(rows)},indent=2))
