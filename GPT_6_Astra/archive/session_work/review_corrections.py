import json
from pathlib import Path
p=Path(__file__).resolve().parent
fixes={
'25-9044':'Postsecondary assistants mainly prepare instructional materials, lead discussions, administer examinations and grade work. Some direct human mentoring can matter, but support for a valued professor is not itself essential demand.',
'25-1193':'Teaching recreation, leisure and fitness studies mainly supplies professional knowledge. Human mentorship and role modeling can carry a premium, while instruction in physiology, program design and facilities management remains largely instrumental.',
'39-9011':'This estimate assumes sustained human care and attachment account for more demand than in instruction-focused preschool teaching. Practical supervision remains derived, while families may specifically want a real human caring for their child.'
}
for name in ['scores_education_media.json','scores_services_production.json']:
    f=p/name
    rows=json.loads(f.read_text(encoding='utf-8-sig'))
    for row in rows:
        if row['code'] in fixes: row['explanation']=fixes[row['code']]
    f.write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
(p/'review_notes.json').write_text(json.dumps({'wording_corrections':fixes,'numeric_changes':[]},indent=2),encoding='utf-8')
print('Applied three explanation corrections; numerical estimates unchanged.')
