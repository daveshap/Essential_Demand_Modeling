import openpyxl,json,copy
from pathlib import Path
p=Path(__file__).resolve().parent
old=p.parent/'outputs/Essential_Demand_ChatGPT_6_Pro_2026-11-16_083431_EST.xlsx'
new=p.parent/'outputs/Essential_vs_Derived_Demand_by_Occupation_GPT_6_Astra_2026-11-16_0834.xlsx'
allowed={('Summary','A4'),('Occupations','A3'),('Method and sources','A5'),('Method and sources','B5'),('Method and sources','B6')}
changes=[]
for data_only in [False,True]:
    a=openpyxl.load_workbook(old,data_only=data_only)
    b=openpyxl.load_workbook(new,data_only=data_only)
    assert a.sheetnames==b.sheetnames
    for sa,sb in zip(a,b):
        assert sa.max_row==sb.max_row and sa.max_column==sb.max_column
        assert sa.freeze_panes==sb.freeze_panes
        assert len(sa.tables)==len(sb.tables)
        assert len(sa.conditional_formatting)==len(sb.conditional_formatting)
        for row in sa:
            for c in row:
                d=sb[c.coordinate]
                if c.value!=d.value:
                    assert (sa.title,c.coordinate) in allowed,(sa.title,c.coordinate,c.value,d.value)
                    if not data_only: changes.append((sa.title,c.coordinate))
                assert c.number_format==d.number_format
                for prop in ['font','alignment','fill','border','protection']:
                    assert copy.copy(getattr(c,prop))==copy.copy(getattr(d,prop)),(sa.title,c.coordinate,prop)
                if isinstance(d.value,str): assert 'ChatGPT 6 Pro' not in d.value and 'GPT-6-based' not in d.value
assert set(changes)==allowed
report={'changed_cells':changes,'all_other_values_formulas_and_cell_styles_unchanged':True}
(p/'attribution_verification.json').write_text(json.dumps(report,indent=2),encoding='utf8')
print(json.dumps(report,indent=2))
