"""Build the essential-vs-derived-demand workbook.

Usage: py -3.14 -I build_workbook.py <Claude_Opus_5.5 folder>

Reads source_data/BLS_EP_Table_1.2_2025-35.csv and estimates/*.txt, validates
them, and writes a timestamped .xlsx (live formulas) plus a flat CSV export.
"""
import csv
import hashlib
import os
import sys
from collections import defaultdict
from datetime import datetime

import xlsxwriter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from load_estimates import load_bls, load_estimates  # noqa: E402

MODEL = 'Claude Opus 5.5'
MODEL_ID = 'claude-opus-5-5'
SRC_PAGE = 'https://www.bls.gov/emp/tables/occupational-projections-and-characteristics.htm'
SRC_XLSX = 'https://www.bls.gov/emp/ind-occ-matrix/occupation.xlsx'


def main(base):
    now = datetime.now().astimezone()
    off = now.strftime('%z')
    tz = {'-0500': 'EST', '-0400': 'EDT'}.get(off, '')
    stamp_human = f"{now:%Y-%m-%d %H:%M} {tz} (UTC{off[:3]}:{off[3:]})".replace('  ', ' ')
    stamp_file = f"{now:%Y-%m-%d_%H%M}"

    bls = load_bls(os.path.join(base, 'source_data', 'BLS_EP_Table_1.2_2025-35.csv'))
    est, errors = load_estimates(os.path.join(base, 'estimates'), bls)
    if errors:
        print('\n'.join(errors))
        sys.exit('validation failed; workbook not written')

    src_xlsx = os.path.join(base, 'source_data', 'BLS_EP_occupation.xlsx')
    sha = hashlib.sha256(open(src_xlsx, 'rb').read()).hexdigest()

    majors = [r for r in bls if r['occupation_type'] == 'Summary' and r['level'] == 1]
    major_title = {r['soc_code']: r['title'] for r in majors}
    items = [r for r in bls if r['occupation_type'] == 'Line item']
    n = len(items)
    total_emp = sum(r['emp_2025_thousands'] for r in items)

    out_name = f'Essential_vs_Derived_Demand_by_Occupation_{MODEL.replace(" ", "_")}_{stamp_file}.xlsx'
    out_path = os.path.join(base, out_name)
    wb = xlsxwriter.Workbook(out_path)
    wb.set_properties({
        'title': 'Essential vs derived demand by U.S. occupation',
        'subject': 'Share of each BLS detailed occupation attributable to Essential Demand vs derived demand',
        'author': f'{MODEL} ({MODEL_ID}), Anthropic',
        'company': 'Essential_Demand_Modeling',
        'comments': f'Generated {stamp_human} by {MODEL} in Claude Code. Baseline: BLS Employment Projections Table 1.2 (2025).',
        'created': now.replace(tzinfo=None),
    })

    # ---------- formats ----------
    F = {}
    F['title'] = wb.add_format({'bold': True, 'font_size': 16})
    F['sub'] = wb.add_format({'italic': True, 'font_color': '#444444'})
    F['stamp'] = wb.add_format({'bold': True, 'font_color': '#1F4E79'})
    F['h2'] = wb.add_format({'bold': True, 'font_size': 12, 'bottom': 1})
    F['hdr'] = wb.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1, 'text_wrap': True, 'valign': 'top'})
    F['txt'] = wb.add_format({'valign': 'top'})
    F['wrap'] = wb.add_format({'text_wrap': True, 'valign': 'top'})
    F['k'] = wb.add_format({'num_format': '#,##0.0', 'valign': 'top'})
    F['pct'] = wb.add_format({'num_format': '0.0%', 'valign': 'top'})
    F['pct_in'] = wb.add_format({'num_format': '0.0%', 'valign': 'top', 'bg_color': '#FFF9E6'})
    F['usd'] = wb.add_format({'num_format': '$#,##0', 'valign': 'top'})
    F['int'] = wb.add_format({'num_format': '#,##0', 'valign': 'top'})
    F['k_b'] = wb.add_format({'num_format': '#,##0.0', 'bold': True, 'top': 1})
    F['pct_b'] = wb.add_format({'num_format': '0.0%', 'bold': True, 'top': 1})
    F['txt_b'] = wb.add_format({'bold': True, 'top': 1})
    F['big_pct'] = wb.add_format({'num_format': '0.0%', 'bold': True, 'font_size': 14, 'font_color': '#1F4E79'})
    F['big_k'] = wb.add_format({'num_format': '#,##0', 'bold': True, 'font_size': 12})
    F['note'] = wb.add_format({'italic': True, 'font_color': '#666666', 'text_wrap': True, 'valign': 'top'})
    F['bad'] = wb.add_format({'bg_color': '#F8CBAD'})

    ws_sum = wb.add_worksheet('Summary')
    ws_occ = wb.add_worksheet('Occupations')
    ws_met = wb.add_worksheet('Method')
    ws_hier = wb.add_worksheet('BLS hierarchy')
    ws_src = wb.add_worksheet('Sources')

    # ---------- Occupations ----------
    HDR_ROW = 3                     # 0-indexed header row
    R0 = HDR_ROW + 1                # first data row (0-indexed)
    first, last = R0 + 1, R0 + n    # 1-indexed Excel rows for formulas
    rng = lambda col: f"Occupations!${col}${first}:${col}${last}"  # noqa: E731

    ws_occ.write(0, 0, 'Essential vs. derived demand: all 831 BLS detailed occupations', F['title'])
    ws_occ.write(1, 0, f'Prepared by {MODEL} ({MODEL_ID}), Anthropic, in Claude Code. Generated {stamp_human}.', F['stamp'])
    ws_occ.write(2, 0, 'Shares are of each occupation\'s current, job-weighted work. Edit any yellow Low/Median/High cell and every total in the workbook updates. Rows where low > median or median > high turn red.', F['sub'])
    headers = ['SOC code', 'Occupation (BLS title)', 'Major group code', 'Major group',
               'Employment 2025 (thousands)', 'Self-employed 2025', 'Median annual wage 2025',
               'Essential share: Low', 'Essential share: Median', 'Essential share: High',
               'Derived share: Median', 'Essential jobs: Low (thousands)', 'Essential jobs: Median (thousands)',
               'Essential jobs: High (thousands)', 'Primary mechanism', 'All mechanisms', 'Rationale']
    for c, h in enumerate(headers):
        ws_occ.write(HDR_ROW, c, h, F['hdr'])
    ws_occ.set_row(HDR_ROW, 45)
    for i, r in enumerate(items):
        row = R0 + i
        xr = row + 1
        e = est[r['soc_code']]
        ws_occ.write_string(row, 0, r['soc_code'], F['txt'])
        ws_occ.write_string(row, 1, r['title'], F['wrap'])
        ws_occ.write_string(row, 2, r['major_group_code'], F['txt'])
        ws_occ.write_string(row, 3, major_title[r['major_group_code']], F['wrap'])
        ws_occ.write_number(row, 4, r['emp_2025_thousands'], F['k'])
        if r['pct_self_employed_2025'] is not None:
            ws_occ.write_number(row, 5, r['pct_self_employed_2025'] / 100, F['pct'])
        else:
            ws_occ.write_blank(row, 5, None, F['pct'])
        if r['median_wage_2025'] is not None:
            ws_occ.write_number(row, 6, r['median_wage_2025'], F['usd'])
        else:
            ws_occ.write_blank(row, 6, None, F['usd'])
        ws_occ.write_number(row, 7, e['low'] / 100, F['pct_in'])
        ws_occ.write_number(row, 8, e['median'] / 100, F['pct_in'])
        ws_occ.write_number(row, 9, e['high'] / 100, F['pct_in'])
        emp = r['emp_2025_thousands']
        ws_occ.write_formula(row, 10, f'=1-I{xr}', F['pct'], 1 - e['median'] / 100)
        ws_occ.write_formula(row, 11, f'=E{xr}*H{xr}', F['k'], emp * e['low'] / 100)
        ws_occ.write_formula(row, 12, f'=E{xr}*I{xr}', F['k'], emp * e['median'] / 100)
        ws_occ.write_formula(row, 13, f'=E{xr}*J{xr}', F['k'], emp * e['high'] / 100)
        tags = e['tags']
        ws_occ.write_string(row, 14, tags[0] if tags else 'None (derived)', F['txt'])
        ws_occ.write_string(row, 15, '; '.join(tags) if tags else 'None', F['wrap'])
        ws_occ.write_string(row, 16, e['rationale'], F['wrap'])
    widths = [9, 40, 9, 26, 12, 10, 11, 9, 9, 9, 9, 11, 11, 11, 15, 22, 100]
    for c, w in enumerate(widths):
        ws_occ.set_column(c, c, w)
    ws_occ.freeze_panes(R0, 2)
    ws_occ.autofilter(HDR_ROW, 0, R0 + n - 1, len(headers) - 1)
    ws_occ.data_validation(R0, 7, R0 + n - 1, 9, {
        'validate': 'decimal', 'criteria': 'between', 'minimum': 0, 'maximum': 1,
        'error_title': 'Share out of range', 'error_message': 'Enter a share between 0% and 100%.'})
    ws_occ.conditional_format(R0, 8, R0 + n - 1, 8, {
        'type': '2_color_scale', 'min_type': 'num', 'min_value': 0, 'max_type': 'num', 'max_value': 1,
        'min_color': '#FFFFFF', 'max_color': '#63BE7B'})
    ws_occ.conditional_format(R0, 7, R0 + n - 1, 9, {
        'type': 'formula', 'criteria': f'=OR($H{first}>$I{first},$I{first}>$J{first})', 'format': F['bad']})
    ws_occ.write(R0 + n + 1, 3, 'Total, all 831 detailed occupations', F['txt_b'])
    ws_occ.write_formula(R0 + n + 1, 4, f'=SUM(E{first}:E{last})', F['k_b'], total_emp)
    for c, col in ((11, 'L'), (12, 'M'), (13, 'N')):
        ws_occ.write_formula(R0 + n + 1, c, f'=SUM({col}{first}:{col}{last})', F['k_b'])
    for c, num in ((7, 'L'), (8, 'M'), (9, 'N')):
        ws_occ.write_formula(R0 + n + 1, c, f'={num}{R0 + n + 2}/E{R0 + n + 2}', F['pct_b'])

    # ---------- Summary ----------
    s = ws_sum
    s.set_column(0, 0, 12)
    s.set_column(1, 1, 52)
    s.set_column(2, 10, 13)
    s.write(0, 0, 'Essential vs. derived demand by U.S. occupation', F['title'])
    s.write(1, 0, f'Prepared by {MODEL} (model ID {MODEL_ID}), Anthropic, working in Claude Code', F['stamp'])
    s.write(2, 0, f'Generated {stamp_human}', F['stamp'])
    s.write(3, 0, 'Baseline: BLS Employment Projections Table 1.2, 2025 employment, all 831 detailed occupations '
                  f'({total_emp:,.1f} thousand jobs, self-employed included; summary rows excluded).', F['sub'])
    s.write(5, 0, 'Question: of all the jobs that exist today, how much is Essential Demand (the human is the point) '
                  'and how much is derived demand (labor wanted as a means to an outcome, Marshall 1890)?', F['txt'])
    s.write(6, 0, 'Essential share = the share of an occupation\'s current work that buyers, audiences or institutions '
                  'would still pay a human to do if a disclosed machine matched the result at equal or lower cost.', F['txt'])
    s.write(7, 0, 'These are judgment estimates, not measurements, and not a forecast of automation. See the Method sheet.', F['note'])

    r = 9
    s.write_row(r, 1, ['', 'Low', 'Median', 'High'], F['hdr'])
    s.write(r + 1, 1, 'Essential-demand share of all jobs', F['txt_b'])
    s.write(r + 2, 1, 'Derived-demand share of all jobs', F['txt'])
    s.write(r + 3, 1, 'Essential-demand jobs (thousands, job-weighted equivalents)', F['txt'])
    tot_cell = f'$C${r + 6}'
    for j, col in enumerate(('L', 'M', 'N')):
        c = 2 + j
        s.write_formula(r + 3, c, f'=SUM({rng(col)})', F['big_k'])
        cl = xlsxwriter.utility.xl_rowcol_to_cell(r + 3, c)
        s.write_formula(r + 1, c, f'={cl}/{tot_cell}', F['big_pct'])
        s.write_formula(r + 2, c, f'=1-{xlsxwriter.utility.xl_rowcol_to_cell(r + 1, c)}', F['pct'])
    s.write(r + 5, 1, 'Total employment covered (thousands)', F['txt'])
    s.write_formula(r + 5, 2, f'=SUM({rng("E")})', F['k'], total_emp)
    s.write(r + 6, 1, 'Detailed occupations scored', F['txt'])
    s.write_formula(r + 6, 2, f'=COUNTA({rng("A")})', F['int'], n)
    s.write(r + 7, 1, 'Occupations with a 0% median (pure derived demand)', F['txt'])
    s.write_formula(r + 7, 2, f'=COUNTIF({rng("I")},0)', F['int'])
    s.write(r + 8, 1, 'Jobs in those occupations (thousands; share of all jobs)', F['txt'])
    s.write_formula(r + 8, 2, f'=SUMIF({rng("I")},0,{rng("E")})', F['k'])
    s.write_formula(r + 8, 3, f'=C{r + 9}/{tot_cell}', F['pct'])
    s.write(r + 9, 1, 'Occupations with a median of 50% or more', F['txt'])
    s.write_formula(r + 9, 2, f'=COUNTIF({rng("I")},">=0.5")', F['int'])
    s.write(r + 10, 1, 'Jobs in those occupations (thousands; share of all jobs)', F['txt'])
    s.write_formula(r + 10, 2, f'=SUMIF({rng("I")},">=0.5",{rng("E")})', F['k'])
    s.write_formula(r + 10, 3, f'=C{r + 11}/{tot_cell}', F['pct'])

    # by major group
    r = 22
    s.write(r, 0, 'By major occupational group', F['h2'])
    gh = ['Group code', 'Major group', 'Employment 2025 (thousands)', 'Essential jobs: Low', 'Essential jobs: Median',
          'Essential jobs: High', 'Essential share: Low', 'Essential share: Median', 'Essential share: High',
          'Share of all essential jobs (median)']
    s.write_row(r + 1, 0, gh, F['hdr'])
    s.set_row(r + 1, 45)
    g0 = r + 2
    for i, g in enumerate(majors):
        rr = g0 + i
        xr = rr + 1
        s.write_string(rr, 0, g['soc_code'], F['txt'])
        s.write_string(rr, 1, g['title'], F['txt'])
        s.write_formula(rr, 2, f'=SUMIF({rng("C")},$A{xr},{rng("E")})', F['k'])
        for j, col in enumerate(('L', 'M', 'N')):
            s.write_formula(rr, 3 + j, f'=SUMIF({rng("C")},$A{xr},{rng(col)})', F['k'])
            dc = xlsxwriter.utility.xl_col_to_name(3 + j)
            s.write_formula(rr, 6 + j, f'={dc}{xr}/$C{xr}', F['pct'])
    gl = g0 + len(majors)
    s.write(gl, 1, 'All occupations', F['txt_b'])
    for c in range(2, 6):
        cn = xlsxwriter.utility.xl_col_to_name(c)
        s.write_formula(gl, c, f'=SUM({cn}{g0 + 1}:{cn}{gl})', F['k_b'])
    for j in range(3):
        dc = xlsxwriter.utility.xl_col_to_name(3 + j)
        s.write_formula(gl, 6 + j, f'={dc}{gl + 1}/$C{gl + 1}', F['pct_b'])
    for i in range(len(majors) + 1):
        rr = g0 + i
        s.write_formula(rr, 9, f'=E{rr + 1}/$E${gl + 1}', F['pct_b'] if i == len(majors) else F['pct'])
    s.conditional_format(g0, 7, gl - 1, 7, {'type': 'data_bar', 'bar_color': '#63BE7B', 'bar_solid': True,
                                            'min_type': 'num', 'min_value': 0, 'max_type': 'num', 'max_value': 0.5})

    # distribution
    r = gl + 3
    s.write(r, 0, 'Distribution of occupations by median essential share', F['h2'])
    s.write_row(r + 1, 0, ['Band', 'Meaning', 'Occupations', 'Employment (thousands)', 'Share of all jobs',
                           'Essential jobs: Median', 'Share of essential total'], F['hdr'])
    s.set_row(r + 1, 30)
    bands = [
        ('0%', 'Pure derived demand', '=0', None),
        ('>0 to <5%', 'A thin human-specific sliver', '>0', '<0.05'),
        ('5 to <20%', 'Minority premium segment or small human core', '>=0.05', '<0.2'),
        ('20 to <50%', 'Mixed bundle with a large human part', '>=0.2', '<0.5'),
        ('50 to <80%', 'The human is mostly the point', '>=0.5', '<0.8'),
        ('80 to 100%', 'The human is the product', '>=0.8', None),
    ]
    b0 = r + 2
    for i, (band, meaning, c1, c2) in enumerate(bands):
        rr = b0 + i
        crit = f'{rng("I")},"{c1}"' + (f',{rng("I")},"{c2}"' if c2 else '')
        s.write(rr, 0, band, F['txt'])
        s.write(rr, 1, meaning, F['txt'])
        s.write_formula(rr, 2, f'=COUNTIFS({crit})', F['int'])
        s.write_formula(rr, 3, f'=SUMIFS({rng("E")},{crit})', F['k'])
        s.write_formula(rr, 4, f'=D{rr + 1}/{tot_cell}', F['pct'])
        s.write_formula(rr, 5, f'=SUMIFS({rng("M")},{crit})', F['k'])
        s.write_formula(rr, 6, f'=F{rr + 1}/$D$13', F['pct'])  # D13 = median essential jobs (headline block)
    bl = b0 + len(bands)
    s.write(bl, 1, 'All occupations', F['txt_b'])
    for c, fmt in ((2, F['k_b']), (3, F['k_b']), (5, F['k_b'])):
        cn = xlsxwriter.utility.xl_col_to_name(c)
        s.write_formula(bl, c, f'=SUM({cn}{b0 + 1}:{cn}{bl})', fmt)
    s.write_formula(bl, 4, f'=D{bl + 1}/{tot_cell}', F['pct_b'])
    s.write_formula(bl, 6, f'=F{bl + 1}/$D$13', F['pct_b'])

    # mechanisms (primary = first tag listed)
    r = bl + 3
    s.write(r, 0, 'Kinds of essential demand (each occupation counted under its primary mechanism)', F['h2'])
    s.write_row(r + 1, 0, ['Mechanism', 'What must remain human', 'Occupations', 'Employment (thousands)',
                           'Essential jobs: Median', 'Share of essential total'], F['hdr'])
    s.set_row(r + 1, 30)
    mech = [('Presence', 'An actual person there'), ('Provenance', 'Human making or authorship'),
            ('Accountability', 'A person who answers for it'), ('Affinity', 'A particular person or relationship'),
            ('None (derived)', 'Nothing: the human is a means to an outcome')]
    m0 = r + 2
    for i, (m, what) in enumerate(mech):
        rr = m0 + i
        s.write(rr, 0, m, F['txt'])
        s.write(rr, 1, what, F['txt'])
        s.write_formula(rr, 2, f'=COUNTIF({rng("O")},A{rr + 1})', F['int'])
        s.write_formula(rr, 3, f'=SUMIF({rng("O")},A{rr + 1},{rng("E")})', F['k'])
        s.write_formula(rr, 4, f'=SUMIF({rng("O")},A{rr + 1},{rng("M")})', F['k'])
        s.write_formula(rr, 5, f'=E{rr + 1}/$D$13', F['pct'])
    s.write(m0 + len(mech), 0, 'The four labels come from Paid to Be Human and are used only as tags; the primary mechanism is the first tag listed for an occupation. Most occupations with essential demand carry more than one.', F['note'])

    # top contributors (static snapshot)
    r = m0 + len(mech) + 3
    s.write(r, 0, 'Largest contributors to the median total (snapshot at generation; sort Occupations by column M for a live view)', F['h2'])
    s.write_row(r + 1, 0, ['SOC code', 'Occupation', 'Employment (thousands)', 'Median share', 'Essential jobs: Median',
                           'Cumulative share of essential total'], F['hdr'])
    s.set_row(r + 1, 30)
    ranked = sorted(items, key=lambda x: -x['emp_2025_thousands'] * est[x['soc_code']]['median'])
    med_total = sum(x['emp_2025_thousands'] * est[x['soc_code']]['median'] / 100 for x in items)
    cum = 0.0
    for i, x in enumerate(ranked[:25]):
        rr = r + 2 + i
        ej = x['emp_2025_thousands'] * est[x['soc_code']]['median'] / 100
        cum += ej
        s.write_string(rr, 0, x['soc_code'], F['txt'])
        s.write_string(rr, 1, x['title'], F['txt'])
        s.write_number(rr, 2, x['emp_2025_thousands'], F['k'])
        s.write_number(rr, 3, est[x['soc_code']]['median'] / 100, F['pct'])
        s.write_number(rr, 4, ej, F['k'])
        s.write_number(rr, 5, cum / med_total, F['pct'])
    s.activate()

    # ---------- Method ----------
    m = ws_met
    m.set_column(0, 0, 130)
    method_md = open(os.path.join(base, 'METHOD.md'), encoding='utf-8').read().split('\n')
    m.write(0, 0, f'Method (prepared by {MODEL}, {MODEL_ID}; generated {stamp_human})', F['title'])
    row = 2
    for line in method_md:
        if line.startswith('# '):
            continue
        if line.startswith('## '):
            row += 1
            m.write(row, 0, line[3:], F['h2'])
        elif line.startswith('|'):
            cells = [c.strip() for c in line.strip('|').split('|')]
            if set(''.join(cells)) <= set('-: '):
                continue
            m.write(row, 0, '   ' + '  ·  '.join(cells), F['wrap'])
        elif line.strip() == '':
            continue
        else:
            text = line.replace('**', '').replace('`', '')
            m.write(row, 0, text, F['wrap'])
        row += 1

    # ---------- BLS hierarchy & reconciliation ----------
    h = ws_hier
    h.write(0, 0, 'BLS Table 1.2 hierarchy, with summary rows reconciled to their detailed occupations', F['title'])
    h.write(1, 0, 'Summary rows are BLS aggregates and are never added into the totals. "Sum of detailed rows" adds the line items beneath each summary row; differences are BLS rounding or occupations BLS does not publish separately.', F['sub'])
    hh = ['SOC code', 'Title', 'Occupation type', 'Level', 'Employment 2025 (thousands)', 'Employment 2035 (thousands)',
          'Sum of detailed rows, 2025', 'Difference (summary − detailed)']
    h.write_row(3, 0, hh, F['hdr'])
    h.set_row(3, 30)
    widths_h = [10, 70, 12, 7, 13, 13, 13, 14]
    for c, w in enumerate(widths_h):
        h.set_column(c, c, w)
    recon = []
    for i, rrow in enumerate(bls):
        rr = 4 + i
        indent = '    ' * rrow['level']
        h.write_string(rr, 0, rrow['soc_code'], F['txt'])
        h.write_string(rr, 1, indent + rrow['title'], F['txt'])
        h.write_string(rr, 2, rrow['occupation_type'], F['txt'])
        h.write_number(rr, 3, rrow['level'], F['int'])
        h.write_number(rr, 4, rrow['emp_2025_thousands'], F['k'])
        h.write_number(rr, 5, rrow['emp_2035_thousands'], F['k'])
        if rrow['occupation_type'] == 'Summary':
            ssum = 0.0
            for later in bls[i + 1:]:
                if later['level'] <= rrow['level']:
                    break
                if later['occupation_type'] == 'Line item':
                    ssum += later['emp_2025_thousands']
            diff = round(rrow['emp_2025_thousands'] - ssum, 1)
            h.write_number(rr, 6, round(ssum, 1), F['k'])
            h.write_number(rr, 7, diff, F['k'])
            recon.append((rrow['soc_code'], rrow['title'], rrow['level'], rrow['emp_2025_thousands'], round(ssum, 1), diff))
    h.freeze_panes(4, 2)
    h.autofilter(3, 0, 3 + len(bls), len(hh) - 1)

    # ---------- Sources ----------
    w = ws_src
    w.set_column(0, 0, 34)
    w.set_column(1, 1, 120)
    big = [x for x in recon if abs(x[5]) >= 1.0]
    lines = [
        ('Prepared by', f'{MODEL} (model ID {MODEL_ID}), Anthropic, working in Claude Code (desktop app)'),
        ('Generated', stamp_human),
        ('Employment baseline', 'U.S. Bureau of Labor Statistics, Employment Projections, Table 1.2: Occupational projections, 2025–35, and worker characteristics, 2025 (National Employment Matrix; employment in thousands; includes self-employed).'),
        ('Table page', SRC_PAGE),
        ('Workbook downloaded', f'{SRC_XLSX} (downloaded through the browser on 2026-11-16; the copy is source_data/BLS_EP_occupation.xlsx, SHA-256 {sha})'),
        ('Rows used', f'{n} detailed ("Line item") occupations, {total_emp:,.1f} thousand jobs. The BLS all-occupation total is 170,280.8 thousand; the {170280.8 - total_emp:,.1f} thousand gap is reconciled on the BLS hierarchy sheet.'),
        ('Reconciliation', '; '.join(f'{c} {t}: summary {e:,.1f} vs detailed {s_:,.1f} (difference {d:,.1f})' for c, t, lv, e, s_, d in big) + '. Every other summary row matches its detailed rows to within ±0.4 thousand (rounding). So the gap sits inside Top executives, whose only detailed rows here are chief executives and general and operations managers; the SOC\'s third occupation in that group, Legislators (11-1031), is not published as a line item, and is most likely the missing 13.5 thousand. Scored as essential as elected representatives would be (roughly 80–95%), it would add about 11–13 thousand essential jobs, under 0.01 percentage points of the total.'),
        ('Definitions', 'C:/PaidToBeHuman/reference/background/Post-Labor Economics - Derived vs Essential Demand v01.docx (David Shapiro, 4 Oct 2026): Essential Demand, the substitution test, and the four dimensions (presence, provenance, accountability, affinity).'),
        ('Course framing', 'C:/PaidToBeHuman/course/Course plan.md: "if you cannot swap any machine out, no matter the quality, and have the same product or service, then that is essential demand."'),
        ('Derived demand', 'Alfred Marshall, Principles of Economics (1890), Book V, ch. 6: demand for a factor of production is derived from demand for what it helps produce.'),
        ('Independence', 'The prior audit\'s aggregate results appear in the white paper text and were seen before scoring; its row-level coefficients were not opened or used.'),
        ('Estimates source files', 'Claude_Opus_5.5/estimates/*.txt (one line per occupation), validated by scripts/check_estimates.py; this workbook is built by scripts/build_workbook.py.'),
    ]
    w.write(0, 0, 'Sources and provenance', F['title'])
    for i, (k, v) in enumerate(lines):
        w.write(2 + i, 0, k, F['txt_b'])
        w.write(2 + i, 1, v, F['wrap'])

    wb.close()

    # ---------- flat CSV for cross-model comparison ----------
    csv_path = os.path.join(base, f'Essential_vs_Derived_Demand_by_Occupation_{MODEL.replace(" ", "_")}_{stamp_file}.csv')
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        wr = csv.writer(f)
        wr.writerow(['soc_code', 'occupation', 'major_group_code', 'major_group', 'employment_2025_thousands',
                     'essential_low_pct', 'essential_median_pct', 'essential_high_pct', 'derived_median_pct',
                     'essential_jobs_median_thousands', 'tags', 'rationale', 'model', 'generated'])
        for x in items:
            e = est[x['soc_code']]
            wr.writerow([x['soc_code'], x['title'], x['major_group_code'], major_title[x['major_group_code']],
                         x['emp_2025_thousands'], e['low'], e['median'], e['high'], round(100 - e['median'], 4),
                         round(x['emp_2025_thousands'] * e['median'] / 100, 3), '; '.join(e['tags']) or 'None',
                         e['rationale'], f'{MODEL} ({MODEL_ID})', stamp_human])
    print(out_path)
    print(csv_path)
    print('reconciliation rows with |diff| >= 1.0:', big)


if __name__ == '__main__':
    main(sys.argv[1])
