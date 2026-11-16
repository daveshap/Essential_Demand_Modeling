#!/usr/bin/env python3
"""Validate this portable run package using only the Python standard library.

Usage: python scripts/validate.py [package_root]
Writes verification/package_validation.json and exits nonzero on failure.
This checks saved data, formulas and cached results; it does not run Excel.
"""

import argparse
import csv
import hashlib
import json
import math
import posixpath
import re
import sys
import traceback
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET


STEM = 'Essential_vs_Derived_Demand_by_Occupation_GPT_6_Astra_2026-11-16_0834'
EXPECTED_ROWS = 831
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'m': MAIN}
CODE = re.compile(r'\d{2}-\d{4}')
MODEL = re.compile(r'GPT[\s_-]*6[\s_-]*Astra', re.I)
OLD_MODEL = re.compile(r'ChatGPT[\s_-]*6[\s_-]*Pro', re.I)
ERRORS = {'#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#N/A', '#NUM!', '#NULL!', '#SPILL!', '#CALC!'}
CSV_FIELDS = [
    'soc_code', 'occupation', 'major_group', 'employment_2025',
    'essential_low_pct', 'essential_median_pct', 'essential_high_pct',
    'derived_low_pct', 'derived_median_pct', 'derived_high_pct',
    'essential_equivalents_low', 'essential_equivalents_median', 'essential_equivalents_high',
    'explanation', 'source_workbook_row', 'bls_description_url',
]


def sha256(path):
    digest = hashlib.sha256()
    with path.open('rb') as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def load_json(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def numeric(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


class Xlsx:
    """Minimal read-only OOXML reader, preserving formulas and saved caches."""

    def __init__(self, path):
        self.path = path
        self.sheets = {}
        self.sheet_xml = {}
        self.tables = {}
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            strings = []
            if 'xl/sharedStrings.xml' in names:
                for item in ET.fromstring(archive.read('xl/sharedStrings.xml')).findall('m:si', NS):
                    strings.append(''.join(node.text or '' for node in item.iter(f'{{{MAIN}}}t')))
            rels = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
            targets = {r.attrib['Id']: r.attrib['Target'] for r in rels}
            book = ET.fromstring(archive.read('xl/workbook.xml'))
            for sheet in book.findall('m:sheets/m:sheet', NS):
                name = sheet.attrib['name']
                target = targets[sheet.attrib[f'{{{DOC_REL}}}id']]
                member = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
                xml = ET.fromstring(archive.read(member))
                self.sheet_xml[name] = xml
                cells = {}
                for cell in xml.findall('m:sheetData/m:row/m:c', NS):
                    kind = cell.attrib.get('t', 'n')
                    value_node = cell.find('m:v', NS)
                    raw = value_node.text if value_node is not None else None
                    formula_node = cell.find('m:f', NS)
                    formula = ('=' + (formula_node.text or '')) if formula_node is not None else None
                    if kind == 's' and raw is not None:
                        value = strings[int(raw)]
                    elif kind == 'inlineStr':
                        value = ''.join(n.text or '' for n in cell.findall('.//m:t', NS))
                    elif kind == 'b' and raw is not None:
                        value = raw == '1'
                    elif kind in ('str', 'e', 'd'):
                        value = raw
                    elif raw is not None:
                        value = float(raw)
                        if value.is_integer():
                            value = int(value)
                    else:
                        value = None
                    cells[cell.attrib['r']] = {'value': value, 'formula': formula, 'type': kind}
                self.sheets[name] = cells
            for member in names:
                if member.startswith('xl/tables/') and member.endswith('.xml'):
                    table = ET.fromstring(archive.read(member))
                    self.tables[table.attrib['name']] = {
                        'ref': table.attrib['ref'],
                        'filter': table.find('m:autoFilter', NS) is not None,
                    }

    def cell(self, sheet, address):
        return self.sheets[sheet].get(address, {'value': None, 'formula': None, 'type': None})

    def value(self, sheet, address):
        return self.cell(sheet, address)['value']

    def formula(self, sheet, address):
        return self.cell(sheet, address)['formula']


class Validation:
    def __init__(self, root):
        self.root = root
        self.failures = []
        self.checks = 0
        self.report = {
            'package_root': str(root),
            'validated_at_utc': datetime.now(timezone.utc).isoformat(),
            'validator': 'scripts/validate.py (Python standard library)',
            'scope': 'Source identity; exact record coverage; individual scores; CSV; saved workbook data, formulas and cached results; weighted summaries; attribution.',
            'limitation': 'Saved formula text and cached values are checked independently. This script does not recalculate in Excel or prove future native-application behavior.',
        }

    def check(self, condition, message):
        self.checks += 1
        if not condition:
            self.failures.append(message)

    def equal(self, actual, expected, message):
        self.check(actual == expected, f'{message}: expected {expected!r}, got {actual!r}')

    def close(self, actual, expected, message, tolerance=1e-7):
        self.check(numeric(actual) and abs(actual - expected) <= tolerance,
                   f'{message}: expected {expected!r}, got {actual!r}')

    def run(self):
        root = self.root
        estimates = load_json(root / 'estimates/combined.json')
        source = load_json(root / 'source_data/supplied_occupations.json')
        raw = load_json(root / 'source_data/supplied_source_rows.json')
        metadata = load_json(root / 'source_data/supplied_metadata.json')
        self.equal(len(estimates), EXPECTED_ROWS, 'Estimate record count')
        self.equal(len(source), EXPECTED_ROWS, 'Source detail count')
        score_codes = [row['code'] for row in estimates]
        source_codes = [row['code'] for row in source]
        self.equal(len(set(score_codes)), len(score_codes), 'Estimate codes are unique')
        self.equal(len(set(source_codes)), len(source_codes), 'Source codes are unique')
        self.equal(score_codes, source_codes, 'Estimates preserve complete source code order')
        self.check(all(CODE.fullmatch(code) for code in score_codes), 'All estimate codes follow SOC/NEM identifier syntax')
        source_map = {row['code']: row for row in source}
        for row in estimates:
            code = row['code']
            reference = source_map.get(code)
            if reference is None:
                self.check(False, f'{code}: missing source row')
                continue
            for key in ['title', 'major_code', 'major_title', 'employment_2025', 'employment_thousands_2025', 'source_row', 'ooh_url']:
                self.equal(row.get(key), reference.get(key), f'{code} source field {key}')
            bounds = [row.get(key) for key in ('low', 'median', 'high')]
            valid = all(numeric(value) for value in bounds)
            self.check(valid, f'{code}: all three estimates must be finite numeric values')
            if valid:
                self.check(0 <= bounds[0] <= bounds[1] <= bounds[2] <= 100, f'{code}: expected 0 <= low <= median <= high <= 100')
            self.check(isinstance(row.get('explanation'), str) and bool(row['explanation'].strip()), f'{code}: nonempty explanation required')
            self.check(numeric(row.get('employment_2025')) and row['employment_2025'] > 0, f'{code}: positive numeric source employment required')
        if any(not all(numeric(row.get(k)) for k in ('low', 'median', 'high', 'employment_2025')) for row in estimates):
            raise ValueError('Cannot recompute results because score/employment values are invalid')

        jobs = sum(row['employment_2025'] for row in estimates)
        equivalent = {key: math.fsum(row['employment_2025'] * row[key] / 100 for row in estimates) for key in ('low', 'median', 'high')}
        shares = {key: equivalent[key] / jobs for key in equivalent}
        published = round(next(row for row in raw if row['code'] == '00-0000')['employment_thousands_2025'] * 1000)
        zero_median = sum(row['median'] == 0 for row in estimates)
        zero_all = sum(row['low'] == row['median'] == row['high'] == 0 for row in estimates)
        self.equal(jobs, metadata['detailed_employment'], 'Recomputed covered jobs versus source metadata')
        self.equal(published, metadata['published_total'], 'Published total versus metadata')
        self.equal(len(estimates), metadata['detailed_count'], 'Detailed count versus metadata')
        self.report['recomputed_summary'] = {
            'occupation_count': len(estimates), 'covered_jobs': jobs, 'published_jobs': published,
            'unallocated_source_difference_jobs': published - jobs,
            'coverage_share': jobs / published,
            'essential_equivalents': equivalent, 'essential_shares': shares,
            'derived_equivalents': {key: jobs - equivalent[opposite] for key, opposite in [('low', 'high'), ('median', 'median'), ('high', 'low')]},
            'derived_shares': {key: 1 - shares[opposite] for key, opposite in [('low', 'high'), ('median', 'median'), ('high', 'low')]},
            'zero_median_occupations': zero_median, 'all_zero_occupations': zero_all,
        }

        # Hash the retained original input, then verify its actual rows against the extracted data.
        candidates = list((root / 'source_data').glob('*.xlsx'))
        matching = [path for path in candidates if sha256(path) == metadata['sha256']]
        self.check(bool(matching), 'A retained source XLSX must match supplied_metadata.json SHA-256')
        if not matching:
            raise ValueError('No source workbook matches recorded input SHA-256')
        original_path = sorted(matching)[0]
        original = Xlsx(original_path)
        source_sheet = metadata['source_sheet']
        actual_details = []
        for row in raw:
            ix = row['source_row']
            title = original.value(source_sheet, f'A{ix}')
            self.equal(title.strip() if isinstance(title, str) else title, row['title'], f'Source original row {ix} title')
            self.equal(original.value(source_sheet, f'B{ix}'), row['code'], f'Source original row {ix} code')
            self.equal(original.value(source_sheet, f'C{ix}'), row['type'], f'Source original row {ix} type')
            self.close(original.value(source_sheet, f'D{ix}'), row['employment_thousands_2025'], f'Source original row {ix} employment', 1e-9)
            if row['type'] == 'Line item':
                actual_details.append(row['code'])
        self.equal(actual_details, source_codes, 'Original source line-item codes versus extracted details')
        raw_by_code = {row['code']: row for row in raw}
        for row in source:
            reference = raw_by_code[row['code']]
            self.equal(reference['type'], 'Line item', f"{row['code']}: source is a detailed line item")
            self.equal(row['source_row'], reference['source_row'], f"{row['code']}: source workbook row mapping")
            self.equal(row['title'], reference['title'], f"{row['code']}: extracted title versus original")
            self.close(row['employment_thousands_2025'], reference['employment_thousands_2025'], f"{row['code']}: extracted employment versus original", 1e-9)
            self.equal(row['employment_2025'], round(reference['employment_thousands_2025'] * 1000), f"{row['code']}: employment conversion to jobs")
            link_value = original.value(source_sheet, f"Q{row['source_row']}") or ''
            link_match = re.search(r'href="([^"]+)"', str(link_value))
            self.equal(row.get('ooh_url') or '', link_match.group(1) if link_match else '', f"{row['code']}: original occupation-description link")
        self.report['source'] = {'file': str(original_path.relative_to(root)), 'sha256': sha256(original_path), 'source_rows_verified': len(raw), 'detail_rows_verified': len(source)}

        csv_path = root / f'{STEM}.csv'
        with csv_path.open(newline='', encoding='utf-8-sig') as handle:
            reader = csv.DictReader(handle)
            self.equal(reader.fieldnames, CSV_FIELDS, 'CSV header schema and order')
            csv_rows = list(reader)
        self.equal(len(csv_rows), len(estimates), 'CSV row count')
        self.equal([row.get('soc_code') for row in csv_rows], score_codes, 'CSV complete code order')
        for ix, (csv_row, row) in enumerate(zip(csv_rows, estimates), 2):
            expected = {
                'soc_code': row['code'], 'occupation': row['title'], 'major_group': row['major_title'],
                'employment_2025': row['employment_2025'],
                'essential_low_pct': row['low'], 'essential_median_pct': row['median'], 'essential_high_pct': row['high'],
                'derived_low_pct': 100 - row['high'], 'derived_median_pct': 100 - row['median'], 'derived_high_pct': 100 - row['low'],
                'essential_equivalents_low': row['employment_2025'] * row['low'] / 100,
                'essential_equivalents_median': row['employment_2025'] * row['median'] / 100,
                'essential_equivalents_high': row['employment_2025'] * row['high'] / 100,
                'explanation': row['explanation'], 'source_workbook_row': row['source_row'],
                'bls_description_url': row.get('ooh_url') or '',
            }
            for key, value in expected.items():
                if numeric(value):
                    try:
                        actual = float(csv_row[key])
                    except (KeyError, TypeError, ValueError):
                        actual = None
                    self.close(actual, value, f'CSV row {ix} field {key}')
                else:
                    self.equal(csv_row.get(key), value, f'CSV row {ix} field {key}')
        self.report['csv'] = {'file': csv_path.name, 'sha256': sha256(csv_path), 'rows_verified': len(csv_rows)}

        path = root / f'{STEM}.xlsx'
        workbook = Xlsx(path)
        self.equal(list(workbook.sheets), ['Summary', 'Occupations', 'BLS source', 'Method and sources'], 'Workbook worksheet order')
        end = 5 + len(estimates)
        source_index = {row['code']: ix for ix, row in enumerate(raw, 5)}
        workbook_codes = [cell['value'] for address, cell in workbook.sheets['Occupations'].items()
                          if re.fullmatch(r'A\d+', address) and isinstance(cell['value'], str) and CODE.fullmatch(cell['value'])]
        self.equal(workbook_codes, score_codes, 'Workbook complete occupation code order')
        for ix, row in enumerate(estimates, 6):
            values = {
                'A': row['code'], 'B': row['title'], 'C': row['employment_2025'],
                'D': row['low'] / 100, 'E': row['median'] / 100, 'F': row['high'] / 100,
                'G': 1 - row['high'] / 100, 'H': 1 - row['median'] / 100, 'I': 1 - row['low'] / 100,
                'J': row['explanation'], 'K': row['major_title'],
                'L': row['employment_2025'] * row['low'] / 100,
                'M': row['employment_2025'] * row['median'] / 100,
                'N': row['employment_2025'] * row['high'] / 100,
                'O': row['source_row'], 'P': row.get('ooh_url') or 'No occupation-specific link supplied',
            }
            for col, value in values.items():
                actual = workbook.value('Occupations', f'{col}{ix}')
                if numeric(value):
                    self.close(actual, value, f'Occupations!{col}{ix}', 1e-12 if col in 'DEFGHI' else 1e-7)
                else:
                    self.equal(actual, value, f'Occupations!{col}{ix}')
            formulas = {
                'C': f"=ROUND('BLS source'!D{source_index[row['code']]}*1000,0)",
                'G': f'=1-F{ix}', 'H': f'=1-E{ix}', 'I': f'=1-D{ix}',
                'L': f'=C{ix}*D{ix}', 'M': f'=C{ix}*E{ix}', 'N': f'=C{ix}*F{ix}',
            }
            for col, formula in formulas.items():
                self.equal(workbook.formula('Occupations', f'{col}{ix}'), formula, f'Occupations!{col}{ix} formula')

        # Every retained BLS source row is checked, including the total and overlapping summaries.
        for ix, row in enumerate(raw, 5):
            for col, value in zip('ABCDE', [row['code'], row['title'], row['type'], row['employment_thousands_2025'], row['source_row']]):
                actual = workbook.value('BLS source', f'{col}{ix}')
                if numeric(value):
                    self.close(actual, value, f'BLS source!{col}{ix}', 1e-9)
                else:
                    self.equal(actual, value, f'BLS source!{col}{ix}')

        for col, key, opposite in [('B', 'low', 'high'), ('C', 'median', 'median'), ('D', 'high', 'low')]:
            self.close(workbook.value('Summary', f'{col}8'), shares[key], f'Summary!{col}8', 1e-12)
            self.close(workbook.value('Summary', f'{col}9'), 1 - shares[opposite], f'Summary!{col}9', 1e-12)
            self.close(workbook.value('Summary', f'{col}10'), equivalent[key], f'Summary!{col}10')
            self.close(workbook.value('Summary', f'{col}11'), jobs - equivalent[opposite], f'Summary!{col}11')
        for address, value in {'B14': len(estimates), 'B15': jobs, 'B16': published, 'B17': published-jobs,
                               'B18': jobs/published, 'B47': zero_median, 'B48': zero_all}.items():
            self.close(workbook.value('Summary', address), value, f'Summary!{address}', 1e-12 if address == 'B18' else 1e-7)
        summary_formulas = {
            'B8': '=B10/$B$15', 'C8': '=C10/$B$15', 'D8': '=D10/$B$15',
            'B9': '=1-D8', 'C9': '=1-C8', 'D9': '=1-B8',
            'B10': f'=SUM(Occupations!L6:L{end})', 'C10': f'=SUM(Occupations!M6:M{end})', 'D10': f'=SUM(Occupations!N6:N{end})',
            'B11': '=$B$15-D10', 'C11': '=$B$15-C10', 'D11': '=$B$15-B10',
            'B14': f'=COUNTA(Occupations!A6:A{end})', 'B15': f'=SUM(Occupations!C6:C{end})',
            'B16': f"='BLS source'!D{source_index['00-0000']}*1000", 'B17': '=B16-B15', 'B18': '=B15/B16',
            'B47': f'=COUNTIF(Occupations!E6:E{end},0)', 'B48': f'=COUNTIF(Occupations!F6:F{end},0)',
        }
        for address, formula in summary_formulas.items():
            self.equal(workbook.formula('Summary', address), formula, f'Summary!{address} formula')
        group_names = list(dict.fromkeys(row['major_title'] for row in estimates))
        self.equal(len(group_names), 22, 'Major group count')
        groups = []
        for ix, title in enumerate(group_names, 23):
            members = [row for row in estimates if row['major_title'] == title]
            employment = sum(row['employment_2025'] for row in members)
            weighted = {key: math.fsum(row['employment_2025'] * row[key] / 100 for row in members) for key in ('low', 'median', 'high')}
            self.equal(workbook.value('Summary', f'A{ix}'), title, f'Summary!A{ix}')
            self.close(workbook.value('Summary', f'B{ix}'), employment, f'Summary!B{ix}')
            for col, key in [('C', 'low'), ('D', 'median'), ('E', 'high')]:
                self.close(workbook.value('Summary', f'{col}{ix}'), weighted[key] / employment, f'Summary!{col}{ix}', 1e-12)
            self.close(workbook.value('Summary', f'F{ix}'), weighted['median'], f'Summary!F{ix}')
            formulas = {
                'B': f'=SUMIF(Occupations!$K$6:$K${end},A{ix},Occupations!$C$6:$C${end})',
                'C': f'=SUMIF(Occupations!$K$6:$K${end},A{ix},Occupations!$L$6:$L${end})/B{ix}',
                'D': f'=F{ix}/B{ix}',
                'E': f'=SUMIF(Occupations!$K$6:$K${end},A{ix},Occupations!$N$6:$N${end})/B{ix}',
                'F': f'=SUMIF(Occupations!$K$6:$K${end},A{ix},Occupations!$M$6:$M${end})',
            }
            for col, formula in formulas.items():
                self.equal(workbook.formula('Summary', f'{col}{ix}'), formula, f'Summary!{col}{ix} formula')
            groups.append({'major_group': title, 'occupation_count': len(members), 'jobs': employment,
                           'essential_equivalents': weighted, 'essential_shares': {key: weighted[key] / employment for key in weighted}})
        self.report['recomputed_groups'] = groups

        formula_errors, empty_formula_caches, forbidden_labels = [], [], []
        formula_count = 0
        for name, cells in workbook.sheets.items():
            for address, cell in cells.items():
                value = cell['value']
                if cell['type'] == 'e' or (isinstance(value, str) and value in ERRORS):
                    formula_errors.append(f'{name}!{address}: {value}')
                if cell['formula'] is not None:
                    formula_count += 1
                    if value is None:
                        empty_formula_caches.append(f'{name}!{address}')
                if isinstance(value, str) and OLD_MODEL.search(value):
                    forbidden_labels.append(f'{name}!{address}: {value}')
        self.check(not formula_errors, f'No saved spreadsheet errors: {formula_errors[:20]}')
        self.check(not empty_formula_caches, f'All formula cells have saved values: {empty_formula_caches[:20]}')
        self.check(not forbidden_labels, f'No obsolete ChatGPT 6 Pro attribution: {forbidden_labels[:20]}')
        for sheet, address in [('Summary', 'A4'), ('Occupations', 'A3'), ('Method and sources', 'B5'), ('Method and sources', 'B6')]:
            value = workbook.value(sheet, address)
            self.check(isinstance(value, str) and bool(MODEL.search(value)), f'{sheet}!{address}: expected GPT-6 Astra model attribution')
        method_cells = workbook.sheets['Method and sources']
        hash_cells = [cell['value'] for cell in method_cells.values() if cell['value'] == metadata['sha256']]
        self.check(bool(hash_cells), 'Workbook records original source SHA-256')
        pane = workbook.sheet_xml['Occupations'].find('m:sheetViews/m:sheetView/m:pane', NS)
        self.check(pane is not None and pane.attrib.get('topLeftCell') == 'C6', 'Occupations freeze pane is C6')
        table = workbook.tables.get('OccupationEstimates')
        self.check(table is not None and table['ref'] == f'A5:P{end}' and table['filter'], 'Occupation table range and filter cover all rows')
        for sheet, xml in workbook.sheet_xml.items():
            view = xml.find('m:sheetViews/m:sheetView', NS)
            self.check(view is not None and view.attrib.get('showGridLines') in ('0', 'false'), f'{sheet}: gridlines are hidden')
        self.report['workbook'] = {
            'file': path.name, 'sha256': sha256(path), 'bytes': path.stat().st_size,
            'occupation_rows_verified': len(estimates), 'source_rows_verified': len(raw),
            'formula_cells_checked': formula_count, 'formula_errors': formula_errors,
            'empty_formula_caches': empty_formula_caches, 'obsolete_model_labels': forbidden_labels,
            'expected_model': 'GPT-6 Astra',
        }

    def finish(self):
        self.report.update(success=not self.failures, checks_performed=self.checks,
                           failure_count=len(self.failures), failures=self.failures)
        target = self.root / 'verification/package_validation.json'
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps(self.report, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
        print(json.dumps({
            'success': self.report['success'], 'checks_performed': self.checks,
            'failure_count': len(self.failures), 'report': str(target),
            'recomputed_summary': self.report.get('recomputed_summary'),
        }, indent=2))
        if self.failures:
            for message in self.failures[:20]:
                print('FAIL: ' + message, file=sys.stderr)
        return 0 if self.report['success'] else 1


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('package_root', nargs='?', type=Path,
                        default=Path(__file__).resolve().parent.parent)
    arguments = parser.parse_args()
    validation = Validation(arguments.package_root.resolve())
    try:
        validation.run()
    except Exception as error:
        validation.failures.append(f'{type(error).__name__}: {error}')
        validation.report['exception_traceback'] = traceback.format_exc()
    return validation.finish()


if __name__ == '__main__':
    sys.exit(main())
