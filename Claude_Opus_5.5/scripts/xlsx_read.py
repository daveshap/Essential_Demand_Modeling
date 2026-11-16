"""Minimal stdlib reader for an .xlsx: list sheets, dump rows of one sheet."""
import sys, zipfile, re, xml.etree.ElementTree as ET
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def load(path):
    z = zipfile.ZipFile(path)
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rmap = {r.get('Id'): r.get('Target') for r in rels}
    sheets = []
    for s in wb.find('m:sheets', NS):
        rid = s.get('{%s}id' % NS['r'])
        tgt = rmap[rid]; tgt = tgt.lstrip('/'); tgt = tgt if tgt.startswith('xl/') else 'xl/' + tgt
        sheets.append((s.get('name'), tgt))
    ss = []
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', NS):
            ss.append(''.join(t.text or '' for t in si.iter('{%s}t' % NS['m'])))
    return z, sheets, ss
def col_idx(ref):
    letters = re.match(r'[A-Z]+', ref).group(0); n = 0
    for ch in letters: n = n * 26 + ord(ch) - 64
    return n - 1
def rows(z, target, ss):
    root = ET.fromstring(z.read(target))
    for row in root.iter('{%s}row' % NS['m']):
        vals = {}
        for c in row.findall('m:c', NS):
            t = c.get('t'); v = c.find('m:v', NS); isel = c.find('m:is', NS)
            if t == 's' and v is not None: val = ss[int(v.text)]
            elif t == 'inlineStr' and isel is not None: val = ''.join(x.text or '' for x in isel.iter('{%s}t' % NS['m']))
            elif v is not None: val = v.text
            else: val = None
            vals[col_idx(c.get('r'))] = val
        if vals:
            n = max(vals) + 1
            yield int(row.get('r')), [vals.get(i) for i in range(n)]
if __name__ == '__main__':
    z, sheets, ss = load(sys.argv[1])
    if len(sys.argv) == 2:
        for i, (n, t) in enumerate(sheets): print(i, n, t)
    else:
        idx = int(sys.argv[2]); lo = int(sys.argv[3]); hi = int(sys.argv[4])
        for r, vals in rows(z, sheets[idx][1], ss):
            if lo <= r <= hi: print(r, vals)
