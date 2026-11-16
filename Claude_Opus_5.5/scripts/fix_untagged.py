"""One-off consistency fix (16 Nov 2026): a nonzero median must name an essential element.
Rows whose rationale names none get median 0 and a high capped at 3 (a general allowance).
Two rows that do name a sliver are tagged instead."""
import glob, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from load_estimates import load_bls, load_estimates
base = sys.argv[1]
bls = load_bls(os.path.join(base, 'source_data', 'BLS_EP_Table_1.2_2025-35.csv'))
est, _ = load_estimates(os.path.join(base, 'estimates'), bls)
keep = {
    '11-3051': ('0', '1', '4', 'Accountability', 'Factory output is the product; production management is derived. A sliver for the manager who answers personally for plant safety decisions.'),
    '39-3092': ('0', '1', '4', 'Affinity', 'Costume handling is instrumental. A sliver for performers who want a trusted dresser with them backstage.'),
}
targets = {c for c, e in est.items() if e['median'] > 0 and not e['tags']}
changed = 0
for path in sorted(glob.glob(os.path.join(base, 'estimates', '*.txt'))):
    lines = open(path, encoding='utf-8').read().split('\n')
    out = []
    for line in lines:
        parts = line.split('|')
        if len(parts) == 6 and parts[0] in targets:
            code = parts[0]
            if code in keep:
                lo, md, hi, tags, why = keep[code]
                parts = [code, lo, md, hi, tags, why]
            else:
                hi = min(float(parts[3]), 3.0)
                parts[1], parts[2], parts[3] = '0', '0', ('%g' % hi)
            changed += 1
            line = '|'.join(parts)
        out.append(line)
    open(path, 'w', encoding='utf-8', newline='\n').write('\n'.join(out))
print('rows changed:', changed, 'of', len(targets))
