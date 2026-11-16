# Claude Opus 5.5: essential vs derived demand by BLS occupation

Claude Opus 5.5 (model ID `claude-opus-5-5`, Anthropic) did this run in Claude Code on 16 Nov 2026. It is one of several independent runs of the same instructions (others by ChatGPT 6 Pro and Fable), meant to be compared afterward.

The task: go through every BLS occupation and estimate how much of it is Essential Demand (the human is the point: nonfungible, or worth a premium) and how much is derived demand, with low, median and high estimates and an explanation for each.

## Files

| File | What it is |
|---|---|
| `Essential_vs_Derived_Demand_by_Occupation_Claude_Opus_5.5_<date>_<time>.xlsx` | The spreadsheet. Summary, all 831 occupations, Method, BLS hierarchy and Sources. Totals are live formulas: edit any Low, Median or High cell and every total updates. |
| `Essential_vs_Derived_Demand_by_Occupation_Claude_Opus_5.5_<date>_<time>.csv` | The same estimates as one flat table, for comparing runs row by row. |
| `Essential_vs_Derived_Demand_by_Occupation_Claude_Opus_5.5_2026-11-16_0809_METHODOLOGY.md` | The full methodology for the CSV and workbook: what was done, in what order, with what inputs and why, the known biases, how to reproduce or re-run it, and file hashes. |
| `METHOD.md` | The scoring standard, written before any occupation was scored (one rule was added after the first pass). |
| `revision_log.csv` | Every estimate changed after the first pass (73 rows), with before and after values, the effect on the total and the reason. |
| `estimates/*.txt` | The judgments themselves, one line per occupation: code, low, median, high, tags and rationale, separated by vertical bars. |
| `source_data/BLS_EP_occupation.xlsx` | BLS's own workbook, downloaded 16 Nov 2026. Table 1.2 is the baseline. |
| `source_data/BLS_EP_Table_1.2_2025-35.csv` | Table 1.2 extracted from it: all 1,112 rows, summary and detailed. |
| `scripts/` | Extraction, validation and build scripts. |

## Rebuild

```
py -3.13 -I scripts/extract_table_1_2.py source_data/BLS_EP_occupation.xlsx source_data/BLS_EP_Table_1.2_2025-35.csv
py -3.13 -I scripts/check_estimates.py .
py -3.14 -I scripts/build_workbook.py .
```

Run them from this folder. The build needs `xlsxwriter`, which is installed for Python 3.14 on this PC. Each build writes a new timestamped workbook and CSV.

## Result (median, with low and high)

Of 170.27 million jobs in 2025, about **7.4%** (12.6 million job-weighted equivalents) is Essential Demand, with a range of 3.3% to 13.8%. The remaining 92.6% is derived demand. 395 of the 831 occupations, holding 31.5% of all jobs, score 0%. The figures are judgments under the substitution test, not measurements or forecasts; see the methodology document.
