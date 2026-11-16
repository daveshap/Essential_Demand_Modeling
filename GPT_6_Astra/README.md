# GPT-6 Astra essential vs derived demand by BLS occupation

GPT-6 Astra performed this assessment in Codex. The original run timestamp is **2026-11-16 08:34:31 EST** (`2026-11-16T13:34:31.236Z`). The package contains estimates for all 831 detailed occupations in BLS Employment Projections Table 1.2, with low, median and high essential-demand shares, complementary derived-demand shares, and an explanation for every occupation.

The user's definition governs: essential demand is **“when the human is the entire point” OR nonfungible OR people are willing to pay a premium for the human OR simply “not derived demand.”** The four categories in *PaidToBeHuman* are optional communication aids, not required tests.

## Result

The 831 occupations cover **170,267,200 jobs in 2025**. Applying the estimates to those employment weights gives:

| Measure | Low | Median | High |
|---|---:|---:|---:|
| Essential demand share | 2.50973% | 7.42008% | 16.19178% |
| Essential job-weighted equivalents | 4,273,247 | 12,633,969 | 27,569,286 |
| Derived demand share | 83.80822% | 92.57992% | 97.49027% |

There are **333 occupations with a 0% median essential estimate**, including **158 with 0% at all three bounds**. These are judgment scenarios, not measured BLS demand shares, full-time equivalents or forecasts of jobs surviving automation. Derived low complements essential high, and derived high complements essential low. The extra decimal places facilitate comparisons between runs; they do not imply corresponding accuracy in the judgments.

## Files

| File or folder | Contents |
|---|---|
| `Essential_vs_Derived_Demand_by_Occupation_GPT_6_Astra_2026-11-16_0834.xlsx` | Main workbook: Summary, Occupations, BLS source, and Method and sources. Editable essential estimates feed formulas for derived shares, weighted equivalents and totals. |
| `Essential_vs_Derived_Demand_by_Occupation_GPT_6_Astra_2026-11-16_0834.csv` | Flat occupation table for comparing runs. It is a snapshot, without spreadsheet formulas. |
| `METHOD.md` | Method and provenance, documented after the analysis. The scoring brief used before scoring is preserved separately. |
| `estimates/` | Four JSON scoring partitions, `combined.json`, the original `scoring_brief.md`, and review notes. |
| `scripts/` | Portable workbook builder, `rebuild.ps1` launcher and `validate.py` checks. |
| `source_data/` | Original supplied BLS workbook, extracted Table 1.2 CSV and JSON, official retrievals and metadata, and framework references. |
| `verification/` | Calculation checks, saved-file checks and visual previews. |
| `archive/` | Original session work and superseded exports, including the incorrectly attributed originals. These are historical records, not the current deliverable. |
| `revision_log.csv` | Corrections and packaging changes, including the attribution correction. |
| `manifest_sha256.csv` | File hashes for checking package integrity. |

## Rebuild

Run from this folder:

```powershell
powershell -NoProfile -File scripts/rebuild.ps1
```

The rebuild uses the saved inputs and writes to `rebuild/`. It requires the bundled Codex runtime and `@oai/artifact-tool`, and creates a local `node_modules` junction to that runtime. It does not download data, rescore occupations or overwrite archived deliverables. The root workbook remains the dated deliverable. Review `scripts/rebuild.ps1` for runtime discovery and validation details.

To validate the package separately, run `python scripts/validate.py .` with a Python interpreter, or substitute the bundled Python executable. The validator uses only Python's standard library.

## Provenance and independence

The employment baseline came from the user's local `occupation.xlsx`, worksheet Table 1.2. All 831 detailed codes, titles and 2025 employment values were checked against a freshly retrieved official BLS table. Its 281 summary rows are retained for reference and excluded from weighting. The published all-occupation total exceeds the sum of detailed rows by 13,600 jobs; the estimates use the covered detailed-row denominator.

The assessment used original User passages in *PaidToBeHuman*, distinguished from Assistant passages, together with current edited lessons and supporting framework material. GPT-6 Astra and three same-model subagents scored disjoint occupation sets, followed by review. No prior aggregate or another model's estimates were used as numerical targets.

The initial export incorrectly attributed the run to **ChatGPT 6 Pro**. The user corrected that attribution, and the current package identifies **GPT-6 Astra**. The original run timestamp and numerical estimates are retained. Superseded files are archived so the correction is visible rather than silently rewriting the record.

Claude Opus 5.5's README and METHOD were first inspected **after this run's scores and final workbook were locked**, at the user's explicit request to follow its naming and documentation layout. No Claude row-level estimates were read. The current METHOD records this run's actual approach; it does not adopt Claude's scoring rules or claim to have preceded scoring.
