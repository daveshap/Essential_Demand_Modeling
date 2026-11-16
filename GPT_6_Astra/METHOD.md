# Method for essential vs derived demand by BLS occupation

Prepared by **GPT-6 Astra in Codex** for the run timestamped **2026-11-16 08:34:31 EST** (`2026-11-16T13:34:31.236Z`). This document was synthesized after the analysis to describe the method actually used. `estimates/scoring_brief.md` preserves the brief written before occupation scoring.

## Question and definition

The question is how much of the labor demand represented by today's occupations is essential rather than derived demand.

The current user's definition is preserved verbatim in substance: **essential demand is “when the human is the entire point” OR nonfungible OR people are willing to pay a premium for the human OR simply “not derived demand.”** These are alternative formulations. They are not cumulative conditions, and no additional checklist is imposed. The four labels previously developed in *PaidToBeHuman* are optional aids to explanation.

Derived demand concerns human labor wanted as a means to an outcome. The original *PaidToBeHuman* thought experiment assumes machines can outperform humans on functional capabilities and then asks where demand for humans persists. This makes current automation difficulty a different question from why the work is demanded.

The originating User passages were read in `C:\PaidToBeHuman\source\Paid_to_Be_Human_Conversation_Export.md`, especially lines 184, 223, 265, 298 and 337. They cover firefighters and warfighters as instrumental labor, fully saturated machine supply, nonfungible people, actual human presence, human provenance, and a contractor who personally stands behind a robot crew's work. User passages were distinguished from Assistant replies. Current edited Word Lessons 02, 04, 06, 07 and 08, the finished derived-versus-essential-demand transcript, and relevant white-paper sections supplied further context. The white paper's broader formulations were not substituted for the current user's wording.

## Employment baseline

The baseline is BLS Employment Projections Table 1.2, *Occupational projections, 2025–35, and worker characteristics, 2025*, from the National Employment Matrix. BLS released the 2025–35 projections on August 27, 2026. Only the base-year 2025 employment column enters the weighting; projected 2035 employment is not used.

The user supplied `C:\Users\daves\Downloads\occupation.xlsx`. Its preserved copy is `source_data/BLS_EP_occupation.xlsx`. Its SHA-256 is:

```text
08ee182f992d972593e64559ce467bad75967d2228f11ad5e2f86d4ec0ee6c00
```

All 831 detailed codes, titles and 2025 employment values matched the official BLS HTML table freshly retrieved on November 16, 2026. This was a verification of the locally supplied workbook; it is not a claim that the assistant separately downloaded that workbook successfully.

| Baseline item | Value |
|---|---:|
| Detailed “Line item” occupations scored | 831 |
| Major occupation groups | 22 |
| Summary rows retained but excluded from weighting | 281 |
| Total source rows retained | 1,112 |
| Covered detailed employment | 170,267,200 jobs |
| Published all-occupation employment | 170,280,800 jobs |
| Published total less detailed sum | 13,600 jobs |

BLS notes that component totals can differ because of rounding and excluded occupational data. The 13,600-job difference is not assigned to an invented occupation or silently distributed over the scored rows. The headline denominator is covered detailed employment. For example, legislators are not published as a separate Line item in this table; no employment weight was invented for them.

The matrix includes civilian wage-and-salary employment and self-employment, with agricultural and private-household employment included in its scope. It is not a count of unique employed people or full-time equivalents, and it is not a real-time November 2026 jobs count. Employment reported in thousands to one decimal place is multiplied by 1,000, retaining the source's 100-job precision. Display rounding of computed equivalents does not alter the underlying calculations.

## What each estimate represents

Each coefficient estimates the proportion of an occupation's present labor demand attributable to essential demand under the user's definition. An occupation can mix both kinds of demand. The estimate is not a measured fraction of customers, a willingness-to-pay premium, a revenue share, an automation probability or an observed share of task minutes.

Every occupation received an explicit low, median and high estimate and an occupation-specific explanation. The explanations distinguish the functional service or output from any reason to want the actual human. No fixed category allocation, automatic positive floor or target overall percentage was used. Zero is valid at one, two or all three bounds. Broad residual occupations receive wider ranges where their mix leaves material uncertainty.

The analysis considered occupational duties and purchasing contexts individually. It did not calculate coefficients by multiplying measured customer preferences by measured human-task time: those data were not available. The coefficients are model judgments, and the workbook makes them editable for comparison and challenge.

Examples of boundaries used in the reasoning are:

- A contractor, judge or executive can be personally accountable without making every supporting worker essential. A nominal signature with no meaningful authority is not automatically the accountable human role described in the source work.
- Care and teaching can include demand for an actual person or a particular continuing relationship alongside instrumental treatment, learning, supervision and administration.
- Human performers, athletes and named creators can be the point of the offering. Their ticketing, equipment, logistics and other support labor must be considered separately.
- Human provenance can matter to a buyer without requiring fresh labor for every reuse of an existing work. Customers' or family members' own human participation is not automatically paid demand for an employee.
- Work can be difficult, important, heroic or currently resistant to automation and still be demanded for its outcome.

These distinctions explain the estimates; they do not replace the user's alternative definitions with new mandatory conditions.

## Low, median and high

**Median** means the subjective central or best estimate. It is not an empirical sample median, a fitted 50th percentile or the arithmetic midpoint of the endpoints. **Low** and **high** are plausible judgment bounds, not statistical confidence limits.

Every occupation satisfies `0 ≤ low ≤ median ≤ high ≤ 100`. Derived-demand bounds reverse the essential endpoints:

```text
Derived low    = 100% − essential high
Derived median = 100% − essential median
Derived high   = 100% − essential low
```

For each scenario, occupation employment is multiplied by that occupation's essential coefficient. Those products are summed and divided by the sum of covered jobs:

```text
Essential job-weighted equivalents = Σ(employment × essential share)
Employment-weighted essential share = essential equivalents ÷ covered jobs
```

The products are **job-weighted scenario equivalents**, not identifiable surviving jobs or full-time equivalents. Equal weighting of occupation percentages would answer a different question and is not used. Summing all lows or all highs produces aligned scenarios, not a probability interval for the economy.

The current occupational mix is held fixed. Changes in prices, service volume, entry, technology, hours per worker, preferences and institutions are not modeled. The displayed precision supports reproducibility rather than empirical precision in the underlying judgments.

## Scoring and review sequence

GPT-6 Astra and three same-model subagents scored disjoint occupation sets. The shared scoring brief preceded those estimates. Existing numerical results in the framework material were encountered as source context but were not used to calibrate the new rows. No prior aggregate result or other model's estimates served as a target.

Review checked occupation scope, the relationship between core human offerings and supporting work, high-impact estimates, and comparables across partitions. Three explanatory revisions were recorded without changing numerical estimates: postsecondary teaching assistants were described as instructional support rather than research staff; recreation and fitness studies teachers were distinguished from direct fitness trainers; and the assumption behind the higher childcare-worker share relative to preschool teachers was made explicit. `estimates/review_notes.json` preserves those revisions. A separate pre-export review changed commercial divers (49-9092) from 0/0/1% to 0/0/0%, removing a speculative demonstration or heritage allowance. `revision_log.csv` records that numerical change.

Claude Opus 5.5's README and METHOD were first opened after this run's scores and final workbook were locked, following the user's explicit request to emulate the other run's naming and documentation organization. No Claude row-level estimates were read. Its documentation was not a scoring input for this run.

## Workbook construction and verification

The workbook was authored with the bundled `@oai/artifact-tool`. It contains Summary, Occupations, BLS source, and Method and sources. Essential percentages are editable numeric inputs. Employment is linked to the source sheet; derived percentages, equivalents, group summaries and headline totals are formulas.

The initial build was recalculated and compared with independently calculated totals. A reversible test increased the first occupation's median by one percentage point. The aggregate median equivalents rose by exactly 2,916, matching that occupation's 291,600 jobs multiplied by 0.01. The original input was restored before export.

The saved XLSX was reopened for read-only verification of all 831 occupation rows, the 1,112 source rows, stored formulas and calculated values, explanations, derived complements, group totals and headline results. Checks also covered formula errors, filter ranges, frozen panes, gridlines and timestamp formatting. This reopen check used `openpyxl` as a reader, not as the workbook author. Representative ranges from every sheet and long explanations were rendered and visually reviewed. These checks demonstrate the saved workbook and the authoring engine's recalculation behavior; they do not claim an independent interactive Excel recalculation test.

Verification outputs and previews are retained in `verification/`. The portable rebuild uses saved local inputs and the bundled runtime without network access, writes to `rebuild/`, and does not regenerate the judgment coefficients. `manifest_sha256.csv` records package file hashes; `revision_log.csv` records subsequent packaging and attribution changes.

## Attribution correction

The initial export used **ChatGPT 6 Pro**, following the original requested label. That was an incorrect attribution of the model performing this run. After the user identified the model as **GPT-6 Astra**, the visible attribution was corrected. The original timestamp, employment data, estimates and formulas were retained. The attribution-only correction was compared with the original export and verified to leave all other values, formulas and cell styles unchanged. Original exports and session work remain under `archive/` as superseded records; the workbook at the package root is the current deliverable.

## Primary data references

- [BLS Table 1.2 occupational projections and characteristics](https://www.bls.gov/emp/tables/occupational-projections-and-characteristics.htm)
- [BLS occupation workbook](https://www.bls.gov/emp/ind-occ-matrix/occupation.xlsx)
- [BLS 2025–35 projections release](https://www.bls.gov/news.release/ecopro.htm)
- [BLS employment projections definitions](https://www.bls.gov/emp/documentation/definitions.htm)
- [BLS projections methodology](https://www.bls.gov/emp/documentation/projections-methods.htm)
- [BLS 2018 SOC definitions](https://www.bls.gov/soc/2018/soc_2018_definitions.pdf)

Occupation-specific OOH links are preserved from the supplied workbook. BLS supplies occupation definitions and employment data; it does not supply this run's essential-demand percentages.
