# Methodology: Essential vs derived demand by BLS occupation

Claude Opus 5.5 run, 16 November 2026

| | |
|---|---|
| Describes | `Essential_vs_Derived_Demand_by_Occupation_Claude_Opus_5.5_2026-11-16_0809.csv` and the `.xlsx` of the same name |
| Prepared by | Claude Opus 5.5 (model ID `claude-opus-5-5`, Anthropic), working in Claude Code in the Claude desktop app |
| Machine | Windows 10 Home 10.0.19045. Python 3.13.0 (standard library only) and Python 3.14.7 with xlsxwriter 3.2.9. Microsoft Excel (Office 16), used read-only for checks. |
| When | 16 Nov 2026. Scoring 07:52–08:03, final build 08:09, this document from 08:26, updated at 10:24 and 10:28 for corrections to `METHOD.md` (Section 8.6). All times are EST (UTC−05:00). |
| Rater | One run of one model. No person or other model contributed estimates. |
| Companion files | `METHOD.md` (the scoring standard, written before any row was scored apart from one rule added after the first pass), `revision_log.csv` (every change made after the first pass), `estimates/`, `scripts/`, `source_data/` |

This document records what I did, in what order, with what inputs, and why. It is written for reproducibility. Anyone should be able to rebuild the outputs exactly from the files in this folder, re-run the judgments with another rater or model under the same procedure, and see where this run's biases and shortcuts lie. Times and file details come from the session's tool-call log and the file system, and every number here was recomputed from the files.

## Contents

1. [Summary](#1-summary)
2. [The instructions, and how I read them](#2-the-instructions-and-how-i-read-them)
3. [Inputs](#3-inputs)
4. [Getting the BLS data](#4-getting-the-bls-data)
5. [Extracting Table 1.2](#5-extracting-table-12)
6. [What each percentage means](#6-what-each-percentage-means)
7. [How the rows were scored](#7-how-the-rows-were-scored)
8. [Checks and changes after the first pass](#8-checks-and-changes-after-the-first-pass)
9. [Building the spreadsheet and CSV](#9-building-the-spreadsheet-and-csv)
10. [Results](#10-results)
11. [Biases and limitations](#11-biases-and-limitations)
12. [How to reproduce](#12-how-to-reproduce)
13. [File formats](#13-file-formats)
14. [File inventory and SHA-256 hashes](#14-file-inventory-and-sha-256-hashes)
15. [Timeline](#15-timeline)

## 1. Summary

- **Baseline.** BLS Employment Projections Table 1.2 (the National Employment Matrix, 2025–35). It has 831 detailed occupations with 170,267.2 thousand jobs in 2025, self-employment included.
- **Quantity.** For each occupation: the share of its current labor that buyers would still pay a human to do at today's prices, if a disclosed machine delivered the same result at equal or lower cost. This is the substitution test from *Paid to Be Human*. The rest of the occupation is derived demand.
- **Method.**
  - There was one rater (this model), judging from general knowledge. I used no surveys, task-time data, O*NET, web searches or other runs.
  - I estimated each row as roughly (share of buyers who want the human) × (share of the worker's time that is the human part), adjusted by judgment.
  - Each row has a low, median and high value, up to four tags, and a written rationale.
- **Order.**
  - The scoring standard (`METHOD.md`) came first.
  - Rows were then scored in SOC order, in 12 batch files, between 07:52 and 08:03.
  - A validation script was written and run afterwards, and one rule was then added to `METHOD.md` (Section 8.2).
  - 73 rows were changed after the first pass, all recorded in `revision_log.csv`.
- **Result.**
  - **7.42%** of jobs at the median: 12.64 million job-weighted equivalents. The range is 3.34% (low) to 13.84% (high).
  - 395 of the 831 occupations, holding 31.5% of jobs, score 0%.

## 2. The instructions, and how I read them

### 2.1 The prompt, verbatim

The user first pasted the instructions they had given ChatGPT. They are reproduced here exactly, typos included:

> Okay I want to repeat some work we did before, but be a little more methodological. I just imported a repo called Essential_Demand_Modeling. what we did before was look at all BLS categories and estimate how mcuh of them was essential demand vs derived demand. you can look at my PaidToBeHuman repo to understand what essential demand is. derived demand was codified by alfred marshall over 100 years ago so you know that that is. Essential demand is "when the human is the entire point" or nonfungible or people are willing to pay a premium for the human. or just "not derived demand" - you don't need to use the four categories I codified in paid to be human - those are categories to make it easier to communicate and use. make sense? anyways, what I want to do is have you download the BLS categories again and go through them, one by one and build a spreadsheet estimating how much of each is derived vs essential demand. many jobs are 0% essential demand. some are mixed. And just do your best. in the spreadsheet give the high, median, and low estimate and then add a field explaining it. and make sure you label the spreadsheet with a datetime stamp and that chatgpt 6 pro is doing the work. I'll have claude opus and fable do the same. there's likely to be some convergence but the point of this exercise is to just look at "okay of all the jobs that exist today, how much is actually essential demand?" - alfred marshall even noted way back in 1890 that most labor was derived demand so the same is likely true today. Okay, does all that make sense?

Then this note:

> note, I gave those instructions to ChatGPT. do the same, except of course swap out chatgpt 6 pro for claude opus 5.5 butoverall same isntructions. please proceed.

### 2.2 How I read it

- **"download the BLS categories again"**
  - *What I did:* used BLS Employment Projections Table 1.2, 2025–35.
  - *Why:* it is BLS's complete list of detailed occupations, with total employment for the latest base year (2025), self-employed included. OEWS, the obvious alternative, leaves out the self-employed. The *Paid to Be Human* white paper uses this table (line 1632 of its text).
- **"go through them, one by one"**
  - *What I did:* scored every "Line item" (detailed) occupation, 831 of them, each on its own line with its own rationale. The 281 summary rows are kept for reference but never scored or added in.
  - *Why:* summary rows contain the detailed rows beneath them, so adding them would double-count.
- **"estimating how much of each is derived vs essential demand"**
  - *What I did:* estimated a share of each occupation's labor (Section 6), not a yes/no label.
  - *Why:* the user's question is "of all the jobs that exist today, how much is actually essential demand?", which needs a quantity that adds up across jobs. Most occupations are mixed bundles. The white paper also says the real unit is "a reason within an offering", not an occupation.
- **Essential means "the human is the entire point", nonfungible, worth a premium, or simply "not derived demand"**
  - *What I did:* counted any of these reasons.
- **"you don't need to use the four categories"**
  - *What I did:* used Presence, Provenance, Accountability and Affinity as tags, not as the definition. Every essential reason I found happened to fit one of them.
  - *Note:* after the first pass, a rule made a tag compulsory for every nonzero median (Section 8.2). So in the final data the tags are not optional, although they never forced a reason out.
- **"give the high, median, and low estimate"**
  - *What I did:* gave three readings of the same judgment: strict, central and generous (Section 7.7).
  - *Why:* they show how much the answer depends on how strictly the definition is read. They are not statistical confidence intervals, because there is no sampling process to put an interval on.
- **"add a field explaining it"**
  - *What I did:* added a rationale for every row.
- **"label the spreadsheet with a datetime stamp and that [Claude Opus 5.5] is doing the work"**
  - *What I did:* the model and build time appear in four places:
    - the file names;
    - the title rows of the Summary, Occupations and Method sheets;
    - the workbook properties (author, comments);
    - the `model` and `generated` columns on every CSV row.
- **"many jobs are 0% essential demand. some are mixed." and "most labor was derived demand"**
  - *What I did:* read these as the user's expectation, not as a target. They are still a prior I was given; see Section 11.

I also fixed these method choices before scoring, in `METHOD.md`:

- **Hours, not job titles.** One essential task does not keep a whole job alive. Only the share of time spent on it counts.
- **Capability gaps excluded.** The test holds the result constant, so work people do only because machines cannot do it yet is derived demand that hasn't been automated.
- **The standard came first.** The definition and the scale anchors were written down before I saw any row. That way each row is scored against the standard, not against a running total.

## 3. Inputs

### 3.1 Data

- `source_data/BLS_EP_occupation.xlsx`: BLS's workbook, as downloaded (Section 4).
- `source_data/BLS_EP_Table_1.2_2025-35.csv`: Table 1.2 extracted from it (Section 5).

### 3.2 Background I read before scoring

All of these are in the user's `C:\PaidToBeHuman` repository, which I only read.

| File | Part read | What it contributed |
|---|---|---|
| `README.md`, `METHODS.md`, `method-sources.json` | The whole files | Orientation only. They describe that repository's production methods and contributed nothing to the scoring. |
| `course/Course plan.md` | The whole file | The course's test: "if you cannot swap any machine out, no matter the quality, and have the same product or service, then that is essential demand." Its lesson examples on the essential side (the barber, the tattoo artist, massage, hospice, doulas) and the derived side (Marshall's house, the plumber, firefighting and rescue machines). A line noting the earlier audit at "about 5 percent" (line 57). |
| White paper text (see below) | The first 12,000 bytes | The title page, the central question and the opening. |
| | Lines 183–338 | The definitions of Essential Demand and the substitution test, and the point that an occupation is too coarse a unit. |
| | Lines 1170–1190 and 1300–1345 | Presence versus provenance, and the paid hours each consumes. |
| | Lines 1625–1680 | Line 1632: the BLS projections table reports 170,280.8 thousand jobs for 2025, and summary and detailed rows must not be added together. Line 1636: the prior audit's aggregate results of 1.59%, 5.47% and 14.85% of covered employment. |
| | Two keyword searches (BLS, SOC, occupation, OEWS, O*NET, "audit", "5 percent") and a list of all headings | Finding the sections above. |

The white paper text is `reference/background/Derived vs Essential Demand v01 (text).md`. It is the text extraction of *Derived and Essential Demand: Human work, income distribution, and the Fourth Sector* (Labor Zero, consolidated research edition, 4 Oct 2026), made from `Post-Labor Economics - Derived vs Essential Demand v01.docx` in the same folder. The workbook's Sources sheet cites that `.docx` for the definitions; I read the text extraction.

For derived demand I relied on Alfred Marshall, *Principles of Economics* (1890), Book V, chapter 6, from general knowledge. I did not re-read it.

### 3.3 What I did not use

- **The prior audit's row-level coefficients.** I never opened them; I saw only the aggregate results the white paper reports.
- **Any other run of this exercise.** That includes the earlier one, ChatGPT's and Fable's. When this run started, the Essential_Demand_Modeling repository held only a one-line README.
- **Outside data.** No O*NET task data, OEWS detail, time-use surveys, consumer surveys or web search. The only network access was to bls.gov.
- **Median wage.** It was shown in the listings but was never a scoring input, and no rationale cites it.
- **Any other repository** on the PC.

## 4. Getting the BLS data

1. **07:44:40.** A direct command-line download of the workbook from bls.gov was refused (HTTP 403). BLS does not serve these files to command-line clients.
2. **07:44:46.** I opened the table's web page in the built-in browser of the Claude desktop app: https://www.bls.gov/emp/tables/occupational-projections-and-characteristics.htm
   - Reading the page confirmed it was the right table. It has one table of 1,114 rows, captioned "Table 1.2 Occupational projections, 2025–35, and worker characteristics, 2025 (Employment in thousands)".
   - Summing that HTML table in the page (07:45:27) gave 831 "Line item" rows totaling 170,267.2 thousand, with no duplicate codes.
3. **07:45:32.** I opened the workbook link in the same browser: https://www.bls.gov/emp/ind-occ-matrix/occupation.xlsx
   - The browser saved the file to the Windows Downloads folder under a temporary name: `5ae2b673-46e4-4ab0-b715-bc3bff6c63d1.tmp`, 413,524 bytes, at 07:45:33.
4. **07:45:56.** I checked that the file was an Excel 2007+ workbook and copied it to a scratch folder.
5. **07:46:49.** I copied it into this folder as `source_data/BLS_EP_occupation.xlsx` and recorded its SHA-256: `08ee182f992d972593e64559ce467bad75967d2228f11ad5e2f86d4ec0ee6c00`.

The temporary file is still in Downloads; nothing was deleted there.

The workbook's sheets are Index and Tables 1.1–1.6 and 1.8–1.12 (there is no 1.7). Only Table 1.2 was used.

**To reproduce this step:** download the same URL in an ordinary web browser and compare the SHA-256. BLS revises these tables with each annual projections release, so a later download may differ. The copy in `source_data/` is this run's baseline.

## 5. Extracting Table 1.2

`scripts/extract_table_1_2.py` does the extraction. It reads the workbook with `scripts/xlsx_read.py`, a small standard-library reader (zipfile plus XML parsing), because openpyxl is not installed for Python 3.13 on this PC.

The extractor:

- opens the sheet "Table 1.2";
- finds the header row, whose first cell is "2025 National Employment Matrix title";
- keeps rows whose occupation type is "Summary" or "Line item", dropping notes and footers;
- sets the hierarchy level to the number of leading spaces in the title divided by 2 (BLS indents titles by level);
- turns BLS's dash for unavailable data into a blank;
- assigns each row to the most recent major-group row (a code ending in -0000);
- writes the columns `soc_code`, `title`, `occupation_type`, `level`, `major_group_code`, `emp_2025_thousands`, `emp_2035_thousands`, `pct_self_employed_2025`, `median_wage_2025` and `typical_entry_education`.

**Checks**

- The output has 1,112 rows: 281 Summary and 831 Line item.
- The line items sum to 170,267.2 thousand, the same as the sum of the web page's table (Section 4).
- BLS's "Total, all occupations" row is 170,280.8 thousand, a gap of 13.6 thousand.
  - 13.5 thousand of it sits in Top executives. That summary row is 3,904.1 thousand, but its only detailed rows here, Chief executives (291.6) and General and operations managers (3,599.0), add to 3,890.6.
  - The missing row is most likely Legislators (11-1031), which does not appear as a detailed row in this table. Scored at 80–95%, it would add about 11–13 thousand essential jobs, which is under 0.01 percentage points.
  - Every other summary row matches the sum of its detailed rows to within ±0.4 thousand, which is rounding.
- The denominator for every share in this run is 170,267.2, the sum of the line items, not 170,280.8.

## 6. What each percentage means

This restates `METHOD.md`, which is the authority.

**Essential share** is my estimate of the share of an occupation's current labor that buyers, audiences or responsible institutions would still pay a human to do, if a disclosed machine alternative delivered the same functional result at equal or lower cost.

- The labor is job-weighted, which is roughly its hours.
- Prices are held at roughly today's level.
- Everything else in the occupation is derived demand.

In the white paper's terms, it is the share of today's hours that would survive the substitution test, with today's level of demand held constant.

What it is not:

- It is not a forecast of when, or whether, automation happens.
- It is not the chance that the job title survives.
- It is not an income estimate.

What the parts of the definition imply:

- **"Disclosed."** The buyer knows it is a machine.
- **"Same functional result."** The machine's capability is assumed, so only a preference for the human counts.
- **"Today's prices."** Payers' budgets matter. For example, a Medicaid program buying home care would take a cheaper equivalent for much of the time.
- **"Job-weighted."** The matrix counts jobs, not full-time equivalents. An occupation scored at 30% contributes 0.3 job-equivalents for each job.

**What counts as essential:** any reason the human is the point, nonfungible, or worth a premium. Four tags label the kinds:

- **Presence:** the buyer wants an actual person there (care, company, live performance, hospitality where the person is the point).
- **Provenance:** human making or authorship is part of what is bought.
- **Accountability:** a person must answer for, authorize, represent or bear consequences. It is counted only for the hours of meaningful accountable work, never for the whole job around it.
- **Affinity:** a particular person, or a continuing relationship with one.

**What does not count:**

- Work done only because machines can't do it yet.
- People skills a job requires, when the buyer would accept an equally good machine.
- Work that is important or heroic but valued for its outcome (rescue, surgery, firefighting).
- A rule that a named person sign, beyond the time a meaningful review takes.
- Customers' wish to be among other customers. That is demand for each other's company, not for the worker's labor.

## 7. How the rows were scored

### 7.1 Order and batches

`METHOD.md` was composed at 07:51:41, before the first row (07:52:08).

Before each batch I printed its occupations with `scripts/list_group.py`, which shows code, title, 2025 employment, percent self-employed and median wage. I then wrote the batch as one file, in BLS order.

| Major groups | File in `estimates/` | Rows | Listed | Written |
|---|---|---|---|---|
| 11 Management | `11_management.txt` | 37 | 07:50:24 | 07:52:08 |
| 13 Business and financial | `13_business_financial.txt` | 32 | 07:50:24 | 07:52:25 |
| 15, 17, 19 Computer, engineering, science | `15_17_19_computer_engineering_science.txt` | 105 | 07:52:27 | 07:53:11 |
| 21, 23 Social service, legal | `21_23_social_legal.txt` | 25 | 07:53:12 | 07:54:35 |
| 25 Education | `25_education.txt` | 64 | 07:53:12 | 07:55:06 |
| 27 Arts, media, sports | `27_arts_media.txt` | 41 | 07:55:08 | 07:56:50 |
| 29 Healthcare practitioners | `29_healthcare_practitioners.txt` | 71 | 07:55:08 | 07:57:20 |
| 31, 33 Healthcare support, protective | `31_33_healthsupport_protective.txt` | 41 | 07:57:22 | 07:59:03 |
| 35, 37, 39 Food, cleaning, personal care | `35_37_39_food_cleaning_personal.txt` | 59 | 07:57:22 | 07:59:29 |
| 41, 43, 45 Sales, office, farming | `41_43_45_sales_office_farming.txt` | 90 | 07:59:30 | 08:00:50 |
| 47, 49 Construction, repair | `47_49_construction_repair.txt` | 111 | 08:00:52 | 08:01:47 |
| 51, 53 Production, transportation | `51_53_production_transport.txt` | 155 | 08:01:48 | 08:03:16 |

The "Written" times are when each file's content was composed, from the session log. Files written in the same step reached the disk together a few seconds later, so the files' own creation times pair up.

The first pass took about 13 minutes, from the first listing to the last file.

### 7.2 The procedure for each row

1. Identify the buyer (customer, patient, family, employer, audience or public institution) and what they are buying.
2. **Question 1 (B):** under the substitution test, what share of those buyers want the human, not just the result?
3. **Question 2 (H):** within the work for those buyers, what share of the worker's time is the human part (the session, the performance, the decision, being there)? The rest is support work a machine could take over while the human stays.
4. Set the median to roughly B × H, adjusted by judgment. Check it against the scale anchors (Section 7.5) and against similar rows already scored.
5. Set the low and high around it (Section 7.7).
6. Tag the kind of essential demand, if any (Section 7.8).
7. Write a rationale of one or two sentences (Section 7.9).

For most rows, steps 2–4 were a single holistic judgment, with no numbers written down.

### 7.3 Rows worked out with explicit arithmetic

For a handful of the largest or most ambiguous rows, I worked the estimate out in numbers.

- **Home health and personal care aides** (31-1120). Result 12 / 28 / 48.
  - Time share × essential share:

    | Part of the job | Time | Essential |
    |---|---|---|
    | Personal care | 35% | 35% |
    | Housekeeping and meals | 30% | 5% |
    | Companionship and supervision | 25% | 70% |
    | Transport | 10% | 5% |

  - That comes to about 32%. I lowered it to 28 because the main payer, Medicaid, buys at today's prices and would take a cheaper equivalent for part of the time.
- **Registered nurses** (29-1141). Result 7 / 18 / 35.
  - Time: direct care 30–35%, documentation 25–35%, medications 10–20%, coordination 10–15%, patient education 5–10%.
  - Essential: about half of direct care (≈16 points), half of education (≈4) and accountable judgment (≈2), so ≈22%.
  - × B ≈ 0.8 gives 18.
- **Elementary teachers** (25-2021). Result 15 / 32 / 50.
  - Instruction 40% × 0.3, supervision 20% × 0.8, relationship 15% × 0.9, administration 25% × 0. That comes to ≈42%.
  - I lowered it to 32. Once instruction can be automated, one adult could supervise more children, so not all of today's supervision hours would survive.
- **Postsecondary teachers** (25-1xxx). Result: field bands (Section 7.4).
  - Lecture 20% × 0.25, grading 10% × 0, mentoring 10% × 0.8, seminar 10% × 0.6, research 35% × 0.05, service 15% × 0.1. That comes to ≈22% for a generic faculty member.
  - The field bands were then set by judgment around that figure, most fields at 16–20%. I did not write down the step from 22% to each band.
- **Waiters and waitresses** (35-3031). Result 6 / 18 / 34.
  - This was re-derived after the first pass (Section 8.3).
  - B from a split of covers: upscale 15% × 0.9, casual 50% × 0.25, other 35% × 0.15. That gives ≈0.31, which I trimmed to 0.25–0.30.
  - × H ≈ 0.70–0.75 gives about 18–21%. I set it at 18.
- **Bartenders** (35-3011). Result 12 / 28 / 48.
  - B 0.35–0.40 × H 0.75 gives 26–30%. The value already written, 28, was kept.
- **Substitute teachers, short-term** (25-3031). Result 12 / 28 / 45.
  - B ≈ 0.6 × H ≈ 0.5 gives about 30%. I set it at 28.
- **Nursing assistants** (31-1131). Result 7 / 18 / 35.
  - Direct care is 65% of the time, 35–40% of that is essential, and B is 0.8. That gives ≈18–21%. The value already written, 18, was kept.

I did not record at the time which of these were worked out before the value was written and which were done as checks afterwards. The exception is waiters, which was changed after the first pass. Every other row is a holistic judgment guided by the two questions and the anchors.

### 7.4 Families scored as bands

The 37 postsecondary teacher rows were scored as six bands by field:

| Low / median / high | Fields (rows) |
|---|---|
| 5 / 14 / 28 | Career/technical education (1) |
| 6 / 16 / 32 | Business; computer science; mathematical science; engineering; agricultural, biological, forestry and conservation, earth and space sciences; chemistry; environmental science; physics; library science; criminal justice; communications; family and consumer sciences; all other (16) |
| 7 / 18 / 35 | Architecture; anthropology; area and cultural studies; economics; geography; political science; psychology; sociology; other social sciences; foreign language; history (11) |
| 8 / 20 / 38 | Health specialties; nursing; education; law; social work; English (6) |
| 10 / 24 / 44 | Philosophy and religion; recreation and fitness studies (2) |
| 15 / 32 / 52 | Art, drama and music (1) |

Elsewhere, I gave identical values on purpose to rows whose work and buyers are alike. For example, the seven construction helper rows (47-3011 to 47-3019) are all 0 / 0 / 0. Identical values are a consistency choice, not evidence that the rows are the same.

### 7.5 Scale anchors (median), from `METHOD.md`

| Median | Meaning | Typical examples |
|---|---|---|
| 0% | The buyer wants the output; no one would pay more for a human | production, freight, IT, clerical, most construction |
| 0.5–5% | A thin segment pays for the human, or a small accountable sign-off | cashiers, customer service, accountants |
| 5–20% | A real but minority premium segment, or a small human core in a large functional job | waiters, real estate agents, nurses |
| 20–50% | A mixed bundle in which the human part is a large share | teachers, hairstylists, personal care aides |
| 50–80% | The human is mostly the point; support work is derived | therapists, clergy, performers, fitness instructors |
| 80–100% | The human is the product | athletes, elected representatives |

### 7.6 Evidence used for each row

- **From BLS:** the occupation's title and code (what the job is) and its 2025 employment (how much the row matters). Also the percent self-employed, as a signal of owner-operators and personal-service businesses; 27 rationales cite it.
- **General knowledge** of the work and its buyers, from the model's training. I did no per-occupation lookups.
- **Revealed preference,** where I knew of a market that already offers a machine or cheaper option. Examples:
  - casino players choosing live dealers over electronic tables (39-3011);
  - people paying for human guides when audio guides and apps exist (39-7010);
  - tips as partial payment for service (35-3031);
  - clients who follow a particular stylist or instructor.
- **Employment size.** It changed how much scrutiny a row got (the largest rows got arithmetic), not its share.
- **"All other" rows.** I scored these by guessing their mix from the title and the occupations BLS usually puts there.

### 7.7 Low, median and high

- **Low** is a strict reading: only clear, durable demand for the human, counted on the narrowest set of hours.
- **Median** is my central judgment.
- **High** is a generous but defensible reading. It takes in softer premium segments and more of the relationship time. It never includes demand that exists only because machines can't do the work yet.

No formula set the ranges. Each was judged per row. In practice they follow a steady pattern: wide relative ranges for small shares, narrowing as the share rises. The pattern below was computed afterwards from the final values, and a replicator can use it as a default.

| Median band | Rows | Low ÷ median: typical (range) | High ÷ median: typical (range) |
|---|---|---|---|
| Above 0, below 5% | 177 | 0.00 (0.00–0.50) | 3.00 (2.00–4.00) |
| 5 to below 20% | 148 | 0.38 (0.20–0.44) | 2.20 (1.78–2.50) |
| 20 to below 50% | 83 | 0.46 (0.40–0.62) | 1.71 (1.42–2.00) |
| 50 to below 80% | 25 | 0.60 (0.58–0.77) | 1.38 (1.15–1.44) |
| 80 to 100% | 3 | 0.81 (0.79–0.91) | 1.13 (1.05–1.15) |

"Typical" is the median of the per-row ratios.

The 395 rows with a 0% median have these highs:

| High | Rows |
|---|---|
| 0 | 168 |
| 0.5% | 83 |
| 1% | 78 |
| 2% | 33 |
| 3% | 33 |

A high above 0 on such a row is a general allowance for segments I may have missed, not a claim.

The values cluster on round numbers. There are 43 distinct medians. The most common nonzero ones are:

| Median | Rows |
|---|---|
| 1% | 45 |
| 3% | 40 |
| 4% | 31 |
| 0.5% | 30 |
| 18% | 26 |
| 6% | 22 |
| 2% | 21 |
| 8% | 19 |
| 10% | 17 |
| 20% | 17 |
| 12% | 16 |
| 16% | 16 |

### 7.8 Tags

The four tags are defined in Section 6, and up to four are listed per row. The first listed is the primary mechanism, which the Summary sheet uses to split the total.

| Tag | Rows carrying it | Rows where it is primary |
|---|---|---|
| Presence | 272 | 177 |
| Affinity | 204 | 91 |
| Accountability | 152 | 106 |
| Provenance | 81 | 65 |
| None | n/a | 392 |

228 rows carry more than one tag.

After the first pass, every nonzero median must carry a tag, and an all-zero row carries none. Three rows have a 0% median but a tag, because their high names a sliver:

- 47-5032 Explosives workers, 0 / 0 / 0.5, Accountability.
- 51-4121 Welders, 0 / 0 / 0.5, Provenance.
- 51-8031 Water treatment operators, 0 / 0 / 0.5, Accountability.

### 7.9 Rationales

- Each is one or two sentences: what part is essential, and why the rest is derived. The mean length is 13 words (3 to 63).
- `METHOD.md` says rationales also explain what drives the low and high. In practice only 6 of the 831 mention the range, so the range reasoning is mostly unwritten.
- A few rationales mention how automatable a task is (parking enforcement, nail art). Under the substitution test, feasibility does not matter. Those phrases describe the job, and the values rest on the buyer preference stated in the same rationale. A reader may still see them as capability reasoning creeping in.

## 8. Checks and changes after the first pass

### 8.1 The validation rules

`scripts/load_estimates.py` was written at 08:03, after all the rows. Both the check and the build use it, and the build refuses to write a workbook if any rule fails. The rules are:

- Each line has exactly 6 fields.
- The code is a BLS detailed occupation, and it is scored only once.
- No occupation is missing.
- The values are numbers, with 0 ≤ low ≤ median ≤ high ≤ 100.
- Tags come from the four allowed names, or "None".
- A nonzero median has at least one tag. An all-zero row has no tag.
- The rationale is at least 20 characters long.

### 8.2 The consistency rule (72 rows)

The first validation run (08:03:42) found 72 errors, all of the same kind: a median of 0.5–2% with no tag. Listing them (08:03:55) showed that all but one (11-3051, below) named nothing essential, for example "Procurement is a means to getting inputs at the right price. Derived." It was a reflexive small allowance, not a judgment that something essential was there.

So I adopted this rule:

- A nonzero median must name what is essential and carry a tag.
- Otherwise the low and median are 0, and the high is at most 3%, as a general allowance.

`scripts/fix_untagged.py` applied the rule at 08:04:17, after which there were 0 errors. The rule was added to `METHOD.md` at 08:05.

- **70 rows:** low 0, median 0, high = the smaller of the old high and 3.
- **11-3051 Industrial production managers.** The first-pass rationale named a sliver: "Small allowance for accountable safety decisions." So I kept its values (0 / 1 / 4), tagged it Accountability and reworded the rationale to name the sliver.
- **39-3092 Costume attendants.** The first-pass rationale was "Costume handling is instrumental.", which names nothing. The rule alone gives 0 / 0 / 3. Instead, I re-judged the row while writing the fix:
  - I added a sliver (performers who want a trusted dresser backstage), tagged it Affinity and set it to 0 / 1 / 4.
  - This was a judgment made inside a mechanical fix.
  - `METHOD.md`'s first description of the fix (08:05:29) said both of these rows "did name a sliver". That is accurate for 11-3051 but not for 39-3092. The description was corrected at 10:24 (Section 8.6).
  - The effect on the total is 6.6 thousand jobs × 1%, or 0.07 thousand.

### 8.3 Waiters

After the first validation run, which lists the 25 largest contributors, I re-derived the waiters row with explicit arithmetic (Section 7.3). It gave about 18–21%, above the first-pass 16%.

At 08:05:26 I changed it from 6 / 16 / 32 to 6 / 18 / 34, and rewrote the rationale to state the arithmetic.

### 8.4 Effect on the totals

| Stage | Low | Median | High |
|---|---|---|---|
| First pass (08:03) | 5,681.8K (3.34%) | 12,668.3K (7.44%) | 23,526.6K (13.82%) |
| After the consistency rule (08:04) | 5,681.8K (3.34%) | 12,593.6K (7.40%) | 23,519.4K (13.81%) |
| After the waiters change (08:05), final | 5,681.8K (3.34%) | 12,639.3K (7.42%) | 23,565.1K (13.84%) |

Both changes came after I had seen the first-pass total of 7.44%. The consistency rule lowered the median total by 74.6K and the waiters change raised it by 45.7K. Neither was made to move the total toward any figure.

### 8.5 Nothing else changed

The first-pass files exactly as first written are recoverable from the session log. Comparing all 831 of their rows with the final files, exactly these 73 rows differ, and only 3 rationales changed (11-3051, 35-3031, 39-3092).

`revision_log.csv` lists every one, with:

- the before and after values and tags;
- the before and after rationales;
- the effect on the median total;
- the reason for the change.

### 8.6 Later corrections to `METHOD.md` (no estimates changed)

After this document was first written, I made two corrections to `METHOD.md`, both at the user's request. No estimate, tag, rationale or total changed.

**1. The account of the fix (10:24:29).** `METHOD.md` gained this sentence at 08:05:29 to describe the consistency fix:

> This rule was enforced by the validation script after the first pass: 72 rows had a reflexive 0.5–2% median with no named essential element, and 70 of them were set to 0% while 2 that did name a sliver were tagged.

What was wrong with it:

- 39-3092 did not name a sliver. It was re-judged during the fix (Section 8.2).
- 11-3051 did name an essential element, so only 71 of the 72 rows named none.
- The validation script only found the rows. `scripts/fix_untagged.py` changed them, and the rule itself was written into `METHOD.md` only after the first pass.

The corrected wording says that the rule was added after the first pass, and covers each group:

- the 70 rows that named nothing were set to 0% (high capped at 3%);
- 11-3051 named a sliver and was tagged;
- 39-3092 named nothing and was re-judged to 0 / 1 / 4.

It now ends with a note that both corrections were made after the 08:09 build.

**2. The opening line (10:28:56).** It made no exception for that rule. It read:

> This file was written before any occupation was scored, and it is the standard every row was scored against.

It now reads:

> This file was written before any occupation was scored, and it is the standard every row was scored against, with one exception: the rule that a nonzero median must name what is essential (under "Low, median and high") was added after the first pass.

**Effect on the other files**

- The 08:09 workbook's Method sheet still has both original wordings, because the build copies `METHOD.md` as it stands at build time. I did not rebuild the workbook, so its name and hash are unchanged.
- A rebuild (Section 12.1) puts the corrected wording on the Method sheet.
- Section 14 lists the corrected file's hash. Putting both original wordings back gives the file as it was at the 08:09 build, SHA-256 `08e8031abe3111b7882dd8a528091fb5f407a8db6c1e8f8738f281ceff650608`. I checked this.

## 9. Building the spreadsheet and CSV

`scripts/build_workbook.py` runs on Python 3.14.7 with xlsxwriter 3.2.9. It:

- loads the BLS CSV and the estimates, and runs every validation rule, stopping if any fails;
- names both outputs with the model and the local build time: `Essential_vs_Derived_Demand_by_Occupation_Claude_Opus_5.5_<YYYY-MM-DD_HHMM>.xlsx` and `.csv`;
- writes the model, model ID and generated time into:
  - the title rows of the Summary, Occupations and Method sheets;
  - the workbook properties (author "Claude Opus 5.5 (claude-opus-5-5), Anthropic", plus comments);
  - every CSV row;
- writes the totals as live Excel formulas, so editing any Low, Median or High cell updates every total (formulas in Section 13);
- records the source workbook's SHA-256 on the Sources sheet;
- builds the Method sheet from `METHOD.md` as it stands at build time, so the 08:09 workbook has the wording from before the corrections in Section 8.6.

**Builds**

- The first build was at 08:07.
- After a text-only change to the Sources sheet, the final build was at 08:09.
- The two CSVs were identical apart from the timestamp, so I deleted the 08:07 files and kept one set.

**Checks on the workbook**

- **08:08.** I opened the 08:07 build read-only in Microsoft Excel through COM automation, recalculated it fully and read the formula results back. All of these matched the Python totals:
  - the headline shares, 3.34% / 7.42% / 13.84%;
  - essential jobs, 5,681.8 / 12,639.3 / 23,565.1 thousand;
  - 831 occupations, 395 of them at 0%;
  - the distribution and mechanism tables;
  - the Occupations total row.

  No cell showed an error value.
- **08:09.** I opened the final build read-only in Excel, recalculated it, and read a few rows back correctly. I exported a PDF of the Summary sheet to a scratch folder for a visual check, but no PDF renderer was available, so the layout was never inspected by eye.

## 10. Results

Of 170,267.2 thousand jobs:

| | Low | Median | High |
|---|---|---|---|
| Share of jobs | 3.34% | 7.42% | 13.84% |
| Essential jobs, in thousands | 5,681.8 | 12,639.3 | 23,565.1 |

At the median, the other 92.58% is derived demand.

The figures are judgments under the substitution test, not measurements or forecasts.

**By major group** (employment is the sum of each group's detailed rows):

| Code | Major group | Jobs (K) | Low | Median | High | Median essential jobs (K) |
|---|---|---|---|---|---|---|
| 11 | Management | 13,670.0 | 2.12% | 5.92% | 13.44% | 809.8 |
| 13 | Business and financial operations | 11,368.5 | 1.10% | 3.35% | 8.43% | 380.9 |
| 15 | Computer and mathematical | 5,493.8 | 0.00% | 0.10% | 1.21% | 5.7 |
| 17 | Architecture and engineering | 2,690.2 | 0.16% | 0.85% | 2.71% | 23.0 |
| 19 | Life, physical, and social science | 1,628.5 | 3.53% | 6.25% | 10.41% | 101.8 |
| 21 | Community and social service | 3,270.1 | 21.15% | 35.89% | 54.17% | 1,173.5 |
| 23 | Legal | 1,445.8 | 5.60% | 12.31% | 22.67% | 178.0 |
| 25 | Educational instruction and library | 9,907.3 | 13.45% | 27.43% | 44.96% | 2,717.9 |
| 27 | Arts, design, entertainment, sports, and media | 2,985.9 | 12.04% | 20.81% | 32.64% | 621.4 |
| 29 | Healthcare practitioners and technical | 10,232.8 | 5.22% | 12.92% | 25.23% | 1,322.2 |
| 31 | Healthcare support | 8,421.3 | 9.03% | 21.00% | 37.07% | 1,768.8 |
| 33 | Protective service | 3,871.4 | 3.33% | 8.15% | 17.47% | 315.7 |
| 35 | Food preparation and serving related | 13,826.5 | 1.99% | 5.99% | 12.29% | 828.5 |
| 37 | Building and grounds cleaning and maintenance | 5,699.6 | 0.00% | 0.43% | 1.85% | 24.3 |
| 39 | Personal care and service | 4,487.4 | 17.36% | 31.12% | 47.61% | 1,396.5 |
| 41 | Sales and related | 14,533.1 | 1.12% | 3.60% | 8.91% | 523.5 |
| 43 | Office and administrative support | 18,792.5 | 0.24% | 1.17% | 3.83% | 220.4 |
| 45 | Farming, fishing, and forestry | 1,036.1 | 0.00% | 0.25% | 0.93% | 2.6 |
| 47 | Construction and extraction | 7,492.9 | 0.01% | 0.41% | 1.37% | 30.7 |
| 49 | Installation, maintenance, and repair | 6,494.1 | 0.00% | 0.29% | 1.13% | 18.5 |
| 51 | Production | 8,780.2 | 0.10% | 0.36% | 0.99% | 31.5 |
| 53 | Transportation and material moving | 14,139.2 | 0.34% | 1.02% | 2.28% | 144.1 |
| | **Total** | **170,267.2** | **3.34%** | **7.42%** | **13.84%** | **12,639.3** |

**By median band:**

| Median | Occupations | Jobs (K) | Share of jobs | Median essential jobs (K) | Share of essential total |
|---|---|---|---|---|---|
| 0% | 395 | 53,709.7 | 31.5% | 0.0 | 0.0% |
| Above 0, below 5% | 177 | 60,289.9 | 35.4% | 1,122.6 | 8.9% |
| 5 to below 20% | 148 | 32,377.0 | 19.0% | 3,621.4 | 28.7% |
| 20 to below 50% | 83 | 19,937.5 | 11.7% | 5,731.9 | 45.3% |
| 50 to below 80% | 25 | 3,674.4 | 2.2% | 1,932.7 | 15.3% |
| 80 to 100% | 3 | 278.7 | 0.2% | 230.8 | 1.8% |

The three occupations at 80% or above are:

- Athletes and sports competitors, 93%;
- Clergy, 82%;
- Dancers, 80%.

28 occupations have a median of 50% or more, holding 3,953.1 thousand jobs (2.3%).

**By primary mechanism:**

| Mechanism | Share of the median essential total | Essential jobs |
|---|---|---|
| Presence | 71.1% | 8,984.2K |
| Affinity | 13.5% | 1,704.6K |
| Accountability | 12.5% | 1,581.9K |
| Provenance | 2.9% | 368.6K |

**Concentration.** The 10 largest contributors hold 37.9% of the median essential total, the top 25 hold 58.3%, and the top 50 hold 72.4%.

The last column shows how far each row alone moves the total if it is set to its own low or high:

| Occupation | Jobs (K) | Low / median / high | Median essential jobs (K) | Total if this row alone is at its low / high |
|---|---|---|---|---|
| 31-1120 Home health and personal care aides | 4,677.1 | 12 / 28 / 48 | 1,309.6 | 6.98% / 7.97% |
| 29-1141 Registered nurses | 3,465.4 | 7 / 18 / 35 | 623.8 | 7.20% / 7.77% |
| 39-9011 Childcare workers | 1,010.3 | 30 / 50 / 72 | 505.1 | 7.30% / 7.55% |
| 25-2021 Elementary school teachers | 1,419.3 | 15 / 32 / 50 | 454.2 | 7.28% / 7.57% |
| 35-3031 Waiters and waitresses | 2,284.4 | 6 / 18 / 34 | 411.2 | 7.26% / 7.64% |
| 25-9045 Teaching assistants | 1,463.0 | 12 / 25 / 45 | 365.8 | 7.31% / 7.60% |
| 25-2011 Preschool teachers | 583.2 | 30 / 50 / 70 | 291.6 | 7.35% / 7.49% |
| 11-1021 General and operations managers | 3,599.0 | 3 / 8 / 18 | 287.9 | 7.32% / 7.63% |
| 25-2031 Secondary school teachers | 1,087.5 | 12 / 25 / 42 | 271.9 | 7.34% / 7.53% |
| 31-1131 Nursing assistants | 1,505.9 | 7 / 18 / 35 | 271.1 | 7.33% / 7.57% |

With all ten at their lows and everything else at the median, the total would be 5.87%. With all ten at their highs, it would be 9.52%. Differences between runs are likely to come mostly from these rows.

## 11. Biases and limitations

The user said bias is fine as long as it is documented. These are the ones I know of.

**Priors and anchoring**

1. The prompt states an expectation: "many jobs are 0% essential demand" and "most labor was derived demand so the same is likely true today".
2. Before scoring, I read the white paper's report of the prior audit's totals (1.59%, 5.47% and 14.85%) and the course plan's "about 5 percent". The final median, 7.42%, is near those figures, and I cannot rule out anchoring.
3. Before writing any row, I sketched provisional medians for the largest occupations and noted the rough total they implied (about 6.8%), as a check that the scale anchors were sensible. So a rough total was in view before scoring, though the first computed total (7.44%) came only after all 831 rows were written.
4. The course's lesson examples may have shaped similar rows. On the essential side these are the barber, the tattoo artist, massage, hospice and doulas; on the derived side, Marshall's house, the plumber, and firefighting and rescue machines.
5. The framework itself is the lens: the substitution test and the four mechanisms. A rater working from a different framework would score differently.

**Evidence**

6. There is one rater, judging from general knowledge, with no data for any row. The values are judgments.
7. Buyers' preferences under the substitution test are hypothetical. I used revealed-preference evidence only where I happened to know of it.
8. I guessed the mix of the "all other" rows.

**Process**

9. **Speed.** The 831 rows were written in about 11 minutes (07:52–08:03), under a second per row on average. Most rows got a quick holistic judgment; only the largest got worked arithmetic.
10. **Effort went to the large rows.** Small rows may have larger relative errors, though they have little effect on the total.
11. **The validator came after scoring.** I adopted the consistency rule after seeing 72 violations. One row (39-3092) was re-judged rather than mechanically fixed.
12. **Both post-pass changes came after I saw the first total.**
13. **No review beyond what is described here.**
    - There was no second rater, no test-retest and no blind re-scoring.
    - There was no systematic cross-group consistency review.
    - The supervision adjustment applied to elementary teachers (fewer adults per child once instruction is automated) was not checked for consistent use on other supervision-heavy rows such as childcare workers, preschool teachers and teaching assistants.
14. **Run-to-run variance is unknown.** The same model given the same prompt would probably give different numbers. Treat this run as one draw.

**Values and ranges**

15. The values cluster on round numbers (Section 7.7).
16. Some families were scored as uniform bands (Section 7.4).
17. Low and high are reading scenarios, not confidence intervals. Their widths follow a habitual pattern rather than a measured uncertainty.
18. The rationales rarely explain the range.

**Definitional choices that move the numbers**

19. **Accountability** counts only the hours of meaningful accountable work. Legal requirements for a licensed or named person, beyond the time a real review takes, count as derived. A rater who counts statutory requirements as essential would score licensed trades and professions higher.
20. **"At today's prices"** lowers rows whose payers are price-constrained, such as home care. A rater who ignored price would score them higher.
21. **Customers' demand for each other's company** is excluded. This matters for bars, gyms and congregations, and a looser reading would raise those rows.
22. **Heroic outcome work** is derived. Firefighters, for example, get 3%.
23. A few rationales mention automatability (Section 7.9).

**Data coverage**

24. Table 1.2 has no military occupations.
25. Legislators are missing (about 13.5 thousand jobs, Section 5).
26. The counts are jobs, not full-time equivalents, so part-time-heavy occupations weigh by job count.
27. The median wages are OEWS data that exclude the self-employed. They were not used.
28. The 2025 employment figures are BLS estimates.

**This document**

29. I wrote it after the scoring (from 08:26), from the files and the session's tool-call log, and recomputed the numbers from the files. Where it describes reasoning, such as which arithmetic was done when, it relies on my own record of the session rather than on a file.

## 12. How to reproduce

### 12.1 Rebuild the outputs exactly

Everything from the committed judgments onward is deterministic. Run these from the `Claude_Opus_5.5` folder:

```
py -3.13 -I scripts/extract_table_1_2.py source_data/BLS_EP_occupation.xlsx source_data/BLS_EP_Table_1.2_2025-35.csv
py -3.13 -I scripts/check_estimates.py .
py -3.14 -I scripts/build_workbook.py .
```

- **Python.** The extraction and the check use only the standard library (tested on 3.13.0). The build needs xlsxwriter (tested with 3.2.9 on Python 3.14.7).
- **The `-I` flag.** It runs Python in isolated mode, which ignores environment variables and the user's site-packages. So xlsxwriter must be installed in the interpreter's own site-packages, or you can drop `-I`.
- **Expected results:**
  - The extraction should reproduce `BLS_EP_Table_1.2_2025-35.csv` byte for byte (SHA-256 below).
  - The check should print `errors: 0` and the totals in Section 10.
  - The build writes a new workbook and CSV named with the current time. Every column of the CSV except `generated` should match the shipped CSV. The workbook's bytes and hash will differ, because they include the build time. Its Method sheet will also show the corrected `METHOD.md` wording (Section 8.6), where the 08:09 workbook shows the original.

### 12.2 Re-run the judgments with another rater or model

1. Give the rater the same inputs:
   - the prompt in Section 2.1;
   - `source_data/BLS_EP_Table_1.2_2025-35.csv`;
   - `METHOD.md`;
   - optionally, the white paper sections in Section 3.2 (they are in a private repository).
2. Either keep `METHOD.md` as the standard or have the rater write their own before scoring. Then record which one was used.
3. Print each batch with `py -3.13 -I scripts/list_group.py source_data/BLS_EP_Table_1.2_2025-35.csv 11 13`, and so on for the other groups. The arguments are two-digit major-group prefixes.
4. Score every row in SOC order, one line per occupation, in `estimates/*.txt` format (Section 13). Write the rationale with the numbers, not afterwards.
5. Run `scripts/check_estimates.py` and fix the validation errors. Log every change made after the first pass, with before and after values and the reason, as in `revision_log.csv`.
6. To label another rater's workbook, edit `MODEL` and `MODEL_ID` at the top of `scripts/build_workbook.py`. Then build.

### 12.3 Compare runs

- **Join on `soc_code`.** These are National Employment Matrix codes. Seven of them are NEM aggregates that end in 0 and differ from the detailed OEWS codes another run might use:
  - 13-1020 Buyers and purchasing agents;
  - 13-2020 Property appraisers and assessors;
  - 29-2010 Clinical laboratory technologists and technicians;
  - 31-1120 Home health and personal care aides;
  - 39-7010 Tour and travel guides;
  - 47-4090 Miscellaneous construction and related workers;
  - 51-2090 Miscellaneous assemblers and fabricators.
- **Compare totals by job weight,** that is employment × share, not by averaging shares across occupations.
- **Look first at the ten largest contributors** in Section 10. They carry most of the difference any two runs can have.

## 13. File formats

**`estimates/*.txt`.** One line per occupation: `code|low|median|high|tags|rationale`.

- Shares are in percent.
- Tags are separated by "; ", or the word `None`.
- Lines starting with `#` are comments.
- UTF-8 encoding, LF line endings.

**`Essential_vs_Derived_Demand_by_Occupation_Claude_Opus_5.5_2026-11-16_0809.csv`.** One row per occupation, in BLS order, UTF-8.

| Column | Meaning |
|---|---|
| `soc_code`, `occupation`, `major_group_code`, `major_group` | From BLS Table 1.2 |
| `employment_2025_thousands` | BLS 2025 employment, in thousands |
| `essential_low_pct`, `essential_median_pct`, `essential_high_pct` | The three essential shares, in percent |
| `derived_median_pct` | 100 minus the median |
| `essential_jobs_median_thousands` | Employment × median ÷ 100 |
| `tags` | The tags, primary first, or `None` |
| `rationale` | The explanation |
| `model` | `Claude Opus 5.5 (claude-opus-5-5)` |
| `generated` | Build time, `2026-11-16 08:09 EST (UTC-05:00)` |

**`revision_log.csv`.** One row per occupation changed after the first pass (73 rows).

| Column | Meaning |
|---|---|
| `soc_code`, `occupation` | The occupation |
| `estimates_file` | Which file it is in |
| `changed_at` | When it was changed |
| `employment_2025_thousands` | BLS 2025 employment |
| `low_before`, `median_before`, `high_before`, `tags_before` | The first-pass values |
| `low_after`, `median_after`, `high_after`, `tags_after` | The final values |
| `median_essential_jobs_change_thousands` | The effect on the median total |
| `rationale_changed`, `rationale_before`, `rationale_after` | Whether and how the rationale changed |
| `reason` | Why the row was changed |

The before values come from the first-pass files as first written. The after values were read from the final files with the repository's own loader.

**`.xlsx` workbook**

- **Summary**
  - Headline shares, low/median/high in cells C11:E11, and essential jobs in C13:E13.
  - Counts in rows 15–20.
  - The major-group table in rows 25–47 (total in row 47).
  - Median bands in rows 52–58.
  - Primary mechanisms in rows 62–66.
  - A static list of the top 25 contributors.
- **Occupations**
  - Header in row 4, data in rows 5–835 and totals in row 837.
  - Columns:
    - A–G: SOC code, occupation, major group code, major group, employment, percent self-employed, median wage;
    - H–J: low, median and high, the yellow input cells, stored as fractions;
    - K: derived share, `=1-I`;
    - L–N: essential jobs, `=E*H`, `=E*I` and `=E*J`;
    - O: primary mechanism;
    - P: all mechanisms;
    - Q: rationale.
  - The totals are `SUM` formulas, and the total shares are essential jobs divided by employment.
  - H–J only accept values from 0 to 1, and those three cells turn red if low > median or median > high. The Median column also has a color scale.
- **Method:** built from `METHOD.md` at build time. The 08:09 copy predates the corrections in Section 8.6.
- **BLS hierarchy:** all 1,112 rows, indented, with each summary row reconciled against the sum of its detailed rows.
- **Sources:** where the data came from, the workbook's SHA-256 and the reconciliation note.

The Summary formulas use `SUMIF` and `COUNTIF` over the Occupations columns: by major group code, by median band (`COUNTIFS` and `SUMIFS`) and by primary mechanism.

## 14. File inventory and SHA-256 hashes

README.md and this document are not listed, because the README changes when this document is added.

| File | SHA-256 |
|---|---|
| `Essential_vs_Derived_Demand_by_Occupation_Claude_Opus_5.5_2026-11-16_0809.csv` | `ce1b8598c82b642042e1885264abeeac186dce9933cf09d36fd3f655915d0b35` |
| `Essential_vs_Derived_Demand_by_Occupation_Claude_Opus_5.5_2026-11-16_0809.xlsx` | `9df80f1f8781702a0913677b6c3b09e8e5b956e8d29b965982d2f4ac48ea9bda` |
| `METHOD.md` | `ad7c408785cc9ea6cb5e882536101ae7612c5e53709987e960be94ede455d063` |
| `revision_log.csv` | `1dec13139cf66d03330eb39b90423edb0ae93c7a2911b564f710c0627c0a56f9` |
| `source_data/BLS_EP_occupation.xlsx` | `08ee182f992d972593e64559ce467bad75967d2228f11ad5e2f86d4ec0ee6c00` |
| `source_data/BLS_EP_Table_1.2_2025-35.csv` | `967777ef4cd59dcd1d9aaaa1dc029775d07c0396a0386ef9d56c867d0d2fca11` |
| `estimates/11_management.txt` | `dbe1ef5854904ed245a0f110cefc2ba57d2e212635bebcc29491c3b7b40d01b9` |
| `estimates/13_business_financial.txt` | `07da2cc1ddcb4bf5ec4ce7b6610c02bc8b60e328d2181892b1486713392bc274` |
| `estimates/15_17_19_computer_engineering_science.txt` | `df91967e576335cc8d5b759f7a20499ba4b0f5ce8d198de622b20a158c38a2e1` |
| `estimates/21_23_social_legal.txt` | `db9df17a269cc4470449b3e8dc37e10639fd6fd5951b53f5215a3b7932609468` |
| `estimates/25_education.txt` | `70a184ecb83b56b560cc32c70460a7b7de3587c4f1778df101aa3615554f1166` |
| `estimates/27_arts_media.txt` | `b9013433b5d35a2c5922bf88268c7d2bab0a66334a5394f47edd3778f5818ad2` |
| `estimates/29_healthcare_practitioners.txt` | `40dca0c250f854aa38a55c1f51f5092cc974c1b38eca73f0c5a5ba62769b6e19` |
| `estimates/31_33_healthsupport_protective.txt` | `55788d4f5a16f27db90d2057670833304916c763ad845ff0af37f7c4f81452ba` |
| `estimates/35_37_39_food_cleaning_personal.txt` | `18001d2c6058408973f4d25e8c7c92dd1d9a43c1c858fec2fd3637981670f412` |
| `estimates/41_43_45_sales_office_farming.txt` | `6beea1a59800d0b6e67480cd908f16e6274605683d68496879591345834747cd` |
| `estimates/47_49_construction_repair.txt` | `016ea8a9e2b6328ea9c832932b0055425c8525a8c031052d589a1f5a150e1f21` |
| `estimates/51_53_production_transport.txt` | `5d1ceaa2e00b849db248ac4021c85153f91b23267fb5e991eeb2b812b5a3468e` |
| `scripts/build_workbook.py` | `a1c288e208f2f4e0dc74463104576b37ab8e33afe7f115be962f6328a6f9e625` |
| `scripts/check_estimates.py` | `5cf5a5f7c4ab4825f630667e290e0cba430b698f6152616297028052e311d7e5` |
| `scripts/extract_table_1_2.py` | `0cb1ae1b981c984f38ed15e0d94ee818830a60ccfe8e7ac265061963c06dff84` |
| `scripts/fix_untagged.py` | `7dee2cd6ebac9cba05781c99d2ed9f077df3bfcb473cabfb2d7ddf6d76e91b64` |
| `scripts/list_group.py` | `26d3af66abfe8e3ba3b59a54194910cd402eb357117d8f8dcf048b39fc8b68da` |
| `scripts/load_estimates.py` | `69ad7ca4082aeedab808910c949b9906fccebe8b2d447900ec9eeb27aed5415c` |
| `scripts/xlsx_read.py` | `7ee71cfaa12ef3f4c43da427a3684a7896b092d9e26b9538d9dfa97ec7a5fbd8` |

`scripts/fix_untagged.py` was a one-off. Running it again on the final files changes nothing, because no untagged nonzero median remains.

The `METHOD.md` hash is for the file as corrected at 10:24 and 10:28. Section 8.6 gives the hash of the file as it was at the 08:09 build.

## 15. Timeline

All times are EST on 16 Nov 2026.

| Time | Step |
|---|---|
| 07:42:48 | Request received. The repository held only a one-line README. |
| 07:43:01–07:43:40 | Read the PaidToBeHuman files listed in Section 3.2. |
| 07:44:40 | Command-line download refused (HTTP 403). |
| 07:44:46–07:45:27 | BLS page opened in the built-in browser; its table read and summed in the page. |
| 07:45:32–07:45:33 | Workbook link opened in the browser; file saved to Downloads. |
| 07:45:56 | File identified as an Excel workbook and copied to a scratch folder. |
| 07:46:03–07:46:31 | openpyxl unavailable; standard-library reader written; sheets listed. |
| 07:46:49 | Workbook and reader copied into this folder; SHA-256 recorded. |
| 07:47:02 | Table 1.2 extracted: 831 line items, 170,267.2 thousand. |
| 07:50:03–07:50:24 | Detailed occupations counted per major group; first batch listed. |
| 07:51:41 | `METHOD.md` composed. |
| 07:52:08–08:03:16 | The 12 estimate files composed (Section 7.1). |
| 08:03:32–08:03:42 | Validator written and run: 72 errors. |
| 08:03:55 | The 72 rows listed. |
| 08:04:17 | Consistency fix applied; 0 errors. |
| 08:05:22–08:05:29 | Waiters revised; the rule added to `METHOD.md`. |
| 08:07:30–08:07:48 | Build script written; first build. |
| 08:08:18 | Excel recalculation check of the first build. |
| 08:08:50 | Summary-row reconciliation checked. |
| 08:08:58–08:09:07 | Sources note updated; README written; final build. |
| 08:09:15 | First build confirmed a duplicate (CSV identical apart from the timestamp) and deleted. |
| 08:09:24 | Final build recalculated in Excel; PDF preview exported but not viewed. |
| 08:09:52 | Concentration and mechanism split computed. |
| 08:26:51 | Request for this document. |
| 08:30 onward | Hashes, file times and statistics; first-pass comparison and `revision_log.csv`; this document. |
| 10:24:29 | At the user's request, the `METHOD.md` sentence describing the consistency fix corrected (Section 8.6); this document updated to match. |
| 10:28:56 | At the user's request, `METHOD.md`'s opening line corrected to note the rule added after the first pass (Section 8.6); this document updated to match. |
