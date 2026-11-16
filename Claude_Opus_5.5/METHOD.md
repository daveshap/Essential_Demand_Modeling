# Method: essential vs derived demand, by BLS occupation

Prepared by Claude Opus 5.5 (model ID `claude-opus-5-5`, Anthropic), running in Claude Code, 16 Nov 2026. This file was written before any occupation was scored, and it is the standard every row was scored against, with one exception: the rule that a nonzero median must name what is essential (under "Low, median and high") was added after the first pass.

## The question

Of all the jobs that exist today, how much is Essential Demand (the human is the point) and how much is derived demand (labor wanted only as a means to an outcome, as Alfred Marshall described it in 1890)?

## The baseline

- BLS Employment Projections, Table 1.2, *Occupational projections, 2025–35, and worker characteristics, 2025* (the National Employment Matrix). Downloaded from bls.gov on 16 Nov 2026 as `occupation.xlsx`; the copy is in `source_data/`.
- Every one of the 831 detailed ("Line item") occupations is scored. Together they cover 170,267.2 thousand jobs in 2025. The 281 summary rows are kept for reference and never added in.
- The matrix counts jobs, including self-employment, not full-time equivalents.

## What each percentage means

**Essential share** is my estimate of the share of an occupation's current labor (job-weighted, which is roughly its hours) that buyers, audiences or responsible institutions would still pay a human to do, at roughly today's prices, if a disclosed machine alternative delivered the same functional result at equal or lower cost. Everything else in the occupation is derived demand.

This is the substitution test from *Paid to Be Human*: hold the result constant, tell the buyer it is a machine, and ask what would be lost by removing the human.

What it is not:

- It is not a forecast of when or whether automation happens.
- It is not the chance that the job title survives. One essential task does not keep a whole job; only the hours of that task are counted.
- It is not an income estimate.

In the white paper's terms it is a share of today's hours that would survive the substitution test, with today's level of demand held constant.

## How each row is estimated

Two questions, for every occupation:

1. Who buys this work, and what share of those buyers want the human, not just the result?
2. In that work, what share of the worker's time is the human part (the session, the performance, the decision, being there), and what share is support work (paperwork, preparation, logistics, analysis) that a machine could take over while the human stays?

The estimate is roughly the product of the two, adjusted by judgment. It is a judgment, not a measurement: no survey or task-time data were used.

## What counts as essential

Any reason the human is the point, nonfungible, or worth a premium. The four labels from *Paid to Be Human* are used as tags where they fit, but a row does not need one:

- **Presence**: the buyer wants an actual person there (care, company, live performance, hospitality where the person is the point).
- **Provenance**: human making or authorship is part of what is bought (art, craft, authored work, human achievement).
- **Accountability**: a person must answer for, authorize, represent or bear consequences (judges, elected officials, final decisions, coercive public authority). Counted only for the hours of meaningful accountable work, never for the whole job around it.
- **Affinity**: a particular person, or a continuing relationship with one.

## What does not count

- Work people do only because machines can't do it yet ("safe for now"). That is derived demand that hasn't been automated.
- People skills that today's job requires, when the buyer would accept an equally good machine.
- Work that is important or heroic but valued for its outcome (rescue, surgery, firefighting). Even heroic work can be derived demand.
- A rule that a named person sign, beyond the time a meaningful review takes.
- Customers' wish to be among other customers. That is demand for each other's company, not for the worker's labor.

## Low, median and high

- **Low**: a strict reading. Only clear, durable human-specific demand, counted on the narrowest set of hours.
- **Median**: my central judgment.
- **High**: a generous but defensible reading. It takes in softer premium segments and more of the relationship time. It never takes in demand that exists only because machines can't do the work yet.

A nonzero median must name what is essential (and carry at least one tag). Where a rationale names nothing essential, the median is 0%, and any high above 0% (at most 3%) is a general allowance for segments I may have missed, not a claim. This rule was added after the first pass, when the validation script found 72 rows with a 0.5–2% median and no tag. Seventy of them named nothing essential: their low and median were set to 0% and their high capped at 3%. Industrial production managers (11-3051) did name a sliver (accountable safety decisions), so that row kept its values, was tagged Accountability and had its rationale reworded to name the sliver. Costume attendants (39-3092) named nothing, so the rule would have set its median to 0%; instead I re-judged it while applying the fix, wrote in a sliver (performers who want a trusted dresser backstage), tagged it Affinity and set it to 0% low, 1% median and 4% high. Every change is listed in `revision_log.csv`. (The account of the fix above, and the file's opening line, were corrected on 16 Nov 2026, after the 08:09 build. The earlier wording, which that workbook's Method sheet still shows, said both rows had named a sliver, and the opening line made no exception for this rule.)

## Scale anchors (median), used to keep rows consistent

| Median | Meaning | Typical examples |
|---|---|---|
| 0% | The buyer wants the output; no one would pay more for a human | production, freight, IT, clerical, most construction |
| 0.5–5% | A thin segment pays for the human, or a small accountable sign-off | cashiers, customer service, accountants |
| 5–20% | A real but minority premium segment, or a small human core in a large functional job | waiters, real estate agents, nurses |
| 20–50% | A mixed bundle in which the human part is a large share | teachers, hairstylists, personal care aides |
| 50–80% | The human is mostly the point; support work is derived | therapists, clergy, performers, fitness instructors |
| 80–100% | The human is the product | athletes, elected representatives |

## Independence

Before starting I read the *Paid to Be Human* white paper, whose text reports the prior audit's aggregate results (1.59%, 5.47% and 14.85% of covered employment). I did not open that audit's row-level coefficients, and none were used here. Each row was judged on its own merits, and the totals are sums of the rows, not targets.

## Columns in the spreadsheet

- SOC code, title, major group, 2025 employment, percent self-employed and median wage come straight from BLS Table 1.2.
- Low, Median and High essential shares are mine. Derived share is 100% minus the median.
- Essential jobs (thousands) are the employment multiplied by each share. They are job-weighted scenario equivalents, not a count of jobs that will survive.
- Tags name the kind of essential demand, if any.
- Rationale explains what part is essential, why the rest is derived, and what drives the low and high.
