# Flux Investigation Playbook

How to investigate and explain month-over-month variances well enough that they
survive a close meeting. This is the *thinking* behind the `flux-analysis-workbook`
skill — the discipline that turns a raw variance into a defensible comment.

## The bar: what a "sufficient" comment looks like

A flagged variance is not explained until the comment names three things:

1. **What** — the specific driver (a named vendor, a specific event, a one-time true-up), not a category ("higher software costs").
2. **How much** — the dollar amount attributable to that driver, reconciled to the total variance.
3. **Why** — the business reason (new contract, annual renewal landed this month, catch-up of a prior-month accrual, headcount change).

Weak: "Professional fees up this month."
Strong: "Professional fees up ~$X because the annual audit fee posted in March (vs. amortized in prior years); one-time, no recurring impact."

## Read the room — who asks what

Different reviewers probe differently. Anticipate the question before the meeting:

| Reviewer role | What they care about | What to pre-answer |
|---|---|---|
| **CFO** | Is this reasonable vs. budget / prior expectation? Anything that "doubles" unexpectedly? | Tie the variance to budget; flag and explain anything that looks like a step-change. Software-by-department coding gets special scrutiny. |
| **Controller / VP Accounting** | Was this expected? Is it coded correctly? | "Expected / not expected" framing on every material line; confirm GL + department coding. |
| **Assistant Controller** | The mechanics — accruals, prepaid amortization, reclasses | Have the supporting JE/accrual logic ready; expect real-time fixes during the call. |
| **FP&A** | Does actual match the plan, by department? | Department-level variance vs. budget; call out reclasses that move spend between departments. |

Reference reviewers by **role**, never by name, in any committed artifact.

## Account-by-account investigation patterns

### Revenue (4xxxx)
Confirm with the revenue owner before commenting. Most "variances" are timing of recognition, not real movement. Watch for deferred-revenue releases and one-time true-ups.

### COGS (5xxxx)
- **Contractors (e.g., 5113xx):** driven by headcount and project ramp. Tie a spike to a specific contractor or project. A drop usually means a contract ended or a bill landed in the wrong month.
- **Amortized software COGS:** should be smooth. A jump means a new capitalized item started amortizing or a catch-up posted.

### Operating expense (6xxxx)
- **Software (67xxxx):** the most-scrutinized line. Every zero-to-balance or balance-to-zero move needs a comment. The usual culprits: an annual renewal landed this month, a subscription got recoded to a different department, or a new tool started.
- **Professional fees (legal / accounting):** lumpy by nature. Tie to the specific engagement (audit, a financing, a legal matter). Distinguish recurring retainer from one-time project fees.
- **Payroll & benefits:** tie to headcount changes, bonus/commission cycles, and employer-tax seasonality (e.g., front-loaded social-tax caps early in the year).

## Discipline that prevents re-work

1. **Three-month lookback.** Always compare against a 3-month trend, not just the immediately prior month. A "spike" is often just last month being abnormally low.
2. **Zero-balance vendors get a mandatory comment.** Any vendor going from a balance to $0 (or $0 to a large balance) will be asked about every time. Comment proactively.
3. **Reclasses move the variance, not the total.** When you recode a vendor to a different department, both departments flux. Comment on both sides so the net is obviously zero.
4. **Reconcile the comment to the number.** If the variance is $X and your named drivers add to less than $X, you are not done — there is an unexplained remainder.

## Escalate vs. silently accept

- **Silently accept** (no comment needed): immaterial variances below your materiality threshold AND below the percentage threshold. Don't manufacture commentary for noise.
- **Comment** (expected, explained): material variance with a clear, named, one-time or expected driver.
- **Escalate** (flag to the controller before the meeting): a material variance you *cannot* explain, a coding error you found while investigating, or anything that suggests a missed accrual or a posting in the wrong period.

## Tuning for your company

The thresholds, account ranges, and reviewer roles above are illustrative. Set your own materiality thresholds (absolute dollar AND percentage), map the account ranges to your chart of accounts, and adjust the reviewer table to your close team's actual roles.
