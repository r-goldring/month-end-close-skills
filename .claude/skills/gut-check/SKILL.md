---
name: gut-check
description: >
  Pre-upload review of a generated payroll JE Import CSV against the prior 2
  same-skill JEs in NetSuite. Catches sign flips, severance dept misroutes,
  abnormal variances, and JB EBITDA reclass drift before the accountant uploads to
  NetSuite UI Import Assistant. Runs automatically as the last step of every
  payroll skill, OR explicitly via /gut-check {pay-run-folder}. Output: chat
  PASS/WARN/FAIL report + timestamped tab in the backup workbook. Triggers
  on: gut-check, gut check, pre-upload review, payroll review, JE review,
  validate JE, verify payroll, check payroll JE, sign flip check.
---

# Payroll Gut-Check Skill

## When to use

- **Auto:** every payroll skill calls this as a step between "Generate CSV"
  and "Upload to NetSuite UI". the accountant sees the report inline.
- **Explicit:** the accountant runs `/gut-check {pay-run-folder}` to re-run after
  editing the CSV (or upstream Reclasses tab), or to validate a CSV that
  was generated in a previous session.

Both paths converge to the same Python orchestration below.

## Why pre-upload (not post)

The CSV is the source of truth for what Import Assistant will create. Editing
a CSV is trivial; reversing a posted JE under the SoD policy is a multi-step
controller-approval cycle. Catch issues at the cheapest moment — before they
hit NetSuite.

The source of truth for the **prior 2 baseline** is still NetSuite (audit_log
finds tranids; SuiteQL fetches the actual posted lines).

## What it checks

### Tier 1 — aggregate variance (account-family roll-up)
Salaries, Commission, 401K Match, Health Benefits, Other Benefits, Payroll
Taxes, and the 5 liability accounts (231200/01/02/250/350). Flag rule:
`abs(delta) > max(family_floor, family_pct * mean(prior_2))`. Bonus and
Severance skip variance (spiky by nature).

### Tier 2 — per-department variance (every (dept, account) line)
Same families, scaled-down per-dept floors. Also surfaces NEW (dept, account)
combos and DISAPPEARED ones.

### Special checks
- **Canonical sign** — every 5xxxx/6xxxx line must be a debit; every 23xxxx
  line must be a credit. Sign violation → **FAIL**. (Catches the FP&A-found
  scenario where a JB Salary line accidentally lands as a credit.)
- **Severance routing (US)** — every 611250/511175 line must have
  Department = `EBITDA Adjustments`. Missing or mis-routed → **FAIL**.
- **JB EBITDA prior-2 (US)** — every `(JB)`-tagged line in EBITDA Adjustments
  is compared to the same component in prior 2. Threshold: $100 floor / 5%.
  Drift → **WARN**, sign flip → **FAIL** (caught by canonical sign).

## Files

- `scripts/_shared/payroll_gut_check.py` — main module
- `scripts/_shared/skill_configs.py` — per-skill configs + thresholds
- `.claude/skills/_shared/netsuite-queries.md` — SuiteQL templates

## Step-by-step (orchestration)

### Step 1 — Phase 1 setup (no MCP calls yet)

```python
import sys, os
sys.path.insert(0, "scripts/_shared")
from payroll_gut_check import run_gut_check
from skill_configs import detect_skill_from_folder

result = run_gut_check(
    folder_path="Monthly Payroll/Pay Runs/US/04.30.2026/",
    audit_log_path="audit_log.json",
    suiteql_runner=None,  # forces phase-1-only
)
```

`run_gut_check` with `suiteql_runner=None` returns one of:

- `{"status": "no_priors", ...}` — no prior JE# found in audit_log for this
  skill. Skip variance and emit an INFO note. (Run only the canonical-sign
  and severance/JB checks if applicable.)
- `{"status": "needs_suiteql", "candidates": [...], "validate_sql": "...",
  "config": {...}, "current_lines": [...], ...}` — proceed to Step 2.

### Step 2 — Validate prior candidates against NetSuite

Call:
```
mcp__claude_ai_NetSuite__ns_runCustomSuiteQL(sqlQuery=result["validate_sql"])
```

The query is pre-built and filters voided/non-posting/reversed candidates out.

### Step 3 — Filter to top 2 valid priors

```python
from payroll_gut_check import filter_validated_priors, build_line_fetch_sql

sieved = filter_validated_priors(
    candidates=result["candidates"],
    validation_rows=<rows from Step 2 SuiteQL>,
    current_trandate=result["current_trandate"],
    cadence_days=result["config"]["cadence_days"],
    n=2,
)
priors = sieved["priors"]               # list of validated priors (up to 2)
baseline_warnings = sieved["warnings"]  # stale-baseline / fewer-than-2 notes
```

### Step 4 — Fetch lines for each valid prior

For each prior in `priors`, run:
```
mcp__claude_ai_NetSuite__ns_runCustomSuiteQL(
    sqlQuery=build_line_fetch_sql(prior["tranid"], result["config"]["subsidiary"])
)
```

Convert each result via `normalize_ns_lines(rows)` to match the line-dict
shape the analyzer expects.

### Step 5 — Analyze, format, write tab

```python
from payroll_gut_check import (analyze, format_chat_report,
                                write_workbook_tab, normalize_ns_lines)

prior_lines_by_je = {
    p["tranid"]: normalize_ns_lines(<rows for that prior>)
    for p in priors
}

gc_result = analyze(result["current_lines"], prior_lines_by_je, result["config"])

header_meta = {
    "skill": result["skill"],
    "pay_run_label": result["folder_name"],
    "csv_basename": os.path.basename(result["csv_path"]),
    "priors": priors,
    "baseline_warnings": baseline_warnings,
    "totals": {
        "line_count": len(result["current_lines"]),
        "debit": round(sum(l["debit"] for l in result["current_lines"]), 2),
        "credit": round(sum(l["credit"] for l in result["current_lines"]), 2),
        "currency": result["config"]["currency"],
    },
}

# Locate the backup workbook for this pay run, then append the gut-check tab
import glob
wb_pattern = result["config"]["workbook_glob"].replace(
    "{folder_name}", result["folder_name"])
wb_matches = [m for m in glob.glob(
    os.path.join(os.path.abspath(<folder>), wb_pattern))
    if not os.path.basename(m).startswith("~")]
if wb_matches:
    tab_name = write_workbook_tab(wb_matches[0], gc_result, header_meta)
    header_meta["workbook_tab_path"] = (
        f"{os.path.basename(wb_matches[0])} -> '{tab_name}'")

chat_report = format_chat_report(gc_result, header_meta)
print(chat_report)
```

### Step 6 — Gate the upload instruction on FAIL

If `gc_result.has_fail`:
> **DO NOT UPLOAD.** Fix the CSV (or Reclasses tab) per the FAIL findings,
> then re-run `/gut-check {pay-run-folder}` before uploading.

Otherwise:
> **Safe to upload.** WARNs are advisory; review and proceed when ready.
> Lists -> Import Assistant -> Transactions -> Journal Entry -> upload CSV.

## Trigger detection

When invoked as `/gut-check {arg}`:
- If `{arg}` looks like a path containing `Monthly Payroll/Pay Runs/...` → that's the folder.
- If `{arg}` is a short form like `04.30 US` or `2026-04 NL` → resolve via `SKILL_CONFIGS`:
  - Match the country word (US / Canada / Germany / Netherlands / Poland / UK / Uruguay) to `folder_root`.
  - Append the date token to the root.
- If `{arg}` is a `JE#####` tranid → **Phase 2 (deferred)**: post-upload mode that compares
  posted JE vs the local CSV. NOT IMPLEMENTED in v1; halt with a note.

When called from inside a payroll skill's flow: pass the same `folder_path`
that the skill just used to generate the CSV.

## What the chat output looks like

```
=========================================================================
GUT-CHECK REVIEW - 04.30.2026 (us-payroll)
CSV: 04.30.2026 US Payroll JE Import.csv
     146 lines, USD 1,710,749.31 debits / 1,710,749.31 credits (balanced)
Compared against: JE##### (2026-04-15)
  ! Only 1 valid prior(s) available (needed 2); comparisons will be best-effort.
=========================================================================
TIER 1 (aggregate):  4 PASS, 1 WARN, 0 FAIL
TIER 2 (per-dept):   0 PASS (suppressed), 6 WARN, 0 FAIL
SPECIAL checks:      5 PASS, 0 WARN, 0 FAIL, 0 INFO
=========================================================================

FAIL (0)

WARN (7)
  T1  Salaries up 6.4% vs prior 2 mean (delta $X,XXX.XX)
       -> Verify expected business change (new hire, comp adj, bonus run)
  T2  Engineering 611100 up 8.1% in Engineering
  ... [more]

PASS (9) [collapsed - full detail in workbook tab]

Workbook tab: 04.30.2026 US Payroll Backup.xlsx -> 'Gut-Check 2026-05-05 14_23'

==> SAFE TO UPLOAD: WARNs are advisory. Review and proceed when ready.
=========================================================================
```

## Phase 2 (deferred) — Optional post-upload mode

After v1 is stable, add `/gut-check JE#####` mode that:
1. Pulls the posted JE from NetSuite by tranid.
2. Compares it against the CSV that was uploaded (stored alongside in the
   pay-run folder).
3. Catches Import-Assistant formatting drift or post-upload manual edits in
   the NS UI.

Not built yet. If a tranid is passed in v1, halt with: "Post-upload gut-check
mode is Phase 2 / not yet implemented; pass a pay-run folder path instead."

## Edge cases

- **No priors in audit_log** → emit INFO; run only canonical-sign and severance/JB checks.
- **Voided/reversed prior** → automatically filtered by the validation SQL; skill reaches further back if needed.
- **Stale baseline** (gap > 1.5 × cadence_days) → emit WARN with explicit note; still use it.
- **Re-run on same CSV** → workbook tab gets a new timestamped name; never overwrites.
- **CSV not found** → halt with the expected glob pattern.
- **Backup workbook missing** → run anyway, skip the tab write, note it in chat.
