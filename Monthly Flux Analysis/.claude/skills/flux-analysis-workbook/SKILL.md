---
name: flux-analysis-workbook
description: >
  Drive the monthly Flux Analysis close cycle for Acme Corp end-to-end:
  (1) pull the 4 detail reports into the Flux Pivot Template, (2) suggest
  accruals based on vendor-level history, (3) build the accrual JE CSV after
  the accountant approves, (4) consume the accountant's IS/BS drop-ins to assemble the full flux
  workbook with column-H variance commentary, and (5) handle post-meeting
  reclass/prepaid JE work. Use this skill whenever the accountant mentions: flux analysis,
  monthly flux, flux workbook, variance analysis, month-end close, MoM
  comparison, income statement flux, balance sheet flux, close workbook, close
  package, COGS detail, contractor detail, professional fees detail, software
  detail, pivot template, refresh pivots, flux pivot, accrual suggestions,
  accrual builder, build the accrual CSV, or asks to pull/build/update/refresh
  the flux file. Also use when reviewing or explaining variances, generating
  first-pass comments, or preparing for a flux meeting. This skill connects to
  NetSuite via MCP tools (ns_runReport, ns_runCustomSuiteQL) to pull live GL
  data.
---

# Flux Analysis Workbook Skill — v2

## Workflow overview (the actual close cycle)

The skill drives a four-pass close cycle. Each pass produces a labeled artifact under `Monthly Flux Analysis/{YYYY}/{YYYY-MM}/`:

| Pass | Trigger | Output | What runs |
|---|---|---|---|
| **v1 — pre-accrual** | the accountant asks to start the close | `Flux Workbook/{YYYY-MM} Flux Pivot Template.xlsx` + `State/accrual_candidates.json` | Pull 4 detail reports, populate pivot template, tie-back to NS, suggest accruals in cols M/N/O |
| **v1 confirmed** | the accountant approves suggestions | `JE Imports/{YYYY-MM} Accruals JE Import.csv` | `build_accrual_csv.py` reads approved JSON, writes upload-ready CSV |
| **v2 — post-approval** | Accrual posted + the accountant drops IS/BS in | `Flux Workbook/{YYYY-MM} Flux Analysis (v2 initial).xlsx` | Re-pull detail reports, assemble IS/BS + 4 detail tabs, write first-pass column-H comments |
| **v3 — post-meeting** | the VP/team comments received | `Flux Workbook/{YYYY-MM} Flux Analysis (v3 post-meeting).xlsx` | Address comments in column H, build follow-up JE CSVs, prepaid schedule items XLSX |
| **final** | All post-meeting JEs posted | `Flux Workbook/{YYYY-MM} Flux Analysis (final).xlsx` | Sign-off copy |

Each pass reads the previous version's notes from `State/flux_notes.json` and the `accrual_candidates.json`. **User-edited notes (`edited_by: ryan`) are never overwritten.**

## Folder layout (mandatory — backfilled May 2026)

```
{YYYY-MM}/
├── Flux Workbook/                              # Main deliverables + transcript
│   ├── Income Statement Drop.xlsx              # the accountant's IS export
│   ├── Balance Sheet Drop.xlsx                 # the accountant's BS export
│   ├── {YYYY-MM} Flux Pivot Template.xlsx      # 4-tab pivot template (v1)
│   ├── {YYYY-MM} Flux Analysis (v2 initial).xlsx
│   ├── {YYYY-MM} Flux Analysis (v3 post-meeting).xlsx
│   ├── {YYYY-MM} Flux Analysis (final).xlsx
│   └── {Meeting transcript}.md
├── JE Imports/                                 # NetSuite CSV uploads
├── Schedules/                                  # Prepaid items, candidates, support xlsx
├── State/                                      # Auto-managed JSON
│   ├── flux_notes.json
│   ├── accrual_candidates.json
│   └── pending_jes.json
├── _cache/                                     # NetSuite report JSON
└── Supporting/                                 # BillFlow export, other team inputs
```

Run `scripts/flux-analysis-workbook/migrate_folder_layout.py {YYYY-MM} --apply` if a month folder is still in the legacy flat layout.

## Inputs the accountant provides

| Pass | File | Where |
|---|---|---|
| v2 onward | `Income Statement Drop.xlsx` | `{YYYY-MM}/Flux Workbook/` |
| v2 onward | `Balance Sheet Drop.xlsx` | `{YYYY-MM}/Flux Workbook/` |

**The skill no longer pulls IS or BS itself** — the accountant can do this in seconds from NetSuite and the formatting comes out clean. If the drop-ins are missing when the skill needs them, fail fast and prompt the accountant with the NetSuite report IDs to run manually.

## What this skill does NOT do

- Does NOT pull Income Statement or Balance Sheet from NetSuite (the accountant handles).
- Does NOT auto-write the accrual JE CSV — the accountant reviews + approves first.
- Does NOT post JEs to NetSuite via MCP (CSV upload only per CLAUDE.md rule 1).
- Does NOT use BillFlow export as a primary signal — false positives are too high. Pivot template vendor-level history is the trusted signal.

## Critical: Subsidiary selection

**Always use subsidiary ID `-2` (Acme, Inc. Consolidated)** for every `ns_runReport` call. Returns all subs FX-converted to USD. Sub `-1` returns empty shells; sub `2` is US-only and misses NL/UK/UY/CA activity.

## Pass v1 — Pre-accrual (pivot template + accrual suggestions)

### Step 1: Pull 4 detail reports

| Report | ID | Title | Account filters |
|---|---|---|---|
| COGS | 537 | Income Statement Detail COGS - BW | 511400, 511425, 511450, 511510, 511520, 511550, 511600 |
| Contractors | 540 | GL Detail Contractor Payroll - BW | 511370, 611700 |
| Professional Fees | 542 | GL Detail Professional Fees - BW | 651100, 651101 |
| Software | 721 | GL Detail Software Expense - BW | 671000, 671100 |

```python
ns_runReport(
    reportId=<id>,
    dateFrom="<first-day-of-M-2>",
    dateTo="<last-day-of-M>",
    subsidiaryId=-2
)
```

3-month window: `[M-2, M-1, M]`. Cache the report output to `{YYYY-MM}/_cache/{id}.json` so subsequent passes can re-use it.

### Step 2: Parse + populate

Use `scripts/flux-analysis-workbook/populate_pivot_template.py`. Key parser features:

- **COGS (537)**: filter to 7 sub-accounts, negate amounts (IS report shows expenses as negative).
- **GL Detail (540/542/721)**: filter to target accounts at the **row level** (the row's Account field must contain a target account number). Reject rows whose Entity name is in the Brex categorical denylist (`Lodging`, `Airfare`, `Meals`, `Transportation`, `Taxi`, `Hotel`, `Uber`) — these are NetSuite report attribution bugs from multi-line Brex JEs.
- **Net amount column**: `_total` key (preferred) or unnamed `""` key (fallback).
- **Date parsing**: `email.utils.parsedate_to_datetime`.
- **Vendor name resolution**: `COALESCE(Entity (Line): Name, Name)`.
- **Account paths**: replace `\x01` with `:`.

### Step 3: Tie-back verification

After populating, verify against the consolidated IS line:

| Detail report | Expected match on IS |
|---|---|
| 537 COGS | "Total - 511000 - Cost of Goods Sold" |
| 540 Contractors | sum("511370 - COGS - Contractor Payroll" + "611700 - Contractor Payroll") |
| 542 Prof Fees | "651100 - Professional Fees" |
| 721 Software | "671100 - Software Subscriptions" |

If the accountant has dropped the IS xlsx in already, tie-back uses that as primary. Otherwise, fall back to a SuiteQL query against `transactionline` joined to `account` and `subsidiary`, including all 6 Acme Corp subsidiaries (US + Canada + Netherlands + Poland + UK + Uruguay; **exclude** Acme Corp Elimination — its balance on these non-IC P&L accounts should be $0).

**Match tolerance**: $0.50 per period. Fail aborts save. Note: tie-back compares (parsed rows + dropped rows from Brex-category filter) so the false-positive filter doesn't break the reconciliation.

### Step 4: Strip pivot cache

After `wb.save(path)`, call `strip_pivot_cache(path)` to empty pivotCacheRecords and force Excel to rebuild on next Refresh All. Eliminates copy-paste crashes on detail tabs.

### Step 5: Build accrual suggestions

Call `build_accrual_suggestions.build_suggestions(wb, month_dir, ...)`. Writes:
- Columns M/N/O on each detail tab's static F-L table (Accrual Suggestion / Confidence / Reason)
- `{YYYY-MM}/State/accrual_candidates.json`

Logic codified in `references/accrual-detection-rules.md`. Cross-references **both** last month's `Accruals JE Import.csv` **AND** NetSuite's actual posted JE (external ID `{YYYY-{MM-1}}-ACCRUALS`) — the accountant sometimes edits accruals directly in NS, so the CSV alone can be stale.

### Step 6: Hand to the accountant

Tell the accountant: the pivot template is ready at `{YYYY-MM}/Flux Workbook/{YYYY-MM} Flux Pivot Template.xlsx`. Open it, Refresh All, review accrual suggestions in cols M/N/O, edit/approve in the workbook or `accrual_candidates.json`, then run `/build-accrual-csv` (or just ask Claude to build it).

## Pass v1 confirmed — Build the accrual CSV

After the accountant flips approvals to `true` in the JSON (or edits in the workbook), run `scripts/flux-analysis-workbook/build_accrual_csv.py {YYYY-MM}`. Writes `{YYYY-MM}/JE Imports/{YYYY-MM} Accruals JE Import.csv` in NetSuite Import Assistant format.

the accountant uploads the CSV via NetSuite UI → Pending Approval → controller approves.

## Pass v2 — Post-approval (full workbook with variance commentary)

### Step 1: Re-pull detail reports

The pending JEs posted, so the data has changed. Re-pull all 4 detail reports (fresh cache).

### Step 2: Verify IS/BS drop-ins

Read `{YYYY-MM}/Flux Workbook/Income Statement Drop.xlsx` and `Balance Sheet Drop.xlsx`. Validate:
- Period columns present (M-2, M-1, M)
- Row count plausibility (90–140 IS rows; 200–250 BS rows)
- Column G flag formula present

If missing, fail fast with a clear pointer to the NetSuite report IDs.

### Step 3: Assemble the full workbook

Combine: IS + BS (from drops) + 4 detail tabs (from re-populated pivot template) → `{YYYY-MM} Flux Analysis (v2 initial).xlsx`. Strip pivot cache. Re-run tie-back.

### Step 4: Write variance commentary (column H)

For every Y-flagged row on IS + BS:

1. Load `State/flux_notes.json`. If the accountant has already edited this row (`edited_by: ryan`), **skip** — preserve his note.
2. Otherwise, consult `references/comment-recipes.md` for the per-account template.
3. Drill the data: cached detail reports + `ns_runCustomSuiteQL` for accounts not covered by the 4 detail reports (revenue, payroll, BS lines).
4. Compose a 1-line comment per the recipe: **named vendor in plain English** + dollar amount + driver. Cite vendors from `references/known-vendors.md` (never bill numbers alone).
5. Write to column H AND persist to `State/flux_notes.json` with `edited_by: claude`.

**Comment rules** (enforced):
- ASCII only (per CLAUDE.md rule 7).
- Arial 8pt.
- Subtotals (Total-rows, Net Ordinary Income, etc.): leave blank.
- Threshold for flagging: `>$10K absolute AND >10%` (already in column G formula).

Read `references/vp-cfo-comment-patterns.md` BEFORE writing — it covers what gets pushback and how to phrase.

### Step 5: Detail-tab static-table notes (column L)

For each vendor on the Software / COGS / Contractors / Prof Fees static F-L table with abs MoM variance > $200:
- Same rules as column H (consult recipes, plain-English vendor, persist to sidecar).

## Pass v3 — Post-meeting

After the accountant shares the v3 workbook (Gemini transcript + the VP's threaded comments):

1. Read the transcript markdown and the VP's threaded `[Threaded comment]` annotations on cells.
2. For each the VP comment, drill the underlying data and write a response in chat (NOT in the cell — the accountant can paste).
3. Generate any post-meeting JE CSVs (reclasses, prepaid moves, accruals). Common categories:
   - Prepaid reclass for future-event spend (Customer Conference, conferences)
   - Account reclass (Software → COGS, T&E re-routing)
   - Department reclass
   - Additional accruals identified in the meeting
4. Generate prepaid schedule items XLSX if controller needs to set up amortization schedules.
5. Update `State/pending_jes.json` with what's outstanding.

## Variance investigation references (load BEFORE writing comments)

| File | Purpose |
|---|---|
| `references/comment-recipes.md` | Per-account-family templates for column H |
| `references/known-vendors.md` | Vendor → plain-English identifier map (BILL-UYMED-001 = Uruguay Medical Vendor Uruguay medical, etc.) |
| `references/recurring-themes.md` | Patterns that repeat across closes (IT Hardware Partner arrears, Q1 transfer pricing, etc.) |
| `references/vp-cfo-comment-patterns.md` | What gets pushback in the flux meeting + how to phrase to preempt |
| `references/accrual-detection-rules.md` | Rule set the accrual builder applies |
| `references/flux-investigation-playbook.md` | Per-account drill paths (mature) |
| `_shared/netsuite-queries.md`, `_shared/subsidiary-constants.md` | Cross-skill references |

## Notes persistence — `State/flux_notes.json`

All manual annotations (the accountant's edits to column H, column L, accrual approvals) survive re-pulls because the sidecar tracks them by row label / vendor name, not by row index. See `scripts/flux-analysis-workbook/notes_sidecar.py`.

**Merge rule**: on any re-pull, user-edited notes always win. Claude-generated comments refresh against current data, but anything marked `edited_by: ryan` is preserved unless `force=True`.

## Pending JE registry — `State/pending_jes.json`

Tracks which JEs are blocking close progression:
```json
[
  {"category": "accruals", "csv": "2026-05 Accruals JE Import.csv", "status": "uploaded|approved|posted", "expected_impact": {"511400": 377091}},
  {"category": "bonus", "csv": null, "status": "waiting", "owner": "the accountant", "note": "Waiting on SPIFF file from the FP&A Lead"}
]
```

Read during v2 to auto-tag SKIP_PENDING rows. Updated as JEs flow through approval.

## Subagent usage — read directly, don't transcribe

**Workbook comparisons (v1 vs v2 diff, stale-comment checks, row-by-row reads): do these directly via `Bash + openpyxl`. Do NOT spawn a subagent to extract and return the data.**

Why: when an agent is asked to transcribe 100+ rows of file data into structured output (Python dict, JSON, table), it pattern-completes after the first ~20 rows instead of reading every one. Observed failure rate on this kind of task: roughly 5-15% — high enough to require manual verification of every output, which defeats the delegation. Direct `openpyxl` reads are deterministic and faster anyway.

**NetSuite variance investigation: call `ns_runCustomSuiteQL` directly. Summarize in your own response.** The MCP returns deterministic data — but if a subagent is asked to "deep-dive NS and report back," its summary of long line-level results can drift the same way.

**When subagents DO add value here:**
- Independent parallel research (e.g., one agent reading the Gemini transcript while another scans the audit log)
- Open-ended synthesis across many sources where the answer isn't a row dump
- Risky / destructive operations that benefit from isolation

**Rule of thumb:** if the deliverable is "extract rows X-Y from a file" or "list every transaction matching a filter," don't delegate. If the deliverable is "figure out why this account moved and recommend a comment," delegating to an agent with NS MCP access is fine — but ask it to write a short narrative, not return a raw data dump.

If you do need an agent to return structured data, end the prompt with: *"After extraction, re-read 3 random rows from the source and confirm they match what you returned. If any mismatch, abort and re-extract."*

## Common issues

| Issue | Solution |
|---|---|
| Copy-paste crash on Software tab in Excel | `strip_pivot_cache` not called after openpyxl save — must run before the accountant opens |
| Pivot misses rows on refresh | Same — cache out-of-sync with data sheet |
| Lodging / Airfare entity appears in Software tab | Brex multi-line JE quirk — `parse_gl_detail` denylist now filters these. Confirm `populate_pivot_template.py` is calling with `warnings=[]` and surface any warnings to the accountant |
| Tie-back fails | Data sheet missing rows OR including extras (e.g., Brex categorical). Cross-check via SuiteQL on consolidated subs |
| Pivot refresh-on-open doesn't trigger | `pivotCacheDefinition.xml` needs `refreshOnLoad="1"` — `strip_pivot_cache` sets this |
| Manual annotations got wiped | Re-pull was run before `notes_sidecar.scan_workbook_user_edits()` captured the accountant's edits — fix is to always sync the sidecar with workbook before populating |
| COGS amounts have wrong sign | COGS IS report shows expenses as negative — negate when parsing |
| GL Detail reports return empty | Verify `subsidiaryId: -2` (Acme, Inc. Consolidated), not 1, -1, or 2 |
| `\x01` in account paths | Report uses `\x01` as hierarchy separator — replace with `:` |
| "- No Entity -" in Final Name | Skip; fall back to Name field |

## File naming convention

```
{YYYY-MM} Flux Pivot Template.xlsx           # v1 pre-accrual deliverable
{YYYY-MM} Accruals JE Import.csv             # v1-confirmed CSV
{YYYY-MM} Flux Analysis (v2 initial).xlsx    # post-accrual workbook
{YYYY-MM} Flux Analysis (v3 post-meeting).xlsx
{YYYY-MM} Flux Analysis (final).xlsx
```

All paths relative to `Monthly Flux Analysis/{YYYY}/{YYYY-MM}/Flux Workbook/`.
