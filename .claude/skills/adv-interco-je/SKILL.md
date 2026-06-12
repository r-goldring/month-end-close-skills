---
name: adv-interco-je
description: >
  Build Advanced Intercompany Journal Entry (Adv IC JE) CSV files for NetSuite import.
  Reclasses expenses booked under the wrong subsidiary to the correct one. Supports
  every direction confirmed in production: US to NL/UK/UY/CA (forward), CA to US,
  CA to UK/NL/UY (reverse from Canada), NL to UK (and other foreign-to-foreign pairs).
  Use this skill whenever the user mentions: intercompany journal entry, IC JE, adv
  interco JE, reclassing to or from Netherlands/UK/Uruguay/Canada/US, accounts
  681100/681110/681130/681140/681150, monthly intercompany reclass, or moving
  expenses from one entity to another. Three-phase workflow: (1) IC-flagged scan
  on each requested subsidiary (lines hitting 681xxx accounts on that sub),
  (2) optional IS sanity check to confirm Phase 1 caught every 681xxx hit,
  (3) generate one upload-ready CSV per (source, target) pair after confirmation.
  Skill ONLY reclasses 681xxx-flagged items; plain wrong-sub expense activity
  is out of scope (handle via je-error-fix or upstream recoding instead).
  Always use this skill for intercompany reclass work.
---

# Advanced Intercompany Journal Entry Skill

## Overview

Each month, vendor invoices and expense JEs land on the wrong subsidiary's books. This skill pulls those transactions from NetSuite, surfaces them for confirmation, then generates upload-ready Adv IC JE CSV files — one per (source sub, target sub) pair that has activity.

**Direction families supported (all confirmed in production):**
- **US to foreign** — US to NL, UK, UY, CA. Driven by 681100/681110/681140/681150 hits on US books.
- **Foreign to US** — CA to US confirmed (681130 on Canada). NL/UK/UY to US possible if 681130 appears on those subs.
- **Foreign to Foreign** — CA to UK, CA to NL, CA to UY (681140/681100/681150 on Canada). NL to UK confirmed (681140 on NL). Same pattern works for any (source, target) foreign pair.

**Out of scope:**
- Poland reclasses.
- Customer-payment wrong-sub reclasses (different 4-line shape — see "Customer Payment Wrong-Sub Reclass" at the bottom; trigger manually, not via this skill).

**References:**
- Subsidiary paths and IC line constants matrix: [`../_shared/subsidiary-constants.md`](../_shared/subsidiary-constants.md)
- SuiteQL query patterns: [`../_shared/netsuite-queries.md`](../_shared/netsuite-queries.md)
- Check-and-balance pattern: [`../_shared/check-and-balance.md`](../_shared/check-and-balance.md)
- Approval / CSV-only policy: [`../_shared/approval-required.md`](../_shared/approval-required.md)

**CSV-only policy:** This skill never calls `ns_createRecord`. It writes CSVs that the accountant imports through the NetSuite UI (Lists → Import Assistant → Adv IC JE), so each JE lands in the controller's Pending Approval queue.

---

## The 681xxx account family — read this first

The 681xxx series is **named by the TARGET subsidiary**, regardless of which sub the line sits on. Same number on US books or Canada books — the account identifies "expense owed to {target}", not where the source sub is.

| Account | Target | Used on which sub's books |
|---------|--------|----------------------------|
| 681100 | Netherlands | source (e.g., US, CA) |
| 681110 | Canada | source |
| 681130 | US | source — used on Canada for CA → US (JE#####, JE#####, JE#####) |
| 681140 | UK | source — used on US, NL, CA for any → UK |
| 681150 | Uruguay | source |

---

## Phase 1 — IC-flagged scan (per subsidiary)

### Step 1 — Determine the period and direction families

Ask the accountant for the month and year if not stated. Date range = first through last day.

Ask which subsidiaries to scan (default all five): US (forward), CA, NL, UK, UY (reverse and F-to-F). Each scan looks for lines hitting 681xxx accounts on that sub's books. Skip a sub's own IC account (e.g., NL scan excludes 681100).

### Step 2 — Run the parameterized IC query

Use the parameterized SuiteQL at [`../_shared/netsuite-queries.md`](../_shared/netsuite-queries.md) ("IC transaction query"). For each requested sub, fill in `{SUBSIDIARY_FULLNAME}` and `{ACCOUNT_LIST}` per the table in that file.

### Step 3 — Display for review

Group results per sub, then by tranid. For each transaction show: tranid, date, record type, vendor, line memo (full — multi-line memos can carry FX hints like "Remittance £100,000"), amount, currency. Subtotal per currency within each (source, target) pair.

Flag lines where the expense account cannot be extracted from the memo — those need the accountant to specify before Phase 2.

Ask: "These look correct? Confirm and I'll move to the IS drift scan (Phase 1B), or note anything to add/exclude."

**Wait for explicit confirmation.**

---

## Phase 1B — IS sanity check (optional, recommended)

> **Scope rule (confirmed 2026-05-06):** This skill ONLY reclasses items already booked to a 681xxx account. Plain expense activity sitting on the wrong sub is NOT in scope — those are JE coding errors that belong upstream (AP recoding) or to the `je-error-fix` skill. Don't drill into non-681xxx accounts looking for cross-sub candidates.

Phase 1B is a quick sanity check that Phase 1 caught every 681xxx hit. For each foreign sub the accountant wants to verify, run the saved IS report and reconcile the 681xxx line(s) against Phase 1 totals.

Same report for every sub, just filter by `subsidiaryId`:

```
reportId: 323
savedSearchId: CUSTOMREPORT_323_678967_591   (Acme Income Statement)
```

Subsidiary IDs (verified 2026-05-06): US/Inc=2, Canada=3, Netherlands=4, Poland=6, UK Ltd=8, Uruguay=13.

Standard `ns_runReport` call:
```json
{
  "reportId": 323,
  "subsidiaryId": <sub_id>,
  "dateFrom": "{YYYY-MM-01}",
  "dateTo": "{YYYY-MM-LAST}"
}
```

Look at the `Other Expenses → 681000 Intercompany Expense → 681xxx` rows in the report. Compare the totals against Phase 1's per-(S,T) buckets. If they reconcile, Phase 1 is complete. If a 681xxx total appears in the IS but is missing from Phase 1, re-run the Phase 1 query for that sub and check why.

Then proceed to Phase 2.

---

## Phase 2 — Build JE CSV files

After the accountant confirms, generate one combined CSV per (source S, target T) pair that has activity. Each source transaction becomes its own JE within the CSV.

### Universal 4-line reclass pattern

For every (S, T) pair, every source transaction expands into this skeleton (more lines if the source has multiple expense components — see "Multi-line variants" below):

| Line | Account | Line Subsidiary | Side | Eliminate | Notes |
|------|---------|-----------------|------|-----------|-------|
| 1 | `681xxx Intercompany Expense : {T}` (number = T's IC account) | S | Credit | No | Number lookup in [`../_shared/subsidiary-constants.md`](../_shared/subsidiary-constants.md) |
| 2..N | Mapped expense / asset / liability account on T | T | Debit | No | One line per source expense component |
| N+1 | `121900 Accounts Receivable : Intercompany Receivables` | S | Debit | Yes | NAME = `IC-{T}`, DUE TO/FROM = T short |
| N+2 | `211900 Accounts Payable : Accounts Payable : Intercompany Payables` | T | Credit | Yes | NAME = `IC-{S}`, DUE TO/FROM = S short |

**Header rules:**
- **Header Subsidiary = Source sub (S).** Not always US — for CA → US, header is Canada; for NL → UK, header is Netherlands.
- **Journal Entry Currency = `t.currency` of the source transaction**, not the source sub's base currency. If a US-booked GBP vendor bill needs to move to NL, the JE currency is GBP (header US, JE currency GBP) — the FX rate column translates GBP back to USD on US-side lines (see FX rule below).
- **Date** — ask which applies for this batch:
  - **Original transaction date** for cash-driven items (wires, ACH debits, customer payments). Production examples: NL → UK Payroll JEs JE#####/19904/19905 use 3/27/2026 (the actual ACH date).
  - **Last day of month** for vendor-bill / Brex / accrual reclasses. Production examples: most US → foreign reclasses, CA → US/UK Brex JEs.
- **Line Subsidiary** alternates per the table above — S on lines 1 and N+1, T on lines 2..N and N+2.

### CSV column format (16 columns — exact order)

```
External ID, Header Memo, Header Subsidiary, Date, Account,
Journal Entry Currency, Debit, Credit, Exchange Rate, Memo,
NAME, VENDOR, Department, DUE TO/FROM SUBSIDIARY, Eliminate, Line Subsidiary
```

**Per-column rules:**
- **External ID**: `{S}-{T}-{MMMYY}-{TRANID}`. Examples: `US-NL-MAR26-A163913`, `CA-US-MAR26-JE#####`, `NL-UK-MAR26-PAYUK1`. The skill emits this convention going forward; historical JEs in NS keep their original IDs.
- **Header Memo**: `{Mon-YY} {S} to {T} InterCo ({Description})`. Examples: `Mar-26 US to NL InterCo (External Auditor)`, `Mar-26 CA to US InterCo (Brex)`, `Mar-26 NL to UK InterCo (UK Payroll)`. Always include both ends of the direction.
- **Header Subsidiary**: full NS path of S on every row.
- **Account**: full NS format where available (`681140 Intercompany Expense : UK`); otherwise just the account number (e.g., `141100`).
- **Journal Entry Currency**: same value on every row (NetSuite enforces single-currency at the JE header). Equals `t.currency` from source.
- **Debit / Credit**: amounts in the JE currency (not in any line's local currency). Never enter the foreign-currency translated amount — the system computes that from the Exchange Rate column.
- **Exchange Rate**: see FX rule below.
- **Memo**: source line memo on lines 1 through N. IC lines (N+1, N+2) use: `Reclass {tranid} from {S} to {T}`.
- **NAME**: populated on 121900 and 211900 lines only. `IC-{T short_name}` on 121900; `IC-{S short_name}` on 211900. Lookup in [`../_shared/subsidiary-constants.md`](../_shared/subsidiary-constants.md).
- **VENDOR**: vendor name on 681xxx and expense lines for **vendor bills only**; blank for JE sources, Brex, and IC lines.
- **Department**: full NS path (`General & Administrative : Legal`, `COGS : Consulting`, etc.).
- **DUE TO/FROM SUBSIDIARY**: T short name on 121900; S short name on 211900. Blank on other lines.
- **Eliminate**: `Yes` on 121900 and 211900 lines only. `No` everywhere else.
- **Date**: `M/D/YYYY` (no leading zeros).

### FX rate placement — the one rule

> The `Exchange Rate` column on a line tells NetSuite how to translate from the JE's transaction currency to that line's subsidiary base currency.
>
> - Lines whose Line Subsidiary's **base currency equals the JE Currency**: leave Exchange Rate blank (or `1`).
> - Lines whose Line Subsidiary's **base currency differs from the JE Currency**: enter the rate = `(JE currency amount) ÷ (line subsidiary base currency amount)`.

Sub base currencies: US=USD, NL=EUR, UK=GBP, UY=UYU, CA=CAD.

**When to leave it blank vs enter explicitly.** Production CSVs sometimes leave the rate blank even on lines where the rule says it's required (e.g., the Mar-26 DPH Legal CSV leaves it blank on NL-side lines, with rate populated only on US-side). When blank, NetSuite uses its consolidated rate table at the JE date — usually close, but not exact. **When precision matters — cash transfers tied to a specific wire, vendor bills with explicit foreign-currency amounts in the memo — enter the rate explicitly on every line that needs translation.**

**Default rate source:** `t.exchangerate` from the source transaction. Pull this into Phase 1's display so the accountant sees it before approving.

**Memo-driven override:** If the source transaction memo contains a target-currency hint (e.g., `Remittance £100,000` on a NL transaction in EUR booking to UK), surface the implied rate alongside `t.exchangerate` and ask the accountant which to use. Computation: implied rate = (JE currency amount) ÷ (target-currency amount in memo). No silent overrides — the accountant picks.

**Worked examples:**

| JE | Header | JE Currency | Side(s) needing rate | Rate origin |
|----|--------|-------------|----------------------|-------------|
| JE##### (US → CA cash transfer, USD wire) | US | USD | CA-side lines | `1.34582` (USD→CAD) — from wire FX |
| `US-NL-MAR26-16892` (DPH Legal vendor bill) | US | GBP | US-side lines (NL-side blank in production; NS auto-derives EUR/GBP) | `1.38370001` — from `t.exchangerate` |
| `NL-UK-MAR26-PAYUK1` (NL → UK payroll, EUR wire) | NL | EUR | UK-side lines | Memo-implied: `116284.88 / 100000 = 1.16284988` (preserves £100,000 exactly) |
| `CA-US-DEC25-JE#####` (CA → US Brex, CAD source) | CA | CAD | US-side lines | `t.exchangerate` from each Brex transaction (often `1` for CAD-only Brex) |

**JE##### single-currency lesson** (verified 2026-04-27, re-verified 2026-05-06): all four lines must be in the JE currency. Don't enter $X,XXX.XX CAD on Canada-side lines for a $X,XXX.XX USD wire. Both the REST API and the CSV import expect USD ($X,XXX.XX) on every line, with the FX rate doing the translation. NS reads back $X,XXX.XX on Canada's books because of its translation, not your input. The local CSV at `Monthly Intercompany JEs/2026-04 April 2026/Canada/2026-04 Canada FX Cash Transfer Adv IC JE.csv` is a stale draft that violates this rule; the actually-posted JE##### follows it.

### Multi-line variants

**Multi-line payroll / vendor-bill JEs.** When a single source transaction has multiple expense components hitting the same 681xxx account, mirror each:

```
For each component: Credit 681xxx (S) + Debit expense (T) — individual amounts
Then ONE net pair: Debit 121900 (S) + Credit 211900 (T) — sum of all components
```

Example: `NL-MAR26-2025574` (KWPS BV) — 681100 split into Legal Fees + VAT, single 121900/211900 net pair.

**Brex expense report JEs (10+ individual line items).** Net the 681xxx and IC lines, expand the expense lines:

```
One net credit to 681xxx (S)
One debit per Brex expense line (T) with full memo + dept
One net debit to 121900 (S)
One net credit to 211900 (T)
```

Production example to model against: `CA-US-DEC25-JE#####` (CA → US, multiple Brex lines) and `CA-MAR26-JE#####` (US → CA, also Brex). Both show the same compression pattern.

### File naming and storage

```
Monthly Intercompany JEs/
  {YYYY-MM} {Month Name} {Year}/
    {Source sub full name}/
      {YYYY-MM} {S full name} to {T full name} Intercompany JE.csv
```

Use the full short names from [`../_shared/subsidiary-constants.md`](../_shared/subsidiary-constants.md) in folder + filename (`Netherlands`, `UK`, `Uruguay`, `Canada`, `US`). The 2-letter codes (`NL`, `UK`, `UY`, `CA`, `US`) are used only in the External ID prefix.

Examples:
- `Monthly Intercompany JEs/2026-04 April 2026/US/2026-04 US to Netherlands Intercompany JE.csv`
- `Monthly Intercompany JEs/2026-04 April 2026/Canada/2026-04 Canada to US Intercompany JE.csv`
- `Monthly Intercompany JEs/2026-04 April 2026/Netherlands/2026-04 Netherlands to UK Intercompany JE.csv`

Only create files for (S, T) pairs that have activity that month.

> **ASCII only** in every CSV field: no Unicode arrows, em-dashes, or special chars. Spell out directional words ("to", "from", "vs").

---

## Pre-import structural diff (recommended for new directions)

The first time the skill emits a CSV for a (S, T) pair that's never been built locally before, **diff its structure against a known-good production JE in NetSuite** before the accountant imports.

Use the "Adv IC JE line detail by tranid" SuiteQL at [`../_shared/netsuite-queries.md`](../_shared/netsuite-queries.md). Canonical references:

| Direction | Canonical historical JE |
|-----------|-------------------------|
| US → NL | JE#####-era NL files; local CSV at `Monthly Intercompany JEs/Historical Examples/Netherlands/Mar-26/2026-03 Netherlands Intercompany JE - EUR.csv` |
| US → UK | local CSV at `Monthly Intercompany JEs/Historical Examples/UK/Mar-26/2026-03 UK Intercompany JE - All.csv` |
| US → UY | local CSV at `Monthly Intercompany JEs/Historical Examples/Uruguay/Mar-26/2026-03 Uruguay Intercompany JE - USD.csv` |
| US → CA | local CSV at `Monthly Intercompany JEs/2026-03 March 2026/Canada/2026-03 Canada Intercompany JE.csv` |
| CA → US | JE##### (Mar-26), JE##### (Dec-25), JE##### (Sept-25) — pull live via SuiteQL |
| CA → UK | JE##### ("CANtoUKF26", Feb-26), JE##### (Mar-26) — pull live |
| NL → UK | JE##### / JE##### / JE##### (Mar-26) and local CSV at `Monthly Intercompany JEs/2026-03 March 2026/Netherlands/2026-03 Netherlands to UK Payroll Intercompany JE.csv` |

If the new CSV's structure (column-by-column on the 4-line skeleton) deviates from the canonical, fix the CSV before import.

---

## Validation before saving

For every JE (grouped by External ID), verify:

- [ ] Debits = Credits (balanced) in the JE currency
- [ ] 121900 and 211900 amounts equal the net IC amount for that JE
- [ ] Header Subsidiary = source sub on every row
- [ ] Journal Entry Currency identical on every row
- [ ] 681xxx account number matches the **target** sub (per the table at the top)
- [ ] Line Subsidiary alternates per the 4-line pattern
- [ ] Exchange Rate placed per the FX rule (lines whose subsidiary base currency != JE currency)
- [ ] Eliminate = Yes on 121900 and 211900 only; No elsewhere
- [ ] DUE TO/FROM SUBSIDIARY and NAME correct on IC lines (per [`../_shared/subsidiary-constants.md`](../_shared/subsidiary-constants.md) F-to-F matrix)
- [ ] Memo on IC lines: `Reclass {tranid} from {S} to {T}`
- [ ] Date format: `M/D/YYYY`, no leading zeros
- [ ] No `$` signs or commas in numeric fields
- [ ] VENDOR blank on 121900, 211900, and JE-sourced lines
- [ ] ASCII only — no Unicode arrows, em-dashes, smart quotes

---

## Check-and-balance

See [`../_shared/check-and-balance.md`](../_shared/check-and-balance.md). Before saving any CSV, display a per-JE summary (External ID, header memo, total debit, total credit, line count) and ask for `yes` before writing files. Writing a local CSV is not a NetSuite write, but the accountant will import it — confirm first.

After saving, no `audit_log.json` entry yet (no NS write). Audit log is appended after the accountant imports and reports the resulting JE numbers (per CLAUDE.md rule #8).

---

## Customer Payment Wrong-Sub Reclass (out of scope — manual)

Different shape than expense reclasses. Triggers when a customer paid `Acme Canada` for a `US` invoice (or similar). Don't auto-detect; AR flags it manually.

Production examples: JE##### (7/2024), JE##### (4/2025) — both for US invoices `INV-US-#####`, `INV-US-#####`.

4-line shape (CA → US example):
| Line | Account | Side | Line Sub | Notes |
|------|---------|------|----------|-------|
| 1 | `111011 TD/Barclays - Canada` | Debit | CA | Cash received on CA bank |
| 2 | `211900 IC Payables` | Credit | CA | Eliminate=Yes, NAME=`IC-Acme, Inc.`, DUE TO/FROM=`Acme, Inc.` |
| 3 | `121100 Accounts Receivable : Accounts Receivable` | Credit | US | Closes out the original AR |
| 4 | `121900 IC Receivables` | Debit | US | Eliminate=Yes, NAME=`IC-Acme Canada`, DUE TO/FROM=`Acme Canada` |

Header sub = CA. JE currency = CAD. Header Memo: `Payment received in Acme Canada for US invoice {INV-XXX}`.

Skip this in the standard run. If the accountant asks for one explicitly, build it ad-hoc using the JE#####/JE##### line detail as the template (pull via SuiteQL).
