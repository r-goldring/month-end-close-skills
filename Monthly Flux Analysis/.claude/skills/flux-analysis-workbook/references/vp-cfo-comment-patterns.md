# VP and CFO Comment Patterns

What gets pushback in the flux meeting, and how to phrase comments to preempt it. Mined from Nov 2025 – Apr 2026 meeting transcripts.

## Phrasing rewrites (Don't write → Write instead)

| Don't write | Write instead | Why |
|---|---|---|
| `Medical Carrier A down -$36K` | `Medical Carrier A billing $90K vs Mar $126K (-$36K), continued FY25 RIF flow-through on enrolled headcount` | Generic "down $X" gets "How come?" pushback. Name the cause. |
| `Mar Cigna stop-loss accrual BILL-UYMED-001 not repeating` | `Mar Uruguay Medical Vendor (Uruguay medical) catch-up bill $123K not repeating` | Bill numbers are opaque. Always spell out vendor + what it is. |
| `Mar Former FSA Administrator refund -$34K and FlexBenefits -$7K not repeating` | `Mar Former FSA Administrator refund -$34K (FSA termination cleanup, likely belongs in 2025 — the Assistant Controller confirming with the AP Specialist) + FlexBenefits -$7K not repeating` | If a number could be wrong, flag the open question rather than leaving it for the VP to surface. |
| `Pending accrual` | `Pending Apr accrual (IT Hardware Partner $377K + DNS Provider A $17K)` | Always name the specific JE(s) and amounts. |
| `Software for Monday purchase` | `Monday.com prepayment for [N] months; capitalized per $1K+/multi-month rule` | the CFO pushes back on $1K+ items that look like one-time expense without prepaid context. |
| `Brex charge hit software` | `Brex {vendor} {amount} on {dept} — {memo first line}; verify receipt by the AP Reviewer/the AP Specialist` | the VP wants vendor + GL + dept + receipt status. |
| `Software contract adjustment` | `BillFlow: $36K annual new (Jan 30 start) replaces $6K/month prior (May 12 end); credit applied per contract, schedule cleaned` | For overlapping contracts always state: original end, new start, credit, resulting amortization. |
| `Contractor one-time charge` | `Contractor_001: recurring Legacy Hosting Stack contractor (not one-time). April accrual TBD pending approver inquiry` | the VP always asks one-time-vs-recurring. State explicitly. |
| `Refund hit the balance sheet` | `Former FSA Administrator refund $34K to 611400. the Assistant Controller confirming with the AP Specialist; expected to push back to 2025 per the VP's policy` | Frame: was this a 2025 cleanup or 2026 benefit? |
| `Commission adjustment` | `Commission Feb $278K USD = $60K CAD + $87K GBP (per SPIFF report, currency-converted). $1K diff vs estimate from FX rate timing` | the CFO hit on Feb commission mismatch — always state currencies and conversion source. |
| `Travel variance` | `Travel +$XK vs forecast; pending Canadian Brex ~$50K (the Executive Assistant) will likely flip variance unfavorable when posted` | Always flag pending Brex impact on T&E. |
| `Software` (no vendor / dept name) | `{Vendor}: ${amount} currently booked to {dept}. Verify per budget owner ({the Financial Analyst / the CFO / dept lead})` | the Financial Analyst (when present) flags wrong-dept software; preempt by tagging owner. |

## Common questions the controller/CFO asks (preempt these in commentary)

### "Why is this jumping month-to-month?"
Provide: prior month $, current $, delta, and the SPECIFIC driver (vendor / contract change / billing timing / FY25 RIF flow-through / etc.). If it's a swing that will continue, say so; if one-time, say "Mar one-off not repeating".

**Examples**:
- Health Premium Billing Intermediary premium swings (Apr meeting the VP: "Why is it bumpy?")
- Software with no contract change (Smart Sheets in Feb — should be flat)
- DocuSign overlap (renewed early → double amortization)
- D and O Broker (should be flat — if not, contract changed)

### "Is this one-time or recurring? Do we accrue next month?"
State: status (recurring / seasonal / one-time / uncertain), cadence (monthly / quarterly / ad-hoc), and whether next-month accrual is needed. Cite the approver if asking them for anticipated commitment.

**Recent examples**:
- Contractor_001 — the CFO confirmed recurring Legacy Hosting Stack
- Tax Firm A / Pension Audit Firm — quarterly/annual cadence
- Health Premium Billing Intermediary premium transfer — random, not predictable

### "Did this belong in [prior year]?"
For refunds / reversals / credits, state which year the related transaction lived in. If it's a true 2026 cost it stays; if it reverses a 2025 accrual, push back.

**Pattern**: Former FSA Administrator refunds → 2025 (per the VP); WC audit refunds → year of the audit period; State tax extension reversals → match the year the extension was for.

### "What's the plan for [vendor / acquisition] transition?"
For ongoing transitions (Legacy Hosting Stack, Uruguay payroll UY Payroll Provider B, Customer Conference rename), state: current status, interim accounting treatment, cutover timeline, external owner.

**Recent examples**:
- Legacy Hosting Stack vendor transition — the VP owns, the Assistant Controller handles Brex reclasses
- Uruguay payroll → UY Payroll Provider B (post-Uruguay Medical Vendor)
- Customer Conference → Customer Conference rename (the CFO worried about coding consistency)

### "Can this process be simpler?"
the VP and the CFO repeatedly pushed for simpler processes:
- Brex accrual by GL/dept/vendor (not by person)
- Prepaid amortization for annual software (not monthly granular entries)
- One Anthropic charge → allocate by usage (not allocate per dept upfront)

When commenting on something complex, propose a simpler approach.

### "Why wasn't this budgeted?"
- If favorable: confirm whether it's a real cost reduction or just timing.
- If unfavorable: explain the driver (acquisition, business change, error) and whether it continues.

**Examples**: IT Hardware Partner Mar accrual drift, Health Premium Billing Intermediary premium swing, Anthropic spend ramp.

## Specific people and their hot-button items

### the VP (Controller / VP Accounting)
- **Asks "was this expected?"** — wants accruals justified.
- **Tracks bill numbers** if missing context — name the vendor.
- **Pushes Former FSA Administrator refunds to 2025** — refunds related to 2025 accruals shouldn't benefit 2026.
- **Wants quick-close mindset** — simplify Brex accrual, don't sweat <$200 items.
- **Owns the close timeline** — Mid-day next-day after flux meeting is target.

### the CFO (CFO)
- **Asks "is this reasonable?"** — directional sanity check.
- **Flags doubled / zero-balance vendors** — wants explanation when a stable vendor jumps or disappears.
- **Software dept-coding accuracy** — the Financial Analyst owns when present; otherwise the CFO asks.
- **Worries about: Anthropic spend ramp, event costs (Customer Conference), forecast misses on PS revenue and travel.
- **Recent insistence**: "no more estimates" for commissions — use actual SPIFF.

### the Director (FP&A)
- **Owns revenue, payroll, and headcount forecast**.
- **Provides budget context** during the call ("100K lower than forecast", etc.).
- **Goes to her for IT Hardware Partner explanation** — she's the VP contact point.

### the FP&A Lead (Sales Leadership)
- **Owns commissions and renewal data** in Salesforce.
- **Confirms vendor classification** (e.g., Executive Content Agency → contractors).
- **Investigates revenue mysteries** (Customer Example A missing ARR, Customer Example B downsell question).

### the Financial Analyst (Financial Analyst — currently OUT)
- **Owns software dept-coding alignment** to budget.
- **While out (Apr 2026)**: do dept coding flags manually in column H.

### the VP (Head of Legal)
- **Owns**: Legacy Hosting Stack vendor transition, cap software infrastructure percentage, acquisition-related expenses.
- **Slow response window** — don't hold close on his updates.

### the Assistant Controller (Assistant Controller)
- **Handles**: Brex accrual, prepaid amort schedules, live close-meeting fixes.
- **Goes to**: the AP Specialist for HR/people/benefits questions, the Tax Accountant at Tax Firm A for tax debits.

## Topics to AVOID in comments unless asked

- Long explanations of FX retranslation mechanics (cite the rate move and dollar impact only).
- Justifying interco eliminations beyond "Interco" (these always tie out by design).
- Reciting investigation procedure ("I pulled 3 months of transactions...") — just give the answer.

## Pre-meeting checklist (preempt mid-meeting Q&A)

Before sending the v3 workbook:
1. Every Y-flagged row has either: (a) a 1-line comment with vendor + amount + driver, OR (b) `Pending {category}` tag.
2. Recurring known-pattern rows (IT Hardware Partner, Health Premium Billing Intermediary, Medical Carrier A, Marsh, etc.) cite the pattern explicitly so the VP/the CFO don't have to ask.
3. Pending JEs are listed in the chat summary by category and expected impact.
4. Anything weird (false-positive Lodging in Software, Brex misattribution) is called out in the chat — not buried in the workbook.
5. Bonus and commission accruals are either uploaded or explicitly tagged as pending.
