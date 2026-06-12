# Recurring Themes — Patterns Across Closes

Use this file to recognize repeating patterns so commentary doesn't re-investigate them from scratch each month. Each theme cites a recent example (month + $ amount) for grounding.

## Monthly cadence patterns (always present)

### IT Hardware Partner Hosting — billed in arrears
- **Pattern**: Accrual booked at month-end based on the AP Specialist's draft report; actual bill from IT Hardware Partner lands 2–3 weeks into the following month; difference creates a small benefit/cost in the bill month.
- **Recent**: Mar 2026 accrued $377K → actual bill $348K = $29K Apr benefit. Mar 2026 conversation: the VP's communication with IT Hardware Partner about missing instance reports + credits caused the $29K gap.
- **Commentary template**: "{Month} IT Hardware Partner accrual $X vs actual $Y; net {benefit/cost} of $Z. {Explanation if available.}"
- **Skill behavior**: Always accrue IT Hardware Partner at month-end at the report estimate.

### Medical Carrier Aical — varies with enrolled headcount
- **Pattern**: Monthly premium tracks active US enrollment. FY25 RIF is still flowing through; expect step-down each month until stabilization.
- **Recent**: Feb $127K → Mar $126K → Apr $90K (-$36K MoM as more RIF terminations land).
- **Commentary template**: "Medical Carrier A billing ${apr}K vs prior ${prior}K (-${delta}K), continued {FY25 RIF flow-through / open enrollment shift / other}."

### Health Premium Billing Intermediary Health Insurance — discretionary billing + renewal surplus carryover
- **Pattern**: Premium transfer line item can cause $80K+ swings between months. Carryover surplus from prior year drives some months up, then absent in following months.
- **Recent**: Mar 2026 had $88K premium transfer; Apr 2026 had none. the VP flagged "why is it bumpy?" in Apr meeting.
- **Commentary**: Always cite if premium transfer is present/absent and amount.

### D and O Broker — flat amortization
- **Pattern**: $X,XXX.XX/month exactly. Any variance signals a contract change or billing error.
- **Skill behavior**: If MoM movement on Marsh ≠ $0, investigate immediately.

### Brex card statement timing — Canadian + UK + US
- **Pattern**: Canadian Brex statement closes ~19th of following month; $50–100K+ regularly outstanding at month-end close. UK and US Brex usually in before close.
- **Commentary template**: When commenting on T&E or Software variance, always flag if pending Brex is expected to flip the variance.
- **Action**: Check with the AP Reviewer and the AP Specialist before declaring T&E variance final.

### Term Loan Interest (711150) — quarterly ACH payment
- **Pattern**: ACH on the 1st–3rd of the quarter-end month (e.g., Apr 2, Jul 1, Oct 1) pays accumulated interest from prior quarter. Accumulated Interest (261277) clears to zero, daily accrual rebuilds.
- **Recent**: Apr 2 ACH paid $1.5M, cleared $1.54M accumulated balance; Apr daily accrual rebuilt $526K toward next quarter.

## Quarterly cadence patterns

### 401k administration — RetireBridge / A-Census quarterly billing
- **Pattern**: ~$6K invoice every 3 months (Nov/Feb/May/Aug or similar). Causes 651180 Payroll Processing Fees to spike in those months.
- **Recent**: Apr 2026 PI-03922723 $X,XXX.XX (Nov–Jan quarterly bill); the VP confirmed quarterly cadence.

### Q1 transfer pricing accrual reversal (April)
- **Pattern**: Mar carries Q1 transfer-pricing accrual ($500K+); Apr reverses out. Income tax expense swings negative in Apr accordingly.
- **Recent**: Mar 2026 JE##### accrued $X,XXX.XX → Apr 2026 reversed via JE##### (-$52K of state extension reversals).

### State tax extension auto-debits (April)
- **Pattern**: Tax Firm A processes state tax extension payments by ACH starting in April. the Assistant Controller follows up with the Tax Accountant at Tax Firm A if they don't appear by close.
- **Recent**: Apr 2026 had auto-debits not yet visible at close — the Assistant Controller owed the Tax Accountant follow-up.

### 401k audit (Pension Audit Firm, ~April)
- **Pattern**: Annual 401k audit fee hits Professional Fees (651100). Timing was July in FY24 but moved earlier to April in FY25.
- **Recent**: Apr 2026 Pension Audit Firm hit + Tax Firm A FY25 tax provision = drove 651100 up.

## Annual / one-time patterns

### Aguinaldo (Uruguay 13-month bonus) — annual payout
- **Pattern**: Accrued monthly at 8% of salary across 231170 / 231171; paid annually mid-year.
- **Recent**: Feb 2026 ~$12.5K USD accrual hit 611150 Bonus.

### Netherlands Holiday Pay (8%) — annual payout
- **Pattern**: 8% statutory holiday allowance, accrued monthly, paid in May.
- **Skill behavior**: Always flag if NL holiday accrual is missing from the bonus run.

### Q1 software reclass (iPaaS Vendor A → COGS, SSO Provider → COGS)
- **Pattern**: Once a year, software subscriptions used for customer-serving infrastructure get reclassed from 671100 OpEx to 511425 COGS. Q1 2026 finished in April via JE#####.
- **Items reclassed Q1 2026**: SSO Provider $17K, iPaaS Vendor A $6.4K.

### Capitalized commissions ST/LT reclass — quarterly
- **Pattern**: Each quarter the commission deferral asset moves between 121370 Current and 161120 Non-Current based on what falls inside/outside the 12-month window.
- **Recent**: Q1 2026 reclass was a $6M move (Apr 2026 saw the reclass-finalization JE).

### Prepaid LT-to-ST reclass — quarterly
- **Pattern**: Each quarter, items in 161112 Prepaids Non-Current that have <12 months remaining move to 121310 / 121340 ST prepaids.

### Term Loan Current/LT reclass — quarterly
- **Pattern**: Each quarter the next 3 months of principal moves from 261270 LT to 245270 Current. Apr 2026 JE##### moved $337K.

## Event-driven patterns

### Legacy Hosting Stack acquisition (Feb 2026 close)
- **Vendor transition timeline**: Feb acquisition → Mar amortization start → Q2 2026 vendor cleanup (Google Cloud, others). the VP owns transition plan.
- **Watch**: Legacy Hosting Stack-era Brex charges hitting 671100 OpEx should reclass to 511425 COGS — the Assistant Controller handles this monthly.
- **Contractor**: Contractor_001 is a recurring Legacy Hosting Stack contractor; needs monthly accrual.

### Customer Conference / Customer Conference customer conference (annually, May)
- **Renamed FY26**: Previously "Customer Conference" / "Customer Conference", now "Customer Conference" (5/14 each year).
- **Spend pattern**: Charges accumulate Mar–Apr (deposits, swag, travel) — should sit in 121340 Prepaid Events until May, then amortize in event month.
- **Site visits in March count as prep**: $500-ish hotel + meals in Mar (Daniel Stanczyk, Ashley Conrad) for May venue scouting — debate-able whether to prepay these.
- **Brex coding**: There's an "Customer Conference" tag in Brex but it doesn't push to NetSuite. Cross-reference manually.

### President's Club Trip / President's Club Trip — President's Club, annually
- **First year executed in 2025–26**: 7 quota-hitters (Amanda, Ross, Stephen, Jamie Haymon, Chad, the AP Reviewer, +1 unknown).
- **Expense pattern**: Sales reps charge to Brex; tags may say "winner circle" / "President's Club Trip" / "President's Club" in memo.
- **Right account**: 641201 Company Events (per the Director's Apr 2026 ask — was hitting 661200 Lodging and 621110 Company Conferences).

### Former FSA Administrator termination — refunds rolling
- **Pattern**: Former FSA Administrator relationship ended rocky in late 2025; FSA refunds posting through 2026.
- **the VP's preference**: Push refunds back to 2025 (where the related accrual lived) rather than benefit 2026 P&L.
- **Recent**: Mar 2026 -$34K refund — the Assistant Controller owes the AP Specialist follow-up on supporting detail.

### AcquiredCo lease cleanup (Feb 2026)
- **Pattern**: AcquiredCo (acquired prior year) lease expired Feb 2026; ROU asset + lease liability cleared.
- **Outstanding**: $18K security deposit (161100) refund pending from landlord.

## Sub/FX patterns

### Foreign sub FX retranslation (CTA-Elimination)
- **Pattern**: Each month, foreign-sub balance sheet retranslates at period-end FX rate; the gain/loss hits Cumulative Translation Adjustment-Elimination (R233 on BS).
- **Recent**: Apr 2026 +$86K from USD/EUR/GBP rate moves.

### Realized vs Unrealized FX (700000 / Unrealized Gain/Loss)
- **Realized**: Hits on settled foreign transactions (e.g., when a foreign-currency AR is collected).
- **Unrealized**: Daily mark-to-market on open foreign-currency balances (e.g., intercompany loans). Apr 2026 swing was +$615K driven by USD/EUR/GBP rate moves on IC balances.

## Revenue recognition patterns

### Late renewals
- **Pattern**: Customer renews late; backdated rev rec creates "catch-up" revenue in the late-recognition month.
- **Recent**: Mosaic backdated to Jan (Apr meeting), AT&T 10-month expansion had 3 months of catch-up in Apr ($26K Apr really runs ~$8K monthly going forward).

### Downsells masked as renewals
- **Pattern**: A "renewal" may actually be a license reduction (read-only / data-storage license). Looks like a churn even though contract was renewed.
- **Recent**: Hitachi $600K → $60K (data-storage license only) in Apr 2026.

### Missing ARR despite contract
- **Pattern**: Salesforce ARR field can be $0 while a $200K renewal sits in renewal pipeline. Usually a data hygiene issue.
- **Recent**: Customer Example A — zero ARR but expected $200K renewal (the FP&A Lead to fix in Salesforce).

### Customer Conference → Customer Conference revenue offsetting
- **Pattern**: Customer Conference ticket revenue (Customer Conference now) historically offsets event expenses — not booked as revenue.

## Materiality / threshold conventions

- **Workbook variance flag**: `>$10K absolute AND >10%`
- **Detail-tab vendor flag**: `>$2K MoM variance` (per investigation playbook)
- **Software accrual minimum**: $500 (lowered from $750 on 2026-05-12)
- **Contractor accrual minimum**: $100
- **Professional Fees / COGS accrual minimum**: $500
- **Prepaid capitalization rule**: $1K+ AND multi-month benefit → capitalize, otherwise expense
