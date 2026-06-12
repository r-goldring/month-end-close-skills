# Comment Recipes — Per-Account Templates

Templates for writing column-H variance commentary. Each recipe enforces the
standards from `vp-cfo-comment-patterns.md`: vendor in plain English (no bill
numbers alone), specific $K with rounded precision, period-comparison context,
and the driver named.

Always pre-load `known-vendors.md` and `recurring-themes.md` when applying
these templates.

## Income Statement recipes

### Revenue (411xxx)

```
Template: "{Direction} {customer name} {dollar context}; {timing/cause}"
Example:  "Higher US PS revenue +$70K (customer milestone timing) plus NL +$8K"
```
- Always name the customer if a single one drives >$10K of the variance.
- If multiple small movers, group as "expansion timing" / "renewal catch-up" / "late rev rec".

### COGS — Hosting (511400)

```
Template: "{Vendor} accrual ${apr}K (vs Mar actual bill ${mar}K) {explanation}"
Example:  "Apr IT Hardware Partner accrual $377K vs Mar actual bill $348K; DNS Provider A $17K"
```
- IT Hardware Partner dominates — name it explicitly.
- If accrual differs from prior bill by >$20K, explain the gap (the VP status, credits, missing instance reports).

### COGS — Software Subs (511425) / Translation (511450)

```
Template: "{Specific bill or accrual event}; {prior month event}"
Example:  "Mar Survey Translation Vendor $11K accrual reversed; Apr bill $10K booked"
```

### COGS — Health Benefits (511200)

```
Template: "{Vendor} {direction} {-$amount}K from {cause}"
Example:  "Medical Carrier A billing $90K vs Mar $126K (-$36K) from continued FY25 RIF flow-through"
```

### OpEx — Bonus (611150)

```
Template: "{Prior-month one-off explanation}; {current month run-rate explanation}"
Example:  "Mar Q1 bonus accrual catch-up not repeating (US $140K higher); Apr back to monthly run-rate"
```

### OpEx — Commission / Commission Deferral (611200 / 611201 / 611202)

```
Template: "Lower/higher Q{X} commission run-rate: {geo breakdowns from SPIFF}"
Example:  "Lower Q2 commission run-rate: Canada -$57K, UK -$28K, US -$14K vs Mar SPIFF"
```
- Use SPIFF as source-of-truth. No more estimates per the CFO's Feb 2026 directive.

### OpEx — Health Benefits (611300)

Cross-reference: 611300 has the Uruguay Uruguay Medical Vendor bills that look like BILL-UYMED-002 / BILL-UYMED-001.

```
Template: "{Primary vendor} {direction} {-$amount}K ({cause}); {prior one-off vendor by plain-English name} not repeating"
Example:  "Medical Carrier A down -$44K (headcount); Mar Uruguay Medical Vendor (Uruguay medical) catch-up bill $123K not repeating"
```

### OpEx — Other Benefits (611400)

```
Template: "Mar {refund/credit names} not repeating; Apr near zero"
Example:  "Mar Former FSA Administrator refund -$34K (FSA termination cleanup, likely belongs in 2025 — the Assistant Controller confirming) and FlexBenefits -$7K not repeating"
```

### OpEx — Payroll Taxes (611450)

```
Template: "{Direction} US payroll tax {-$amount}K (Mar carried higher bonus payroll tax with Q1 catch-up)"
Example:  "Lower US payroll tax -$30K (Mar carried higher bonus payroll tax with Q1 catch-up)"
```

### OpEx — Events (621100) / Conferences (621110)

```
Template: "{Vendor/event} {amortized/deposit applied/new spend}; {prepaid vs expense reasoning if event is future}"
Example:  "Q2 live event ramp (UK Leadership Programme Apr 14+21, CHRO Assembly Series Q2 event); Customer Conference charges in 621110 should reclass to prepaid (event 5/14)"
```

### OpEx — Marketing (621200 / 621250)

```
Template: "{Mar campaign vendor + $} not repeating; Apr {credit/lower spend}"
Example:  "Mar Lead Gen Vendor A lead gen $16K and Lead Gen Vendor B ABM $9K not repeating; Apr net credit from reclass adjustments"
```

### OpEx — Legal (651150)

```
Template: "Mar {legal matter names} closed; Apr only ${residual}"
Example:  "Mar legal matters closed (Legal Firm A, Legal Firm B); Apr only $3K residual"
```

### OpEx — Lodging (661200)

```
Template: "Higher/lower Apr T&E lodging volume {+/-$amount} plus {any one-time JE if material}"
Example:  "Higher Apr T and E lodging volume plus $16K Apr re-coding JE"
```
- If pending Canadian Brex, always flag.

### OpEx — Software Subs (671100)

```
Template: "{Pattern from comment-recipes.md per vendor} or "See Software tab"
Example:  "See Software tab"
```
- Always defer to detail tab if the line has one.

### Other Expense — Interest (711xxx)

```
Template: "Q{X} interest paid {ACH date} via ${amount}M ACH cleared accumulated balance; Apr daily accrual {status}"
Example:  "Q1 interest paid 4/2 via $1.5M ACH cleared accumulated balance; Apr daily accrual still rebuilding"
```

### Other Expense — FX (700000 / Unrealized Gain/Loss)

```
Template (Realized): "FX rate movement on settled foreign transactions"
Template (Unrealized): "USD/EUR/GBP rate movement on intercompany balances"
```

### Income Tax (720000)

```
Template: "{Prior period} {transfer pricing accrual / state estimates / true-up} not repeating; {current period} only {residual} {payments/reversals}"
Example:  "Mar Q1 tax estimate true-up $571K not repeating; Apr only $104K state estimates after reversals"
```

## Balance Sheet recipes

### Bank Accounts (111xxx)

```
Template (operating): "{Sub} operational cash flow"
Template (sweep):     "Cash sweep from {source} to {target}"
Template (inflow):    "{Sub} customer collections inflow"
```

### Accounts Receivable (121100)

```
Template: "See AR tab"
```

### Prepaid Events (121340)

```
Template: "Apr event {amortization|deposit accumulation} ({vendors + $K applied})"
Example:  "Apr event amortization (CHRO Assembly Series $30K + UK Leadership Programme $37K deposits applied)"
```

### Capitalized Commissions (121370 / 161120)

```
Template: "Pending commissions ST/LT reclass" (or describe the quarterly reclass)
```

### Intercompany (681xxx / 121900 / 211900 / 211999 / 121999)

```
Template: "Interco" (single word — these always reconcile; don't over-explain)
```

### Accounts Payable (211100)

```
Template: "See AP tab"
```

### Brex Credit Card (231110)

```
Template: "Higher/lower Apr Brex US card spend (statement closes {date})"
Example:  "Higher Apr Brex US card spend (statement closes mid-May)"
```

### Sublease Deposits (231155) / Deferred Sublease Rent (261100)

```
Template: "{Tenant} {action} {-$amount} ({location/context})"
Examples: "Premier Plans Springfield sublease deposit returned via $534K ACH on 4/30"
          "Springfield sublease rent +$11.7K (consistent monthly)"
```

### Accrued Commissions (231160)

```
Template: "Q{X} commission payouts: US ${}K + Canada ${}K + UK ${}K cleared from accrual"
Example:  "Q1 commission payouts: US $30K + Canada $58K + UK $32K cleared from accrual"
```

### Accrued Income Taxes — ST (231490)

```
Template: "Mar Q1 tax estimate {amount} not repeating; Apr only {residual} state estimates"
Example:  "Mar Q1 tax estimate $571K not repeating; Apr only $104K state estimates"
```

### Sales Taxes (235100 Avatax)

```
Template: "Apr Avatax remittance timing (state filings)"
```

### Term Loan ST/LT (245270 / 261270 / 261277)

```
Template (current portion bump): "Quarterly principal reclass from LT to ST (JE{XXX} ${}K)"
Template (accumulated interest): "Q{X} interest paid {date} via ${}M ACH cleared accumulated balance"
```

### DR LT (251100 / 251150 / 251160)

```
Template: "Apr LT to ST DR reclass"
```

### Cumulative Translation Adjustment (R233)

```
Template: "FX retranslation of foreign sub equity (USD/EUR/GBP rate moves)"
```

## Anti-patterns (avoid these)

| Phrase | Why it's bad | Use instead |
|---|---|---|
| "Costs increased" / "Higher spend" | No specificity | "${vendor} {+$amount}K from {cause}" |
| "See above" / "See related" | Doesn't help meeting reader | Repeat the key driver inline |
| "FY25 accrual reversal" without vendor | Vague | Name the vendor and amount |
| Bill / JE / Doc numbers without context | Opaque | Look up vendor in `known-vendors.md` |
| Em-dashes (—) | Non-ASCII per CLAUDE.md rule 7 | Use ` - ` or `;` |
| Arrows (→ / ↑ / ↓) | Non-ASCII | "from X to Y", "up", "down" |

## Subtotal rows

**Leave column H blank on subtotals**. The driver lives on the underlying detail rows. This matches Mar 2026 final flux convention.

Subtotal rows to leave blank:
- All "Total - {acct}" rows
- "Net Ordinary Income", "Net Other Income", "Net Income"
- "Total Current Assets / Liabilities", "Total Bank", "Total Equity"
- "Total ASSETS", "Total Liabilities & Equity"
