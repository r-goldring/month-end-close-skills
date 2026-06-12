# Subsidiary Constants

Use these values in all SuiteQL queries and JE CSV fields. These are the exact NetSuite full-name strings.

## Full subsidiary paths

| Short name | NetSuite fullname | Currency | Reconciliation Tool folder |
|------------|------------------|----------|----------------|
| Acme Inc (US) | `Acme Holdings : Acme, Inc.` | USD | `Acme Inc Reconciliations` |
| Acme Canada | `Acme Holdings : Acme, Inc. : Acme Canada` | CAD | `Acme Canada Reconciliations` |
| Acme Netherlands | `Acme Holdings : Acme, Inc. : Acme Netherlands` | EUR | `Acme Netherlands Reconciliations` |
| Acme Poland | `Acme Holdings : Acme, Inc. : Acme Poland` | PLN | `Acme Poland Reconciliations` |
| Acme UK Ltd | `Acme Holdings : Acme, Inc. : Acme UK Ltd` | GBP | `Acme UK Ltd Reconciliations` |
| Acme Uruguay | `Acme Holdings : Acme, Inc. : Acme Uruguay` | UYU | `Acme Uruguay Reconciliations` |
| Consolidated (USD) | Subsidiary ID `-2` in ns_runReport | USD | n/a |

## SuiteQL WHERE clause snippets

```sql
-- US subsidiary only
sub.fullname = 'Acme Holdings : Acme, Inc.'

-- Canada subsidiary only
sub.fullname = 'Acme Holdings : Acme, Inc. : Acme Canada'

-- Netherlands
sub.fullname = 'Acme Holdings : Acme, Inc. : Acme Netherlands'

-- Poland
sub.fullname = 'Acme Holdings : Acme, Inc. : Acme Poland'

-- UK
sub.fullname = 'Acme Holdings : Acme, Inc. : Acme UK Ltd'

-- Uruguay
sub.fullname = 'Acme Holdings : Acme, Inc. : Acme Uruguay'
```

## Intercompany constants (used by adv-interco-je skill)

The IC line NAMEs and DUE TO/FROM values name the **other** subsidiary, regardless of which sub is source vs target. So on a 121900 line (the source-sub side), the NAME is `IC-{Target}` and DUE TO/FROM is the target's short name.

Reference table — short names and shorthand used in External IDs / Header Memos:

| Subsidiary | Short | NS short name (for NAME / DUE TO/FROM) |
|-----------|-------|----------------------------------------|
| Acme, Inc. (US) | `US` | `Acme, Inc.` |
| Netherlands | `NL` | `Acme Netherlands` |
| UK | `UK` | `Acme UK Ltd` |
| Uruguay | `UY` | `Acme Uruguay` |
| Canada | `CA` | `Acme Canada` |

The `IC-{short_name}` NAME values are formed by prefixing `IC-` to the NS short name above (e.g., `IC-Acme Netherlands`, `IC-Acme, Inc.`).

## IC Accounts — global, named by TARGET subsidiary

The 681xxx series accounts identify the **target** subsidiary, not the source. The same account number is used regardless of which sub's books it sits on. Confirmed via production line data on 2026-05-06 (see JE#####, JE#####, JE#####, JE#####, JE#####, JE##### in NetSuite).

| Account | Full format | Identifies target | Production examples (header sub on left) |
|---------|-------------|-------------------|------------------------------------------|
| 681100 | `681100 Intercompany Expense : Netherlands` | Netherlands | US → NL |
| 681110 | `681110 Intercompany Expense : Canada` | Canada | US → CA |
| **681130** | **`681130 Intercompany Expense : US`** | **US** | **CA → US** (JE#####, JE#####, JE#####) |
| 681140 | `681140 Intercompany Expense : UK` | UK | US → UK; NL → UK (JE#####); CA → UK (JE#####, JE#####) |
| 681150 | `681150 Intercompany Expense : Uruguay` | Uruguay | US → UY |
| 121900 | `121900 Accounts Receivable : Intercompany Receivables` | — (source-side IC line) | All directions |
| 211900 | `211900 Accounts Payable : Accounts Payable : Intercompany Payables` | — (target-side IC line) | All directions |

## F-to-F (and any direction) IC line constants matrix

For any (Source S, Target T) pair, the four-line reclass JE uses the constants below. The 681xxx account number is determined by T (the target). The 121900 line goes on S; the 211900 line goes on T.

| Source S → Target T | 681xxx on S | 121900 NAME (on S) | 121900 DUE TO/FROM (on S) | 211900 NAME (on T) | 211900 DUE TO/FROM (on T) |
|---------------------|-------------|--------------------|---------------------------|--------------------|---------------------------|
| US → NL | 681100 | `IC-Acme Netherlands` | `Acme Netherlands` | `IC-Acme, Inc.` | `Acme, Inc.` |
| US → UK | 681140 | `IC-Acme UK Ltd` | `Acme UK Ltd` | `IC-Acme, Inc.` | `Acme, Inc.` |
| US → UY | 681150 | `IC-Acme Uruguay` | `Acme Uruguay` | `IC-Acme, Inc.` | `Acme, Inc.` |
| US → CA | 681110 | `IC-Acme Canada` | `Acme Canada` | `IC-Acme, Inc.` | `Acme, Inc.` |
| CA → US | 681130 | `IC-Acme, Inc.` | `Acme, Inc.` | `IC-Acme Canada` | `Acme Canada` |
| CA → UK | 681140 | `IC-Acme UK Ltd` | `Acme UK Ltd` | `IC-Acme Canada` | `Acme Canada` |
| CA → NL | 681100 | `IC-Acme Netherlands` | `Acme Netherlands` | `IC-Acme Canada` | `Acme Canada` |
| CA → UY | 681150 | `IC-Acme Uruguay` | `Acme Uruguay` | `IC-Acme Canada` | `Acme Canada` |
| NL → US | 681130 | `IC-Acme, Inc.` | `Acme, Inc.` | `IC-Acme Netherlands` | `Acme Netherlands` |
| NL → UK | 681140 | `IC-Acme UK Ltd` | `Acme UK Ltd` | `IC-Acme Netherlands` | `Acme Netherlands` |
| NL → CA | 681110 | `IC-Acme Canada` | `Acme Canada` | `IC-Acme Netherlands` | `Acme Netherlands` |
| NL → UY | 681150 | `IC-Acme Uruguay` | `Acme Uruguay` | `IC-Acme Netherlands` | `Acme Netherlands` |
| UK → US | 681130 | `IC-Acme, Inc.` | `Acme, Inc.` | `IC-Acme UK Ltd` | `Acme UK Ltd` |
| UK → NL | 681100 | `IC-Acme Netherlands` | `Acme Netherlands` | `IC-Acme UK Ltd` | `Acme UK Ltd` |
| UK → CA | 681110 | `IC-Acme Canada` | `Acme Canada` | `IC-Acme UK Ltd` | `Acme UK Ltd` |
| UK → UY | 681150 | `IC-Acme Uruguay` | `Acme Uruguay` | `IC-Acme UK Ltd` | `Acme UK Ltd` |
| UY → US | 681130 | `IC-Acme, Inc.` | `Acme, Inc.` | `IC-Acme Uruguay` | `Acme Uruguay` |
| UY → NL | 681100 | `IC-Acme Netherlands` | `Acme Netherlands` | `IC-Acme Uruguay` | `Acme Uruguay` |
| UY → UK | 681140 | `IC-Acme UK Ltd` | `Acme UK Ltd` | `IC-Acme Uruguay` | `Acme Uruguay` |
| UY → CA | 681110 | `IC-Acme Canada` | `Acme Canada` | `IC-Acme Uruguay` | `Acme Uruguay` |

Eliminate = `Yes` on 121900 and 211900 lines for every direction. Eliminate = `No` on the 681xxx line and on the actual expense / asset / liability lines.

## Header Subsidiary

The Header Subsidiary equals the **Source** subsidiary S of the reclass (set on the JE header). Production examples:
- US-headered (US → foreign): `Acme Holdings : Acme, Inc.`
- CA-headered (CA → US/UK/...): `Acme Holdings : Acme, Inc. : Acme Canada`
- NL-headered (NL → UK/...): `Acme Holdings : Acme, Inc. : Acme Netherlands`
- UK / UY headered: corresponding full path from the table at the top of this file.
