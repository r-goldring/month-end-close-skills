# Shared utilities (`scripts/_shared/`)

Reusable building blocks the skills import. If you're adapting these skills to
your own company — or building a new payroll/JE integration — these are the
pieces worth reusing rather than rewriting.

## Modules

### `je_csv_writer.py`
Writes a NetSuite Journal Entry Import CSV in the exact column order the NetSuite
UI Import Assistant expects (`External ID, Date, Journal Entry Memo, Currency,
Account, Debit, Credit, Line Memo, Subsidiary, Department`).

Handles the gotchas that otherwise cause failed imports:
- Date normalization (`YYYY-MM-DD` string or `date`/`datetime` → `M/D/YYYY`)
- Account-number extraction (splits `"611100 Salaries"` → `611100`)
- Empty (not zero) debit/credit cells
- Quoting the subsidiary path (it contains commas)
- **ASCII enforcement** — raises `ValueError` on any non-ASCII character so an
  em-dash or curly quote can never be silently uploaded

Used by every JE-producing skill. The CSV-upload path (vs. API posting) is the
separation-of-duties control documented in `_shared/approval-required.md`.

```python
from je_csv_writer import write_je_csv
write_je_csv(rows, out_path, header_memo="...", currency="USD")
# Uruguay-style multi-JE bundles: pass external_id_fn so rows sharing an
# External ID get bundled into one JE by the Import Assistant.
```

### `payroll_gut_check.py`
Pre-upload sanity check: compares a freshly generated payroll JE CSV against the
prior 2 same-skill JEs and flags sign flips, abnormal variances, and routing
mistakes before you upload.

Designed for a two-phase flow where the *caller* runs the SuiteQL (e.g., via the
NetSuite MCP) and this module does all the parsing, classification, variance math,
and report formatting. Public API is listed in the module docstring
(`parse_csv`, `find_prior_candidates`, `build_validation_sql`,
`filter_validated_priors`, `build_line_fetch_sql`, `analyze`).

### `preflight_report.py`
Consistent formatted output for the payroll pre-flight mapping check — the step
that confirms every cost center and pay code in a raw payroll export maps to a
department + GL account *before* you run the mapper. Each country's
`check_mappings.py` collects the scan results and hands them here for printing.

### `skill_configs.py`
Per-skill configuration for the gut-check review: subsidiary path + internal ID,
currency, pay cadence (bi-weekly vs. monthly), folder/file naming conventions,
and the per-skill special checks. Variance thresholds live here too, in
`THRESHOLDS`, shared across skills because the account families are the same.

**Adapt this first** when porting to your company: replace the subsidiary paths,
internal IDs, and folder conventions with your own.

## Design patterns worth stealing

Even if you don't use this code directly, these patterns transfer to any
ERP-automation project:

- **CSV upload over API write** for anything requiring approval (separation of duties).
- **Preflight before transform** — validate all mappings exist before producing a JE.
- **Gut-check against history** — compare new output to the prior N runs to catch anomalies.
- **ASCII-only output** — fail loudly on non-ASCII rather than corrupting an upload.
