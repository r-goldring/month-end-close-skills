"""
Per-skill configs for the payroll gut-check review.

Each config defines:
  - subsidiary       : full NetSuite path (matches `sub.fullname` in SuiteQL)
  - subsidiary_id    : numeric NetSuite internal ID
  - currency         : ISO code; gut-check pulls debitforeignamount/creditforeignamount
  - cadence_days     : 14 (bi-weekly) or 31 (monthly); used for stale-baseline gate
  - folder_root      : "Monthly Payroll/Pay Runs/{Country}/" prefix
  - folder_pattern   : MM.DD.YYYY (bi-weekly) | MM-YYYY (most monthlies) | YYYY-MM (Uruguay)
  - csv_glob         : how to find the JE Import CSV inside a pay-run folder
  - workbook_glob    : how to find the backup workbook to write the gut-check tab into
  - special_checks   : list of skill-specific check names (severance_routing, jb_ebitda, ...)

Variance thresholds are in THRESHOLDS (shared across all skills since the
account families are the same Acme Corp COA).
"""

# ============================================================================
# Variance thresholds — applied at both Tier 1 (aggregate) and Tier 2 (per-dept)
# ============================================================================
#
# Flag rule: WARN if abs(delta) > max(floor, pct * mean(prior_2))
#
# Tier 1 floor uses "agg_floor"; Tier 2 floor uses "dept_floor" (~10% of agg).
# For families marked "skip", we do NOT run variance — only sign + dept-routing
# checks. (Bonus and Severance are spiky by nature.)

THRESHOLDS = {
    # account-family key -> dict of (matchers, agg_floor, dept_floor, pct, skip_variance)
    "salaries":         {"prefixes": ["611100", "511100"], "agg_floor": 10000, "dept_floor": 1000, "pct": 0.10, "skip_variance": False},
    "bonus":            {"prefixes": ["611150", "511150"], "agg_floor": 0,     "dept_floor": 0,    "pct": 0,    "skip_variance": True},
    "severance":        {"prefixes": ["611250", "511175"], "agg_floor": 0,     "dept_floor": 0,    "pct": 0,    "skip_variance": True},
    "commission":       {"prefixes": ["611200"],           "agg_floor": 5000,  "dept_floor": 500,  "pct": 0.15, "skip_variance": False},
    "401k_match":       {"prefixes": ["611350", "511250"], "agg_floor": 1000,  "dept_floor": 250,  "pct": 0.12, "skip_variance": False},
    "health_benefits":  {"prefixes": ["611300", "511200"], "agg_floor": 1500,  "dept_floor": 250,  "pct": 0.12, "skip_variance": False},
    "other_benefits":   {"prefixes": ["611400", "511300"], "agg_floor": 500,   "dept_floor": 100,  "pct": 0.15, "skip_variance": False},
    "payroll_tax":      {"prefixes": ["611450", "511350"], "agg_floor": 3000,  "dept_floor": 300,  "pct": 0.10, "skip_variance": False},
    # Liability accounts — aggregate-only (no dept on most US liab lines)
    "liab_payroll":     {"prefixes": ["231200"], "agg_floor": 10000, "dept_floor": None, "pct": 0.05, "skip_variance": False},
    "liab_hsa":         {"prefixes": ["231201"], "agg_floor": 500,   "dept_floor": None, "pct": 0.08, "skip_variance": False},
    "liab_fsa":         {"prefixes": ["231202"], "agg_floor": 500,   "dept_floor": None, "pct": 0.08, "skip_variance": False},
    "liab_401k":        {"prefixes": ["231250"], "agg_floor": 1500,  "dept_floor": None, "pct": 0.08, "skip_variance": False},
    "liab_payroll_tax": {"prefixes": ["231350"], "agg_floor": 5000,  "dept_floor": None, "pct": 0.05, "skip_variance": False},
}

# JB EBITDA reclass thresholds (US-only). Tighter than the general Tier 2 because
# JB amounts are very stable cadence-over-cadence.
JB_THRESHOLDS = {"floor": 100, "pct": 0.05}

# Account-family classifier prefix-list, sorted longest-first to avoid 511 matching
# something that should be a sub-family. Built once from THRESHOLDS.
def classify_account(account_str: str):
    """Return the family key for a Acme Corp GL account number. None if unknown."""
    if not account_str:
        return None
    # account_str typically starts with "611100 ..." — first whitespace-bounded token
    num = account_str.split()[0] if account_str else ""
    for family, spec in THRESHOLDS.items():
        for prefix in spec["prefixes"]:
            if num == prefix or num.startswith(prefix):
                return family
    return None


# ============================================================================
# Per-skill configs
# ============================================================================

SKILL_CONFIGS = {
    "us-payroll": {
        "subsidiary": "Acme Holdings : Acme, Inc.",
        "subsidiary_id": "2",
        "currency": "USD",
        "cadence_days": 14,
        "folder_root": "Monthly Payroll/Pay Runs/US/",
        "folder_pattern": "MM.DD.YYYY",
        "csv_glob": "{folder_name} US Payroll JE Import.csv",
        "workbook_glob": "{folder_name} US Payroll Backup.xlsx",
        "memo_prefix": "US PAYROLL",
        "special_checks": ["severance_routing", "jb_ebitda"],
        "liability_dept_optional": True,  # US liab lines have no dept
    },
    "canada-payroll": {
        "subsidiary": "Acme Holdings : Acme, Inc. : Acme Canada",
        "subsidiary_id": "3",
        "currency": "CAD",
        "cadence_days": 14,
        "folder_root": "Monthly Payroll/Pay Runs/Canada/",
        "folder_pattern": "MM.DD.YYYY",
        "csv_glob": "{folder_name} Canada Payroll JE Import.csv",
        "workbook_glob": "{folder_name} Canada Payroll Backup.xlsx",
        "memo_prefix": "Canada Payroll",
        "special_checks": [],
        "liability_dept_optional": True,
    },
    "germany-payroll": {
        # Germany books to the Netherlands subsidiary
        "subsidiary": "Acme Holdings : Acme, Inc. : Acme Netherlands",
        "subsidiary_id": "4",
        "currency": "EUR",
        "cadence_days": 31,
        "folder_root": "Monthly Payroll/Pay Runs/Germany/",
        "folder_pattern": "MM-YYYY",
        "csv_glob": "{folder_name} Germany Payroll JE Import.csv",
        "workbook_glob": "{folder_name} Germany Payroll Backup.xlsx",
        "memo_prefix": "Germany Payroll",
        "special_checks": [],
        "liability_dept_optional": True,
    },
    "netherlands-payroll": {
        "subsidiary": "Acme Holdings : Acme, Inc. : Acme Netherlands",
        "subsidiary_id": "4",
        "currency": "EUR",
        "cadence_days": 31,
        "folder_root": "Monthly Payroll/Pay Runs/Netherlands/",
        "folder_pattern": "MM-YYYY",
        "csv_glob": "{folder_name} Netherlands Payroll JE Import.csv",
        "workbook_glob": "{folder_name} Netherlands Payroll Backup.xlsx",
        "memo_prefix": "Netherlands Payroll",
        "special_checks": [],
        "liability_dept_optional": True,
    },
    "poland-payroll": {
        "subsidiary": "Acme Holdings : Acme, Inc. : Acme Poland",
        "subsidiary_id": "6",
        "currency": "PLN",
        "cadence_days": 31,
        "folder_root": "Monthly Payroll/Pay Runs/Poland/",
        "folder_pattern": "MM-YYYY",
        "csv_glob": "{folder_name} Poland Payroll JE Import.csv",
        "workbook_glob": "{folder_name} Poland Payroll Backup.xlsx",
        "memo_prefix": "Poland Payroll",
        "special_checks": [],
        "liability_dept_optional": True,
    },
    "uk-payroll": {
        "subsidiary": "Acme Holdings : Acme, Inc. : Acme UK Ltd",
        "subsidiary_id": "8",
        "currency": "GBP",
        "cadence_days": 31,
        "folder_root": "Monthly Payroll/Pay Runs/UK/",
        "folder_pattern": "MM-YYYY",
        "csv_glob": "{folder_name} UK Payroll JE Import.csv",
        "workbook_glob": "{folder_name} UK Payroll Backup.xlsx",
        "memo_prefix": "UK Payroll",
        "special_checks": [],
        "liability_dept_optional": True,
    },
    "uruguay-payroll": {
        "subsidiary": "Acme Holdings : Acme, Inc. : Acme Uruguay",
        "subsidiary_id": "13",
        "currency": "UYU",
        "cadence_days": 31,
        "folder_root": "Monthly Payroll/Pay Runs/Uruguay/",
        "folder_pattern": "YYYY-MM",
        # Uruguay generates multiple CSVs per month (Main / Aguinaldo / Egreso /
        # Extra Run / Extra Run Aguinaldo). Each is gut-checked independently
        # against its memo-bucket peer in prior month.
        "csv_glob": "*Uruguay*Import*.csv",
        "workbook_glob": "{folder_name} Uruguay Payroll Backup.xlsx",
        "memo_prefix": "Uruguay Payroll",
        "memo_buckets": [
            {"key": "main",          "match": ["Uruguay Payroll"], "exclude": ["Aguinaldo", "Egreso", "Extra Run"]},
            {"key": "aguinaldo",     "match": ["Aguinaldo Accrual"], "exclude": ["Extra Run"]},
            {"key": "egreso",        "match": ["Egreso"], "exclude": [], "skip_variance": True},
            {"key": "extra_run",     "match": ["Extra Run"], "exclude": ["Aguinaldo"], "skip_variance": True},
            {"key": "extra_run_agu", "match": ["Extra Run Aguinaldo"], "exclude": [], "skip_variance": True},
        ],
        "special_checks": [],
        "liability_dept_optional": True,
    },
}


def detect_skill_from_folder(folder_path: str) -> str:
    """Resolve a pay-run folder path to its originating payroll skill.

    Examples:
      'Monthly Payroll/Pay Runs/US/04.30.2026/'         -> 'us-payroll'
      'Monthly Payroll/Pay Runs/Canada/04.15.2026/'     -> 'canada-payroll'
      'Monthly Payroll/Pay Runs/Uruguay/2026-04/'       -> 'uruguay-payroll'
    """
    norm = folder_path.replace("\\", "/").rstrip("/")
    for skill, cfg in SKILL_CONFIGS.items():
        root = cfg["folder_root"].rstrip("/")
        if root in norm:
            return skill
    raise ValueError(
        f"Cannot detect payroll skill from folder: {folder_path!r}. "
        f"Expected path containing one of: "
        f"{[c['folder_root'] for c in SKILL_CONFIGS.values()]}"
    )


# ============================================================================
# Account canonical-sign rules (5xxxx/6xxxx debit, 23xxxx credit)
# ============================================================================

def canonical_sign(account_str: str) -> str:
    """Return 'debit' | 'credit' | 'unknown' for the given account.

    Used by the sign-flip check: any line whose actual sign opposes its
    canonical sign is a FAIL.

    Note: NOT every account has a fixed canonical sign.
    - Health Benefits (611300/511200) and Other Benefits (611400/511300) are
      NET accounts — the employer-paid premium debits are netted against
      employee-deduction credits, so they can legitimately net to either side.
    - Accrual-style liabilities (231206 NL Holiday Pay Accrual, 231207 UY
      Accrued PTO, 231171 UY Accrued Bonus PR Tax Liability, 231490 Accrued
      Income Taxes) BUILD via credit and CLEAR via debit; both are valid.
    These accounts return 'unknown' so they're caught by the prior-2 sign-flip
    check instead, not by the static canonical-sign assertion.
    """
    num = account_str.split()[0] if account_str else ""
    if not num:
        return "unknown"
    # Gross-expense accounts — always debit (employer cost; no offsetting EE credit)
    debit_only = {
        "611100", "511100",   # Salaries
        "611150", "511150",   # Bonus
        "611175", "511175",   # Severance
        "611200",             # Commission
        "611250", "511250",   # 401K Match (employer match)
        "611450", "511350",   # Payroll Tax expense
    }
    if num in debit_only:
        return "debit"
    # Standard payroll liability accounts — always credit (only ever
    # accumulated during payroll; cleared later by separate cash payment, not
    # by debits in the same JE).
    credit_only = {
        "231200",   # Payroll Liability
        "231201",   # HSA Payable
        "231202",   # FSA Payable
        "231250",   # 401K Payable
        "231350",   # Payroll Tax Liability
        "231205",   # NL Pension Payable
    }
    if num in credit_only:
        return "credit"
    # Other 23xxxx (accruals like 231206/231207/231171/231490) and everything
    # else falls through to prior-2 sign-flip detection.
    return "unknown"


# ============================================================================
# Department-name normalization (mirrors payroll_mapper._normalize_dept_token
# so dept renames across pay runs don't break comparison)
# ============================================================================

def normalize_dept(dept: str) -> str:
    """Lowercase, collapse whitespace, normalize 'Mgmt' <-> 'Management'."""
    if not dept:
        return ""
    s = dept.strip().lower()
    s = s.replace("management", "mgmt")
    return " ".join(s.split())
