"""Per-entity configuration for the monthly-cfs skill.

Source of truth for: file naming, tab naming, subsidiary IDs, currencies,
NetSuite report layout grammar (US-style vs UK-style editions), FX cell
placement, and manual-cell derivation rules.

All amounts in the _data JSON files are RAW GL convention (debit - credit).
Display sign conventions are applied by the tab builders via the layout
grammar below.
"""

import datetime
import os

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", ".."))
CFS_ROOT = os.path.join(REPO_ROOT, "Monthly CFS")
DATA_DIR = os.path.join(CFS_ROOT, "_data")

# ---------------------------------------------------------------------------
# Months / periods
# ---------------------------------------------------------------------------

def ym(year, month):
    return f"{year:04d}-{month:02d}"

def prev_ym(ym_str):
    y, m = int(ym_str[:4]), int(ym_str[5:7])
    return ym(y - 1, 12) if m == 1 else ym(y, m - 1)

def month_end(ym_str):
    y, m = int(ym_str[:4]), int(ym_str[5:7])
    if m == 12:
        return datetime.date(y, 12, 31)
    return datetime.date(y, m + 1, 1) - datetime.timedelta(days=1)

def period_name(ym_str):
    """NetSuite accounting period name, e.g. 'Apr 2026'."""
    return month_end(ym_str).strftime("%b %Y")

def fiscal_year_start(ym_str):
    """Acme Corp FY = calendar year (validated by regen_history_test)."""
    return f"{ym_str[:4]}-01"

# ---------------------------------------------------------------------------
# Entities
# ---------------------------------------------------------------------------
# layout: "us"  -> Income / Cost Of Sales / Expense / Other Income and Expenses
#         "uk"  -> Sales / Purchases / Overheads / Other Expenses
# bs_layout: "us" -> ASSETS .. Liabilities & Equity .. Total Equity
#            "uk" -> Current Assets .. Capital and Reserves
# fx_cells: where the period FX rates live on the USD CFS tab of the
#           subsidiary workbook: (prior_cell, current_cell). BV keeps labels in
#           row 1 and rates in row 2; everyone else keeps rates in row 1.
# data_row_offset: first data row of the USD CFS tab is row 3 + offset.

ENTITIES = {
    "US": {
        "file_index": 6,
        "file_label": "US",
        "subsidiary_id": 2,
        "entity_name": "Acme, Inc.",
        "entity_path": "Acme Holdings : Acme, Inc.",
        "currency": "USD",
        "currency_symbol": "$",
        "is_layout": "us",
        "bs_layout": "us",
        "is_tab": "Income Statement",
        "bs_tab": "ACSBalanceSheetFiscalYear",
        # (template_pattern, currency_suffix). US has a single CFS tab in USD.
        "cfs_tabs": [("{ym} CFS", "USD")],
        "usd_cfs_tab": "{ym} CFS",
        "fx_cells": None,
        "legacy_tabs": [],
    },
    "BV": {
        "file_index": 2,
        "file_label": "BV",
        "subsidiary_id": 4,
        "entity_name": "Acme, Inc.",
        "entity_path": "Acme Holdings : Acme, Inc. : Acme Netherlands",
        "currency": "EUR",
        "currency_symbol": "€",
        "is_layout": "uk",
        "bs_layout": "uk",
        "is_tab": "3. Income Statement",
        "bs_tab": "4. Balance Sheet",
        "cfs_tabs": [("2. {ym} CFS EUR", "EUR")],
        "usd_cfs_tab": "1. {ym} CFS USD",
        "fx_cells": {"prior": "B2", "current": "C2",
                     "prior_label": "B1", "current_label": "C1"},
        "legacy_tabs": ["2023-02 CFS"],
    },
    "CAD": {
        "file_index": 3,
        "file_label": "CAD",
        "subsidiary_id": 3,
        "entity_name": "Acme, Inc.",
        "entity_path": "Acme Holdings : Acme, Inc. : Acme Canada",
        "currency": "CAD",
        "currency_symbol": "C$",
        "is_layout": "us",
        "bs_layout": "us",
        "is_tab": "Income Statement",
        "bs_tab": "ACSBalanceSheetFiscalYear",
        "cfs_tabs": [("{ym} CFS CAD", "CAD")],
        "usd_cfs_tab": "{ym} CFS USD",
        "fx_cells": {"prior": "B1", "current": "C1"},
        "legacy_tabs": ["2023-04 CFS"],
    },
    "PL": {
        "file_index": 4,
        "file_label": "Poland",
        "subsidiary_id": 6,
        "entity_name": "Acme, Inc.",
        "entity_path": "Acme Holdings : Acme, Inc. : Acme Poland",
        "currency": "PLN",
        "currency_symbol": "zł",
        "is_layout": "uk",
        "bs_layout": "uk",
        "is_tab": "Income Statement",
        "bs_tab": "ACSBalanceSheetFiscalYear",
        "cfs_tabs": [("{ym} CFS PLN", "PLN")],
        "usd_cfs_tab": "{ym} CFS USD",
        "fx_cells": {"prior": "B1", "current": "C1"},
        "legacy_tabs": [],
    },
    "UK": {
        "file_index": 5,
        "file_label": "UK",
        "subsidiary_id": 8,
        "entity_name": "Acme, Inc.",
        "entity_path": "Acme Holdings : Acme, Inc. : Acme UK Ltd",
        "currency": "GBP",
        "currency_symbol": "£",
        "is_layout": "uk",
        "bs_layout": "uk",
        "is_tab": "Income Statement",
        "bs_tab": "ACSBalanceSheetFiscalYear",
        "cfs_tabs": [("{ym} CFS GBP", "GBP")],
        "usd_cfs_tab": "{ym} CFS USD",
        "fx_cells": {"prior": "B1", "current": "C1"},
        "legacy_tabs": [],
    },
    "UY": {
        "file_index": 7,
        "file_label": "UY",
        "subsidiary_id": 13,
        "entity_name": "Acme, Inc.",
        "entity_path": "Acme Holdings : Acme, Inc. : Acme Uruguay",
        "currency": "UYU",
        "currency_symbol": "$U",
        "is_layout": "uk",
        "bs_layout": "uk",
        "is_tab": "Income Statement",
        "bs_tab": "ACSBalanceSheetFiscalYear",
        "cfs_tabs": [("{ym} CFS UYU", "UYU")],
        "usd_cfs_tab": "{ym} CFS USD",
        "fx_cells": {"prior": "B1", "current": "C1"},
        "legacy_tabs": ["2023-02 CFS"],
    },
}

# Consolidated workbook (file index 1). Per-subsidiary USD tabs mirror the
# subsidiary workbooks via external links; BV tab data is offset +1 row
# because its FX labels occupy row 1 and rates row 2.
CONSOLIDATED = {
    "file_index": 1,
    "file_label": "Consolidated",
    "tab_order": ["{ym} CFS Consolidated", "{ym} CFS US", "{ym} CFS BV USD",
                  "{ym} CFS CAN USD", "{ym} CFS PL USD", "{ym} CFS UK USD",
                  "{ym} CFS UY USD"],
    # entity key -> (tab pattern, fx prior cell, fx current cell, row offset)
    "sub_tabs": {
        "BV":  ("{ym} CFS BV USD",  "B2", "C2", 1),
        "CAD": ("{ym} CFS CAN USD", "B1", "C1", 0),
        "PL":  ("{ym} CFS PL USD",  "B1", "C1", 0),
        "UK":  ("{ym} CFS UK USD",  "B1", "C1", 0),
        "UY":  ("{ym} CFS UY USD",  "B1", "C1", 0),
    },
}

# Currency code -> NetSuite consolidatedexchangerate fromsubsidiary name
FX_SUBSIDIARY = {
    "EUR": "Acme Netherlands",
    "CAD": "Acme Canada",
    "PLN": "Acme Poland",
    "GBP": "Acme UK Ltd",
    "UYU": "Acme Uruguay",
}

# ---------------------------------------------------------------------------
# Report layout grammar
# ---------------------------------------------------------------------------
# Display sign: +1 means displayed = (debit - credit), -1 means (credit - debit).
# Sections are emitted in order; a section is omitted when it has no accounts
# (and its slot in computed formulas becomes a literal 0, matching NetSuite).

IS_LAYOUTS = {
    "us": {
        "preamble": "Ordinary Income/Expense",   # bold row before first section
        "sections": [
            # (label, accttypes, sign)
            ("Income",        ("Income",),     -1),
            ("Cost Of Sales", ("COGS",),       +1),
            # Gross Profit = Total Income - Total COS
            ("Expense",       ("Expense",),    +1),
            # Net Ordinary Income = Gross Profit - Total Expense
        ],
        "other_preamble": "Other Income and Expenses",
        "other_sections": [
            ("Other Income",  ("OthIncome",),  -1),
            ("Other Expense", ("OthExpense",), +1),
        ],
        "gross_profit_label": "Gross Profit",
        "operating_label": "Net Ordinary Income",
        # Net Other Income = [Total Other Income | 0] - [Total Other Expense | 0]
        "net_other_label": "Net Other Income",
        "net_label": "Net Income",   # = operating + net other
    },
    "uk": {
        "preamble": None,
        "sections": [
            ("Sales",     ("Income",),  -1),
            ("Purchases", ("COGS",),    +1),
            ("Overheads", ("Expense",), +1),
        ],
        "other_preamble": None,
        "other_sections": [
            ("Other Income",   ("OthIncome",),  -1),
            ("Other Expenses", ("OthExpense",), -1),   # credit-positive, ADDED
        ],
        "gross_profit_label": "Gross Profit",
        "operating_label": "Operating Profit",
        "net_other_label": None,
        # Net Profit/(Loss) = Operating + [Total Other Income | 0] + [Total Other Expenses | 0]
        "net_label": "Net Profit/(Loss)",
    },
}

# BS subsection labels in display order, with the accttypes they hold.
# Sign: assets +1 (debit-positive), liabilities/equity -1 (credit-positive).
BS_SUBSECTIONS = [
    ("Bank",                 "Bank",            +1),
    ("Accounts Receivable",  "AcctRec",         +1),
    ("Unbilled Receivable",  "UnbilledRec",     +1),
    ("Other Current Asset",  "OthCurrAsset",    +1),
    ("Fixed Assets",         "FixedAsset",      +1),
    ("Other Assets",         "OthAsset",        +1),
    ("Accounts Payable",     "AcctPay",         -1),
    ("Credit Card",          "CreditCard",      -1),
    ("Other Current Liability", "OthCurrLiab",  -1),
    ("Long Term Liabilities", "LongTermLiab",   -1),
    ("Equity",               "Equity",          -1),
]

ACCTTYPE_TO_BS_SUBSECTION = {t: lbl for lbl, t, _ in BS_SUBSECTIONS}
BS_SIGN = {t: sign for _, t, sign in BS_SUBSECTIONS}

IS_ACCTTYPES = ("Income", "COGS", "Expense", "OthIncome", "OthExpense")

# ---------------------------------------------------------------------------
# Number formats / styling
# ---------------------------------------------------------------------------

def is_number_format(symbol):
    s = symbol
    return f'"{s}"#,##0.00_);\\("{s}"#,##0.00\\)'

def bs_number_format(symbol):
    return f'"{symbol}"#,##0.00'

STYLE = {
    "title_font": ("Arial", 12, True),
    "report_font": ("Arial", 14, True),
    "header_font": ("Arial", 7, True),
    "header_fill": "FFD0D0D0",
    "body_font": ("Arial", 8, False),
    "bold_font": ("Arial", 8, True),
    "col_a_width_is": 50.78,
    "col_a_width_bs": 59.33,
    # foreign-currency amount columns need width 11+ so the symbol never clips
    "amount_col_width_usd": 14.11,
    "amount_col_width_foreign": 17.11,
}

YELLOW_FILL = "FFFFFF00"   # manual-cell highlight

# ---------------------------------------------------------------------------
# Manual-cell derivation rules for the CFS tabs
# ---------------------------------------------------------------------------
# Scanned manual (constant) cells inside the CFS matrix get derived as the
# value that zeroes that column's Check Figure (row 49) -- the "column plug"
# rule the accountant approved -- EXCEPT cells listed here, which have specific rules
# applied first. Any remaining residual goes to the column's designated plug
# cell. Cells with rule "carry" keep the prior-month value and get flagged.
#
#   bs_value:   value = sign * BS balance of acct at (which) month
#   bs_delta:   value = sign * (BS current - BS prior) of acct
#   column_plug: value that zeroes the column check (default for lone cells)
#   carry:      carry prior value forward, flag for review
# sign is applied to the RAW GL value (debit - credit), which is negative for
# credit balances; signs below are calibrated against the accountant's 2026-04 workbook.
MANUAL_CELL_RULES = {
    "US": {
        "S14": {"rule": "bs_value", "acct": "261277", "which": "current", "sign": +1,
                "note": "Non-cash accumulated interest = -(261277 Accumulated Interest, current month)"},
        "S28": {"rule": "bs_value", "acct": "261277", "which": "prior", "sign": -1,
                "note": "Reversal of prior accumulated interest = +(261277, prior month)"},
        "W40": {"rule": "bs_delta", "acct": "261270", "which": "delta", "sign": +1,
                "note": "Payment of term loans = change in 261270 Term Loan - Non-current"},
        "B53": {"rule": "carry", "note": "Note row: cash interest paid during month"},
        "B54": {"rule": "is_value", "acct": "711000", "which": "current", "sign": -1,
                "note": "Note row: cash interest income received = IS 711000 current month"},
    },
}

# Default plug cells per column when a residual remains after specific rules.
# F = PP&E column (capex plug, row 34); row 47 = FX effect plug for the
# cash column (B) on foreign subs.
DEFAULT_PLUGS = {
    "F": 34,    # capital expenditures
}
FX_EFFECT_ROW = 47

# ---------------------------------------------------------------------------
# File naming
# ---------------------------------------------------------------------------

def workbook_name(entity_key, ym_str, prep_date):
    if entity_key == "CONS":
        idx, label = CONSOLIDATED["file_index"], CONSOLIDATED["file_label"]
        suffix = " (FX embedded)"
    else:
        e = ENTITIES[entity_key]
        idx, label = e["file_index"], e["file_label"]
        suffix = ""
    d = f"{prep_date.month}.{prep_date.day}.{prep_date.year}"
    return f"{idx}. {ym_str} CFS {label} {d}{suffix}.xlsx"

def month_dir(ym_str):
    return os.path.join(CFS_ROOT, ym_str)

def find_workbook(entity_key, ym_str):
    """Locate the workbook for an entity+month regardless of prep date."""
    idx = CONSOLIDATED["file_index"] if entity_key == "CONS" else ENTITIES[entity_key]["file_index"]
    d = month_dir(ym_str)
    if not os.path.isdir(d):
        return None
    label = CONSOLIDATED["file_label"] if entity_key == "CONS" else ENTITIES[entity_key]["file_label"]
    matches = [f for f in os.listdir(d)
               if f.startswith(f"{idx}. {ym_str} CFS {label} ") and f.endswith(".xlsx")
               and not f.startswith("~$")]
    # exclude one-off variants like "- Claude Update"
    clean = [f for f in matches if "Claude" not in f]
    pick = sorted(clean or matches)
    return os.path.join(d, pick[-1]) if pick else None
