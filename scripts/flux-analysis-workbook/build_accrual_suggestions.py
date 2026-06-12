"""
build_accrual_suggestions.py

Pivot-driven accrual identification. After populate_pivot_template.py loads
the 4 detail report data sheets, this script analyzes vendor-level activity
patterns across M-2 / M-1 / M and emits a draft accrual suggestion list for
the accountant to review.

Rules are codified in:
    Monthly Flux Analysis/.claude/skills/flux-analysis-workbook/references/accrual-detection-rules.md

Output:
1. JSON at  {YYYY-MM}/State/accrual_candidates.json
2. Annotations on the Software / COGS / Contractors / ProfFees pivot tabs'
   static F-L tables in columns M (Suggested $), N (Confidence), O (Reason)

Usage (called from populate_pivot_template.py):
    from build_accrual_suggestions import build_suggestions
    build_suggestions(wb, month_dir, csv_path_prior_accrual=..., posted_je_lookup=...)
"""
from __future__ import annotations

import csv
import json
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable

# ----- Configuration: account-family thresholds -----

THRESHOLDS = {
    "software": 500,
    "contractors": 100,
    "professional_fees": 500,
    "cogs_general": 500,
}

# Always-accrue vendors: hardcoded high-confidence baseline. Maps to account
# family for threshold + the typical monthly amount as a sanity fallback.
ALWAYS_ACCRUE = {
    "IT Hardware Partner Networked Solutions Group, LLC": {"account": "511400", "family": "cogs_general", "typical": 377000},
    "DNS Provider A": {"account": "511400", "family": "cogs_general", "typical": 2166},
    "D and O Broker": {"account": "651250", "family": "professional_fees", "typical": 20259},
    # Health Premium Billing Intermediary flows via vendor bills directly, not always accrued; intentionally not in this list.
}

# Always-skip vendors: never suggest these regardless of pattern.
ALWAYS_SKIP = {
    "Former FSA Administrator",
    "AcquiredCo",
}

# Always-verify vendors: low-confidence suggestion, requires the accountant's judgment.
ALWAYS_VERIFY = {
    "Anthropic",
    "OpenAI, LLC",
    "MICROSOFT",
    "SEMRUSH",
}

# Brex categorical entity labels — filtered upstream by parse_gl_detail, but
# we keep the set here as a defensive check.
BREX_CATEGORICAL = {"lodging", "airfare", "meals", "transportation", "taxi", "hotel", "uber"}

# Map detail-tab sheet name to (account_family, account_filters)
DETAIL_TABS = {
    "Software":            ("software",          ("671000", "671100")),
    "Contractors":         ("contractors",       ("511370", "611700")),
    "Professional Fees":   ("professional_fees", ("651100", "651101")),
    "COGS":                ("cogs_general",      ("511400", "511425", "511450", "511510", "511520", "511550", "511600")),
}

# Mapping from visible-tab name to the hidden data sheet that backs it
DATA_SHEET_FOR = {
    "Software": "GeneralLedgerdetailSoftwa",
    "Contractors": "GeneralLedgerdetailContra",
    "Professional Fees": "GeneralLedgerdetailProfes",
    "COGS": "IncomeStatementDetailCOGS",
}

# Static F-L table column positions on each pivot tab
STATIC_TABLE_COLS = {
    "vendor": 6,        # F
    "m_minus_2": 7,     # G
    "m_minus_1": 8,     # H
    "m_current": 9,     # I
    "variance": 10,     # J
    "abs_variance": 11, # K
    "note": 12,         # L (existing Notes column - left alone)
    "accrual_amount": 13, # M (NEW)
    "confidence": 14,     # N (NEW)
    "reason": 15,         # O (NEW)
}


@dataclass
class AccrualSuggestion:
    vendor: str
    account: str
    department: str
    amount: float
    confidence: str  # "high" | "medium" | "low"
    rule: str        # which rule fired
    reason: str      # short explanation
    approved: bool = False
    edited_by: str = "claude"
    ts: str = ""

    def __post_init__(self):
        if not self.ts:
            self.ts = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


# ----- Data extraction from pivot template -----


def vendor_periods_from_data_sheet(ws, m_minus_2: str, m_minus_1: str, m_current: str) -> dict[str, dict]:
    """
    Walk the hidden data sheet and produce per-vendor (Final Name) sums per period.
    Skip Brex categorical entity labels.
    Returns: { vendor_name: {"m2": $, "m1": $, "m": $, "account_path": "...", "department": "...", "memos": [...] } }
    """
    out: dict[str, dict] = {}

    # Columns per SKILL.md GL Detail layout:
    # A=Account, B=Type, C=Date, D=Doc#, E=Name, F=Entity, G=Final Name,
    # H=Debit, I=Credit, J=Net, K=Memo, L=Message, M=Dept, N=Sub, O=Period
    COL_ACCT, COL_FINAL, COL_NET, COL_MEMO, COL_DEPT, COL_PERIOD = 1, 7, 10, 11, 13, 15
    # For COGS data sheet the schema is different (Amount in col J already negated):
    # A=Financial Row, ..., G=Final Name, J=Amount, M=Acct (Line):Name, N=Department, O=Period
    # Net column J works for both since COGS J is negated already.

    for r in range(8, ws.max_row + 1):
        final_name = ws.cell(r, COL_FINAL).value
        if not final_name:
            continue
        vendor = str(final_name).strip()
        if vendor.lower() in BREX_CATEGORICAL:
            continue
        period = ws.cell(r, COL_PERIOD).value
        if not period:
            continue
        net_val = ws.cell(r, COL_NET).value
        try:
            net = float(net_val) if net_val is not None else 0.0
        except (TypeError, ValueError):
            net = 0.0
        if vendor not in out:
            out[vendor] = {
                "m2": 0.0,
                "m1": 0.0,
                "m": 0.0,
                "account_path": ws.cell(r, COL_ACCT).value or "",
                "department": ws.cell(r, COL_DEPT).value or "",
                "memos": [],
            }
        if period == m_minus_2:
            out[vendor]["m2"] += net
        elif period == m_minus_1:
            out[vendor]["m1"] += net
        elif period == m_current:
            out[vendor]["m"] += net
        memo = ws.cell(r, COL_MEMO).value
        if memo:
            out[vendor]["memos"].append(str(memo)[:200])
    return out


def _extract_account_number(account_path: str) -> str:
    """From '671000 - Software:671100 - Software Subscriptions' return '671100' (most specific)."""
    if not account_path:
        return ""
    # Take the rightmost segment with a 6-digit number
    segments = account_path.split(":")
    for seg in reversed(segments):
        seg = seg.strip()
        if seg and seg[:6].isdigit():
            return seg[:6]
    return ""


# ----- Prior-month accrual lookup -----


def load_prior_csv_amounts(csv_path: Path) -> dict[str, float]:
    """Read last month's Accruals JE Import CSV; return {vendor: total_debit}."""
    if not csv_path.exists():
        return {}
    out: dict[str, float] = {}
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                debit = float(row.get("Debit") or 0)
            except ValueError:
                debit = 0
            if debit <= 0:
                continue  # skip credit lines (they're the offset to 231100)
            vendor = (row.get("Name") or "").strip()
            if not vendor:
                continue
            out[vendor] = out.get(vendor, 0) + debit
    return out


# ----- Rule engine -----


def _classify(vendor: str, periods: dict, family: str, threshold: float,
              prior_amount: float | None) -> AccrualSuggestion | None:
    """Apply rules A-H. Returns a suggestion or None to skip."""
    m2, m1, m = periods["m2"], periods["m1"], periods["m"]
    avg_prior = (m2 + m1) / 2
    acct = _extract_account_number(periods["account_path"])
    dept = periods["department"]

    # Rule G: always-skip
    if vendor in ALWAYS_SKIP:
        return None

    # Rule A: always-accrue
    if vendor in ALWAYS_ACCRUE:
        cfg = ALWAYS_ACCRUE[vendor]
        amount = prior_amount if prior_amount else cfg["typical"]
        return AccrualSuggestion(
            vendor=vendor,
            account=cfg["account"],
            department=dept,
            amount=round(amount, 2),
            confidence="high",
            rule="always_accrue",
            reason=f"Hardcoded recurring monthly accrual. Last month: ${prior_amount:,.0f}." if prior_amount else "Hardcoded recurring monthly accrual.",
        )

    # Rule B: last-month-accrual recurrence (highest signal for non-Rule-A vendors)
    if prior_amount and prior_amount > threshold:
        return AccrualSuggestion(
            vendor=vendor,
            account=acct,
            department=dept,
            amount=round(prior_amount, 2),
            confidence="high",
            rule="recurring_accrual",
            reason=f"Accrued last month ${prior_amount:,.0f}; re-suggest same amount.",
        )

    # Rule H: always-verify
    if any(v.lower() in vendor.lower() for v in ALWAYS_VERIFY):
        if avg_prior > threshold:
            return AccrualSuggestion(
                vendor=vendor,
                account=acct,
                department=dept,
                amount=round(avg_prior, 2),
                confidence="low",
                rule="always_verify",
                reason=f"Avg M-2/M-1 ${avg_prior:,.0f}; verify with the accountant (annual vs monthly).",
            )

    # Rule C: recurring missing-current
    # vendor in M-2 AND M-1, M near zero, gap > threshold
    if m2 > 0 and m1 > 0 and abs(m) < 0.05 * max(m1, 1) and avg_prior > threshold:
        # Confidence: high if M-2 and M-1 are within 20% of each other
        spread_ok = m1 > 0 and abs(m2 - m1) / max(m1, 1) < 0.2
        return AccrualSuggestion(
            vendor=vendor,
            account=acct,
            department=dept,
            amount=round(avg_prior, 2),
            confidence="high" if spread_ok else "medium",
            rule="missing_current",
            reason=f"M-2 ${m2:,.0f} + M-1 ${m1:,.0f} but current month $0; avg ${avg_prior:,.0f}.",
        )

    # Rule D: step-down from prior
    if m2 > 0 and m1 > 0 and m > 0 and m < 0.3 * avg_prior and (avg_prior - m) > threshold:
        return AccrualSuggestion(
            vendor=vendor,
            account=acct,
            department=dept,
            amount=round(avg_prior - m, 2),
            confidence="medium",
            rule="step_down",
            reason=f"M-2 ${m2:,.0f}, M-1 ${m1:,.0f}, current ${m:,.0f} (~{m/avg_prior:.0%} of avg); accrue gap.",
        )

    # Rule E: quarterly cadence (memo-driven)
    memo_blob = " ".join(periods["memos"]).lower()
    if any(kw in memo_blob for kw in ("quarterly", "annual", "renewal")) and m2 == 0 and m1 == 0 and m == 0:
        # Vendor invoiced last quarter but not this 3-month window; no obvious accrual
        return None  # cadence-aware but no signal in current 3-month window

    # Nothing fired
    return None


# ----- Public API -----


def build_suggestions(
    wb,
    month_dir: str | Path,
    m_minus_2_label: str,
    m_minus_1_label: str,
    m_current_label: str,
    prior_csv_path: str | Path | None = None,
    posted_je_lookup: Callable[[str], dict[str, float]] | None = None,
) -> dict:
    """
    Run the suggester. Mutates `wb` in place (writes columns M/N/O on each
    detail tab). Also writes JSON to {month_dir}/State/accrual_candidates.json.

    posted_je_lookup: optional callable that, given the prior month's external
    ID (e.g. "2026-04-ACCRUALS"), returns {vendor: posted_amount}. If both
    posted and CSV amounts exist, posted wins and the diff is noted.

    Returns: summary dict.
    """
    month_dir = Path(month_dir)
    state_dir = month_dir / "State"
    state_dir.mkdir(parents=True, exist_ok=True)

    # Load last month's accrual amounts (CSV + posted JE)
    prior_csv = load_prior_csv_amounts(Path(prior_csv_path)) if prior_csv_path else {}
    prior_posted: dict[str, float] = {}
    if posted_je_lookup:
        try:
            # External ID format: 2026-03-ACCRUALS for the month before m_current
            # We don't know the exact label here; caller is responsible for the lookup.
            prior_posted = posted_je_lookup("prior") or {}
        except Exception:
            prior_posted = {}

    # Merge: posted wins if both exist
    prior_amounts: dict[str, float] = dict(prior_csv)
    for vendor, amt in prior_posted.items():
        prior_amounts[vendor] = amt
    prior_diffs: dict[str, tuple[float, float]] = {
        v: (prior_csv.get(v, 0), prior_posted.get(v, 0))
        for v in prior_amounts
        if v in prior_csv and v in prior_posted and abs(prior_csv[v] - prior_posted[v]) > 0.5
    }

    all_suggestions: list[AccrualSuggestion] = []

    for tab_name, (family, _filters) in DETAIL_TABS.items():
        data_sheet_name = DATA_SHEET_FOR[tab_name]
        if data_sheet_name not in wb.sheetnames:
            continue
        data_ws = wb[data_sheet_name]
        vendors = vendor_periods_from_data_sheet(
            data_ws, m_minus_2_label, m_minus_1_label, m_current_label
        )
        threshold = THRESHOLDS[family]
        tab_suggestions: list[AccrualSuggestion] = []
        for vendor, periods in vendors.items():
            prior = prior_amounts.get(vendor)
            suggestion = _classify(vendor, periods, family, threshold, prior)
            if suggestion is None:
                continue
            # Append CSV-vs-posted diff note if relevant
            if vendor in prior_diffs:
                csv_amt, posted_amt = prior_diffs[vendor]
                suggestion.reason += (
                    f" (Note: last month CSV ${csv_amt:,.0f} != NS posted "
                    f"${posted_amt:,.0f}; using posted.)"
                )
            tab_suggestions.append(suggestion)
            all_suggestions.append(suggestion)

        # Write annotations to the static F-L table on the corresponding visible tab
        if tab_name in wb.sheetnames:
            _annotate_static_table(wb[tab_name], tab_suggestions)

    # Write JSON
    payload = {
        "version": 1,
        "generated_at": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "m_minus_2": m_minus_2_label,
        "m_minus_1": m_minus_1_label,
        "m_current": m_current_label,
        "candidates": [asdict(s) for s in all_suggestions],
    }
    json_path = state_dir / "accrual_candidates.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)

    return {
        "candidates_written": len(all_suggestions),
        "by_confidence": _count_by(all_suggestions, "confidence"),
        "by_rule": _count_by(all_suggestions, "rule"),
        "json_path": str(json_path),
    }


def _count_by(items: list, attr: str) -> dict:
    out: dict[str, int] = {}
    for it in items:
        v = getattr(it, attr)
        out[v] = out.get(v, 0) + 1
    return out


def _annotate_static_table(ws, suggestions: list[AccrualSuggestion]) -> None:
    """Find each suggestion's vendor in column F of the static table and write M/N/O."""
    # Build a vendor → row map from column F (positions 5..~131)
    vendor_to_row: dict[str, int] = {}
    for r in range(5, ws.max_row + 1):
        v = ws.cell(r, STATIC_TABLE_COLS["vendor"]).value
        if v:
            vendor_to_row[str(v).strip()] = r

    # Write headers if not already present
    hdr_row = 4
    if not ws.cell(hdr_row, STATIC_TABLE_COLS["accrual_amount"]).value:
        ws.cell(hdr_row, STATIC_TABLE_COLS["accrual_amount"]).value = "Accrual Suggestion"
        ws.cell(hdr_row, STATIC_TABLE_COLS["confidence"]).value = "Confidence"
        ws.cell(hdr_row, STATIC_TABLE_COLS["reason"]).value = "Reason"

    for s in suggestions:
        row = vendor_to_row.get(s.vendor)
        if not row:
            continue
        ws.cell(row, STATIC_TABLE_COLS["accrual_amount"]).value = s.amount
        ws.cell(row, STATIC_TABLE_COLS["accrual_amount"]).number_format = '$#,##0.00'
        ws.cell(row, STATIC_TABLE_COLS["confidence"]).value = s.confidence.capitalize()
        ws.cell(row, STATIC_TABLE_COLS["reason"]).value = s.reason
