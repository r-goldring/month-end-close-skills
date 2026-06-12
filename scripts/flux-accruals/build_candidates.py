"""
Build Accrual_Reclass_Candidates_{YYYY-MM}.xlsx for the requested closing month.

CLI usage:
    python build_candidates.py 2026-04 [--repo-root .]

Inputs (all under the repo root):
  Monthly Flux Analysis/Vendor Budget v*.xlsx
  Monthly Flux Analysis/{YYYY}/{YYYY-MM}/_cache/{537|540|542|721}.json
  Monthly Flux Analysis/{YYYY}/{YYYY-MM}/BillFlow Export/TransactionListByExportStatus*.csv  (optional)
  Monthly Flux Analysis/{YYYY}/{YYYY-MM}/_cache/billcom_already_in_ns.json                   (optional)

Outputs:
  Monthly Flux Analysis/{YYYY}/{YYYY-MM}/Accrual_Reclass_Candidates_{YYYY-MM}.xlsx

Claude is responsible for populating the _cache/ JSONs by running ns_runReport
and ns_runCustomSuiteQL before invoking this script. See SKILL.md.
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, List, Optional, Tuple

SCRIPT_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(SCRIPT_DIR))

from load_vendor_budget import (
    BudgetRow, VendorBudget, load_vendor_budget, find_latest_budget_file,
    _normalize_name, MONTH_COLS, classify_category,
)
from parse_reports import (
    TransactionRecord, load_cache_dir, aggregate_vendor_period,
    ACCOUNT_FILTERS,
)
from load_billcom_export import (
    BillcomBill, find_billcom_csv, load_billcom_csv, mark_already_in_netsuite,
    aggregate_by_vendor as aggregate_billcom_by_vendor,
)
from candidate_workbook import (
    AccrualCandidate, ReclassCandidate, CandidateBundle, write_workbook,
)


# Thresholds mirror references/accrual-thresholds.md (CLAUDE.md authoritative)
THRESHOLDS = {
    "Software":    750.0,
    "Contractors": 100.0,
    "ProfFees":    500.0,
    "COGS":        500.0,
}
PARTIAL_BASELINE_RATIO = 0.50    # actual < 50% of budget triggers partial signal
OVER_BUDGET_RATIO = 1.20         # actual > 120% of budget triggers review (informational)
SOFTWARE_ACCOUNTS = {"671100", "671000", "511425"}
FUZZY_THRESHOLD = 0.85           # SequenceMatcher.ratio()


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Build flux-accruals candidate workbook.")
    p.add_argument("period", help="Closing period as YYYY-MM (e.g., 2026-04).")
    p.add_argument("--repo-root", default=".", help="Repo root path (default: cwd).")
    return p.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    period = args.period
    if not _is_valid_period(period):
        print(f"ERROR: invalid period {period!r}; expected YYYY-MM", file=sys.stderr)
        return 2

    repo = Path(args.repo_root).resolve()
    flux = repo / "Monthly Flux Analysis"
    year = period[:4]
    month_dir = flux / year / period
    cache_dir = month_dir / "_cache"

    if not cache_dir.exists():
        print(f"ERROR: cache dir not found: {cache_dir}\n"
              f"  Claude must populate {{537,540,542,721}}.json there first "
              f"(via ns_runReport).", file=sys.stderr)
        return 2

    # 1) Vendor budget
    budget_path = find_latest_budget_file(flux)
    print(f"[1/5] Loading vendor budget: {budget_path.name}")
    vb = load_vendor_budget(budget_path)
    print(f"      {len(vb.rows)} budget vendors loaded.")

    # 2) NetSuite report actuals (4 reports, multi-month data)
    print(f"[2/5] Loading NetSuite report cache from: {cache_dir}")
    records = load_cache_dir(cache_dir)
    print(f"      {len(records)} transaction records parsed.")
    if not records:
        print("      WARNING: no records parsed - check cache JSONs.", file=sys.stderr)

    by_period_cat = aggregate_vendor_period(records)
    # Also keep raw records for dept-drift (need original dept per posting)
    current_period_records = [r for r in records if r.period_yyyymm == period]
    print(f"      {len(current_period_records)} records in {period}.")

    # 3) BillFlow export (optional)
    bc_csv = find_billcom_csv(month_dir)
    bc_open: List[BillcomBill] = []
    bc_already: List[BillcomBill] = []
    if bc_csv:
        print(f"[3/5] Loading BillFlow export: {bc_csv.name}")
        bc_bills = load_billcom_csv(bc_csv)
        print(f"      {len(bc_bills)} bills (Type=Bill, Amount>0).")
        ns_bills = _load_already_in_ns(cache_dir)
        if ns_bills:
            bc_open, bc_already = mark_already_in_netsuite(bc_bills, ns_bills)
            print(f"      {len(bc_already)} already in NetSuite, {len(bc_open)} truly open.")
        else:
            print(f"      (no billcom_already_in_ns.json cached - skipping NS validation; "
                  f"all bills treated as open)")
            bc_open = bc_bills
    else:
        print(f"[3/5] No BillFlow Export folder/CSV found - skipping BillFlow merge.")

    bc_by_vendor = aggregate_billcom_by_vendor(bc_open) if bc_open else {}

    # 4) Compute candidates
    print(f"[4/5] Computing accrual + reclass candidates for {period}.")
    bundle = _build_bundle(period, vb, records, by_period_cat,
                           current_period_records, bc_by_vendor, bc_already)

    # 5) Write workbook
    out_path = month_dir / f"Accrual_Reclass_Candidates_{period}.xlsx"
    print(f"[5/5] Writing candidate workbook: {out_path}")
    write_workbook(bundle, out_path)

    _print_summary(bundle)
    return 0


def _is_valid_period(p: str) -> bool:
    if len(p) != 7 or p[4] != "-":
        return False
    try:
        y, m = int(p[:4]), int(p[5:])
        return 1 <= m <= 12 and 2020 <= y <= 2099
    except ValueError:
        return False


def _load_already_in_ns(cache_dir: Path) -> List[dict]:
    """
    Read cache/billcom_already_in_ns.json (Claude populates from SuiteQL).
    Format: list of {"tranid": ..., "companyname": ..., "trandate": ...}
    """
    path = cache_dir / "billcom_already_in_ns.json"
    if not path.exists():
        return []
    import json
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict) and "rows" in data:
        return data["rows"]
    return data if isinstance(data, list) else []


def _period_label(period: str) -> str:
    y, m = int(period[:4]), int(period[5:])
    return dt.date(y, m, 1).strftime("%B %Y")


def _previous_periods(period: str, n: int) -> List[str]:
    """Return [M-n, ..., M-1] as YYYY-MM strings."""
    y, m = int(period[:4]), int(period[5:])
    out = []
    for i in range(n, 0, -1):
        prev_month = ((m - 1 - i) % 12) + 1
        prev_year = y + (m - 1 - i) // 12
        out.append(f"{prev_year:04d}-{prev_month:02d}")
    return out


def _build_bundle(period: str,
                  vb: VendorBudget,
                  records: List[TransactionRecord],
                  by_period_cat: Dict[Tuple[str, str, str, str], float],
                  current_period_records: List[TransactionRecord],
                  bc_by_vendor: Dict[str, dict],
                  bc_already: List[BillcomBill]) -> CandidateBundle:
    period_idx = int(period[5:])
    label = _period_label(period)
    prior2 = _previous_periods(period, 2)   # [M-2, M-1] - matches 3-month NS pull window

    accruals: Dict[str, List[AccrualCandidate]] = {
        "Software": [], "Contractors": [], "ProfFees": [], "COGS": [],
    }
    seen_budget_vendor_ids: set[int] = set()

    # --- Accrual identification (budget-driven) ---
    for budget_row in vb.rows:
        cat = budget_row.category
        if cat not in accruals:
            continue
        seen_budget_vendor_ids.add(budget_row.vendor_id)

        budget_amt = budget_row.budget_for(period_idx)
        actual_amt = _sum_actuals_for_budget_row(budget_row, period, by_period_cat)
        gap = budget_amt - actual_amt
        thresh = THRESHOLDS[cat]

        signal = _classify_signal(budget_amt, actual_amt, gap, thresh)
        if signal is None:
            continue   # vendor is fine

        suggested = _suggested_amount(signal, budget_amt, actual_amt, gap)
        confidence = "MED"
        notes_parts = []

        # BillFlow confirmation override
        bc_match = _match_billcom(budget_row, bc_by_vendor)
        if bc_match is not None:
            signal = "billcom_confirmed"
            suggested = bc_match["total"]
            confidence = "HIGH"
            notes_parts.append(f"bill.com inv: {', '.join(bc_match['invoice_numbers'])}")

        if signal == "missing_budgeted" and budget_amt >= thresh * 5:
            confidence = "HIGH"

        # Trailing 2-month actuals for context (M-2, M-1)
        trailing = [_sum_actuals_for_budget_row(budget_row, p, by_period_cat) for p in prior2]

        if budget_row.contract_end_date:
            notes_parts.append(f"contract ends {budget_row.contract_end_date}")
        if budget_row.comment:
            notes_parts.append(budget_row.comment[:120])

        accruals[cat].append(AccrualCandidate(
            vendor_name=budget_row.vendor_name,
            vendor_id=budget_row.vendor_id,
            account=f"{budget_row.account_number} - {budget_row.account_name}".strip(" -"),
            department=budget_row.department,
            budget_amount=round(budget_amt, 2),
            actual_amount=round(actual_amt, 2),
            gap=round(gap, 2),
            trailing_3mo=[round(x, 2) for x in trailing],
            signal=signal,
            suggested_amount=round(suggested, 2),
            confidence=confidence,
            notes=" | ".join(notes_parts),
        ))

    # Sort each category by suggested_amount descending
    for cat in accruals:
        accruals[cat].sort(key=lambda x: (-x.suggested_amount, x.vendor_name))

    # --- Dept drift (non-software accounts only) ---
    dept_drift = _build_dept_drift(vb, current_period_records)

    # --- Software reclass (all software-account activity) ---
    software_reclass = _build_software_reclass(vb, current_period_records)

    # --- Unmatched actuals (vendors in actuals but not in budget) ---
    unmatched_actuals = _build_unmatched_actuals(vb, current_period_records)

    # --- Billcom_AlreadyInNS audit tab rows ---
    billcom_already = [
        {"vendor": b.vendor_name,
         "invoice_number": b.invoice_number,
         "amount": b.amount,
         "netsuite_tranid": b.netsuite_tranid,
         "approval_status": b.approval_status}
        for b in bc_already
    ]

    # --- Billcom_Unmatched: open bills whose vendor didn't match any budget vendor ---
    billcom_unmatched = []
    matched_keys = {_normalize_name(r.vendor_name) for r in vb.rows}
    for norm_name, info in bc_by_vendor.items():
        # Try fuzzy match against any budget vendor
        if norm_name in matched_keys or _fuzzy_match_in(norm_name, matched_keys, FUZZY_THRESHOLD):
            continue
        billcom_unmatched.append({
            "vendor": info["vendor"],
            "invoice_number": ", ".join(info["invoice_numbers"]),
            "amount": info["total"],
            "approval_status": "",
        })

    return CandidateBundle(
        period=period,
        period_label=label,
        accruals=accruals,
        dept_drift=dept_drift,
        software_reclass=software_reclass,
        unmatched_actuals=unmatched_actuals,
        billcom_already_in_ns=billcom_already,
        billcom_unmatched=billcom_unmatched,
    )


def _sum_actuals_for_budget_row(budget_row: BudgetRow, period: str,
                                by_period_cat: Dict[Tuple[str, str, str, str], float]) -> float:
    """Sum all actual transactions for this vendor in this period within their category."""
    cat = budget_row.category
    norm = _normalize_name(budget_row.vendor_name)
    total = 0.0
    # Match on (period, category, normalized_vendor) ignoring account — vendor's
    # actuals might be split across accounts within the category (e.g., 671100 + 511425).
    for (p, c, v, _acct), amt in by_period_cat.items():
        if p == period and c == cat and v == norm:
            total += amt
    return total


def _classify_signal(budget_amt: float, actual_amt: float, gap: float,
                     threshold: float) -> Optional[str]:
    if budget_amt > 0 and abs(actual_amt) < 1.0:
        return "missing_budgeted"
    if budget_amt > 0 and actual_amt > 0:
        if actual_amt < budget_amt * PARTIAL_BASELINE_RATIO and gap > threshold:
            return "partial_below_baseline"
        if actual_amt > budget_amt * OVER_BUDGET_RATIO and (actual_amt - budget_amt) > threshold:
            return "over_budget_review"
    return None


def _suggested_amount(signal: str, budget_amt: float, actual_amt: float, gap: float) -> float:
    if signal == "missing_budgeted":
        return budget_amt
    if signal == "partial_below_baseline":
        return max(gap, 0.0)
    return 0.0


def _match_billcom(budget_row: BudgetRow,
                   bc_by_vendor: Dict[str, dict]) -> Optional[dict]:
    if not bc_by_vendor:
        return None
    norm = _normalize_name(budget_row.vendor_name)
    if norm in bc_by_vendor:
        return bc_by_vendor[norm]
    # Fuzzy fallback
    best = None
    best_score = 0.0
    for k, v in bc_by_vendor.items():
        score = SequenceMatcher(None, norm, k).ratio()
        if score >= FUZZY_THRESHOLD and score > best_score:
            best, best_score = v, score
    return best


def _fuzzy_match_in(name: str, candidates: set[str], threshold: float) -> bool:
    for c in candidates:
        if SequenceMatcher(None, name, c).ratio() >= threshold:
            return True
    return False


def _build_dept_drift(vb: VendorBudget,
                      current_period_records: List[TransactionRecord]) -> List[ReclassCandidate]:
    """
    Per (vendor, account, dept) bucket of current-month activity, flag where
    actual dept != expected dept. Skip software accounts (handled separately).
    """
    bucket: Dict[Tuple[str, str, str, str], float] = defaultdict(float)
    for r in current_period_records:
        if r.account_number in SOFTWARE_ACCOUNTS:
            continue
        if not r.vendor or not r.department:
            continue
        bucket[(_normalize_name(r.vendor), r.vendor, r.account_number, r.department)] += r.amount

    out: List[ReclassCandidate] = []
    for (norm, display, acct, actual_dept), amt in bucket.items():
        if abs(amt) < 1.0:
            continue   # debit + credit cancelled in the same dept; nothing to reclass
        budget_row = vb.lookup(vendor_name=norm)
        if budget_row is None:
            continue
        expected_dept = budget_row.department
        if not expected_dept or expected_dept == actual_dept:
            continue
        out.append(ReclassCandidate(
            vendor_name=display,
            account=f"{acct} - {budget_row.account_name}".strip(" -") if budget_row.account_number == acct else acct,
            actual_dept=actual_dept,
            expected_dept=expected_dept,
            actual_amount=round(amt, 2),
            reason=f"FP&A budget dept = {expected_dept}; posted to {actual_dept}",
        ))
    out.sort(key=lambda x: -abs(x.actual_amount))
    return out


def _build_software_reclass(vb: VendorBudget,
                            current_period_records: List[TransactionRecord]) -> List[ReclassCandidate]:
    """
    All current-month activity on software accounts (671xxx + 511425) where
    actual dept != budget dept. Source-agnostic - bills, amortization JEs, etc.
    """
    bucket: Dict[Tuple[str, str, str, str], float] = defaultdict(float)
    for r in current_period_records:
        if r.account_number not in SOFTWARE_ACCOUNTS:
            continue
        if not r.vendor or not r.department:
            continue
        bucket[(_normalize_name(r.vendor), r.vendor, r.account_number, r.department)] += r.amount

    out: List[ReclassCandidate] = []
    for (norm, display, acct, actual_dept), amt in bucket.items():
        if abs(amt) < 1.0:
            continue   # debit + credit cancelled; nothing to reclass
        budget_row = vb.lookup(vendor_name=norm)
        if budget_row is None:
            continue
        expected_dept = budget_row.department
        if not expected_dept or expected_dept == actual_dept:
            continue
        # account drift bonus reason
        account_drift = budget_row.account_number not in SOFTWARE_ACCOUNTS and acct in SOFTWARE_ACCOUNTS
        reason = f"FP&A budget dept = {expected_dept}; posted to {actual_dept}"
        if account_drift:
            reason += f" (also account drift: budget={budget_row.account_number}, actual={acct})"
        out.append(ReclassCandidate(
            vendor_name=display,
            account=f"{acct} - software",
            actual_dept=actual_dept,
            expected_dept=expected_dept,
            actual_amount=round(amt, 2),
            reason=reason,
        ))
    out.sort(key=lambda x: -abs(x.actual_amount))
    return out


def _build_unmatched_actuals(vb: VendorBudget,
                             current_period_records: List[TransactionRecord]) -> List[dict]:
    bucket: Dict[Tuple[str, str, str, str], float] = defaultdict(float)
    for r in current_period_records:
        if not r.vendor:
            continue
        bucket[(r.category, r.vendor, r.account_number, r.department)] += r.amount

    out: List[dict] = []
    for (cat, vendor, acct, dept), amt in bucket.items():
        if vb.lookup(vendor_name=vendor) is not None:
            continue
        out.append({
            "category": cat,
            "vendor": vendor,
            "account": acct,
            "department": dept,
            "actual_amount": round(amt, 2),
        })
    out.sort(key=lambda x: -abs(x["actual_amount"]))
    return out


def _print_summary(bundle: CandidateBundle) -> None:
    print()
    print(f"=== Summary - {bundle.period_label} ===")
    grand_count, grand_total = 0, 0.0
    for cat in ("Software", "Contractors", "ProfFees", "COGS"):
        rows = bundle.accruals.get(cat, [])
        approve_rows = [r for r in rows if r.signal not in ("over_budget_review", "unbudgeted_vendor")]
        total = sum(r.suggested_amount for r in approve_rows)
        print(f"  {cat:13s}  {len(approve_rows):3d} candidates  ${total:>14,.2f}")
        grand_count += len(approve_rows)
        grand_total += total
    print(f"  {'TOTAL':13s}  {grand_count:3d} candidates  ${grand_total:>14,.2f}")
    print()
    print(f"  Reclass_Dept_Drift:    {len(bundle.dept_drift)} candidates")
    print(f"  Reclass_Software:      {len(bundle.software_reclass)} candidates")
    print(f"  Unmatched_Actuals:     {len(bundle.unmatched_actuals)} vendors")
    print(f"  Billcom_AlreadyInNS:   {len(bundle.billcom_already_in_ns)} bills (should be ~0)")
    print(f"  Billcom_Unmatched:     {len(bundle.billcom_unmatched)} vendors")


if __name__ == "__main__":
    sys.exit(main())
