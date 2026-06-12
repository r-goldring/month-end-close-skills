"""
Shared formatting helpers for payroll pre-flight mapping checks.

Each country's check_mappings.py produces a dict of scan results and
passes it here for consistent formatted output.
"""

BAR = "=" * 72


def print_report(
    country: str,
    pay_date_folder: str,
    input_file: str,
    dept_map_file: str,
    gl_map_file: str,
    dept_map_entries: int,
    gl_map_entries: int,
    cost_centers_found: list,
    cost_centers_mapped: list,
    cost_centers_unmapped: list,
    cost_centers_special_case: list,
    codes_by_category: dict,
    codes_mapped: set,
    codes_unmapped_by_category: dict,
    category_info: dict,
) -> int:
    """
    Print a formatted pre-flight report. Returns an exit code:
      0 = all required items mapped (safe to proceed)
      1 = unmapped items found in categories that require mapping
         (new codes/centers that would be silently dropped)

    category_info: {category_name: {"note": str, "required": bool}}
      - required=True  : unmapped codes here BREAK the JE (must fix)
      - required=False : unmapped codes here are intentionally ignored
                         by the main script (safe but flagged for awareness)
    """
    print(BAR)
    print(f"  PAYROLL PREFLIGHT - {country} {pay_date_folder}")
    print(BAR)
    print(f"  Input file:   {input_file}")
    print(f"  Dept map:     {dept_map_file}  ({dept_map_entries} entries)")
    print(f"  GL map:       {gl_map_file}  ({gl_map_entries} entries)")
    print(BAR)

    # ─── Cost centers ─────────────────────────────────────────────
    print(f"\n  COST CENTERS - {len(cost_centers_found)} found")
    print("  " + "-" * 68)
    for cc in sorted(cost_centers_mapped):
        print(f"    OK     {cc}")
    for cc in sorted(cost_centers_special_case):
        print(f"    NOTE   {cc}   (handled by script fallback - no CSV mapping needed)")
    for cc in sorted(cost_centers_unmapped):
        print(f"    MISS   {cc}   <-- NOT MAPPED (would be dropped)")

    cc_problems = len(cost_centers_unmapped)
    print(f"\n  Summary: {len(cost_centers_mapped)} mapped"
          f" | {len(cost_centers_special_case)} special-case"
          f" | {cc_problems} UNMAPPED (blocking)")

    # ─── Payroll codes ────────────────────────────────────────────
    total_codes = sum(len(codes) for codes in codes_by_category.values())
    print(f"\n  PAYROLL CODES - {total_codes} found across {len(codes_by_category)} categories")
    print("  " + "-" * 68)

    total_blocking_codes = 0
    total_noted_codes = 0

    for category, codes in codes_by_category.items():
        info = category_info.get(category, {"note": "", "required": True})
        required = info["required"]
        note = info["note"]
        req_tag = "MAPPING REQUIRED" if required else "informational only"
        print(f"\n    {category} ({len(codes)} codes)  [{req_tag}: {note}]:")
        unmapped_in_cat = codes_unmapped_by_category.get(category, set())
        for code in sorted(codes):
            if code in unmapped_in_cat:
                if required:
                    print(f"      MISS   {code}   <-- NOT MAPPED (would be dropped)")
                    total_blocking_codes += 1
                else:
                    print(f"      NOTE   {code}   (new code in ignored category - review if it should start mattering)")
                    total_noted_codes += 1
            else:
                print(f"      OK     {code}")

    print(f"\n  Summary: {total_codes - total_blocking_codes - total_noted_codes} mapped"
          f" | {total_noted_codes} informational"
          f" | {total_blocking_codes} UNMAPPED (blocking)")

    # ─── Verdict ──────────────────────────────────────────────────
    print()
    print(BAR)
    blocking = cc_problems + total_blocking_codes
    if blocking == 0:
        print("  RESULT: All required cost centers and payroll codes are mapped.")
        print("          Safe to proceed with the main payroll script.")
        if total_noted_codes:
            print(f"          ({total_noted_codes} new code(s) in ignored categories noted above;")
            print("          no action required unless you want the main script to start")
            print("          handling them.)")
        print(BAR)
        return 0
    else:
        print(f"  RESULT: {blocking} UNMAPPED item(s) found that would be SILENTLY DROPPED.")
        print()
        print("  Next steps (decide with your accountant):")
        print("    1. Add missing mappings to the CSV file(s), then re-run preflight")
        print("    2. Proceed anyway - unmapped items will be silently dropped,")
        print("       making the JE under-stated or out of balance")
        print("    3. Abort this pay run")
        print(BAR)
        return 1
