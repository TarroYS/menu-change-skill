#!/usr/bin/env python3
"""
part2_report.py — deterministic Part 2 report renderer (anti-hallucination, G12).

WHY THIS EXISTS.
    The user-facing Part 2 summary (staged-draft report + Slack feedback block)
    used to be a fill-in-the-blank template — `[count] cells`, `N dish / M
    modifier cells`, `Total cells written: [total]` — that the assistant filled
    from memory or mental arithmetic. That is a hallucination source: in a past
    run the same change was variously reported as "194", "623", and "338 items",
    and the draft was announced as "staged" when none existed.

    This script removes ALL model judgment from the numbers. `action5` and
    `modifier_rule_pricing.py apply` each write their EXACT computed counts to
    `part2_summary.json` in the content dir. This script reads that JSON and
    prints the finished report + Slack block verbatim. The assistant copies this
    stdout into chat — it must NEVER author, adjust, or recall a count itself.

USAGE.
    .venv/bin/python scripts/part2_report.py <dir> [--draft-id <uuid>]

    <dir> is the content dir containing part2_summary.json (the same dir the
    xlsx outputs / TSVs live in). --draft-id, if known, is echoed into the
    staged-draft line; otherwise a placeholder reminds you to fill it from the
    real `wsheet draft create` output (never invent one).

EXIT CODES.
    0 -> report printed.
    2 -> part2_summary.json missing/unreadable (run action5 first). The script
         refuses to fabricate a report from nothing.
"""
import argparse
import json
import os
import sys


def _load_summary(content_dir):
    path = os.path.join(content_dir, "part2_summary.json")
    if not os.path.isfile(path):
        print(f"part2_report: no part2_summary.json in {content_dir}. "
              "Run action5_to_raw_gsheet.py (and modifier_rule_pricing.py apply) "
              "first — this report is rendered ONLY from their computed output, "
              "never hand-authored.", file=sys.stderr)
        sys.exit(2)
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f), path
    except (OSError, ValueError) as e:
        print(f"part2_report: cannot read {path}: {e}", file=sys.stderr)
        sys.exit(2)


def _drift_lines(a5, mr):
    lines = []
    d5 = (a5.get("drift") or {})
    for n in d5.get("dish_price", []):
        lines.append(f"    - dish price: {n}")
    for n in d5.get("dish_code", []):
        lines.append(f"    - dish code:  {n}")
    for n in d5.get("modifier", []):
        lines.append(f"    - modifier:   {n}")
    dmr = (mr.get("drift") or {})
    for t in dmr.get("triples", []):
        lines.append(f"    - modifier_rule: (dish='{t.get('dish','')}', "
                     f"variant='{t.get('variant','')}', modifier='{t.get('modifier','')}')")
    return lines


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("content_dir")
    ap.add_argument("--draft-id", default=None,
                    help="Draft uuid from `wsheet draft create` output. If omitted, "
                         "a fill-me placeholder is printed — never invent a uuid.")
    ap.add_argument("--restaurant-name", default=None,
                    help="Human restaurant name for the header (optional).")
    args = ap.parse_args()

    summary, path = _load_summary(args.content_dir)
    r_number = summary.get("r_number", "RXXXXX")
    a5 = summary.get("action5", {})
    mr = summary.get("modifier_rule", {})

    dish_items = a5.get("dish_item_items_changed", 0)
    dish_cells = a5.get("dish_item_cells", 0)
    code_enabled = a5.get("dish_code_enabled", False)
    code_items = a5.get("dish_code_items_changed", 0)
    code_cells = a5.get("dish_code_cells", 0)
    mod_items = a5.get("modifier_item_items_changed", 0)
    mod_cells = a5.get("modifier_item_cells", 0)
    mr_cells = mr.get("out_unit_price_cells_written", 0)
    mr_triples = mr.get("matched_live_rows", 0)
    mr_model = mr.get("pricing_model", "n/a")

    total_cells = dish_cells + code_cells + mod_cells + mr_cells
    drift_total = (a5.get("drift", {}).get("total", 0)
                   + mr.get("drift", {}).get("total", 0))
    drift_lines = _drift_lines(a5, mr)

    rest = f"{args.restaurant_name} ({r_number})" if args.restaurant_name else r_number
    draft_line = (f"Staged draft {args.draft_id}" if args.draft_id
                  else "Staged draft <FILL FROM `wsheet draft create` output — do NOT invent>")
    codes_str = (f"{code_items} items, {code_cells} cells" if code_enabled else "DISABLED")

    # ── Staged-draft report (Step 6) ─────────────────────────────────
    print("=" * 60)
    print(f"PART 2 STAGED-DRAFT REPORT — {rest}")
    print(f"(rendered from {path} — do not edit the numbers by hand)")
    print("=" * 60)
    print(f"{draft_line} — NOT committed.")
    print(f"  dish_item prices:        {dish_items} items, {dish_cells} cells")
    print(f"  dish_item codes:         {codes_str}")
    print(f"  modifier_item prices:    {mod_items} items, {mod_cells} cells")
    print(f"  modifier_rule out_price: {mr_cells} cells ({mr_triples} triples matched, model {mr_model})")
    print(f"  TOTAL cells written:     {total_cells}")
    print(f"  Drift (NOT applied):     {drift_total}")
    if drift_lines:
        print("  Drift detail (resolve before commit):")
        for ln in drift_lines:
            print(ln)
    print("  All other tabs preserved untouched.")
    print("  Reply 'push to prod' to commit via the gated `wsheet draft commit`.")
    print("=" * 60)

    # ── Slack feedback block (Step 4) ────────────────────────────────
    print()
    print("----- copy the block below to #proj-agentic-content-editor -----")
    print(f"Menu Change Part 2 Complete - {rest}")
    print(f"Action 5 - Raw Gsheet:")
    print(f"  dish_item prices: {dish_items} items changed, {dish_cells} cells")
    print(f"  dish_item codes: {codes_str}")
    print(f"  modifier_item: {mod_items} items changed, {mod_cells} cells")
    print(f"  modifier_rule out_unit_price: {mr_cells} cells ({mr_triples} triples matched)")
    print(f"  Total cells written: {total_cells}")
    print(f"Drift warnings (NOT applied): {drift_total}")
    print(f"All other tabs preserved untouched. Draft staged for review; NOT committed.")
    print("---------------------------------------------------------------")


if __name__ == "__main__":
    main()