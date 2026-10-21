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
            data = json.load(f)
    except (OSError, ValueError) as e:
        print(f"part2_report: cannot read {path}: {e}", file=sys.stderr)
        sys.exit(2)
    if "action5" not in data:
        # specialist_requests (step 0b) creates the file first — without
        # action5's block there are no counts to report, never render zeros
        print(f"part2_report: {path} has no action5 result. Run action5_to_raw_gsheet.py with its "
              "-o in this directory first — this report is rendered ONLY from its output.",
              file=sys.stderr)
        sys.exit(2)
    return data, path


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
        lines.append(f"    - modifier_rule: {_rule_label(t)}")
    return lines


def _rule_label(t):
    who = (f"dish='{t.get('dish','')}'" if t.get("dish")
           else f"dish_group='{t.get('dish_group','')}' (group-level)")
    return f"({who}, variant='{t.get('variant','')}', modifier='{t.get('modifier','')}')"


def _request_lines(a5, req):
    """x removals done by action5 + what the specialist asked Wonka to do."""
    removed = [f"    - {c.get('dish','')} [{c.get('variant','')}] was {c.get('old','')} → removed"
               for c in (a5.get("removed") or {}).get("cells", []) if c.get('kind') != 'code']
    unavailable = [f"    - {u.get('dish','')} ({u.get('dish_group','')})"
                   + (f" — only \"{u.get('struck_text','')}\" struck" if u.get('partial') else '')
                   for u in (req.get("make_unavailable") or {}).get("dishes", [])]
    notes_blk = req.get("notes_for_wonka") or {}
    notes = [f"    - {n.get('tab') or 'Questionnaire'} row {n.get('row')} "
             f"[{n.get('question_id') or n.get('dish') or '—'}]: {n.get('note','')}"
             for n in notes_blk.get("notes", [])]
    return removed, unavailable, notes, notes_blk.get("skipped_reason")


def _entry_lines(a5, mr=None):
    """New prices in empty cells, ignored input, duplicate-row conflicts,
    dish codes left behind, new-cell drift — from action5 (2026-09-30)."""
    added = [f"    + {c.get('dish','')} [{c.get('variant','')}] = {c.get('new','')}"
             for c in (a5.get("added") or {}).get("cells", []) if c.get('kind') != 'code']
    ignored = [f"    - {c.get('tab','')} {('row ' + str(c['row']) + ' ') if c.get('row') else ''}"
               f"{c.get('name','')} [{c.get('variant','')}]: {c.get('value','')!r} — {c.get('why','')}"
               for c in (a5.get("ignored_input") or {}).get("cells", [])]
    dup = a5.get("duplicates") or {}
    conflicts = []
    for c in dup.get("conflicts", []):
        if c.get("rows"):
            detail = '; '.join(f"row {x.get('row')} ({x.get('group','')}) {x.get('value')}"
                               for x in c["rows"])
        else:
            detail = c.get("detail", "")
        conflicts.append(f"    - {c.get('tab','')}: {c.get('name','')} [{c.get('variant','')}] — {detail}")
    codes = []
    for c in (a5.get("code_followups") or {}).get("items", []):
        if c.get("action") == "move":
            codes.append(f"    - {c.get('dish','')}: code '{c.get('code','')}' still on [{c.get('from','')}]"
                         f" — price moved to [{c.get('to','')}]")
        else:
            codes.append(f"    - {c.get('dish','')}: code '{c.get('code','')}' still on [{c.get('from','')}]"
                         f" — its price was removed")
    drift_new = [f"    - {c.get('dish','')} [{c.get('variant','')}] report {c.get('report','')} — "
                 + (f"typed into an empty cell, but live now has {c['live']} (not overwritten)" if c.get('live')
                    else "the live cell has no price any more")
                 for c in a5.get("new_cell_drift", [])]
    ignored += [f"    - x on {c.get('dish','')} [{c.get('variant','')}] — that cell has no "
                f"{'code' if c.get('what') == 'code' else 'price'} to remove"
                + (" (Dish Code tab)" if c.get('what') == 'code' else '')
                for c in (a5.get("removed") or {}).get("x_on_empty_cell_ignored", [])]
    ignored += [f"    - {t.get('tab') or 'Modifier Rule'} {_rule_label(t)}: {t.get('value','')!r} — {t.get('why','')}"
                for t in (mr or {}).get("ignored_input", [])]
    # the same rule edited differently on the Modifier Rule and Fried Specialty tabs
    conflicts += [f"    - Modifier Rule / Fried Specialty: {_rule_label(t)} — {t.get('modifier_rule_tab','')} on "
                  f"Modifier Rule, {t.get('fried_specialty_tab','')} on Fried Specialty — not applied"
                  for t in (mr or {}).get("tab_conflicts", [])]
    return added, ignored, conflicts, codes, drift_new, len(dup.get("resolved_from_edited_copy", []))


def _held_lines(a5, mr):
    """Edits deliberately NOT applied by a guard (v4.4) — never silent."""
    lines = []
    for c in (a5.get("absolute_guard") or {}).get("cells", []):
        lines.append(f"    - ABSOLUTE GUARD (price lives on modifiers): {c.get('dish','')} "
                     f"[{c.get('variant','')}] report {c.get('report_price','')} — kept as it was")
    for t in mr.get("held", []):
        lines.append(f"    - modifier_rule HELD (G11): {_rule_label(t)} {t.get('old','')} → {t.get('new','')} "
                     f"— {t.get('why','')}")
    for t in mr.get("no_live_price", []):
        lines.append(f"    - modifier_rule has no live price (disabled / gate rule — never written): "
                     f"{_rule_label(t)}")
    na = mr.get("not_applied") or {}
    for t in na.get("ambiguous", []):
        lines.append(f"    - modifier_rule AMBIGUOUS (matches several live rows): {_rule_label(t)}")
    for t in na.get("group_level_unkeyed", []):
        lines.append(f"    - modifier_rule group-level row, no dish_group column: {_rule_label(t)}")
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
    req = summary.get("specialist_requests", {})

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

    removed, unavailable, notes, notes_skipped = _request_lines(a5, req)
    added, ignored, conflicts, codes, drift_new, dup_used = _entry_lines(a5, mr)
    # dish codes written from the Dish Code tab (fills / corrections / removals)
    cd = a5.get("codes") or {}
    c_fill = (cd.get("filled") or {}).get("items", [])
    c_corr = (cd.get("corrected") or {}).get("items", [])
    c_rem = (cd.get("removed") or {}).get("items", [])
    c_drift = (cd.get("drift") or {}).get("items", [])
    code_lines = ([f"    + {c.get('dish','')} [{c.get('variant','')}] code {c.get('new','')} (was none)" for c in c_fill]
                  + [f"    ~ {c.get('dish','')} [{c.get('variant','')}] code {c.get('old','')} → {c.get('new','')}"
                     for c in c_corr]
                  + [f"    - {c.get('dish','')} [{c.get('variant','')}] code {c.get('old','')} → removed" for c in c_rem])
    drift_new += [f"    - {c.get('dish','')} [{c.get('variant','')}] code: report {c.get('report','')!r} — the "
                  f"code changed in wsheet since Part 1 ({c.get('was','')!r} → {c.get('live','')!r}), not written"
                  for c in c_drift]
    total_cells = dish_cells + mod_cells + mr_cells + len(removed) + len(added) + len(code_lines)
    drift_total = (a5.get("drift", {}).get("total", 0)
                   + mr.get("drift", {}).get("total", 0) + len(drift_new))
    drift_lines = _drift_lines(a5, mr)
    held_lines = _held_lines(a5, mr)

    rest = f"{args.restaurant_name} ({r_number})" if args.restaurant_name else r_number
    draft_line = (f"Staged draft {args.draft_id}" if args.draft_id
                  else "Staged draft <FILL FROM `wsheet draft create` output — do NOT invent>")
    codes_str = (f"{len(c_fill)} filled, {len(c_corr)} corrected, {len(c_rem)} removed (Dish Code tab)"
                 if code_enabled else "none (no Dish Code tab in the report)")

    # ── Staged-draft report (Step 6) ─────────────────────────────────
    print("=" * 60)
    print(f"PART 2 STAGED-DRAFT REPORT — {rest}")
    print(f"(rendered from {path} — do not edit the numbers by hand)")
    print("=" * 60)
    print(f"{draft_line} — NOT committed.")
    if a5.get("rerun_in_same_dir"):
        print("  !! Part 2 ran twice in this directory — if the TSVs were already patched, these counts "
              "are relative to that patched base. Re-fetch into a fresh directory and re-run.")
    for t in a5.get("twin_rows_skipped", []):
        print(f"  NOTE: {t.get('tab','')} row {t.get('row')} '{t.get('name','')}' — a second live row with "
              f"that name, not written")
    print(f"  dish_item prices:        {dish_items} items, {dish_cells} cells")
    print(f"  dish_item codes:         {codes_str}")
    print(f"  modifier_item prices:    {mod_items} items, {mod_cells} cells")
    print(f"  modifier_rule out_price: {mr_cells} cells ({mr_triples} triples matched, model {mr_model})")
    if added:
        print(f"  new prices in empty cells (typed in the report): {len(added)}")
        for ln in added:
            print(ln)
    if removed:
        print(f"  prices removed (x in the report): {len(removed)}")
        for ln in removed:
            print(ln)
    if code_lines:
        print(f"  dish codes written (confirmed on the Dish Code tab): {len(code_lines)}")
        for ln in code_lines:
            print(ln)
    print(f"  TOTAL cells written:     {total_cells}")
    if dup_used:
        print(f"  Dish in several groups — edited copy used over untouched copies: {dup_used} cell(s)")
    kept = ((a5.get("not_reverted") or {}).get("total", 0)
            + (mr.get("not_reverted") or {}).get("total", 0))
    if kept:
        print(f"  Live prices changed since Part 1, untouched in the report — kept, NOT reverted: {kept}"
              + ("  (many: the report may be older than the live menu)" if kept >= 20 else ""))
    print(f"  Drift (NOT applied):     {drift_total}")
    if drift_lines or drift_new:
        print("  Drift detail (resolve before commit):")
        for ln in drift_lines + drift_new:
            print(ln)
    if conflicts:
        print(f"  CONFLICT — edited to different values in two rows (NOT applied, ask): {len(conflicts)}")
        for ln in conflicts:
            print(ln)
    if ignored:
        print(f"  Not a valid price / code — ignored (NOT applied): {len(ignored)}")
        for ln in ignored:
            print(ln)
    if codes:
        print(f"  Dish codes left on a moved / removed price — ask the specialist (never moved or removed "
              f"automatically): {len(codes)}")
        for ln in codes:
            print(ln)
    for c in a5.get("no_price_left", []):
        print(f"  CONFIRM: {c.get('dish','')} has NO price left (x on {', '.join(c.get('removed', []))}) — "
              f"take it off the menu with a strikethrough instead?")
    if held_lines:
        print(f"  Held by a guard (NOT applied): {len(held_lines)}")
        for ln in held_lines:
            print(ln)
    if req:
        print(f"  Handed to Wonka — make permanently unavailable: {len(unavailable)}")
        for ln in unavailable:
            print(ln)
        if notes_skipped:
            print(f"  Handed to Wonka — notes: not read ({notes_skipped})")
        else:
            print(f"  Handed to Wonka — notes (Questionnaire Notes → Wonka): {len(notes)}")
            for ln in notes:
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
    if added:
        print(f"New prices in empty cells: {len(added)}")
    if removed:
        print(f"Prices removed on request (x): {len(removed)}")
    if conflicts or ignored:
        print(f"Not applied — conflicts: {len(conflicts)}, not a valid price / code: {len(ignored)}")
    if code_lines:
        print(f"Dish codes written: {len(code_lines)}")
    if codes:
        print(f"Dish codes to settle: {len(codes)}")
    if req:
        print(f"Handed to Wonka: {len(unavailable)} to make permanently unavailable, "
              f"{len(notes)} note(s)")
    if held_lines:
        print(f"Held by a guard (NOT applied): {len(held_lines)}")
    print(f"All other tabs preserved untouched. Draft staged for review; NOT committed.")
    print("---------------------------------------------------------------")


if __name__ == "__main__":
    main()