#!/usr/bin/env python3
"""
action4_apply_modifier_changes.py — Excel 3 (with specialist edits) → Excel 4
===============================================================================

Standalone script. Dependencies: openpyxl only.

Usage:
    python action4_apply_modifier_changes.py <excel_3.xlsx> [-o <output.xlsx>]

What it does:
    Takes Excel 3 (returned from specialist after client callback) and:
    1. Reads the Modifier Summary tab where the specialist edited variant price
       cells directly (NOT the Notes column — that's for reference only)
    2. Compares each Modifier Summary price against the original modifier tab
       to detect what the specialist changed
    3. Expands those changes into the full modifier tab — one summary item maps
       to every dish-group-replicated instance in the modifier tab
    4. Highlights modifier changes: Blue = increase, Pink = decrease
    5. Carries ALL tabs forward (Dish Item, Modifier, Dish Code, Summary,
       Questionnaire, Dish Code Questionnaire, Modifier Summary) into Excel 4

    The Questionnaire and Dish Code Questionnaire tabs are NOT read — specialists may
    write interview answers/notes there but Claude ignores them (G5). Only the
    Modifier Summary variant columns are read.

Guardrails:
    G1. Never write to blank/empty cells — only overwrite existing values
    G2. Blue fill = increase, Pink fill = decrease
    G3. Never add or remove rows — only change pricing
"""

import argparse
import sys
from typing import Optional
from collections import defaultdict

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import PatternFill, Font
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl required. Install: pip install openpyxl")
    sys.exit(1)

BLUE_FILL = PatternFill('solid', fgColor='ADD8E6')
PINK_FILL = PatternFill('solid', fgColor='FFB6C1')
HEADER_FILL = PatternFill('solid', fgColor='D9E1F2')
HEADER_FONT = Font(bold=True, size=11)

PRICE_VARIANTS = [
    'Order', 'Specialty', 'Small', 'Large', 'Lunch', 'Combo',
    'Party Tray', '2 Ounce', '8 Ounce', '13 Ounce', '15 Ounce',
    '16 Ounce', 'Can', 'Bottle',
    '8 Pieces', '12 Pieces', '20 Pieces', '30 Pieces', '2 Liter'
]


NAME_META_COLS = {
    'Dish Group', 'Dish Name:Primary', 'Dish Name:Secondary',
    'Modifier Group', 'Modifier Name:Primary', 'Modifier Name:Secondary',
    'Modifier Item', 'Notes', ''
}


def _detect_extra_variants(headers: list) -> list:
    """Header-driven: every column that is not a known name/metadata column (and
    not a hidden _ref_ tracking column) is treated as a variant. Agnostic to the
    template's column set (Entree, Piece, Pint, Tray, Dinner, ...)."""
    variants = list(PRICE_VARIANTS)
    for h in headers:
        h = str(h).strip()
        if h and h not in NAME_META_COLS and not h.startswith('_ref') and h not in variants:
            variants.append(h)
    return variants


def _parse(val) -> Optional[float]:
    if val is None or str(val).strip() in ('', 'None', 'NULL'):
        return None
    s = str(val).strip().replace('$', '').replace(',', '')
    try:
        return float(s)
    except ValueError:
        return None

def _fmt(val) -> str:
    if val is None: return ''
    try: return f'${float(val):.2f}'
    except: return ''


def main():
    parser = argparse.ArgumentParser(
        description='Apply Modifier Summary edits from Excel 3 → Excel 4',
        epilog='Example: python action4_apply_modifier_changes.py R88312_Excel_3.xlsx'
    )
    parser.add_argument('excel_3', help='Path to Excel 3 (with specialist edits in Modifier Summary)')
    parser.add_argument('--output', '-o', help='Output path (default: auto-named as Excel_4)')
    args = parser.parse_args()

    if args.output:
        out = args.output
    else:
        base = args.excel_3.replace('_Excel_3', '').replace('.xlsx', '')
        out = f"{base}_Excel_4.xlsx"

    print(f"Excel 3: {args.excel_3}")

    # ── Load Excel 3 ──────────────────────────────────────────────
    wb = load_workbook(args.excel_3)

    # Tabs are named beginner-friendly ('Modifier'); older reports used the raw
    # lowercase name ('modifier'). Accept either so both vintages of report load.
    def _pick_tab(wb, *names):
        for n in names:
            if n in wb.sheetnames:
                return n
        return None

    if 'Modifier Summary' not in wb.sheetnames:
        print("ERROR: No 'Modifier Summary' tab found in Excel 3")
        sys.exit(1)
    mod_tab = _pick_tab(wb, 'Modifier', 'modifier')
    if mod_tab is None:
        print("ERROR: No 'Modifier' tab found in Excel 3")
        sys.exit(1)

    ws_ms = wb['Modifier Summary']
    ws_mod = wb[mod_tab]

    # ── Read Modifier Summary headers ─────────────────────────────
    ms_rows = list(ws_ms.iter_rows(values_only=True))
    ms_headers = [str(c).strip() if c else '' for c in ms_rows[0]]

    # Find variant columns in Modifier Summary (auto-detect extra ounce/piece sizes)
    active_variants = _detect_extra_variants(ms_headers)
    ms_var_cols = {}  # variant_name → column_index_in_ms_rows
    for v in active_variants:
        if v in ms_headers:
            ms_var_cols[v] = ms_headers.index(v)

    # ── Read original modifier tab ────────────────────────────────
    mod_headers = [str(c).strip() if c else '' for c in list(ws_mod.iter_rows(min_row=1, max_row=1, values_only=True))[0]]
    mod_var_cols = {}
    mod_active_variants = _detect_extra_variants(mod_headers)
    for v in mod_active_variants:
        if v in mod_headers:
            mod_var_cols[v] = mod_headers.index(v)

    # Build original modifier price index: modifier_item_name → {variant: price_str}
    # This is the ORIGINAL price before specialist edits
    original_mod_prices = {}
    for row in ws_mod.iter_rows(min_row=2, values_only=True):
        mn = str(row[2]).strip() if row[2] else ''
        if not mn:
            continue
        if mn in original_mod_prices:
            continue  # take first occurrence (de-duplicated)
        prices = {}
        for v, ci in mod_var_cols.items():
            if ci < len(row) and row[ci] and str(row[ci]).strip():
                prices[v] = str(row[ci]).strip()
        original_mod_prices[mn] = prices

    # ── Read specialist edits from Modifier Summary ───────────────
    # Compare each Modifier Summary row against original prices
    # to detect what the specialist changed.
    # IGNORES the "Notes" column — only reads variant price columns.

    modifier_changes = {}  # modifier_item_name → {variant: new_price_str}
    mod_diffs = []  # for summary output

    for ms_row in ms_rows[1:]:
        mg = str(ms_row[0]).strip() if ms_row[0] else ''
        mn = str(ms_row[1]).strip() if ms_row[1] else ''
        if not mn:
            continue

        original = original_mod_prices.get(mn, {})
        changes_for_item = {}

        for v, ci in ms_var_cols.items():
            new_val = str(ms_row[ci]).strip() if ci < len(ms_row) and ms_row[ci] else ''
            old_val = original.get(v, '')

            if not new_val or not old_val:
                continue  # G1: don't write to blank cells

            new_p = _parse(new_val)
            old_p = _parse(old_val)

            if new_p is not None and old_p is not None and abs(new_p - old_p) > 0.001:
                changes_for_item[v] = _fmt(new_p)
                mod_diffs.append({
                    'group': mg, 'name': mn, 'variant': v,
                    'old': old_p, 'new': new_p, 'diff': round(new_p - old_p, 2)
                })

        if changes_for_item:
            modifier_changes[mn] = changes_for_item

    # ── Apply changes to the full modifier tab ────────────────────
    # Each modifier item in the summary maps to EVERY row with that
    # item name in the modifier tab (across all dish groups).

    cells_changed = 0
    for row_idx, row in enumerate(ws_mod.iter_rows(min_row=2), 2):
        mn_cell = row[2]
        mn = str(mn_cell.value).strip() if mn_cell.value else ''
        if not mn or mn not in modifier_changes:
            continue

        changes = modifier_changes[mn]
        for v, new_price_str in changes.items():
            ci = mod_var_cols.get(v)
            if ci is None:
                continue
            cell = row[ci]
            old_val = str(cell.value).strip() if cell.value else ''
            if not old_val:
                continue  # G1: don't write to blank cells

            old_p = _parse(old_val)
            new_p = _parse(new_price_str)
            if old_p is not None and new_p is not None and abs(old_p - new_p) > 0.001:
                cell.value = new_price_str
                cell.fill = BLUE_FILL if new_p > old_p else PINK_FILL
                cells_changed += 1

    # ── Resolve any Excel formulas in price tab variant columns ───
    # Specialists may use formulas (e.g. =concat("$",G23*5) for Party Tray).
    # openpyxl cannot evaluate these, so Action 5 would skip them.
    # We resolve them here by reading cached values from the original file
    # (saved by Excel) and replacing formula cells with static values.
    import re as _re
    try:
        # Price tab is named 'Dish Item' (beginner-friendly); older reports used
        # the raw lowercase 'price'. Accept either.
        price_tab = _pick_tab(wb, 'Dish Item', 'price')
        wb_cached = load_workbook(args.excel_3, read_only=True, data_only=True)
        ws_cached = wb_cached[price_tab]
        cached_rows = list(ws_cached.iter_rows(values_only=True))
        wb_cached.close()

        ws_price = wb[price_tab]
        p_headers = [c.value for c in next(ws_price.iter_rows(min_row=1, max_row=1))]
        # Find all variant columns (header-driven, excludes name/meta + _ref_)
        variant_cols = set()
        for ci, h in enumerate(p_headers):
            if h and str(h).strip() in _detect_extra_variants(
                    [str(x).strip() if x else '' for x in p_headers]):
                variant_cols.add(ci)

        formulas_resolved = 0
        for row_idx in range(2, ws_price.max_row + 1):
            for ci in variant_cols:
                cell = ws_price.cell(row=row_idx, column=ci + 1)
                val = cell.value
                if val and str(val).startswith('='):
                    # Get cached value from data_only read
                    cached_row = cached_rows[row_idx - 1] if row_idx - 1 < len(cached_rows) else None
                    if cached_row and ci < len(cached_row) and cached_row[ci] is not None:
                        cached_val = str(cached_row[ci]).strip()
                        cell.value = cached_val
                        formulas_resolved += 1
        if formulas_resolved > 0:
            print(f"\n  Formulas resolved in price tab: {formulas_resolved}")
    except Exception as e:
        print(f"\n  Warning: Could not resolve formulas: {e}")

    # ── Save Excel 4 ──────────────────────────────────────────────
    wb.save(out)
    wb.close()

    # ── Print summary ─────────────────────────────────────────────
    mod_inc = [d for d in mod_diffs if d['diff'] > 0]
    mod_dec = [d for d in mod_diffs if d['diff'] < 0]

    print(f"\n{'='*50}")
    print(f"  Modifier items changed: {len(modifier_changes)}")
    print(f"  Modifier price edits: {len(mod_diffs)} ({len(mod_inc)} increases, {len(mod_dec)} decreases)")
    print(f"  Cells updated in modifier tab: {cells_changed}")

    if mod_diffs:
        print(f"\n  Modifier changes:")
        for d in sorted(mod_diffs, key=lambda x: (x['group'], x['name'], x['variant'])):
            direction = '↑' if d['diff'] > 0 else '↓'
            print(f"    {direction} [{d['group']}] {d['name']} [{d['variant']}]: {_fmt(d['old'])} → {_fmt(d['new'])}")

    if mod_dec:
        print(f"\n  ⚠ MODIFIER DECREASES DETECTED:")
        for d in mod_dec:
            print(f"    {d['name']} [{d['variant']}]: {_fmt(d['old'])} → {_fmt(d['new'])}")

    print(f"\n  Output: {out}")
    print(f"{'='*50}")
    print("Done.")


if __name__ == '__main__':
    main()