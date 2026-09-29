#!/usr/bin/env python3
"""
action5_to_raw_gsheet.py — Excel 4 + Raw Excel 1 → Final Raw Gsheet
=====================================================================

Standalone script. Dependencies: openpyxl only.

Usage:
    python action5_to_raw_gsheet.py <excel_4.xlsx> <raw_excel_1.xlsx> [-o <output.xlsx>]

What it does:
    Takes Excel 4 (fully updated price + modifier tabs from the workflow) and
    the original Raw Excel 1 (system spreadsheet), then:
    1. Copies the ENTIRE Raw Excel 1 as a base (all tabs preserved untouched)
    2. For each dish in Excel 4's price tab: finds it in Raw Excel 1's dish_item
       tab by name, locates its unit_price columns, and writes the new price
       ONLY if it differs from the original
    3. For each modifier item in Excel 4's modifier tab: finds it in Raw Excel 1's
       modifier_item tab by name, writes the new price ONLY if it differs
    4. Prices are written in $X.XX format (matching Raw Excel 1's original format)
    5. Output: RXXXXX_final_raw_gsheet.xlsx — ready to paste into production

Guardrails:
    G1. NEVER add or remove rows — only change existing price cells
    G2. NEVER write to blank/empty cells in Raw Excel 1
    G3. Only touches dish_item and modifier_item tabs — all other tabs untouched
    G4. Only overwrites cells where the price actually changed
"""

import argparse
import re
import sys
from typing import Optional, Dict, Tuple, List
from collections import OrderedDict

try:
    from openpyxl import load_workbook
except ImportError:
    print("ERROR: openpyxl required. Install: pip install openpyxl")
    sys.exit(1)


# DISABLED (user request): dish-code reconciliation produced too many mistakes,
# so it is turned off end-to-end. Part 1 no longer emits a 'Dish Code' tab, and
# here in Part 2 action5 will NOT read that tab or write any unit_code cells.
# When False: e4_codes stays empty, no code cells are written to dish_item, and
# the dish-code drift check is skipped. Price patching is unaffected.
APPLY_DISH_CODES = False


def _parse(val) -> Optional[float]:
    if val is None:
        return None
    s = str(val).strip()
    if not s or s in ('None', 'NULL', ''):
        return None
    s = s.replace('$', '').replace(',', '')
    try:
        return float(s)
    except ValueError:
        return None

def _fmt(val) -> str:
    """Format price as $X.XX (matching Raw Excel 1 format)."""
    if val is None: return ''
    try: return f'${float(val):.2f}'
    except: return ''

def _is_root_row(row) -> bool:
    """Check if this is a root entity row (JSON ID in column 0)."""
    val = str(row[0]).strip() if row[0] else ''
    return val.startswith('{')


NAME_META_COLS = {
    'Dish Group', 'Dish Name:Primary', 'Dish Name:Secondary',
    'Modifier Group', 'Modifier Name:Primary', 'Modifier Name:Secondary',
    'Modifier Item', 'Notes', ''
}


def _is_variant_header(h: str) -> bool:
    """Header-driven: any non-name, non-meta, non-tracking column is a variant."""
    h = str(h).strip()
    return bool(h) and h not in NAME_META_COLS and not h.startswith('_ref')


def _norm_code(val) -> str:
    """Normalize a dish code for comparison/writing (handles '38', 38, '38.0')."""
    if val is None:
        return ''
    s = str(val).strip()
    if s.endswith('.0') and s[:-2].isdigit():
        s = s[:-2]
    return s


def _find_unit_price_cols(headers: list, row2: list) -> Dict[str, int]:
    """Find all unit_price variant columns in a Raw Excel 1 tab.
    Row 2 (path) contains 'wprop.unit_price'; row 3 (header) is the variant name.
    Header-driven: any non-name column under a unit_price path counts (Entree,
    Piece, Pint, Tray, ... — not just a hardcoded allow-list)."""
    cols = {}
    for i, h in enumerate(headers):
        path = str(row2[i]).lower() if i < len(row2) else ''
        if 'unit_price' in path and _is_variant_header(h):
            cols[str(h).strip()] = i
    return cols


def _find_unit_code_cols(headers: list, row2: list) -> Dict[str, int]:
    """Find all unit_code variant columns in the Raw Excel 1 dish_item tab.
    Row 2 (path) contains 'wprop.unit_code'; row 3 (header) is the variant name.
    These hold the POS dish codes (one per variant slot)."""
    cols = {}
    for i, h in enumerate(headers):
        path = str(row2[i]).lower() if i < len(row2) else ''
        if 'unit_code' in path and _is_variant_header(h):
            cols[str(h).strip()] = i
    return cols


def main():
    parser = argparse.ArgumentParser(
        description='Excel 4 + Raw Excel 1 → Final Raw Gsheet',
        epilog='Example: python action5_to_raw_gsheet.py R88312_Excel_4.xlsx R88312_raw.xlsx'
    )
    parser.add_argument('excel_4', help='Path to Excel 4 (fully updated price + modifier)')
    parser.add_argument('raw_excel_1', help='Path to original Raw Excel 1')
    parser.add_argument('--output', '-o', help='Output path (default: RXXXXX_final_raw_gsheet.xlsx)')
    args = parser.parse_args()

    print(f"Excel 4:     {args.excel_4}")
    print(f"Raw Excel 1: {args.raw_excel_1}")

    # ── Load Excel 4 (read-only for price data) ──────────────────
    wb_4 = load_workbook(args.excel_4, read_only=True)

    # Tabs are named beginner-friendly ('Dish Item'/'Modifier'/'Dish Code');
    # older reports used raw lowercase ('price'/'modifier'/'dish_code'). Accept
    # either vintage so Part 2 keeps working on both.
    def _pick_tab(wb, *names):
        for n in names:
            if n in wb.sheetnames:
                return n
        return names[-1]

    # Read Excel 4 price tab → {dish_name: {variant: price_str}}
    e4_price_rows = list(wb_4[_pick_tab(wb_4, 'Dish Item', 'price')].iter_rows(values_only=True))
    e4_price_headers = [str(c).strip() if c else '' for c in e4_price_rows[0]]
    e4_price_name_col = e4_price_headers.index('Dish Name:Primary') if 'Dish Name:Primary' in e4_price_headers else 1

    e4_prices = {}  # dish_name → {variant: price_str}
    # Auto-detect variant columns from headers (header-driven)
    e4_price_variants = {h.strip() for h in e4_price_headers if _is_variant_header(h)}
    for r in e4_price_rows[1:]:
        name = str(r[e4_price_name_col]).strip() if e4_price_name_col < len(r) and r[e4_price_name_col] else ''
        if not name:
            continue
        prices = {}
        for v in e4_price_variants:
            if v in e4_price_headers:
                ci = e4_price_headers.index(v)
                if ci < len(r) and r[ci] and str(r[ci]).strip():
                    prices[v] = str(r[ci]).strip()
        e4_prices[name] = prices

    # Read Excel 4 modifier tab → {modifier_item_name: {variant: price_str}}
    e4_mod_rows = list(wb_4[_pick_tab(wb_4, 'Modifier', 'modifier')].iter_rows(values_only=True))
    e4_mod_headers = [str(c).strip() if c else '' for c in e4_mod_rows[0]]

    e4_mod_prices = {}  # modifier_name → {variant: price_str}
    e4_mod_variants = {h.strip() for h in e4_mod_headers if _is_variant_header(h)}
    for r in e4_mod_rows[1:]:
        mn = str(r[2]).strip() if len(r) > 2 and r[2] else ''
        if not mn or mn in e4_mod_prices:
            continue
        prices = {}
        for v in e4_mod_variants:
            if v in e4_mod_headers:
                ci = e4_mod_headers.index(v)
                if ci < len(r) and r[ci] and str(r[ci]).strip():
                    prices[v] = str(r[ci]).strip()
        e4_mod_prices[mn] = prices

    # Read Excel 4 dish_code tab → {dish_name: {variant: code_str}}
    # (Codes the workflow resolved: blue = Claude change, plus any specialist
    #  hand-edits. Action 5 writes those that differ from the raw gsheet.)
    e4_codes = {}  # dish_name → {variant: code_str}
    dc_tab = _pick_tab(wb_4, 'Dish Code', 'dish_code')
    if APPLY_DISH_CODES and dc_tab in wb_4.sheetnames:
        dc_rows = list(wb_4[dc_tab].iter_rows(values_only=True))
        if dc_rows:
            dc_headers = [str(c).strip() if c else '' for c in dc_rows[0]]
            dc_name_col = dc_headers.index('Dish Name:Primary') if 'Dish Name:Primary' in dc_headers else 1
            dc_variants = {h for h in dc_headers if _is_variant_header(h)}
            for r in dc_rows[1:]:
                nm = str(r[dc_name_col]).strip() if dc_name_col < len(r) and r[dc_name_col] else ''
                if not nm:
                    continue
                codes = {}
                for v in dc_variants:
                    ci = dc_headers.index(v)
                    if ci < len(r) and r[ci] not in (None, ''):
                        codes[v] = _norm_code(r[ci])
                if codes:
                    e4_codes[nm] = codes

    wb_4.close()

    # ── Load Raw Excel 1 (writable copy) ──────────────────────────
    wb_raw = load_workbook(args.raw_excel_1)

    # Auto-detect R-number from restaurant_enterprise tab or filename
    r_number = 'RXXXXX'
    if 'restaurant_enterprise' in wb_raw.sheetnames:
        ws_re = wb_raw['restaurant_enterprise']
        re_rows = list(ws_re.iter_rows(values_only=True))
        if len(re_rows) >= 4:
            headers_re = [str(c).strip() if c else '' for c in re_rows[2]]
            if 'e_number' in headers_re:
                ei = headers_re.index('e_number')
                for r in re_rows[3:]:
                    if ei < len(r) and r[ei] and str(r[ei]).strip():
                        e_num = str(r[ei]).strip()  # e.g. 'ER88312'
                        r_number = e_num.replace('E', '').replace('e', '')
                        if not r_number.startswith('R'):
                            r_number = 'R' + r_number
                        break
    # Fallback: extract from filename
    if r_number == 'RXXXXX':
        import re as _re
        m = _re.search(r'[RE]?(\d{5})', args.raw_excel_1.split('/')[-1])
        if m:
            r_number = f'R{m.group(1)}'

    # Auto-name output
    if args.output:
        out = args.output
    else:
        out = f"{r_number}_final_raw_gsheet.xlsx"

    print(f"R-Number:    {r_number}")

    # ── Update dish_item tab ──────────────────────────────────────
    ws_di = wb_raw['dish_item']
    di_rows_all = list(ws_di.iter_rows(values_only=True))
    di_headers = [str(c).strip() if c else '' for c in di_rows_all[2]]
    di_row2 = [str(c).strip() if c else '' for c in di_rows_all[1]]

    di_name_col = None
    for i, h in enumerate(di_headers):
        if h == 'dish_item':
            di_name_col = i
            break

    di_price_cols = _find_unit_price_cols(di_headers, di_row2)
    di_code_cols = _find_unit_code_cols(di_headers, di_row2)

    dish_changes = 0
    dish_cells = 0
    code_cells = 0
    code_items = 0

    # Drift detection: track which report names matched a live dish_item row.
    # Any report name never matched = the dish was renamed/removed in live
    # content since Part 1, so its price/code edit silently goes unapplied.
    di_price_seen = set()
    di_code_seen = set()

    # Walk every data row in dish_item (row 4+, 1-indexed = row 4 onwards)
    for row_idx in range(4, ws_di.max_row + 1):  # openpyxl is 1-indexed
        row = [ws_di.cell(row=row_idx, column=c+1).value for c in range(ws_di.max_column)]

        if not _is_root_row(row):
            continue

        name = str(row[di_name_col]).strip() if di_name_col is not None and di_name_col < len(row) and row[di_name_col] else ''
        if not name:
            continue

        # ── Prices ──
        if name in e4_prices:
            di_price_seen.add(name)
            e4 = e4_prices[name]
            item_changed = False
            for v, ci in di_price_cols.items():
                old_val = ws_di.cell(row=row_idx, column=ci+1).value
                old_p = _parse(old_val)
                new_str = e4.get(v, '')
                new_p = _parse(new_str)

                if old_p is None or new_p is None:
                    continue  # G1/G2: don't write to blank cells
                if abs(old_p - new_p) < 0.001:
                    continue  # no change

                ws_di.cell(row=row_idx, column=ci+1, value=new_str)
                dish_cells += 1
                item_changed = True

            if item_changed:
                dish_changes += 1

        # ── Dish codes (variant.wprop.unit_code.*) ──
        # DISABLED (user request) when APPLY_DISH_CODES is False: e4_codes is
        # empty so this block never runs. Kept intact for the (off) code path.
        # G1: only overwrite a code cell that already has a value; blank code
        # slots were flagged in Part 1 (P-CODE), never auto-filled.
        if APPLY_DISH_CODES and name in e4_codes:
            di_code_seen.add(name)
            ec = e4_codes[name]
            code_changed = False
            for v, ci in di_code_cols.items():
                old_code = _norm_code(ws_di.cell(row=row_idx, column=ci+1).value)
                new_code = ec.get(v, '')
                if not old_code or not new_code:
                    continue  # G1: never fill a blank code slot
                if old_code == new_code:
                    continue  # no change
                ws_di.cell(row=row_idx, column=ci+1, value=new_code)
                code_cells += 1
                code_changed = True
            if code_changed:
                code_items += 1

    # ── Update modifier_item tab ──────────────────────────────────
    ws_mi = wb_raw['modifier_item']
    mi_rows_all = list(ws_mi.iter_rows(values_only=True))
    mi_headers = [str(c).strip() if c else '' for c in mi_rows_all[2]]
    mi_row2 = [str(c).strip() if c else '' for c in mi_rows_all[1]]

    mi_name_col = None
    for i, h in enumerate(mi_headers):
        if h == 'modifier_item':
            mi_name_col = i
            break

    mi_price_cols = _find_unit_price_cols(mi_headers, mi_row2)

    mod_changes = 0
    mod_cells = 0
    mi_seen = set()

    for row_idx in range(4, ws_mi.max_row + 1):
        row = [ws_mi.cell(row=row_idx, column=c+1).value for c in range(ws_mi.max_column)]

        if not _is_root_row(row):
            continue

        name = str(row[mi_name_col]).strip() if mi_name_col is not None and mi_name_col < len(row) and row[mi_name_col] else ''
        if not name or name not in e4_mod_prices:
            continue
        mi_seen.add(name)

        e4 = e4_mod_prices[name]
        item_changed = False

        for v, ci in mi_price_cols.items():
            old_val = ws_mi.cell(row=row_idx, column=ci+1).value
            old_p = _parse(old_val)
            new_str = e4.get(v, '')
            new_p = _parse(new_str)

            if old_p is None or new_p is None:
                continue
            if abs(old_p - new_p) < 0.001:
                continue

            ws_mi.cell(row=row_idx, column=ci+1, value=new_str)
            mod_cells += 1
            item_changed = True

        if item_changed:
            mod_changes += 1

    # ── Save ──────────────────────────────────────────────────────
    wb_raw.save(out)
    wb_raw.close()

    # ── Drift detection ───────────────────────────────────────────
    # Report names present in the report (Excel 4) but NOT matched to any live
    # row are content drift: the dish/modifier was renamed or removed in wsheet
    # between Part 1 and Part 2, so its edit silently never landed.
    drift_price = sorted(set(e4_prices) - di_price_seen)
    drift_code = sorted(set(e4_codes) - di_code_seen) if APPLY_DISH_CODES else []
    drift_mod = sorted(set(e4_mod_prices) - mi_seen)
    drift_total = len(drift_price) + len(drift_code) + len(drift_mod)

    # ── Deterministic machine-readable summary (anti-hallucination, G12) ──
    # The assistant must NEVER hand-count or recall these numbers. This block
    # is the single source of truth: it writes the exact computed counts (the
    # same values printed below) to part2_summary.json next to the output, so
    # scripts/part2_report.py can render the user-facing report verbatim from
    # ground truth. Merged (not overwritten) so a later modifier_rule run can
    # add its own key without clobbering these.
    import json as _json, os as _os
    _summary_path = _os.path.join(_os.path.dirname(_os.path.abspath(out)) or '.',
                                  'part2_summary.json')
    _existing = {}
    if _os.path.isfile(_summary_path):
        try:
            with open(_summary_path, encoding='utf-8') as _f:
                _existing = _json.load(_f)
        except (OSError, ValueError):
            _existing = {}
    _existing['r_number'] = r_number
    _existing['action5'] = {
        'dish_item_items_changed': dish_changes,
        'dish_item_cells': dish_cells,
        'dish_code_enabled': bool(APPLY_DISH_CODES),
        'dish_code_items_changed': code_items,
        'dish_code_cells': code_cells,
        'modifier_item_items_changed': mod_changes,
        'modifier_item_cells': mod_cells,
        'total_cells_written': dish_cells + code_cells + mod_cells,
        'drift': {
            'total': drift_total,
            'dish_price': drift_price,
            'dish_code': drift_code,
            'modifier': drift_mod,
        },
        'output_xlsx': out,
    }
    with open(_summary_path, 'w', encoding='utf-8') as _f:
        _json.dump(_existing, _f, indent=2)

    # ── Print summary ─────────────────────────────────────────────
    print(f"\n{'='*50}")
    print(f"  dish_item tab: {dish_changes} items changed, {dish_cells} cells updated")
    if APPLY_DISH_CODES:
        print(f"  dish_item codes: {code_items} items changed, {code_cells} code cells updated")
    else:
        print(f"  dish_item codes: DISABLED (dish-code reconciliation turned off)")
    print(f"  modifier_item tab: {mod_changes} items changed, {mod_cells} cells updated")
    print(f"  Total cells written: {dish_cells + code_cells + mod_cells}")
    print(f"  All other tabs: preserved untouched")

    if drift_total:
        print(f"\n  !! DRIFT WARNING: {drift_total} report item(s) not found in live content.")
        print(f"     These edits were NOT applied (dish renamed/removed since Part 1):")
        for n in drift_price:
            print(f"       - dish price: {n}")
        for n in drift_code:
            print(f"       - dish code:  {n}")
        for n in drift_mod:
            print(f"       - modifier:   {n}")
        print(f"     Resolve each before committing (rename match, or confirm intentionally dropped).")
    else:
        print(f"  Drift check: all report items matched a live row. OK.")

    print(f"\n  Output: {out}")
    print(f"  Summary JSON: {_summary_path}  (source of truth for part2_report.py)")
    print(f"{'='*50}")
    print("Done.")


if __name__ == '__main__':
    main()