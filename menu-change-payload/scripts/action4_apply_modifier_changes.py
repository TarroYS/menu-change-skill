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

    Edits (2026-09-30): an edit is a cell that differs from its hidden _ref_
    baseline. A modifier item listed under several groups is ONE item: its
    edits are resolved per (item, variant) across all its rows — an edited row
    wins over untouched ones, rows edited to different prices are a CONFLICT
    (nothing applied). x / X = remove that price (every instance); a price
    typed into an EMPTY cell = a new price (every instance, all empty). Not a
    valid price (letters, comma, negative, 1000+, 0 in an empty cell) = ignored.
    New cells / removals / conflicts / ignored input go to a hidden
    'Part 2 Intent' sheet that action5 reads.

Guardrails:
    G1. Never write to blank/empty cells — only overwrite existing values —
        except a new price the specialist typed into an empty cell.
    G2. Blue fill = increase, Pink fill = decrease (yellow = new, grey = x)
    G3. Never add or remove rows — only change pricing
"""

import argparse
import re
import sys
from typing import Optional
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import PatternFill, Font
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl required. Install: pip install openpyxl")
    sys.exit(1)

BLUE_FILL = PatternFill('solid', fgColor='ADD8E6')
PINK_FILL = PatternFill('solid', fgColor='FFB6C1')
NEW_FILL = PatternFill('solid', fgColor='FFFF00')     # new price in an empty cell
X_FILL = PatternFill('solid', fgColor='595959')       # x = remove this price
INTENT_TAB = 'Part 2 Intent'    # hidden sheet action5 reads (new cells, removals, …)
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


# A price entry after every '$' is removed (like the sheet's SUBSTITUTE): an
# optional sign, ASCII digits, optional decimals. Kept in step with the report's
# red-text rule (action3), so the sheet and Part 2 agree on what is a price.
_ENTRY_RE = re.compile(r'^([+-]?)([0-9]+(?:\.[0-9]*)?|\.[0-9]+)$')
PRICE_CAP = 1000.0      # no dish / modifier costs this — a missing decimal point (1395)


def _round2(x) -> float:
    """Round to cents the way a spreadsheet does — half away from zero, on the
    decimal value as typed (2.675 → 2.68, 15.995 → 16.00; Python's round()
    works on the binary value and gives 2.67 / 15.99, so Part 2 would write a
    cent less than the report shows)."""
    f = float(x)
    if f != f or abs(f) >= 1e12:          # NaN / huge: no cents to round (fails the cap anyway)
        return f
    return float(Decimal(x if isinstance(x, str) else repr(f)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))


def _entry(val, allow_negative=True):
    """(kind, number) for a specialist-edited Modifier Summary cell: 'blank',
    'x', 'price' or 'invalid' — never applied, so a stray keystroke never ships.
    Same rules as action5's parser and the report's red-text rules: a number is
    checked as a number; text is trimmed of plain spaces only, '$' removed, then
    must be a plain decimal. A modifier may be negative ("Less Whole Wing"
    −$1.70); 1000 or more (after rounding) is invalid."""
    if val is None:
        return 'blank', None
    if isinstance(val, bool):
        return 'invalid', None
    if isinstance(val, (int, float)):
        f = _round2(val)
    else:
        s = str(val).strip(' ')
        if not s:
            return 'blank', None
        if s.lower() == 'x':
            return 'x', None
        m = _ENTRY_RE.fullmatch(s.replace('$', '').strip(' '))
        if not m:
            return 'invalid', None
        f = _round2(m.group(2)) * (-1 if m.group(1) == '-' else 1)
    if f != f or abs(f) >= PRICE_CAP or (f < 0 and not allow_negative):
        return 'invalid', None
    return 'price', f + 0.0


# A price 5× or a fifth of the SYSTEM price it replaces is a slipped decimal
# point (195 for 1.95, 0.95 for 9.50), whether the price went up or down. No
# real edit in 101 edited reports (9,943 cells) or Part 1 change on the real
# report pairs comes within 3×. Kept in step with the report's red-text rule.
MISPLACED_RATIO = 5


def _misplaced(new, base) -> bool:
    """5× or a fifth of the system price it replaces — a misplaced decimal
    point. The report shows the same cells red."""
    if new is None or base is None or abs(base) <= 0.001 or abs(new) <= 0.001:
        return False
    return (abs(new) >= MISPLACED_RATIO * abs(base) - 1e-9
            or abs(new) * MISPLACED_RATIO <= abs(base) + 1e-9)


def _same(a, b) -> bool:
    """Report cell equals its baseline (numbers compared as numbers)."""
    pa, pb = _parse(a), _parse(b)
    if pa is not None and pb is not None:
        return abs(pa - pb) < 0.001
    return str(a if a is not None else '').strip().lower() == \
        str(b if b is not None else '').strip().lower()


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

    if INTENT_TAB in wb.sheetnames:
        print(f"ERROR: {args.excel_3} is already a Final Review (it has the '{INTENT_TAB}' sheet). "
              f"Run action4 on the specialist-edited report, not on its own output — a second pass "
              f"would read the applied edits as untouched and drop them.")
        sys.exit(1)
    if 'Modifier Summary' not in wb.sheetnames:
        print("ERROR: No 'Modifier Summary' tab found in Excel 3")
        sys.exit(1)
    mod_tab = _pick_tab(wb, 'Modifier', 'modifier')
    if mod_tab is None:
        # Kept as a stop (maintainer, 2026-10-05). Seen on real reports re-saved by
        # another app: the hidden 'Modifier' tab AND every hidden _ref_ column are
        # gone, so Part 2 cannot tell what the specialist changed.
        lost_refs = not any(str(c.value or '').startswith('_ref_') for c in wb['Modifier Summary'][1])
        print(f"ERROR: {args.excel_3} has lost the report's hidden data — the hidden 'Modifier' tab"
              + (" and the hidden baseline (_ref_) columns" if lost_refs else '') + ". It was probably "
              "saved again by another app. Without it Part 2 cannot tell what the specialist changed, "
              "so nothing is applied.\n  Re-make the report with Part 1 (or use the original file "
              "Part 1 produced), copy the specialist's edits into it, and run Part 2 on that.")
        sys.exit(1)

    ws_ms = wb['Modifier Summary']
    ws_mod = wb[mod_tab]

    # ── Read Modifier Summary headers ─────────────────────────────
    ms_rows = list(ws_ms.iter_rows(values_only=True))
    ms_headers = [str(c).strip() if c else '' for c in ms_rows[0]]
    # A formula typed in Modifier Summary counts by its calculated value (like
    # the Dish Item tab below); one with no saved value is reported, not used.
    ms_formula_blank = set()
    try:
        _wbc = load_workbook(args.excel_3, read_only=True, data_only=True)
        _cached = list(_wbc['Modifier Summary'].iter_rows(values_only=True))
        _wbc.close()
        _fixed = []
        for ri, row in enumerate(ms_rows):
            row = list(row)
            for ci, v in enumerate(row):
                if isinstance(v, str) and v.startswith('='):
                    cv = _cached[ri][ci] if ri < len(_cached) and ci < len(_cached[ri]) else None
                    if cv is None:
                        ms_formula_blank.add((ri + 1, ci))
                    else:
                        row[ci] = cv
            _fixed.append(tuple(row))
        ms_rows = _fixed
    except Exception as e:
        print(f"  Warning: could not read Modifier Summary formula values: {e}")

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
    # Sizes = the Modifier Summary columns Part 1 made (in the hidden Modifier
    # tab, with a _ref_ column) — not a column the specialist added (notes …).
    _ms_refs = {h[len('_ref_'):] for h in ms_headers if h.startswith('_ref_')}
    ms_var_cols = {v: ci for v, ci in ms_var_cols.items()
                   if v in mod_var_cols and (not _ms_refs or v in _ms_refs)}
    # group / item columns by header (a column inserted in front must not shift them)
    ms_g = ms_headers.index('Modifier Group') if 'Modifier Group' in ms_headers else 0
    ms_n = ms_headers.index('Modifier Item') if 'Modifier Item' in ms_headers else 1

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
            # 0 is a real price (report prices are plain numbers) — not blank
            if ci < len(row) and row[ci] is not None and str(row[ci]).strip():
                prices[v] = str(row[ci]).strip()
        original_mod_prices[mn] = prices

    # ── Read specialist edits from Modifier Summary ───────────────
    # An EDIT is a cell that differs from its hidden _ref_ baseline (what Part 1
    # wrote; without _ref_ columns: the original modifier price). The same
    # modifier item can be listed under several modifier groups — it is ONE
    # item (every instance is the same live row), so edits are resolved per
    # (item, variant) across all its rows: an edited row wins over untouched
    # ones; rows edited to different values → conflict, nothing applied.
    # IGNORES the "Notes" column — only reads variant price columns.
    ms_ref_cols = {h[len('_ref_'):]: i for i, h in enumerate(ms_headers)
                   if h.startswith('_ref_') and not h.startswith(('_ref_sys_', '_ref_key:'))}
    # Baselines by each row's own key (hidden '_ref_key:group' / ':name'), not by
    # position — a sort in Excel may have moved only the visible columns.
    ms_hidden = None
    if '_ref_key:group' in ms_headers and '_ref_key:name' in ms_headers:
        _gi, _ni = ms_headers.index('_ref_key:group'), ms_headers.index('_ref_key:name')
        ms_hidden = {(str(r[_gi] or '').strip(), str(r[_ni]).strip()): r for r in ms_rows[1:]
                     if _ni < len(r) and r[_ni] not in (None, '')}
    edits = defaultdict(lambda: defaultdict(list))   # mn → v → [(row, group, kind, num)]
    ignored = []        # not a valid price — never applied
    for ri, ms_row in enumerate(ms_rows[1:], 2):
        mg = str(ms_row[ms_g]).strip() if ms_g < len(ms_row) and ms_row[ms_g] else ''
        mn = str(ms_row[ms_n]).strip() if ms_n < len(ms_row) and ms_row[ms_n] else ''
        if not mn:
            continue
        original = original_mod_prices.get(mn, {})
        hid = ms_hidden.get((mg, mn), ms_row) if ms_hidden is not None else ms_row
        for v, ci in ms_var_cols.items():
            raw = ms_row[ci] if ci < len(ms_row) else None
            if (ri, ci) in ms_formula_blank:
                ignored.append({'tab': 'Modifier Summary', 'row': ri, 'name': mn, 'group': mg,
                                'variant': v, 'value': str(raw),
                                'why': 'a formula with no calculated value — open and save the '
                                       'report in Excel / Google Sheets, or type the number'})
                continue
            rc = ms_ref_cols.get(v)
            ref = (hid[rc] if rc is not None and rc < len(hid) else None) if ms_ref_cols \
                else original.get(v)
            if _same(raw, ref):
                continue                        # untouched
            kind, num = _entry(raw)
            if kind == 'price' and '%' in str(ws_ms.cell(row=ri, column=ci + 1).number_format or ''):
                kind = 'invalid'                # typed as a percentage (Excel stores 10% as 0.1)
            if kind == 'blank':
                continue                        # a cleared cell is no instruction (G1)
            if kind == 'invalid' or (kind == 'price' and abs(num) <= 0.001
                                     and _parse(original.get(v)) is None):
                ignored.append({'tab': 'Modifier Summary', 'row': ri, 'name': mn, 'group': mg,
                                'variant': v, 'value': str(raw),
                                'why': ('0 in an empty cell — a new price must be above 0'
                                        if kind == 'price' else 'not a valid price')})
                continue
            if kind == 'price':
                base = _parse(ref) if ms_ref_cols else _parse(original.get(v))
                if base is None:                # a new price: against the item's other sizes
                    # 10× the largest (by size), or a tenth of the smallest when all
                    # are above 0 — the report's red-text rule is the same
                    others = [p for p in (_parse(x) for x in original.values()) if p is not None]
                    top = max((abs(p) for p in others), default=0)
                    bad = top > 0.001 and (abs(num) >= 10 * top - 1e-9
                                           or (min(others) > 0.001 and abs(num) * 10 <= min(others) + 1e-9))
                else:
                    bad = _misplaced(num, base)
                if bad:
                    ignored.append({'tab': 'Modifier Summary', 'row': ri, 'name': mn, 'group': mg,
                                    'variant': v, 'value': str(raw),
                                    'why': ('looks like a misplaced decimal point (10× or a tenth of the '
                                            "item's other prices)" if base is None else
                                            f'looks like a misplaced decimal point ({MISPLACED_RATIO}× or '
                                            f'a fifth of the system price {base:.2f})')
                                           + ' — not applied; retype it, or confirm in Notes → Wonka'})
                    continue
            edits[mn][v].append((ri, mg, kind, num))

    modifier_changes = {}  # modifier_item_name → {variant: ('price'|'new'|'x', value)}
    mod_diffs = []  # for summary output
    conflicts, x_empty, new_cells, removals = [], [], [], []
    for mn, by_v in edits.items():
        original = original_mod_prices.get(mn, {})
        for v, rows in by_v.items():
            distinct = sorted({(k, n) for _r, _g, k, n in rows}, key=str)
            if len(distinct) > 1:
                conflicts.append({'name': mn, 'variant': v,
                                  'rows': [{'row': r, 'group': g, 'value': 'x' if k == 'x' else n}
                                           for r, g, k, n in rows]})
                continue
            kind, new_p = distinct[0]
            old_p = _parse(original.get(v))
            where = ', '.join(f'row {r}' for r, _g, _k, _n in rows)
            if kind == 'x':
                if old_p is None:
                    x_empty.append({'name': mn, 'variant': v, 'where': where})
                    continue
                modifier_changes.setdefault(mn, {})[v] = ('x', None)
                removals.append({'name': mn, 'variant': v, 'old': old_p, 'where': where})
            elif old_p is None:
                modifier_changes.setdefault(mn, {})[v] = ('new', new_p)
                new_cells.append({'name': mn, 'variant': v, 'new': new_p, 'where': where})
            elif abs(new_p - old_p) > 0.001:
                modifier_changes.setdefault(mn, {})[v] = ('price', new_p)
                mod_diffs.append({
                    'group': rows[0][1], 'name': mn, 'variant': v,
                    'old': old_p, 'new': new_p, 'diff': round(new_p - old_p, 2)
                })

    # ── Apply changes to the full modifier tab ────────────────────
    # Each modifier item in the summary maps to EVERY row with that
    # item name in the modifier tab (across all dish groups).

    cells_changed = 0
    for row_idx, row in enumerate(ws_mod.iter_rows(min_row=2), 2):
        mn_cell = row[2]
        mn = str(mn_cell.value).strip() if mn_cell.value else ''
        if not mn or mn not in modifier_changes:
            continue

        for v, (kind, new_p) in modifier_changes[mn].items():
            ci = mod_var_cols.get(v)
            if ci is None:
                continue
            cell = row[ci]
            old_p = _parse(cell.value) if cell.value is not None and str(cell.value).strip() else None
            if kind == 'new':
                if old_p is not None:
                    continue    # this instance already has a price — not an empty cell
                cell.value = _fmt(new_p)
                cell.fill = NEW_FILL
                cells_changed += 1
            elif old_p is None:
                continue        # G1: don't write to blank cells
            elif kind == 'x':
                cell.value = 'x'
                cell.fill = X_FILL
                cells_changed += 1
            elif abs(old_p - new_p) > 0.001:
                cell.value = _fmt(new_p)
                cell.fill = BLUE_FILL if new_p > old_p else PINK_FILL
                cells_changed += 1

    # What action5 needs besides the Modifier tab (new cells / removals it may
    # write) and what the report must show (conflicts, ignored input): a hidden
    # sheet in the Final Review — action5 copies it into part2_summary.json.
    if INTENT_TAB in wb.sheetnames:
        del wb[INTENT_TAB]
    ws_in = wb.create_sheet(INTENT_TAB)
    ws_in.sheet_state = 'hidden'
    ws_in.append(['kind', 'item', 'variant', 'value', 'detail'])
    # Every resolved edit — action5 writes modifier_item ONLY for these (an
    # untouched snapshot value never overwrites the live price).
    for d in mod_diffs:
        ws_in.append(['price', d['name'], d['variant'], d['new'], ''])
    for c in new_cells:
        ws_in.append(['new', c['name'], c['variant'], c['new'], c['where']])
    for c in removals:
        ws_in.append(['x', c['name'], c['variant'], c['old'], c['where']])
    for c in conflicts:
        ws_in.append(['conflict', c['name'], c['variant'], '',
                      '; '.join(f"row {x['row']} ({x['group']}): {x['value']}" for x in c['rows'])])
    for c in ignored:
        ws_in.append(['ignored', c['name'], c['variant'], c['value'],
                      f"Modifier Summary row {c['row']} ({c['group']}): {c['why']}"])
    for c in x_empty:
        ws_in.append(['x_empty', c['name'], c['variant'], '', c['where']])

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

    # The same for the Dish Code tab: a formula (=G5 pointing at a paper code)
    # becomes the value the specialist saw. One without a saved value stays a
    # formula and action5 refuses it (a code never starts with =).
    if 'Dish Code' in wb.sheetnames:
        try:
            ws_dc = wb['Dish Code']
            wb_cached = load_workbook(args.excel_3, read_only=True, data_only=True)
            dc_cached = list(wb_cached['Dish Code'].iter_rows(values_only=True))
            wb_cached.close()
            dc_resolved = 0
            for row in ws_dc.iter_rows(min_row=2):
                for cell in row:
                    if cell.data_type != 'f':
                        continue
                    r_, c_ = cell.row - 1, cell.column - 1
                    cached = dc_cached[r_][c_] if r_ < len(dc_cached) and c_ < len(dc_cached[r_]) else None
                    if cached is not None:
                        if isinstance(cached, float) and cached.is_integer():
                            cached = int(cached)     # shown as 12, not 12.0
                        cell.value = cached
                        if isinstance(cached, str):
                            cell.data_type = 's'     # text, even when it reads '=…'
                        dc_resolved += 1
            if dc_resolved:
                print(f"\n  Formulas resolved in the Dish Code tab: {dc_resolved}")
        except Exception as e:
            print(f"\n  Warning: Could not resolve Dish Code formulas: {e}")

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
    if new_cells:
        print(f"\n  NEW prices typed into empty cells (Part 2 writes them): {len(new_cells)}")
        for c in new_cells:
            print(f"    + {c['name']} [{c['variant']}]: {_fmt(c['new'])} ({c['where']})")
    if removals:
        print(f"\n  Removed with x (Part 2 clears them): {len(removals)}")
        for c in removals:
            print(f"    - {c['name']} [{c['variant']}] was {_fmt(c['old'])} ({c['where']})")
    if x_empty:
        print(f"\n  x on an empty cell — nothing to remove: "
              f"{', '.join(c['name'] + ' [' + c['variant'] + ']' for c in x_empty)}")
    if conflicts:
        print(f"\n  ⚠ CONFLICT — one modifier item edited to different prices in two groups "
              f"(NOT applied, ask the specialist): {len(conflicts)}")
        for c in conflicts:
            print(f"    {c['name']} [{c['variant']}]: "
                  + '; '.join(f"row {x['row']} ({x['group']}) {x['value']}" for x in c['rows']))
    if ignored:
        print(f"\n  ⚠ Not a valid price — ignored (never applied): {len(ignored)}")
        for c in ignored:
            print(f"    Modifier Summary row {c['row']} {c['name']} [{c['variant']}]: "
                  f"{c['value']!r} — {c['why']}")

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