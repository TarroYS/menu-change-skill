#!/usr/bin/env python3
"""
action3_compare_and_merge.py — Compare Excel 2a vs 2b → Excel 3 (v2)
=====================================================================

Standalone script. Dependencies: openpyxl only.
Works for any restaurant cuisine — variant columns are auto-discovered from
the Menu Preview header row. Tested on Chinese-American and sushi/Japanese.

Usage:
    python action3_compare_and_merge.py <excel_2a.xlsx> <excel_2b.xlsx> [--output <output.xlsx>]

What it does (v2):
    Since Excel 2b now mirrors 2a's structure exactly (same rows, same names,
    same columns), this script is a simple row-by-row, cell-by-cell comparison.
    No name matching, no column mapping — all that complexity lives in Action 2.

    Produces Excel 3 with these sheets:
      - price: 2a structure with 2b prices where available, highlighted changes
      - modifier: copied from 2a unchanged
      - Summary: stats + full change detail table
      - Questionnaire: (DISHES) — P1 (+ P1-TRAY when the restaurant has Party
        Tray prices) + P-COLUMN + P-ABS + P2. Columns A–H are the script's
        (ID / Category / Dish Group / Dish / Current / Expected / Menu Price /
        Details); column I "Notes → Wonka" is the specialist's (instructions
        read in Part 2 by specialist_requests.py). No dropdowns. No P3 /
        Additional Questions / P5 / P6 (removed 2026-09-30).
        P-ABS (v4.4) lists paper prices action2 held off a $0.00 dish whose
        price lives on modifiers (OCR Output 'Absolute Held' tab); its Details
        say exactly which tab / row / column to edit instead.
      Every price is written as a plain number (no "$"), so a specialist types
      only the number. An "x" over a Dish Item price = remove it (dark grey).

Guardrails:
    G1. Never write to blank cells
    G3. Blue fill = increase, Pink fill = decrease
    G9. (v4.3) Party Tray is compared like any other variant — printed tray
        prices go blue/pink/green. Unprinted trays keep the system value
        here; price_inference.py (T-TRAY) may derive them afterwards.
"""

import argparse
import re
import sys
from typing import Optional
from collections import defaultdict

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import PatternFill, Font
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.hyperlink import Hyperlink
except ImportError:
    print("ERROR: openpyxl required. Install: pip install openpyxl")
    sys.exit(1)


PRICE_VARIANTS = [
    'Order', 'Specialty', 'Small', 'Large', 'Lunch', 'Combo',
    'Party Tray', '2 Ounce', '8 Ounce', '13 Ounce', '15 Ounce',
    '16 Ounce', 'Can', 'Bottle'
]
# Kept for backward reference. NOT used to gate which columns are
# processed — see _detect_extra_variants() below.

NAME_COLUMNS = {'Dish Group', 'Dish Name:Primary', 'Dish Name:Secondary'}
MOD_NAME_COLUMNS = NAME_COLUMNS | {'Modifier Group', 'Modifier Item', 'Modifier Name:Primary',
                                   'Modifier Name:Secondary', 'Notes'}

BLUE_FILL = PatternFill('solid', fgColor='ADD8E6')
PINK_FILL = PatternFill('solid', fgColor='FFB6C1')
GREEN_FILL = PatternFill('solid', fgColor='C6EFCE')
HEADER_FILL = PatternFill('solid', fgColor='D9E1F2')
# Flags a dish whose prices never changed (no blue/pink variant edit) so the
# team can see at a glance which rows still need action. Applied to col B only.
YELLOW_FILL = PatternFill('solid', fgColor='FCF7DE')
HEADER_FONT = Font(bold=True, size=11)
# Dish Item: a price cell the specialist replaces with x / X = remove that
# price (Part 2 clears the cell). Dark grey + white — no other report colour.
X_FILL = PatternFill('solid', fgColor='595959')
X_FONT = Font(color='FFFFFF', bold=True)
# Dish Item / Modifier Summary: a price typed into an EMPTY cell = a new price
# (adds that size; paste + x elsewhere = move a price). Vivid yellow + bold so
# it cannot go unnoticed — no other report colour.
NEW_FILL = PatternFill('solid', fgColor='FFFF00')
NEW_FONT = Font(color='000000', bold=True)
# Not a valid price (letters, comma, %, negative, 1000+, 0 in an empty cell):
# red text on white — Part 2 ignores the cell (a stray keystroke never ships).
INVALID_FILL = PatternFill('solid', fgColor='FFFFFF')
INVALID_FONT = Font(color='C00000', bold=True, italic=True)
APRICOT_CODE_FILL = PatternFill('solid', fgColor='FBCEB1')   # an edited code (as an edited price)
# Questionnaire: column I is the specialist's instruction column for Wonka.
NOTES_HEADER = 'Notes → Wonka (write an instruction)'
NOTES_FILL = PatternFill('solid', fgColor='FFF2CC')
EDIT_FONT = Font(bold=True, color='C00000')
PABS_RULES = 'P-ABS — price on modifier rules'
PABS_CHOICE = 'P-ABS — price on a required choice'
# Catering-tray columns; keep in sync with action2 / price_inference TRAY_VARIANTS.
TRAY_VARIANTS = {'Party Tray'}
MONEY_RE = re.compile(r'^\$([+-]?)(\d+(?:\.\d+)?)$')


def _numeric_display(wb):
    """Every price in the report is a plain number shown as 0.00 — never a
    '$X.XX' string — so a specialist just types 12.95 (no '$ prefix) and the
    hidden _ref_ copies compare number to number. A signed delta ('$+1.20')
    keeps its sign. Questionnaire price columns that list several prices lose
    the '$' ('Small 8.95 / Large 12.95'). Part 2 reads either form (_parse)
    and writes the raw back in its own $X.XX format."""
    for ws in wb.worksheets:
        if ws.title == 'Dish Code':
            continue                    # codes are text, never prices
        questionnaire = ws.title == 'Questionnaire'
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str):
                    continue
                m = MONEY_RE.match(v.strip())
                if m:
                    cell.value = round(float(m.group(1).replace('+', '') + m.group(2)), 2)
                    cell.number_format = '+0.00;-0.00' if m.group(1) == '+' else '0.00'
                elif questionnaire and cell.column in (5, 6, 7) and '$' in v:
                    cell.value = v.replace('$', '')


def _detect_extra_variants(headers: list) -> list:
    """
    Return all variant column headers found in the row.

    Auto-discovers every non-name column. The Menu Preview's header row
    is the source of truth — if a column is there, it's a variant
    worth processing.

    Older versions used a hardcoded PRICE_VARIANTS whitelist plus a
    narrow regex for "N Ounce/Pieces/Liter" columns, which silently
    dropped restaurant-specific variants like 'Roll' (sushi),
    'Dinner' (bento), 'Seapark' (combo sets), 'Bowl' (poke), etc.
    """
    variants = []
    seen = set()
    for h in headers:
        h = str(h).strip() if h else ''
        if not h:
            continue
        if h in NAME_COLUMNS:
            continue
        if h in seen:
            continue
        seen.add(h)
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


# ── Dish Code tab (maintainer, 2026-10-05) ─────────────────────────────
# Each size's dish code (unit_code) next to the code printed on the paper.
# Nothing is written unless the specialist puts a code in the left cell —
# fills AND corrections are each confirmed by a person (Wonka's dish-code
# brief). Same conventions as prices: x removes, a code typed into an empty
# cell adds one. Part 2 (action5) reads the left block only.
CODE_TAB = 'Dish Code'
PAPER_SUFFIX = ' · paper'
NO_PRICE_FILL = PatternFill('solid', fgColor='EDEDED')


def _code_key(code) -> str:
    """Search key of a code: letters and digits (any script), lower-case
    ('L-12' = 'l12'; '春1' ≠ '夏1')."""
    return ''.join(ch for ch in str(code or '').casefold() if ch.isalnum())


def _code_text(v) -> str:
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def build_dish_code_tab(wb_out, rows_2a, headers, var_cols, sys_rows, paper_rows, force=False):
    """Add the Dish Code tab, row-aligned with Dish Item (rows_2a[ri] → row
    ri + 1). sys_rows / paper_rows: the 2a 'System Codes' / 2b 'Paper Codes'
    rows (same layout as 'price'). Returns (added, stats); no tab when neither
    side has a code — unless `force` (the paper printed two codes for a size:
    the specialist types the right one here)."""
    from openpyxl.formatting.rule import FormulaRule
    name_col = headers.index('Dish Name:Primary') if 'Dish Name:Primary' in headers else 1

    def cell_of(rows, ri, ci):
        return _code_text(rows[ri][ci]) if ri < len(rows) and ci < len(rows[ri]) else ''
    priced = {v for v, ci in var_cols.items()
              if any(ci < len(r) and str(r[ci] or '').strip() for r in rows_2a[1:])}
    coded = {v for v, ci in var_cols.items()
             if any(cell_of(rows, ri, ci) for rows in (sys_rows, paper_rows) for ri in range(1, len(rows)))}
    if not coded and not force:
        return False, {}
    sizes = [v for v in var_cols if v in priced or v in coded]
    # the paper block only when the paper printed codes (else nothing to compare)
    paper_sizes = (sizes if any(cell_of(paper_rows, ri, var_cols[v]) for v in sizes
                                for ri in range(1, len(paper_rows))) else [])
    ws = wb_out.create_sheet(CODE_TAB)
    n_names = 3
    hdr = ['Dish Group', 'Dish Name:Primary', 'Dish Name:Secondary'] + sizes + [v + PAPER_SUFFIX for v in paper_sizes]
    for ci, h in enumerate(hdr, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    for ci in range(n_names + len(sizes) + 1, n_names + len(sizes) + len(paper_sizes) + 1):
        ws.cell(row=1, column=ci).fill = PatternFill('solid', fgColor='D9D9D9')    # paper block: read-only
    stats = {'differs': 0, 'new': 0, 'same': 0, 'system_codes': 0}
    max_row = len(rows_2a)
    no_price = {}                                          # sheet row → '|k|…|' sizes with no price
    for ri in range(1, len(rows_2a)):
        r = list(rows_2a[ri])
        row = ri + 1
        for ci in range(min(n_names, len(r))):
            if r[ci] not in (None, ''):
                ws.cell(row=row, column=ci + 1, value=str(r[ci]).strip())
        if not (len(r) > name_col and r[name_col]):
            continue                                       # group separator row
        for k, v in enumerate(sizes):
            ci = var_cols[v]
            sys_c, pap_c = cell_of(sys_rows, ri, ci), cell_of(paper_rows, ri, ci)
            has_price = ci < len(r) and str(r[ci] or '').strip() != ''
            cc = ws.cell(row=row, column=n_names + 1 + k, value=sys_c or None)
            cc.number_format = '@'                         # a typed 012 stays "012"
            if not has_price:
                cc.fill = NO_PRICE_FILL                    # a code needs a price on its size
                no_price[row] = no_price.get(row, '|') + f'{k}|'
            if sys_c:
                stats['system_codes'] += 1
            if not paper_sizes:
                continue
            pc = ws.cell(row=row, column=n_names + len(sizes) + 1 + k, value=pap_c or None)
            pc.number_format = '@'
            if pap_c:
                if not sys_c:
                    pc.fill, stats['new'] = NEW_FILL, stats['new'] + 1
                elif _code_key(pap_c) != _code_key(sys_c):
                    pc.fill, stats['differs'] = BLUE_FILL, stats['differs'] + 1
                else:
                    pc.fill, stats['same'] = GREEN_FILL, stats['same'] + 1

    # hidden baselines (the system code) + row keys, right after the data
    ref_start = n_names + len(sizes) + len(paper_sizes) + 1
    grey_font, grey_fill = Font(color='808080', size=8), PatternFill('solid', fgColor='D9D9D9')
    for k, v in enumerate(sizes):
        rc = ref_start + k
        ws.cell(row=1, column=rc, value=f'_ref_{v}').font = grey_font
        for row in range(2, max_row + 1):
            val = ws.cell(row=row, column=n_names + 1 + k).value
            if val not in (None, ''):
                c = ws.cell(row=row, column=rc, value=val)
                c.font, c.number_format = grey_font, '@'
        ws.column_dimensions[get_column_letter(rc)].hidden = True
    key_ci = ref_start + len(sizes)
    for kc, label, srcc in ((key_ci, '_ref_key:group', 1), (key_ci + 1, '_ref_key:name', 2)):
        ws.cell(row=1, column=kc, value=label).font = grey_font
        for row in range(2, max_row + 1):
            kv = ws.cell(row=row, column=srcc).value
            if kv not in (None, ''):
                ws.cell(row=row, column=kc, value=str(kv)).font = grey_font
        ws.column_dimensions[get_column_letter(kc)].hidden = True
    # not a row key: the row's sizes (by place in the code block) that had no
    # price at Part 1 — read only by the red "no price" rule below
    np_ci = key_ci + 2
    ws.cell(row=1, column=np_ci, value='_ref_key:noprice').font = grey_font
    for row, ks in no_price.items():
        ws.cell(row=row, column=np_ci, value=ks).font = grey_font
    ws.column_dimensions[get_column_letter(np_ci)].hidden = True

    # conditional formats on the code cells — the same verdicts as action5,
    # which reads a code like TRIM does, no-break spaces as spaces
    for k, v in enumerate(sizes):
        col = get_column_letter(n_names + 1 + k)
        ref = get_column_letter(ref_start + k)
        rng = f'{col}2:{col}{max(max_row, 2)}'
        c, r_, np_ = f'{col}2', f'{ref}2', f'${get_column_letter(np_ci)}2'
        t = f'TRIM(SUBSTITUTE({c}&"",CHAR(160)," "))'
        rt = f'TRIM(SUBSTITUTE({r_}&"",CHAR(160)," "))'
        stripped = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({t},"#",""),"-",""),".","")," ","")'
        rules = [
            (f'AND(LOWER({t})="x",NOT(EXACT({t},{rt})))', X_FILL, X_FONT),    # x = remove the code
            # not a code: a line break / tab, longer than 128, a leading = (a
            # formula), or only # - . and spaces
            (f'OR(ISNUMBER(SEARCH(CHAR(10),{c}&"")),ISNUMBER(SEARCH(CHAR(13),{c}&"")),'
             f'ISNUMBER(SEARCH(CHAR(9),{c}&"")),LEN({c}&"")>128)', INVALID_FILL, INVALID_FONT),
            (f'AND({t}<>"",OR(LEFT({t},1)="=",LEN({stripped})=0))', INVALID_FILL, INVALID_FONT),
            # a code typed on a size with no price (grey) — Part 2 refuses it
            # unless a price for that size is added on Dish Item too
            (f'AND({t}<>"",NOT(EXACT({t},{rt})),ISNUMBER(FIND("|{k}|",{np_}&"")))',
             INVALID_FILL, INVALID_FONT),
            (f'AND({rt}="",{t}<>"")', NEW_FILL, NEW_FONT),                     # a new code
            (f'AND({t}<>"",NOT(EXACT({t},{rt})))', APRICOT_CODE_FILL, None),   # edited
        ]
        for f, fill, font in rules:
            assert len(f) < 255, (len(f), f)
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f], fill=fill, font=font, stopIfTrue=True))

    ws.freeze_panes = 'D2'
    for col, w in (('A', 22), ('B', 34), ('C', 16)):
        ws.column_dimensions[col].width = w
    for k in range(len(sizes) + len(paper_sizes)):
        ws.column_dimensions[get_column_letter(n_names + 1 + k)].width = 12
    ws.sheet_properties.tabColor = '00B050'
    return True, stats


def main():
    parser = argparse.ArgumentParser(
        description='Compare Excel 2a vs 2b → Excel 3 (simple cell-by-cell)',
        epilog='Example: python action3_compare_and_merge.py R88312_Excel_2a.xlsx R88312_Excel_2b.xlsx'
    )
    parser.add_argument('excel_2a', help='Path to Excel 2a (system prices)')
    parser.add_argument('excel_2b', help='Path to Excel 2b (paper menu prices, mirrors 2a structure)')
    parser.add_argument('--output', '-o', help='Output path (default: auto-named)')
    args = parser.parse_args()

    # Auto-name output
    if args.output:
        out = args.output
    else:
        base = args.excel_2a.replace('_Excel_2a', '').replace('.xlsx', '')
        out = f"{base}_Excel_3.xlsx"

    print(f"Excel 2a: {args.excel_2a}")
    print(f"Excel 2b: {args.excel_2b}")

    wb_2a = load_workbook(args.excel_2a, read_only=True)
    wb_2b = load_workbook(args.excel_2b, read_only=True)
    wb_out = Workbook()

    # ── Read headers ──────────────────────────────────────────────
    rows_2a = list(wb_2a['price'].iter_rows(values_only=True))
    rows_2b = list(wb_2b['price'].iter_rows(values_only=True))
    headers = [str(c).strip() if c else '' for c in rows_2a[0]] if rows_2a else []

    name_col = headers.index('Dish Name:Primary') if 'Dish Name:Primary' in headers else 1
    group_col = headers.index('Dish Group') if 'Dish Group' in headers else 0
    var_cols = {v: headers.index(v) for v in _detect_extra_variants(headers) if v in headers}

    # v4.4 — $0.00 dish cells whose price lives on modifiers (Menu Preview
    # 'absolute_pricing', from action1), and the paper prices action2 held off
    # them ('Absolute Held'). A held dish WAS printed: it counts as matched and
    # is asked about in P-ABS, never as unmatched.
    def _sheet_dicts(wb, tab):
        if tab not in wb.sheetnames:
            return []
        rows = list(wb[tab].iter_rows(values_only=True))
        hdr = [str(c).strip() if c else '' for c in rows[0]] if rows else []
        return [{h: (str(v).strip() if v is not None else '') for h, v in zip(hdr, r)}
                for r in rows[1:] if r and any(v not in (None, '') for v in r)]
    absolute_cells = {(d['Dish Name'], d['Variant'])
                      for d in _sheet_dicts(wb_2a, 'absolute_pricing') if d.get('Dish Name')}
    absolute_held = [d for d in _sheet_dicts(wb_2b, 'Absolute Held') if d.get('Dish Name')]
    held_dishes = {d['Dish Name'] for d in absolute_held}
    # two different prices printed for one size (action2 'Price Conflicts'):
    # neither written — the dish WAS printed, asked in P-PRICE (2026-10-05)
    price_conflict_rows = [d for d in _sheet_dicts(wb_2b, 'Price Conflicts') if d.get('Dish Name')]
    held_dishes |= {d['Dish Name'] for d in price_conflict_rows}

    # ── Dish Item tab: row-by-row comparison ──────────────────────
    # (output tab named 'Dish Item' — beginner-friendly. The INPUT 2a/2b tab it
    #  is built from is still the raw lowercase 'price' tab, read above.)
    ws = wb_out.active
    ws.title = 'Dish Item'
    ws.freeze_panes = 'A2'  # freeze header row
    for ci, h in enumerate(headers):
        c = ws.cell(row=1, column=ci+1, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT

    diffs = []
    items_with_2b = 0
    items_without_2b = 0
    confirmed_no_change = 0

    # These groups are excluded from match rate — rarely on paper menus
    EXCLUDED_FROM_MATCH_RATE = {'Drink', 'Sauces', 'Side Order'}
    excluded_items = 0

    sys_prices = {}     # (sheet row, sheet col) → the SYSTEM price at Part 1 (hidden _ref_sys_ mirror)
    for ri in range(1, len(rows_2a)):
        r2a = list(rows_2a[ri]) if ri < len(rows_2a) else []
        r2b = list(rows_2b[ri]) if ri < len(rows_2b) else []

        dish = str(r2a[name_col]).strip() if len(r2a) > name_col and r2a[name_col] else ''
        group = str(r2a[group_col]).strip() if len(r2a) > group_col and r2a[group_col] else ''

        # Copy columns 1-3 from 2a (structure)
        for ci in range(len(r2a)):
            val = r2a[ci]
            if val is not None:
                ws.cell(row=ri+1, column=ci+1, value=str(val).strip() if val else '')

        if not dish:
            continue  # separator row

        # Compare each price variant column
        has_any_2b = False
        row_has_activity = False  # any blue/pink (change) OR green (confirmed) cell
        for v, ci in var_cols.items():
            p2a = _parse(r2a[ci]) if ci < len(r2a) else None
            p2b = _parse(r2b[ci]) if ci < len(r2b) else None
            if p2a is not None:
                sys_prices[(ri + 1, ci + 1)] = round(p2a, 2)

            if p2b is not None:
                has_any_2b = True
                if p2a is not None and abs(p2a - p2b) > 0.001:
                    # Price changed
                    ws.cell(row=ri+1, column=ci+1, value=_fmt(p2b))
                    ws.cell(row=ri+1, column=ci+1).fill = BLUE_FILL if p2b > p2a else PINK_FILL
                    row_has_activity = True
                    diffs.append({
                        'group': group, 'name': dish, 'variant': v,
                        'old': p2a, 'new': p2b, 'diff': round(p2b - p2a, 2)
                    })
                elif p2a is not None:
                    # Same price — confirmed match, no change
                    ws.cell(row=ri+1, column=ci+1, value=_fmt(p2a))
                    ws.cell(row=ri+1, column=ci+1).fill = GREEN_FILL
                    confirmed_no_change += 1
                    row_has_activity = True
                # G1: if p2a is None, don't write
            else:
                # No 2b price — keep 2a
                if p2a is not None:
                    ws.cell(row=ri+1, column=ci+1, value=_fmt(p2a))
        if dish in held_dishes:
            has_any_2b = True   # printed; its price was held (P-ABS: a $0 cell; P-PRICE: printed twice)

        # Flag rows with NO blue/pink/green activity yellow (cols B AND C) so the
        # team sees which dishes still need action. A green (confirmed-match) row
        # counts as handled, so it is NOT flagged. name_col+1 = 1-based col B
        # (Dish Name:Primary); +2 = col C (Dish Name:Secondary).
        if not row_has_activity:
            ws.cell(row=ri+1, column=name_col+1).fill = YELLOW_FILL
            ws.cell(row=ri+1, column=name_col+2).fill = YELLOW_FILL

        # Track match rate — exclude operational groups
        if group in EXCLUDED_FROM_MATCH_RATE:
            excluded_items += 1
        elif has_any_2b:
            items_with_2b += 1
        else:
            items_without_2b += 1

    # ── Modifier tab: copy from 2a ────────────────────────────────
    # (output tab named 'Modifier'; INPUT read from raw lowercase 'modifier'.)
    ws_m = wb_out.create_sheet('Modifier')
    mod_rows = list(wb_2a['modifier'].iter_rows(values_only=True)) if 'modifier' in wb_2a.sheetnames else []
    for ri, row in enumerate(mod_rows):
        for ci, val in enumerate(row):
            cell = ws_m.cell(row=ri+1, column=ci+1, value=str(val).strip() if val else '')
            if ri == 0:
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT

    # ── Summary sheet ─────────────────────────────────────────────
    # ── Dish Code tab: system codes next to the printed ones (2026-10-05) ──
    sys_code_rows = (list(wb_2a['System Codes'].iter_rows(values_only=True))
                     if 'System Codes' in wb_2a.sheetnames else [])
    paper_code_rows = (list(wb_2b['Paper Codes'].iter_rows(values_only=True))
                       if 'Paper Codes' in wb_2b.sheetnames else [])
    code_conflict_rows = ([r for r in wb_2b['Code Conflicts'].iter_rows(min_row=2, values_only=True) if r and r[1]]
                          if 'Code Conflicts' in wb_2b.sheetnames else [])
    code_tab, code_stats = build_dish_code_tab(wb_out, rows_2a, headers, var_cols, sys_code_rows,
                                               paper_code_rows, force=bool(code_conflict_rows))

    ws_s = wb_out.create_sheet('Summary')

    inc = [d for d in diffs if d['diff'] > 0]
    dec = [d for d in diffs if d['diff'] < 0]

    # --- Dashboard section ---
    ws_s.cell(row=1, column=1, value='DASHBOARD').font = Font(bold=True, size=14)
    ws_s.cell(row=1, column=1).fill = PatternFill('solid', fgColor='4472C4')
    ws_s.cell(row=1, column=1).font = Font(bold=True, size=14, color='FFFFFF')
    for ci in range(2, 7):
        ws_s.cell(row=1, column=ci).fill = PatternFill('solid', fgColor='4472C4')

    # Match rate (excluding Drinks, Sauces, Side Orders)
    total_items = items_with_2b + items_without_2b
    match_pct = round(items_with_2b / total_items * 100, 1) if total_items else 0
    ws_s.cell(row=3, column=1, value='Match rate').font = Font(bold=True)
    ws_s.cell(row=3, column=2, value=f'{items_with_2b} / {total_items} ({match_pct}%)')
    ws_s.cell(row=4, column=1, value='Excluded from match rate').font = Font(color='808080')
    ws_s.cell(row=4, column=2, value=f'{excluded_items} items (Drinks, Sauces, Side Orders)')
    ws_s.cell(row=4, column=2).font = Font(color='808080')

    # Coverage by dish group
    from collections import Counter as _Ctr
    group_matched = _Ctr()
    group_total = _Ctr()
    for ri in range(1, len(rows_2a)):
        r2a = list(rows_2a[ri]) if ri < len(rows_2a) else []
        r2b = list(rows_2b[ri]) if ri < len(rows_2b) else []
        group = str(r2a[group_col]).strip() if len(r2a) > group_col and r2a[group_col] else ''
        dish = str(r2a[name_col]).strip() if len(r2a) > name_col and r2a[name_col] else ''
        if not dish or not group:
            continue
        group_total[group] += 1
        has_2b_any = dish in held_dishes
        for v, ci in var_cols.items():
            p2b = _parse(r2b[ci]) if ci < len(r2b) else None
            if p2b is not None:
                has_2b_any = True
                break
        if has_2b_any:
            group_matched[group] += 1

    ws_s.cell(row=5, column=1, value='Coverage by dish group').font = Font(bold=True)
    gr = 6
    for ci, h in enumerate(['Dish Group', 'Matched', 'Total', 'Coverage'], 1):
        ws_s.cell(row=gr, column=ci, value=h).font = HEADER_FONT
        ws_s.cell(row=gr, column=ci).fill = HEADER_FILL
    gr += 1
    for g in sorted(group_total.keys()):
        m = group_matched.get(g, 0)
        t = group_total[g]
        pct = round(m / t * 100) if t else 0
        ws_s.cell(row=gr, column=1, value=g)
        ws_s.cell(row=gr, column=2, value=m)
        ws_s.cell(row=gr, column=3, value=t)
        ws_s.cell(row=gr, column=4, value=f'{pct}%')
        if g in EXCLUDED_FROM_MATCH_RATE:
            # Grey out — excluded from match rate
            ws_s.cell(row=gr, column=4, value=f'{pct}% (excluded)')
            for ci in range(1, 5):
                ws_s.cell(row=gr, column=ci).font = Font(color='808080')
        elif pct == 0:
            for ci in range(1, 5):
                ws_s.cell(row=gr, column=ci).font = Font(color='FF0000')
        gr += 1

    # --- Color Legend ---
    gr += 1
    ws_s.cell(row=gr, column=1, value='COLOR LEGEND').font = Font(bold=True, size=12)
    gr += 1
    legend = [
        (BLUE_FILL, 'Price increase — update in system'),
        (PINK_FILL, 'Price decrease — verify during callback'),
        (GREEN_FILL, 'Confirmed match, no change — no action needed'),
        (None, 'No color = needs attention (unmatched or unverified)'),
        (X_FILL, 'x — type x (or X) over a Dish Item / Modifier Summary price = remove that price '
                 '(Part 2 clears the cell)'),
        (NEW_FILL, 'NEW price typed into an EMPTY Dish Item / Modifier Summary cell = adds that size. '
                   'To MOVE a price: paste it into the empty cell, then type x over the old one'),
    ]
    for fill, desc in legend:
        if fill:
            ws_s.cell(row=gr, column=1, value='    ').fill = fill
        else:
            ws_s.cell(row=gr, column=1, value='    ')
        ws_s.cell(row=gr, column=2, value=desc)
        gr += 1
    ws_s.cell(row=gr, column=1, value='9,5').font = INVALID_FONT
    ws_s.cell(row=gr, column=2, value=('Red text = not a valid price (letters, comma, %, a negative dish '
                                       'price, 1000 or more, 0 in an empty cell, 5× / a fifth of the system '
                                       'price, or a new price at 10× / a tenth of the other sizes — a '
                                       'misplaced decimal point) — Part 2 does not apply it'))
    gr += 1
    ws_s.cell(row=gr, column=2, value=('A dish or modifier listed under two groups is ONE item — edit '
                                       'either copy: the edited copy wins over an untouched one. Two '
                                       'copies edited to different prices are not applied (Part 2 asks)'))
    gr += 1
    ws_s.cell(row=gr, column=1, value='Abc').font = Font(strike=True)
    ws_s.cell(row=gr, column=2, value=('Strikethrough on a Dish Item dish name (column B) = make that '
                                       'dish permanently unavailable (Part 2 hands it to Wonka)'))
    gr += 1
    ws_s.cell(row=gr, column=1, value='Notes').fill = NOTES_FILL
    ws_s.cell(row=gr, column=2, value=('Questionnaire column I (Notes → Wonka) = write an instruction on '
                                       'any row; Part 2 hands every note to Wonka'))
    gr += 1
    if code_tab:
        ws_s.cell(row=gr, column=1, value='Code')
        ws_s.cell(row=gr, column=2, value=(
            "Dish Code tab: each size's code (left block) next to the code printed on the paper (right "
            "block: blue = differs, yellow = the system has none, green = same). Nothing is written unless "
            "you put the code in the left cell — copy a paper code over to accept it (a whole paper column "
            "pastes safely, header and all: its empty cells change nothing). x removes a code; grey = that "
            "size has no price (a code needs one: a code typed there is red, applied only if you add that "
            "price on Dish Item too); red = not a code."))
        gr += 1
    ws_s.cell(row=gr, column=2, value=('Prices are plain numbers (12.95) everywhere — type them without '
                                       '$ or an apostrophe')).font = Font(italic=True, color='808080')
    gr += 1

    # --- Stats section ---
    gr += 1
    ws_s.cell(row=gr, column=1, value='PRICE CHANGE STATS').font = Font(bold=True, size=12)
    gr += 1
    # Count inferred-column routings (P-COLUMN) from the 2b passthrough sheet.
    inferred_col_count = 0
    if 'Inferred Columns' in wb_2b.sheetnames:
        inferred_col_count = sum(
            1 for r in wb_2b['Inferred Columns'].iter_rows(min_row=2, values_only=True)
            if r and r[0]
        )
    # printed prices for a size the dish does not have (action2 could not place them)
    unplaced_rows = ([r for r in wb_2b['Unplaced Prices'].iter_rows(min_row=2, values_only=True)
                      if r and len(r) > 1 and r[1]]
                     if 'Unplaced Prices' in wb_2b.sheetnames else [])

    stats = [
        ('Total price changes', len(diffs)),
        ('Price increases (blue)', len(inc)),
        ('Price decreases (pink)', len(dec)),
        ('Confirmed no change (green)', confirmed_no_change),
        ('Items with menu prices', items_with_2b),
        ('Items without menu prices (kept from system)', items_without_2b),
        ('Inferred column routings (P-COLUMN — verify)', inferred_col_count),
        ('Printed prices for a size the dish does not have (P-SIZE — not compared)', len(unplaced_rows)),
    ]
    if absolute_held:
        stats.append(('Held — $0 dish, price on modifiers (P-ABS — not written)', len(absolute_held)))
    if price_conflict_rows:
        stats.append(('Two different prices printed for one size (P-PRICE — not written)',
                      len(price_conflict_rows)))
    for label, val in stats:
        ws_s.cell(row=gr, column=1, value=label)
        ws_s.cell(row=gr, column=2, value=val)
        gr += 1

    # --- Decreases alert ---
    if dec:
        gr += 1
        ws_s.cell(row=gr, column=1, value='⚠ PRICE DECREASES DETECTED').font = Font(bold=True, color='FF0000', size=12)
        gr += 1
        for d in dec:
            ws_s.cell(row=gr, column=1, value=f"{d['name']} [{d['variant']}]: {_fmt(d['old'])} → {_fmt(d['new'])} ({_fmt(d['diff'])})")
            ws_s.cell(row=gr, column=1).fill = PINK_FILL
            gr += 1

    # --- Metadata ---
    gr += 2
    ws_s.cell(row=gr, column=1, value='REPORT METADATA').font = Font(bold=True, size=12)
    gr += 1
    import datetime
    ws_s.cell(row=gr, column=1, value='Generated')
    ws_s.cell(row=gr, column=2, value=datetime.datetime.now().strftime('%Y-%m-%d %H:%M'))
    gr += 1
    ws_s.cell(row=gr, column=1, value='Menu Preview')
    ws_s.cell(row=gr, column=2, value=str(args.excel_2a))
    gr += 1
    ws_s.cell(row=gr, column=1, value='OCR Output')
    ws_s.cell(row=gr, column=2, value=str(args.excel_2b))
    gr += 1
    ws_s.cell(row=gr, column=1, value='Script version')
    ws_s.cell(row=gr, column=2, value='action3 v2.1 (green highlights, dashboard)')

    # --- Full change detail table ---
    gr += 2
    ws_s.cell(row=gr, column=1, value='ALL PRICE CHANGES').font = Font(bold=True, size=12)
    gr += 1
    sr = gr
    for ci, h in enumerate(['Dish Group', 'Dish Name', 'Variant', 'Old', 'New', 'Change'], 1):
        ws_s.cell(row=sr, column=ci, value=h).font = HEADER_FONT
        ws_s.cell(row=sr, column=ci).fill = HEADER_FILL
    sr += 1
    for d in sorted(diffs, key=lambda x: (x['group'], x['name'], x['variant'])):
        ws_s.cell(row=sr, column=1, value=d['group'])
        ws_s.cell(row=sr, column=2, value=d['name'])
        ws_s.cell(row=sr, column=3, value=d['variant'])
        ws_s.cell(row=sr, column=4, value=_fmt(d['old']))
        ws_s.cell(row=sr, column=5, value=_fmt(d['new']))
        ws_s.cell(row=sr, column=6, value=_fmt(d['diff']))
        ws_s.cell(row=sr, column=6).fill = BLUE_FILL if d['diff'] > 0 else PINK_FILL
        sr += 1

    ws_s.column_dimensions['A'].width = 50
    ws_s.column_dimensions['B'].width = 42
    for col in 'CDEF':
        ws_s.column_dimensions[col].width = 14

    # ── Questionnaire sheet ───────────────────────────────────────
    # Build group pricing intelligence: for each (group, variant),
    # what's the most common NEW price from matched items?
    from collections import Counter as _Counter
    group_new_prices = defaultdict(list)  # (group, variant) → [new_price, ...]
    for ri in range(1, len(rows_2a)):
        r2a = list(rows_2a[ri]) if ri < len(rows_2a) else []
        r2b = list(rows_2b[ri]) if ri < len(rows_2b) else []
        group = str(r2a[group_col]).strip() if len(r2a) > group_col and r2a[group_col] else ''
        dish = str(r2a[name_col]).strip() if len(r2a) > name_col and r2a[name_col] else ''
        if not dish or not group: continue
        for v, ci in var_cols.items():
            p2b = _parse(r2b[ci]) if ci < len(r2b) else None
            if p2b is not None:
                group_new_prices[(group, v)].append(round(p2b, 2))

    # Compute dominant (most common) new price per (group, variant)
    group_expected = {}  # (group, variant) → expected_price
    for key, prices in group_new_prices.items():
        if prices:
            counter = _Counter(prices)
            most_common_price, most_common_count = counter.most_common(1)[0]
            # Only use as "expected" if it covers at least 30% of items in that group+variant
            if most_common_count >= max(2, len(prices) * 0.3):
                group_expected[key] = most_common_price

    ws_q = wb_out.create_sheet('Questionnaire')
    # Columns A–H are the script's (H "Details" carries the context and, for
    # P-ABS, exactly where to edit). Column I is the SPECIALIST's: an
    # instruction written in "Notes → Wonka" is read by Part 2
    # (specialist_requests.py) and handed to Wonka. Part 1 never writes to it.
    q_headers = ['Question ID', 'Category', 'Dish Group', 'Dish Name',
                 'Current Price', 'Expected Price', 'Menu Price', 'Details', NOTES_HEADER]
    DETAILS_COL = 8
    NOTES_COL = 9
    for ci, h in enumerate(q_headers, 1):
        c = ws_q.cell(row=1, column=ci, value=h)
        c.font = HEADER_FONT
        c.fill = NOTES_FILL if ci == NOTES_COL else HEADER_FILL
    ws_q.cell(row=1, column=NOTES_COL).comment = Comment(
        'Write an instruction for Wonka on any row (e.g. "this is the Sesame Chicken '
        'dish, rename it"). Part 2 reads every note in this column and hands it to Wonka.',
        'menu-change')
    ws_q.freeze_panes = 'A2'
    qr = 2

    def _qprice(row, col, value):
        """A Questionnaire price is a plain number (0.00) — never '$X.XX'."""
        if value is None:
            return
        c = ws_q.cell(row=row, column=col, value=round(float(value), 2))
        c.number_format = '0.00'

    def _paper_prices(text):
        """action2's 'Order=$7.95, Large=$9.95' → (single number or None, readable text)."""
        parts = [p.strip() for p in str(text or '').split(',') if p.strip()]
        vals = [(p.split('=', 1)[0].strip(), _parse(p.split('=', 1)[1]))
                for p in parts if '=' in p]
        vals = [(k, v) for k, v in vals if v is not None]
        if len(vals) == 1:
            return vals[0][1], vals[0][0]
        return None, ' / '.join(f'{k} {v:.2f}' for k, v in vals)

    # --- Section: P1 (on menu, not matched to system) --- highest priority
    ws_q.cell(row=qr, column=1, value='P1 — ON MENU, NOT MATCHED TO SYSTEM').font = Font(bold=True, size=11, color='C00000')
    qr += 1

    # Party Tray pricing is always asked when the restaurant prices trays
    # (moved up from the old 'Additional Questions' list).
    tray_cols = [ci for v, ci in var_cols.items() if v in TRAY_VARIANTS]
    tray_dishes = set()
    for ri in range(1, len(rows_2a)):
        r2a = list(rows_2a[ri])
        r2b = list(rows_2b[ri]) if ri < len(rows_2b) else []
        dish = str(r2a[name_col]).strip() if len(r2a) > name_col and r2a[name_col] else ''
        if dish and any((ci < len(r2a) and _parse(r2a[ci]) is not None)
                        or (ci < len(r2b) and _parse(r2b[ci]) is not None) for ci in tray_cols):
            tray_dishes.add(dish)
    if tray_dishes:
        ws_q.cell(row=qr, column=1, value='P1-TRAY')
        ws_q.cell(row=qr, column=2, value='P1 (Party Tray pricing — confirm or update?)')
        ws_q.cell(row=qr, column=3, value='(all)')
        ws_q.cell(row=qr, column=4, value=f'Party Tray — {len(tray_dishes)} dish(es)')
        ws_q.cell(row=qr, column=DETAILS_COL, value=(
            f"{len(tray_dishes)} dish(es) have a Party Tray price (Dish Item, column 'Party "
            f"Tray'). Confirm or update them with the client; edit them in Dish Item. P-TRAY "
            f"at the bottom shows the tray formula and each tray that needs an answer."))
        qr += 1

    if 'Unmatched' in wb_2b.sheetnames:
        ws_u = wb_2b['Unmatched']
        u_rows = list(ws_u.iter_rows(min_row=2, values_only=True))
        for row in u_rows:
            if row[0]:
                dish_name = str(row[3]).strip() if row[3] else ''
                single, text = _paper_prices(row[4])
                ws_q.cell(row=qr, column=1, value=str(row[0]).strip() if row[0] else '')
                ws_q.cell(row=qr, column=2, value=str(row[1]).strip() if row[1] else '')
                ws_q.cell(row=qr, column=3, value=str(row[2]).strip() if row[2] else '')
                ws_q.cell(row=qr, column=4, value=dish_name)
                if single is not None:
                    _qprice(qr, 7, single)
                elif text:
                    ws_q.cell(row=qr, column=7, value=text)
                details = [f'Paper price ({text})'] if single is not None and text else []
                details += [str(row[k]).strip() for k in (5, 6) if len(row) > k and row[k]]
                details.append('Not matched to a system dish — match it, add it or ignore it; '
                               'write what to do in Notes → Wonka.')
                ws_q.cell(row=qr, column=DETAILS_COL, value=' | '.join(details))
                qr += 1

    # --- Section: P-COLUMN (inferred variant-column routing) --- verify
    # Paper price's variant column was blank in the system row, so it was
    # routed to the single populated column instead of being silently
    # dropped (G1). Specialist should confirm the column is correct.
    if 'Inferred Columns' in wb_2b.sheetnames:
        ic_rows = [r for r in wb_2b['Inferred Columns'].iter_rows(min_row=2, values_only=True)
                   if r and r[0]]
        if ic_rows:
            qr += 1
            ws_q.cell(row=qr, column=1, value='P-COLUMN — INFERRED COLUMN ROUTING (VERIFY)').font = Font(bold=True, size=11, color='C00000')
            qr += 1
            for i, row in enumerate(ic_rows, 1):
                dish = str(row[0]).strip() if row[0] else ''
                paper_v = str(row[1]).strip() if len(row) > 1 and row[1] else ''
                routed_v = str(row[2]).strip() if len(row) > 2 and row[2] else ''
                ws_q.cell(row=qr, column=1, value=f'P-COLUMN-{i}')
                ws_q.cell(row=qr, column=2, value='P-COLUMN (inferred column, verify)')
                ws_q.cell(row=qr, column=4, value=dish)
                _qprice(qr, 7, _parse(row[3]) if len(row) > 3 else None)
                ws_q.cell(row=qr, column=DETAILS_COL, value=(
                    f"Paper '{paper_v}' price was put in '{routed_v}' (the only price column this "
                    f"dish has) — check that column is right."))
                qr += 1

    # --- Section: P-CODE (printed dish codes to review on the Dish Code tab) ---
    if code_tab and (code_stats.get('differs') or code_stats.get('new') or code_conflict_rows):
        qr += 1
        ws_q.cell(row=qr, column=1, value='P-CODE — DISH CODES (REVIEW ON THE DISH CODE TAB)').font = Font(
            bold=True, size=11, color='C00000')
        qr += 1
        n = 1
        if code_stats.get('differs') or code_stats.get('new'):
            ws_q.cell(row=qr, column=1, value=f'P-CODE-{n}')
            ws_q.cell(row=qr, column=2, value='P-CODE (printed codes)')
            ws_q.cell(row=qr, column=DETAILS_COL, value=(
                f"{code_stats.get('differs', 0)} printed code(s) differ from the system and "
                f"{code_stats.get('new', 0)} size(s) have a printed code the system lacks — see the Dish Code "
                f"tab. Nothing is written until you copy a paper code into its code cell."))
            qr += 1
            n += 1
        for row in code_conflict_rows:
            ws_q.cell(row=qr, column=1, value=f'P-CODE-{n}')
            ws_q.cell(row=qr, column=2, value='P-CODE (two codes printed)')
            ws_q.cell(row=qr, column=3, value=str(row[0] or ''))
            ws_q.cell(row=qr, column=4, value=f'{row[1]} [{row[2]}]')
            ws_q.cell(row=qr, column=DETAILS_COL, value=(
                f"The paper prints two different codes for this size ({row[3]}) — ask which is right and "
                f"type it into the Dish Code tab."))
            qr += 1
            n += 1

    # --- Section: P-SIZE (a printed price for a size the dish does not have) ---
    # action2 could not place it (the dish has several other sizes, or it is a
    # tray on a dish without one): never written (G1) and never dropped
    # silently — the specialist decides whether the dish should get that size.
    if unplaced_rows:
        qr += 1
        ws_q.cell(row=qr, column=1, value=('P-SIZE — PAPER PRICE FOR A SIZE THIS DISH DOES NOT HAVE '
                                           '(NOT COMPARED)')).font = Font(bold=True, size=11, color='C00000')
        qr += 1
        for i, row in enumerate(unplaced_rows, 1):
            grp, dish = str(row[0] or '').strip(), str(row[1]).strip()
            size = str(row[2] or '').strip() if len(row) > 2 else ''
            ws_q.cell(row=qr, column=1, value=f'P-SIZE-{i}')
            ws_q.cell(row=qr, column=2, value='P-SIZE (paper size not in the system)')
            ws_q.cell(row=qr, column=3, value=grp)
            ws_q.cell(row=qr, column=4, value=f'{dish} [{size}]')
            _qprice(qr, 7, _parse(row[3]) if len(row) > 3 else None)
            sizes = str(row[5] or '').strip() if len(row) > 5 else ''
            note = str(row[4] or '').strip() if len(row) > 4 and str(row[4] or '').startswith('probably') else ''
            how = (f"type the price into this dish's empty '{size}' cell on the Dish Item tab (Part 2 adds "
                   f"a new price there)" if size in var_cols else
                   f"this menu has no '{size}' column — ask Wonka in Notes → Wonka")
            ws_q.cell(row=qr, column=DETAILS_COL, value=(
                f"The paper prints a '{size}' price for this dish, but the system dish has no '{size}' "
                f"price{f' (its sizes: {sizes})' if sizes else ''} — nothing was compared or written."
                + (f" It is {note}." if note else '')
                + f" If it is one of its sizes under another name, edit that size; if the dish should get "
                  f"a new '{size}' size, {how}; otherwise leave it."))
            qr += 1

    # --- Section: P-PRICE (two different prices printed for one size) ---
    # Neither price was written (action2 left the cell blank: Dish Item keeps
    # the system price, price_inference leaves it alone). The client says which
    # is right; the specialist types it into Dish Item (maintainer, 2026-10-05).
    if price_conflict_rows:
        qr += 1
        ws_q.cell(row=qr, column=1, value=('P-PRICE — TWO DIFFERENT PRICES PRINTED FOR ONE SIZE (ASK THE CLIENT '
                                           'WHICH IS RIGHT)')).font = Font(bold=True, size=11, color='C00000')
        qr += 1
        for i, d in enumerate(price_conflict_rows, 1):
            ws_q.cell(row=qr, column=1, value=f'P-PRICE-{i}')
            ws_q.cell(row=qr, column=2, value='P-PRICE (two prices printed)')
            ws_q.cell(row=qr, column=3, value=d.get('Dish Group', ''))
            ws_q.cell(row=qr, column=4, value=f"{d['Dish Name']} [{d['Variant']}]")
            _qprice(qr, 5, _parse(d.get('System Price')))
            ws_q.cell(row=qr, column=7, value=d.get('Printed Prices', '').replace('$', ''))
            ws_q.cell(row=qr, column=DETAILS_COL, value=(
                f"The paper prints two different prices for this size ({d.get('Printed As', '')}) — neither "
                f"was written. Ask the client which is right and type it into this dish's "
                f"'{d['Variant']}' cell on the Dish Item tab."))
            qr += 1

    # --- Section: P-ABS ($0 dish, price lives on modifiers) --- high priority
    # The paper printed a price for a dish the system keeps at $0.00 because
    # its price lives on modifiers (ABSOLUTE). Writing it onto dish_item would
    # charge it on top of the modifier price (G11), so action2 held it back.
    # "Details" names the exact tab / row / column to edit: the By Itself row on
    # Modifier Rule (row number filled in by modifier_rule_pricing.py compare,
    # which builds that tab), or the choice's rows on Modifier Summary.
    pabs_group_rows = []    # (questionnaire row, modifier group, variant, evidence)
    if absolute_held:
        qr += 1
        ws_q.cell(row=qr, column=1, value=('P-ABS — $0 DISH, PRICE LIVES ON MODIFIERS · EDIT WHERE '
                                           '"DETAILS" SAYS (NOT IN DISH ITEM)')).font = Font(bold=True, size=11, color='C00000')
        qr += 1
        for i, h in enumerate(absolute_held, 1):
            label = f"{h['Dish Name']} [{h['Variant']}]"
            evidence = (f"Paper price not written to Dish Item (it stays 0.00 — a price there "
                        f"would be charged on top). Evidence: {h.get('Evidence', '')}")
            if h.get('Source') == 'modifier_group':
                category = PABS_CHOICE
                where = f"EDIT → 'Modifier Summary' tab (the '{h.get('Modifier Group', '')}' rows)"
                pabs_group_rows.append((qr, h.get('Modifier Group', ''), h['Variant'], evidence))
            else:
                category = PABS_RULES
                where = (f"EDIT → 'Modifier Rule' tab · the By Itself row of {label} · column B "
                         f"(out_unit_price) — the price of the plain dish. No Modifier Rule tab in "
                         f"this report? Write the price you want in Notes → Wonka on this row.")
            ws_q.cell(row=qr, column=1, value=f'P-ABS-{i}')
            ws_q.cell(row=qr, column=2, value=category)
            ws_q.cell(row=qr, column=3, value=h.get('Dish Group', ''))
            ws_q.cell(row=qr, column=4, value=label)
            _qprice(qr, 5, 0.0)
            _qprice(qr, 7, _parse(h.get('Menu Price')))
            c = ws_q.cell(row=qr, column=DETAILS_COL, value=f'{where} | {evidence}')
            c.font = EDIT_FONT
            qr += 1

    # --- Section: P2 (price decreases) --- high priority
    qr += 1
    ws_q.cell(row=qr, column=1, value='P2 — PRICE DECREASES (VERIFY)').font = Font(bold=True, size=11, color='C00000')
    qr += 1

    p2_count = 0
    for d in dec:
        p2_count += 1
        ws_q.cell(row=qr, column=1, value=f'P2-{p2_count}')
        ws_q.cell(row=qr, column=2, value='P2 (price decrease — verify)')
        ws_q.cell(row=qr, column=3, value=d['group'])
        ws_q.cell(row=qr, column=4, value=d['name'])
        _qprice(qr, 5, d['old'])
        _qprice(qr, 6, group_expected.get((d['group'], d['variant'])))
        _qprice(qr, 7, d['new'])
        ws_q.cell(row=qr, column=DETAILS_COL, value=(
            f"Decrease in {d['variant']}: {d['old']:.2f} → {d['new']:.2f} — confirm with the "
            f"client, or put the old price back in Dish Item."))
        for ci in range(1, DETAILS_COL + 1):
            ws_q.cell(row=qr, column=ci).fill = PINK_FILL
        qr += 1
    if not dec:
        ws_q.cell(row=qr, column=1, value='(none)')
        qr += 1

    ws_q.column_dimensions['A'].width = 12   # Question ID
    ws_q.column_dimensions['B'].width = 36   # Category
    ws_q.column_dimensions['C'].width = 22   # Dish Group
    ws_q.column_dimensions['D'].width = 38   # Dish Name
    ws_q.column_dimensions['E'].width = 14   # Current Price
    ws_q.column_dimensions['F'].width = 14   # Expected Price
    ws_q.column_dimensions['G'].width = 16   # Menu Price
    ws_q.column_dimensions['H'].width = 80   # Details
    ws_q.column_dimensions['I'].width = 48   # Notes → Wonka (specialist)

    # ── Fried Specialty tab (multi-column side pricing) ───────────
    # If the OCR output has a 'Fried Specialty Sides' sheet (passed through
    # from action2), build a dedicated tab showing the full side pricing
    # and the price difference (combo price - plain price) per item.
    if 'Fried Specialty Sides' in wb_2b.sheetnames:
        ws_fs_src = wb_2b['Fried Specialty Sides']
        fs_rows = list(ws_fs_src.iter_rows(values_only=True))
        # One block per multi-column paper section: a 'Section' | <name> row,
        # then that section's own 'Code', 'Name', 'Plain', <columns> header
        # (v4.4 — action2 passes every section). An older 2b has a single block
        # and no 'Section' row.
        fs_blocks = []          # [(section_name, headers, item_rows)]
        sec_name = ''
        for row in fs_rows:
            row = list(row)
            if not row or all(v in (None, '') for v in row):
                continue
            c0 = str(row[0]).strip() if row[0] is not None else ''
            c1 = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ''
            if c0 == 'Section':
                sec_name = c1
                continue
            if c0 == 'Code' and c1 == 'Name':
                hdr = list(row)
                while hdr and hdr[-1] in (None, ''):
                    hdr.pop()
                fs_blocks.append((sec_name, hdr, []))
                continue
            if fs_blocks:
                fs_blocks[-1][2].append(row)
        fs_blocks = [b for b in fs_blocks if b[2]]
        if fs_blocks:
            ws_fs = wb_out.create_sheet('Fried Specialty')

            # --- Section 1: Full prices ---
            r = 1
            ws_fs.cell(row=r, column=1, value='FULL PRICES (from paper menu)').font = Font(bold=True, size=12)
            r += 1
            for sec_name, fs_headers, items in fs_blocks:
                if sec_name:
                    ws_fs.cell(row=r, column=1, value=f'Section: {sec_name}').font = Font(bold=True)
                    r += 1
                for ci, h in enumerate(fs_headers, 1):
                    ws_fs.cell(row=r, column=ci, value=h).font = HEADER_FONT
                    ws_fs.cell(row=r, column=ci).fill = HEADER_FILL
                r += 1
                for row in items:
                    for ci, val in enumerate(row, 1):
                        if val is not None:
                            if ci >= 3 and isinstance(val, (int, float)):
                                ws_fs.cell(row=r, column=ci, value=f'${val:.2f}')
                            else:
                                ws_fs.cell(row=r, column=ci, value=val)
                    r += 1
                r += 1

            # --- Section 2: Deltas (combo price - plain price) ---
            ws_fs.cell(row=r, column=1, value='PRICE DIFFERENCE (column price minus Plain)').font = Font(bold=True, size=12)
            r += 1
            for sec_name, fs_headers, items in fs_blocks:
                if sec_name:
                    ws_fs.cell(row=r, column=1, value=f'Section: {sec_name}').font = Font(bold=True)
                    r += 1
                delta_headers = ['Code', 'Name', 'Plain'] + [f'Δ {h}' for h in fs_headers[3:]]
                for ci, h in enumerate(delta_headers, 1):
                    ws_fs.cell(row=r, column=ci, value=h).font = HEADER_FONT
                    ws_fs.cell(row=r, column=ci).fill = HEADER_FILL
                r += 1

                for row in items:
                    code = row[0] if row[0] else ''
                    name = row[1] if row[1] else ''
                    plain = _parse(row[2]) if len(row) > 2 else None     # a number, or nothing
                    row = list(row[:3]) + [_parse(x) for x in row[3:]]

                    ws_fs.cell(row=r, column=1, value=code)
                    ws_fs.cell(row=r, column=2, value=name)
                    if plain is not None:
                        ws_fs.cell(row=r, column=3, value=f'${plain:.2f}')

                    for si, side_price in enumerate(row[3:], 0):
                        col_idx = 4 + si
                        if side_price is not None and plain is not None:
                            delta = side_price - plain
                            ws_fs.cell(row=r, column=col_idx, value=f'${delta:+.2f}')
                            if delta > 0:
                                ws_fs.cell(row=r, column=col_idx).fill = BLUE_FILL
                            elif delta < 0:
                                ws_fs.cell(row=r, column=col_idx).fill = PINK_FILL
                    r += 1

            # --- Section 3: Observation ---
            r += 1
            ws_fs.cell(row=r, column=1, value='NOTE:').font = Font(bold=True)
            ws_fs.cell(row=r, column=2, value='Side columns show FLAT combo prices, not fixed upcharges. The delta varies per item because the combo total is the same regardless of the base item price.')
            r += 1
            # Check if all side columns are the same value
            side_vals = set()
            for _n, _h, items in fs_blocks:
                for row in items:
                    for val in row[3:]:
                        if val is not None:
                            side_vals.add(val)
            if len(side_vals) <= 3:
                ws_fs.cell(row=r, column=2, value=f'Flat combo prices observed: {", ".join(f"${v:.2f}" for v in sorted(side_vals))}')

            ws_fs.column_dimensions['A'].width = 8
            ws_fs.column_dimensions['B'].width = 32
            for ci in range(3, max(len(h) for _n, h, _i in fs_blocks) + 1):
                ws_fs.column_dimensions[get_column_letter(ci)].width = 18

    # ── Modifier Summary tab ─────────────────────────────────────
    # One row per unique modifier item (de-duplicated from per-dish-group
    # replication in the modifier tab). Specialists use this during the
    # client callback to confirm/update modifier upcharge prices.
    #
    # Filters out:
    #   - Entire modifier groups where ALL items are $0.00 (e.g., Taste)
    #   - Individual items that are $0.00 across all variants (e.g., free sauces)
    # This typically reduces 500+ rows to ~30-70 actionable rows.

    ws_ms = wb_out.create_sheet('Modifier Summary')
    ws_ms.freeze_panes = 'A2'  # freeze header row

    # Read modifier data from 2a
    mod_headers = [str(c).strip() if c else '' for c in mod_rows[0]] if mod_rows else []
    mod_var_indices = {}
    # the modifier tab's own name columns are not sizes (they used to show up as
    # 'Modifier Group' / 'Modifier Name:Primary' price columns in the summary)
    for v in [h for h in _detect_extra_variants(mod_headers) if h not in MOD_NAME_COLUMNS]:
        if v in mod_headers:
            mod_var_indices[v] = mod_headers.index(v)

    # Collect unique (modifier_group, modifier_item) → prices
    # De-duplicates the per-dish-group replication
    from collections import OrderedDict as _OD
    seen_items = _OD()  # (mg, mn) → {variant: price_str}
    for r in mod_rows[1:]:
        mg = str(r[1]).strip() if r[1] else ''
        mn = str(r[2]).strip() if r[2] else ''
        if not mn or not mg:
            continue
        key = (mg, mn)
        if key in seen_items:
            continue
        prices = {}
        for v, ci in mod_var_indices.items():
            if ci < len(r) and r[ci] and str(r[ci]).strip().startswith('$'):
                prices[v] = str(r[ci]).strip()
        seen_items[key] = prices

    # Identify all-$0 groups (skip entirely)
    group_all_prices = defaultdict(set)
    for (mg, mn), prices in seen_items.items():
        group_all_prices[mg].update(prices.values())
    all_zero_groups = {mg for mg, vals in group_all_prices.items()
                       if all(v == '$0.00' for v in vals)}

    # Write headers — Modifier Group, Modifier Item, then variant columns, then Notes
    var_list = list(mod_var_indices.keys())
    ms_headers = ['Modifier Group', 'Modifier Item'] + var_list + ['Notes']
    for ci, h in enumerate(ms_headers, 1):
        # Req 1: standardise header row 1 to the same fill as the price tab.
        c = ws_ms.cell(row=1, column=ci, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL

    ms_row = 2
    ms_item_count = 0
    ms_group_rows = defaultdict(list)   # modifier group → its Modifier Summary rows

    for (mg, mn), prices in seen_items.items():
        # Skip all-$0 groups
        if mg in all_zero_groups:
            continue
        # Skip individual items where all prices are $0.00
        if all(v == '$0.00' for v in prices.values()):
            continue

        ms_item_count += 1

        # Column A: Modifier Group
        ws_ms.cell(row=ms_row, column=1, value=mg)
        # Column B: Modifier Item
        ws_ms.cell(row=ms_row, column=2, value=mn)
        # Variant price columns — specialist edits these directly
        for vi, v in enumerate(var_list):
            price_val = prices.get(v, '')
            if price_val:
                ws_ms.cell(row=ms_row, column=3 + vi, value=price_val)
        # Notes column left blank for specialist (Claude does NOT read this in Part 2)
        ms_group_rows[mg].append(ms_row)
        ms_row += 1

    # P-ABS rows priced by a required choice: point at that choice's rows here.
    for q_row, mg, variant, evidence in pabs_group_rows:
        rows = ms_group_rows.get(mg, [])
        if rows and variant in var_list:
            col = get_column_letter(3 + var_list.index(variant))
            span = f'row {rows[0]}' if len(rows) == 1 else f'rows {rows[0]}–{rows[-1]}'
            where = (f"EDIT → 'Modifier Summary' tab · {span} (group '{mg}') · column "
                     f"'{variant}' ({col}) — this dish is priced by that required choice; "
                     f"change the option prices there.")
            cell = ws_q.cell(row=q_row, column=DETAILS_COL, value=f'{where} | {evidence}')
            cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'Modifier Summary'!{col}{rows[0]}")
        else:
            cell = ws_q.cell(row=q_row, column=DETAILS_COL, value=(
                f"The required choice '{mg}' is not listed in Modifier Summary — write the change "
                f"you want in Notes → Wonka on this row. | {evidence}"))
        cell.font = EDIT_FONT

    # Column widths
    ws_ms.column_dimensions['A'].width = 22
    ws_ms.column_dimensions['B'].width = 35
    for idx in range(3, 3 + len(var_list)):
        ws_ms.column_dimensions[get_column_letter(idx)].width = 12
    notes_col_idx = 3 + len(var_list)
    ws_ms.column_dimensions[get_column_letter(notes_col_idx)].width = 30

    # Column widths for price/modifier
    for sheet in [ws, ws_m]:
        sheet.column_dimensions['A'].width = 22
        sheet.column_dimensions['B'].width = 38
        sheet.column_dimensions['C'].width = 20
        for idx in range(4, 16):
            sheet.column_dimensions[get_column_letter(idx)].width = 12

    # ── Conditional Formatting for Specialist Edit Tracking ─────────
    # Google Sheets compatible: uses hidden columns on the SAME sheet
    # (not cross-sheet references, which Google Sheets ignores).
    #
    # For each price column, a mirror "reference" column is created far
    # to the right with the original action3 output values. The column is
    # hidden. A CF rule fires when the visible cell ≠ the reference cell:
    #   IF cell ≠ same-sheet reference → APRICOT (specialist edit)
    #
    # Blue/pink/green remain as direct cell fills from the comparison loop.
    # When a specialist types a new value in Google Sheets, the direct fill
    # is cleared, and the apricot CF rule kicks in automatically.
    from openpyxl.formatting.rule import FormulaRule

    APRICOT_FILL = PatternFill(start_color='FBCEB1', end_color='FBCEB1', fill_type='solid')

    def _apply_edit_tracking_gsheets(ws_target, var_col_indices, entry_rules=False, sys_values=None,
                                     allow_negative=False):
        """
        Add hidden reference columns on the same sheet and apply apricot CF.
        Google Sheets compatible (no cross-sheet references).
        
        For each price column at index ci, creates a reference column at
        (ref_start + offset) with the same values. CF compares visible vs reference.
        entry_rules: also flag, ahead of the apricot rule, a value that is not a
        valid price (red text — Part 2 ignores it) and a price typed into a cell
        that was empty (vivid yellow — a new price). Each formula stays under
        Excel's 255-character limit for conditional-format formulas.
        sys_values: {(row, col): system price} → also a hidden `_ref_sys_<header>`
        mirror of the SYSTEM price. Part 2 reads "Part 1 changed this price" as
        _ref_ ≠ _ref_sys_ — never from the cell colour, which a specialist may
        change. (The `_ref_` prefix makes every reader skip the column.)
        """
        max_row = ws_target.max_row
        if max_row < 2 or not var_col_indices:
            return

        # Reference columns start RIGHT AFTER the data, with no gap: Excel's quick
        # sort takes the contiguous block, so the hidden columns move with their
        # rows (a gap left them behind — every cell then compared against another
        # dish's baseline). Part 2 also finds baselines by the row key below.
        max_col = ws_target.max_column or 1
        ref_start = max_col + 1

        GREY_FILL = PatternFill('solid', fgColor='D9D9D9')
        GREY_FONT = Font(color='808080', size=8)

        # Each row's own key (column A group, column B name) in the hidden block,
        # so Part 2 finds a row's baselines by dish, not by position: Excel's
        # quick sort only takes the visible block and would leave the hidden
        # columns behind.
        key_ci = ref_start + len(var_col_indices) * (2 if sys_values is not None else 1)
        for kc, label, src in ((key_ci, '_ref_key:group', 1), (key_ci + 1, '_ref_key:name', 2)):
            hc = ws_target.cell(row=1, column=kc, value=label)
            hc.font = GREY_FONT
            hc.fill = GREY_FILL
            for ri in range(2, max_row + 1):
                kv = ws_target.cell(row=ri, column=src).value
                if kv not in (None, ''):
                    ws_target.cell(row=ri, column=kc, value=str(kv)).font = GREY_FONT
            ws_target.column_dimensions[get_column_letter(kc)].hidden = True

        # The row's _ref_ cells of its non-tray sizes ("$R2:$T2,$V2"), for the
        # new-price rule: a price typed into an empty cell at 10× the item's
        # other prices is a misplaced decimal point (Part 2 holds it; a tray is
        # naturally many times a size, so Dish Item trays are exempt).
        def _ref_block(exempt, first):
            cols = [first + o for o, c in enumerate(var_col_indices)
                    if str(ws_target.cell(row=1, column=c).value or '').strip() not in exempt]
            runs, out = [], []
            for c in cols:
                if runs and c == runs[-1][1] + 1:
                    runs[-1][1] = c
                else:
                    runs.append([c, c])
            for a_, b_ in runs:
                out.append(f'${get_column_letter(a_)}2' + (f':${get_column_letter(b_)}2' if b_ > a_ else ''))
            return ','.join(out)
        # A new price is judged against the item's other prices as the report
        # shows them (its _ref_ block — what the specialist sees); an edit
        # against the SYSTEM price (Dish Item: the hidden _ref_sys_ cell)
        sys_start = ref_start + len(var_col_indices)
        block = _ref_block(set() if allow_negative else TRAY_VARIANTS, ref_start)

        # Header row for reference columns
        for offset, ci in enumerate(var_col_indices):
            ref_ci = ref_start + offset
            src_header = ws_target.cell(row=1, column=ci).value or ''
            cell = ws_target.cell(row=1, column=ref_ci, value=f'_ref_{src_header}')
            cell.font = GREY_FONT
            cell.fill = GREY_FILL

            # Copy values from the price column to the reference column
            for ri in range(2, max_row + 1):
                val = ws_target.cell(row=ri, column=ci).value
                ref_cell_obj = ws_target.cell(row=ri, column=ref_ci, value=val)
                ref_cell_obj.fill = GREY_FILL
                ref_cell_obj.font = GREY_FONT

            # Hide the reference column
            ref_letter = get_column_letter(ref_ci)
            ws_target.column_dimensions[ref_letter].hidden = True

            # Apply CF: if visible cell ≠ reference cell → apricot
            src_letter = get_column_letter(ci)
            cell_range = f'{src_letter}2:{src_letter}{max_row}'
            # Relative refs: a pasted price carries rules that point at its new
            # column's _ref_ (the _ref_ block keeps the price columns' order).
            ref_cell = f'{ref_letter}2'
            this_cell = f'{src_letter}2'

            if entry_rules:
                c, r_ = this_cell, ref_cell
                t = f'TRIM({c}&"")'
                val = f'VALUE(SUBSTITUTE(TRIM({c}),"$",""))'       # a text cell's number
                has = lambda ch: f'ISNUMBER(SEARCH("{ch}",{c}))'
                if allow_negative:                                  # Modifier Summary
                    big_n, big_t = f'ROUND(ABS({c}),2)>=1000', f'ROUND(ABS({val}),2)>=1000'
                else:                                               # Dish Item
                    big_n = f'OR({c}<0,ROUND({c},2)>=1000)'
                    big_t = f'OR({val}<0,ROUND({val},2)>=1000)'
                # 5× / a fifth is judged against the SYSTEM price (Dish Item: the
                # hidden _ref_sys_ cell) — so an untouched misread (19.50 printed
                # for 1.95) shows red and its correction does not. Rounded first,
                # like Part 2 (0.161 is 0.16 — a tenth of 1.60).
                b_ = f'{get_column_letter(sys_start + offset)}2' if sys_values is not None else r_
                ratio = lambda x: (f'{b_}<>0,ROUND({x},2)<>0,'
                                   f'OR(ROUND(ABS({x}),2)>=5*ABS({b_}),ROUND(ABS({x}),2)*5<=ABS({b_}))')
                rules = (
                    # text that is not a number (a number cell is never text-
                    # tested — its decimal comma in some locales is fine)
                    f'AND(ISTEXT({c}),{t}<>"",LOWER({t})<>"x",ISERROR({val}))',
                    # text with characters a price never has
                    f'AND(ISTEXT({c}),OR({has(",")},{has("%")},{has(":")},{has("(")},{has("e")},{has("/")}))',
                    # out of range (after rounding, like Part 2), or 0 typed into an empty cell
                    f'IFERROR(AND(ISNUMBER({c}),OR({big_n},AND(TRIM({r_}&"")="",ROUND({c},2)=0))),FALSE)',
                    f'AND(ISTEXT({c}),{t}<>"",IFERROR(OR({big_t},AND(TRIM({r_}&"")="",ROUND({val},2)=0)),FALSE))',
                    # 5× or a fifth of the system price — a misplaced decimal
                    # point (a number, or a number typed as text)
                    f'IFERROR(AND(ISNUMBER({c}),ISNUMBER({b_}),{ratio(c)}),FALSE)',
                    f'IFERROR(AND(ISTEXT({c}),ISNUMBER({b_}),{ratio(val)}),FALSE)',
                )
                if block and (allow_negative or str(src_header).strip() not in TRAY_VARIANTS):
                    # a NEW price (empty cell) at 10× the item's highest other price,
                    # or a tenth of its lowest (all above 0)
                    m = (f'MAX(MAX({block}),-MIN({block}))' if allow_negative else f'MAX({block})')
                    big = lambda x: f'{m}>0,ROUND(ABS({x}),2)>=10*{m}'
                    small = lambda x: f'MIN({block})>0,ROUND(ABS({x}),2)*10<=MIN({block})'
                    rules += (f'IFERROR(AND(TRIM({r_}&"")="",ISNUMBER({c}),{big(c)}),FALSE)',
                              f'IFERROR(AND(TRIM({r_}&"")="",ISTEXT({c}),{t}<>"",{big(val)}),FALSE)',
                              f'IFERROR(AND(TRIM({r_}&"")="",ISNUMBER({c}),{small(c)}),FALSE)',
                              f'IFERROR(AND(TRIM({r_}&"")="",ISTEXT({c}),{t}<>"",{small(val)}),FALSE)')
                for f in rules:
                    assert len(f) < 255, (len(f), f)
                    ws_target.conditional_formatting.add(cell_range, FormulaRule(
                        formula=[f], fill=INVALID_FILL, font=INVALID_FONT, stopIfTrue=True))
                ws_target.conditional_formatting.add(cell_range, FormulaRule(
                    formula=[f'AND(TRIM({r_}&"")="",{t}<>"",LOWER({t})<>"x")'],
                    fill=NEW_FILL, font=NEW_FONT, stopIfTrue=True))

            # an edit = a value that differs from the baseline, compared as
            # numbers (text "9.45" over 9.45 is no edit; nor is a lone space)
            ws_target.conditional_formatting.add(cell_range, FormulaRule(
                formula=[f'AND(TRIM({this_cell}&"")<>"",IF(ISNUMBER({this_cell}),{this_cell}<>{ref_cell},'
                         f'IFERROR(VALUE(SUBSTITUTE(TRIM({this_cell}),"$",""))<>{ref_cell},TRUE)))'],
                fill=APRICOT_FILL,
                stopIfTrue=True
            ))

            if sys_values is not None:
                sys_ci = ref_start + len(var_col_indices) + offset
                cell = ws_target.cell(row=1, column=sys_ci, value=f'_ref_sys_{src_header}')
                cell.font = GREY_FONT
                cell.fill = GREY_FILL
                for ri in range(2, max_row + 1):
                    sv = sys_values.get((ri, ci))
                    if sv is not None:
                        c2 = ws_target.cell(row=ri, column=sys_ci, value=sv)
                        c2.fill = GREY_FILL
                        c2.font = GREY_FONT
                ws_target.column_dimensions[get_column_letter(sys_ci)].hidden = True

    # Dish Item: x / X typed over a price = remove that price (Part 2 clears the
    # cell). Added BEFORE the apricot edit rule so it wins (lower priority number).
    price_var_indices = [ci + 1 for ci in var_cols.values()]  # 1-indexed
    for ci in price_var_indices:
        letter = get_column_letter(ci)
        ws.conditional_formatting.add(f'{letter}2:{letter}{max(ws.max_row, 2)}', FormulaRule(
            formula=[f'LOWER(TRIM({letter}2&""))="x"'], fill=X_FILL, font=X_FONT, stopIfTrue=True))

    # Apply edit tracking to price tab
    _apply_edit_tracking_gsheets(ws, price_var_indices, entry_rules=True, sys_values=sys_prices)

    # Apply edit tracking to modifier summary tab
    if 'Modifier Summary' in [s.title for s in wb_out.worksheets]:
        ws_ms_target = wb_out['Modifier Summary']
        ms_headers_out = [c.value for c in next(ws_ms_target.iter_rows(min_row=1, max_row=1))]
        # Auto-discover variant columns from the Modifier Summary headers.
        # Treat any non-metadata, non-reference column as a variant — this
        # adapts to whatever variant columns the restaurant actually uses
        # (sushi: 'Roll', bento: 'Dinner', combo sets: 'Seapark', etc.).
        ms_meta_cols = {'Modifier Group', 'Modifier Item', 'Notes'}
        ms_var_indices = []
        for i, h in enumerate(ms_headers_out):
            h_str = str(h).strip() if h else ''
            if not h_str:
                continue
            if h_str in ms_meta_cols:
                continue
            if h_str.startswith('_ref_'):
                continue
            ms_var_indices.append(i + 1)
        # x / X over a Modifier Summary price = remove it (all its instances) —
        # ahead of the entry and apricot rules, like on Dish Item.
        for ci in ms_var_indices:
            letter = get_column_letter(ci)
            ws_ms_target.conditional_formatting.add(
                f'{letter}2:{letter}{max(ws_ms_target.max_row, 2)}', FormulaRule(
                    formula=[f'LOWER(TRIM({letter}2&""))="x"'], fill=X_FILL, font=X_FONT,
                    stopIfTrue=True))
        # modifier prices may be negative ("Less Whole Wing" −$1.70 — real restaurants)
        _apply_edit_tracking_gsheets(ws_ms_target, ms_var_indices, entry_rules=True, allow_negative=True)

    # ── Green sheet-tab color for the tabs the specialist EDITS ──
    # 'Modifier Rule' is appended later by modifier_rule_pricing.py
    # (it colors its own tab green); here we color the two Action-3 owns.
    for _edit_tab in ('Dish Item', 'Modifier Summary'):
        if _edit_tab in wb_out.sheetnames:
            wb_out[_edit_tab].sheet_properties.tabColor = '00B050'

    # ── Canonical tab order (missing sheets skipped; extras kept at end) ──
    # Final order: Summary, Dish Item, Questionnaire, Modifier Summary,
    # Modifier Rule (added by compare), Fried Specialty.
    # The raw 'Modifier' tab (Part 2 input) and any others fall after these.
    _tab_order = ['Summary', 'Dish Item', 'Questionnaire', 'Modifier Summary',
                  'Modifier Rule', 'Fried Specialty', CODE_TAB]
    _present = wb_out.sheetnames
    _ordered = [n for n in _tab_order if n in _present]
    _rest = [n for n in _present if n not in _ordered]
    wb_out._sheets.sort(key=lambda s: (_ordered + _rest).index(s.title))

    _numeric_display(wb_out)   # every price a plain number — no '$X.XX' strings
    wb_out.save(out)
    wb_2a.close()
    wb_2b.close()
    wb_out.close()

    # Print summary
    print(f"\n{'='*50}")
    print(f"  Match rate: {items_with_2b} / {total_items} ({match_pct}%)")
    print(f"  (Excluded from match rate: {excluded_items} items — Drinks, Sauces, Side Orders)")
    print(f"  Price changes: {len(diffs)} ({len(inc)} increases, {len(dec)} decreases)")
    print(f"  Confirmed no change (green): {confirmed_no_change}")
    print(f"  Modifier Summary items: {ms_item_count}")
    if dec:
        print(f"\n  ⚠ DECREASES DETECTED:")
        for d in dec:
            print(f"    {d['name']} [{d['variant']}]: {_fmt(d['old'])} → {_fmt(d['new'])}")
    print(f"\n  Output: {out}")
    print(f"{'='*50}")
    print("Done.")


if __name__ == '__main__':
    main()