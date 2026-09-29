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
      - Questionnaire: guided-interview (DISHES) — P1 + P-COLUMN + P2 + P3 +
        standing questions. Free-text interview columns (Question /
        Recommendation / Answer / Answer Detail), no dropdowns.

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
    from openpyxl.styles import PatternFill, Font
    from openpyxl.utils import get_column_letter
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

BLUE_FILL = PatternFill('solid', fgColor='ADD8E6')
PINK_FILL = PatternFill('solid', fgColor='FFB6C1')
GREEN_FILL = PatternFill('solid', fgColor='C6EFCE')
HEADER_FILL = PatternFill('solid', fgColor='D9E1F2')
# Flags a dish whose prices never changed (no blue/pink variant edit) so the
# team can see at a glance which rows still need action. Applied to col B only.
YELLOW_FILL = PatternFill('solid', fgColor='FCF7DE')
HEADER_FONT = Font(bold=True, size=11)


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
        has_2b_any = False
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
    ]
    for fill, desc in legend:
        if fill:
            ws_s.cell(row=gr, column=1, value='    ').fill = fill
        else:
            ws_s.cell(row=gr, column=1, value='    ')
        ws_s.cell(row=gr, column=2, value=desc)
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

    stats = [
        ('Total price changes', len(diffs)),
        ('Price increases (blue)', len(inc)),
        ('Price decreases (pink)', len(dec)),
        ('Confirmed no change (green)', confirmed_no_change),
        ('Items with menu prices', items_with_2b),
        ('Items without menu prices (kept from system)', items_without_2b),
        ('Inferred column routings (P-COLUMN — verify)', inferred_col_count),
    ]
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
    # Compact layout: the lengthy free-text interview columns
    # (Question / Recommendation / Answer / Answer Detail) were removed per
    # specialist feedback. Each row is identified by ID + Category, and the
    # Notes column carries the one-line context. Notes is the last column.
    q_headers = ['Question ID', 'Category', 'Dish Group', 'Dish Name',
                 'Current Price', 'Expected Price', 'Menu Price', 'Notes']
    NOTES_COL = 8
    for ci, h in enumerate(q_headers, 1):
        ws_q.cell(row=1, column=ci, value=h).font = HEADER_FONT
        ws_q.cell(row=1, column=ci).fill = HEADER_FILL
    qr = 2

    def _interview(row, question='', recommendation=''):
        """No-op: the Question/Recommendation/Answer/Answer Detail columns were
        removed. Retained as a stub so existing call sites need no change."""
        return

    # Low-priority dish groups — these are rarely on paper menus and rarely change
    LOW_PRIORITY_GROUPS = {'Drink', 'Sauces', 'Side Order'}

    # --- Section: P1 (on menu, not matched to system) --- highest priority
    ws_q.cell(row=qr, column=1, value='P1 — ON MENU, NOT MATCHED TO SYSTEM').font = Font(bold=True, size=11, color='C00000')
    qr += 1

    if 'Unmatched' in wb_2b.sheetnames:
        ws_u = wb_2b['Unmatched']
        u_rows = list(ws_u.iter_rows(min_row=2, values_only=True))
        for row in u_rows:
            if row[0]:
                dish_name = str(row[3]).strip() if row[3] else ''
                paper_price = str(row[4]).strip() if row[4] else ''
                ws_q.cell(row=qr, column=1, value=str(row[0]).strip() if row[0] else '')
                ws_q.cell(row=qr, column=2, value=str(row[1]).strip() if row[1] else '')
                ws_q.cell(row=qr, column=3, value=str(row[2]).strip() if row[2] else '')
                ws_q.cell(row=qr, column=4, value=dish_name)
                ws_q.cell(row=qr, column=7, value=paper_price)
                _interview(
                    qr,
                    question=(f"Paper lists '{dish_name}'"
                              + (f" ({paper_price})" if paper_price else '')
                              + ' with no system match — existing dish renamed, a new dish, or noise (typo/dupe)?'),
                    recommendation='Match to existing dish (put the system name in Answer Detail), or New dish / Ignore.')
                if len(row) > 5 and row[5]:
                    ws_q.cell(row=qr, column=NOTES_COL, value=str(row[5]).strip())
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
                price = str(row[3]).strip() if len(row) > 3 and row[3] else ''
                ws_q.cell(row=qr, column=1, value=f'P-COLUMN-{i}')
                ws_q.cell(row=qr, column=2, value='P-COLUMN (inferred column, verify)')
                ws_q.cell(row=qr, column=4, value=dish)
                ws_q.cell(row=qr, column=7, value=price)
                _interview(
                    qr,
                    question=f"Paper '{paper_v}' price for '{dish}' was routed to '{routed_v}' (only populated system column) — is that the right column?",
                    recommendation='Confirm column (leave Answer = confirm), or name the correct column in Answer Detail.')
                ws_q.cell(row=qr, column=NOTES_COL, value=f"Paper '{paper_v}' price routed to '{routed_v}' (only populated column) — verify column is correct")
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
        ws_q.cell(row=qr, column=5, value=_fmt(d['old']))
        exp = group_expected.get((d['group'], d['variant']))
        ws_q.cell(row=qr, column=6, value=_fmt(exp) if exp else '')
        ws_q.cell(row=qr, column=7, value=_fmt(d['new']))
        _interview(
            qr,
            question=f"Paper shows a price DECREASE for '{d['name']}' in {d['variant']}: {_fmt(d['old'])} → {_fmt(d['new'])}. Confirm the decrease?",
            recommendation='Confirm decrease, or Reject (keep system price).')
        ws_q.cell(row=qr, column=NOTES_COL, value=f"Decrease in {d['variant']}: {_fmt(d['old'])} → {_fmt(d['new'])}")
        for ci in range(1, NOTES_COL + 1):
            ws_q.cell(row=qr, column=ci).fill = PINK_FILL
        qr += 1
    if not dec:
        ws_q.cell(row=qr, column=1, value='(none)')
        qr += 1

    # --- Section: P3a (in system, not on menu — ENTREES/DISHES) --- medium priority
    qr += 1
    ws_q.cell(row=qr, column=1, value='P3 — IN SYSTEM, NOT ON MENU (DISHES)').font = Font(bold=True, size=11, color='4472C4')
    ws_q.cell(row=qr, column=NOTES_COL, value='Ask client if discontinued or still available')
    qr += 1

    p3_entree_count = 0
    p3_operational = []  # collect low-priority items for later
    for ri in range(1, len(rows_2a)):
        r2a = list(rows_2a[ri]) if ri < len(rows_2a) else []
        r2b = list(rows_2b[ri]) if ri < len(rows_2b) else []
        dish = str(r2a[name_col]).strip() if len(r2a) > name_col and r2a[name_col] else ''
        group = str(r2a[group_col]).strip() if len(r2a) > group_col and r2a[group_col] else ''
        if not dish:
            continue

        has_2b = False
        for v, ci in var_cols.items():
            p2b = _parse(r2b[ci]) if ci < len(r2b) else None
            if p2b is not None:
                has_2b = True
                break

        if not has_2b:
            has_2a_prices = {}
            for v, ci in var_cols.items():
                p2a = _parse(r2a[ci]) if ci < len(r2a) else None
                if p2a is not None:
                    has_2a_prices[v] = p2a

            if has_2a_prices:
                current = ', '.join(f"{v}={_fmt(p)}" for v, p in has_2a_prices.items())
                expected_parts = []
                for v, old_p in has_2a_prices.items():
                    exp = group_expected.get((group, v))
                    if exp is not None and abs(exp - old_p) > 0.001:
                        expected_parts.append(f"{v}={_fmt(exp)}")
                expected_str = ', '.join(expected_parts) if expected_parts else '(no change expected)'

                if group in LOW_PRIORITY_GROUPS:
                    p3_operational.append((group, dish, current, expected_str))
                else:
                    p3_entree_count += 1
                    ws_q.cell(row=qr, column=1, value=f'P3-{p3_entree_count}')
                    ws_q.cell(row=qr, column=2, value='P3 (in system, not on menu)')
                    ws_q.cell(row=qr, column=3, value=group)
                    ws_q.cell(row=qr, column=4, value=dish)
                    ws_q.cell(row=qr, column=5, value=current)
                    ws_q.cell(row=qr, column=6, value=expected_str)
                    _interview(
                        qr,
                        question=f"'{dish}' is in the system but not on the paper menu — keep it as-is or discontinue?",
                        recommendation='Keep as-is, or Discontinue (route discontinues to a human — Part 2 cannot remove rows).')
                    qr += 1
    if p3_entree_count == 0:
        ws_q.cell(row=qr, column=1, value='(none)')
        qr += 1

    # --- Section: Standing additional questions ---
    qr += 1
    ws_q.cell(row=qr, column=1, value='ADDITIONAL QUESTIONS').font = Font(bold=True, size=11, color='4472C4')
    qr += 1
    # v4.3: the static 'A-1 Party Tray pricing — confirm or update?' row is
    # gone. price_inference.py writes a data-driven P-TRAY section instead
    # (detected formula + one row per tray that needs a callback answer).
    for qid, note in [('A-1', 'Triple Combo / Special Combo pricing?'),
                       ('A-2', 'Fried Specialties combo column pricing?'),
                       ('A-3', 'Side choice upcharge tiers — confirm?')]:
        ws_q.cell(row=qr, column=1, value=qid)
        ws_q.cell(row=qr, column=2, value='Additional')
        ws_q.cell(row=qr, column=NOTES_COL, value=note)
        qr += 1

    # --- Section: P5 (operational — Drinks, Sauces, Side Orders) ---
    qr += 1
    ws_q.cell(row=qr, column=1, value='P5 — DRINKS, SAUCES, SIDE ORDERS').font = Font(bold=True, size=11, color='808080')
    ws_q.cell(row=qr, column=NOTES_COL, value='Rarely on paper menu. Skip unless client raises.')
    qr += 1
    p3_op_count = 0
    for group, dish, current, expected_str in p3_operational:
        p3_op_count += 1
        ws_q.cell(row=qr, column=1, value=f'P5-{p3_op_count}')
        ws_q.cell(row=qr, column=2, value='P5 (drinks, sauces, side orders)')
        ws_q.cell(row=qr, column=3, value=group)
        ws_q.cell(row=qr, column=4, value=dish)
        ws_q.cell(row=qr, column=5, value=current)
        ws_q.cell(row=qr, column=6, value=expected_str)
        _interview(
            qr,
            question=f"'{dish}' ({group}) is in the system but not on the paper menu — keep as-is or discontinue?",
            recommendation='Keep as-is (low priority — skip unless client raises).')
        for ci in range(1, NOTES_COL + 1):
            ws_q.cell(row=qr, column=ci).font = Font(color='808080')
        qr += 1

    p3_count = p3_entree_count + p3_op_count

    ws_q.column_dimensions['A'].width = 12   # Question ID
    ws_q.column_dimensions['B'].width = 34   # Category
    ws_q.column_dimensions['C'].width = 22   # Dish Group
    ws_q.column_dimensions['D'].width = 38   # Dish Name
    ws_q.column_dimensions['E'].width = 20   # Current Price
    ws_q.column_dimensions['F'].width = 20   # Expected Price
    ws_q.column_dimensions['G'].width = 16   # Menu Price
    ws_q.column_dimensions['H'].width = 55   # Notes

    # ── Fried Specialty tab (multi-column side pricing) ───────────
    # If the OCR output has a 'Fried Specialty Sides' sheet (passed through
    # from action2), build a dedicated tab showing the full side pricing
    # and the price difference (combo price - plain price) per item.
    if 'Fried Specialty Sides' in wb_2b.sheetnames:
        ws_fs_src = wb_2b['Fried Specialty Sides']
        fs_rows = list(ws_fs_src.iter_rows(values_only=True))
        if len(fs_rows) > 1:
            ws_fs = wb_out.create_sheet('Fried Specialty')
            fs_headers = list(fs_rows[0])  # Code, Name, Plain, col2, col3, ...
            side_cols = fs_headers[3:]  # Everything after Code, Name, Plain

            # --- Section 1: Full prices ---
            r = 1
            ws_fs.cell(row=r, column=1, value='FULL PRICES (from paper menu)').font = Font(bold=True, size=12)
            r += 1
            for ci, h in enumerate(fs_headers, 1):
                ws_fs.cell(row=r, column=ci, value=h).font = HEADER_FONT
                ws_fs.cell(row=r, column=ci).fill = HEADER_FILL
            r += 1

            for row in fs_rows[1:]:
                for ci, val in enumerate(row, 1):
                    if val is not None:
                        if ci >= 3 and isinstance(val, (int, float)):
                            ws_fs.cell(row=r, column=ci, value=f'${val:.2f}')
                        else:
                            ws_fs.cell(row=r, column=ci, value=val)
                r += 1

            # --- Section 2: Deltas (combo price - plain price) ---
            r += 1
            ws_fs.cell(row=r, column=1, value='PRICE DIFFERENCE (column price minus Plain)').font = Font(bold=True, size=12)
            r += 1
            delta_headers = ['Code', 'Name', 'Plain'] + [f'Δ {h}' for h in side_cols]
            for ci, h in enumerate(delta_headers, 1):
                ws_fs.cell(row=r, column=ci, value=h).font = HEADER_FONT
                ws_fs.cell(row=r, column=ci).fill = HEADER_FILL
            r += 1

            for row in fs_rows[1:]:
                code = row[0] if row[0] else ''
                name = row[1] if row[1] else ''
                plain = row[2] if len(row) > 2 and row[2] else None

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
            for row in fs_rows[1:]:
                for val in row[3:]:
                    if val is not None:
                        side_vals.add(val)
            if len(side_vals) <= 3:
                ws_fs.cell(row=r, column=2, value=f'Flat combo prices observed: {", ".join(f"${v:.2f}" for v in sorted(side_vals))}')

            ws_fs.column_dimensions['A'].width = 8
            ws_fs.column_dimensions['B'].width = 32
            for ci in range(3, len(fs_headers) + 1):
                ws_fs.column_dimensions[chr(64 + ci)].width = 18

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
    for v in _detect_extra_variants(mod_headers):
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
        ms_row += 1

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

    def _apply_edit_tracking_gsheets(ws_target, var_col_indices):
        """
        Add hidden reference columns on the same sheet and apply apricot CF.
        Google Sheets compatible (no cross-sheet references).
        
        For each price column at index ci, creates a reference column at
        (ref_start + offset) with the same values. CF compares visible vs reference.
        """
        max_row = ws_target.max_row
        if max_row < 2 or not var_col_indices:
            return

        # Start reference columns well to the right of all data
        # Find the rightmost used column first
        max_col = ws_target.max_column or 1
        ref_start = max_col + 3  # leave a gap

        GREY_FILL = PatternFill('solid', fgColor='D9D9D9')
        GREY_FONT = Font(color='808080', size=8)

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
            ref_cell = f'${ref_letter}2'
            this_cell = f'${src_letter}2'

            ws_target.conditional_formatting.add(cell_range, FormulaRule(
                formula=[f'AND({this_cell}<>"",{this_cell}<>{ref_cell})'],
                fill=APRICOT_FILL,
                stopIfTrue=True
            ))

    # Apply edit tracking to price tab
    price_var_indices = [ci + 1 for ci in var_cols.values()]  # 1-indexed
    _apply_edit_tracking_gsheets(ws, price_var_indices)

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
        _apply_edit_tracking_gsheets(ws_ms_target, ms_var_indices)

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
                  'Modifier Rule', 'Fried Specialty']
    _present = wb_out.sheetnames
    _ordered = [n for n in _tab_order if n in _present]
    _rest = [n for n in _present if n not in _ordered]
    wb_out._sheets.sort(key=lambda s: (_ordered + _rest).index(s.title))

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
    print(f"  Questionnaire P3 items: {p3_count}")
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