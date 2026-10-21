#!/usr/bin/env python3
"""
specialist_requests.py — what the specialist asked for OUTSIDE price cells
==========================================================================

Standalone script. Dependencies: openpyxl only. Part 2, run first, on the
specialist-edited Price Comparison Report.

Usage:
    python specialist_requests.py <edited_report.xlsx> [--dir <content_dir>]

Reads two conventions the scripts cannot carry out themselves and hands them
to Wonka (part2.md §0b):

  * Dish Item, column B (Dish Name:Primary): a dish name with strikethrough —
    the whole cell or any part of it — = make that dish permanently
    unavailable.
  * Questionnaire, column I (Notes → Wonka): any note the specialist wrote =
    an instruction for Wonka, listed with its row's context (ID, dish,
    prices, Details). Only a report that has the "Details" column is read —
    in an older report column H "Notes" holds the script's own text.
  * Fried Specialty tab (2026-10-16): the Notes → Wonka column of its price
    tables (one note per dish and size) and the NOTES → WONKA box under them
    (instructions or formulas for the whole tab).

Prices are NOT touched here: price edits (and x = remove a price) are applied
by action5 / modifier_rule_pricing apply as usual.

Writes the lists into <content_dir>/part2_summary.json (key
"specialist_requests", merged — the same file action5 / apply / part2_report
use) and prints them. Read-only on the report; ALWAYS exits 0.
"""
import argparse
import json
import os
import sys

try:
    from openpyxl import load_workbook
    from openpyxl.cell.rich_text import CellRichText
except ImportError:
    print("ERROR: openpyxl >= 3.1 required. Install: pip install openpyxl")
    sys.exit(1)


def _text(v):
    """Plain text of a cell value (rich text included)."""
    if v is None:
        return ''
    return str(v).strip()


def _strike(cell):
    """(struck?, struck part) — whole-cell strikethrough or any struck run."""
    v = cell.value
    if isinstance(v, CellRichText):
        parts = [str(b.text) for b in v
                 if not isinstance(b, str) and getattr(b.font, 'strike', False)]
        if parts:
            return True, ''.join(parts).strip()
    if cell.font is not None and cell.font.strike:
        return True, _text(v)
    return False, ''


def _unavailable(wb):
    """Dish Item rows whose dish name (column B) carries strikethrough."""
    tab = 'Dish Item' if 'Dish Item' in wb.sheetnames else ('price' if 'price' in wb.sheetnames else None)
    if tab is None:
        return []
    ws = wb[tab]
    hdr = [_text(c.value) for c in ws[1]]
    name_col = hdr.index('Dish Name:Primary') + 1 if 'Dish Name:Primary' in hdr else 2
    group_col = hdr.index('Dish Group') + 1 if 'Dish Group' in hdr else 1
    second_col = hdr.index('Dish Name:Secondary') + 1 if 'Dish Name:Secondary' in hdr else 3
    out = []
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(row=r, column=name_col)
        name = _text(cell.value)
        if not name:
            continue
        struck, part = _strike(cell)
        if struck:
            out.append({'row': r, 'dish_group': _text(ws.cell(row=r, column=group_col).value),
                        'dish': name, 'secondary': _text(ws.cell(row=r, column=second_col).value),
                        'struck_text': part, 'partial': part != name})
    return out


def _notes(wb):
    """(notes, skipped_reason) — every Questionnaire row with a note in the
    Notes → Wonka column, with its context."""
    if 'Questionnaire' not in wb.sheetnames:
        return [], 'no Questionnaire tab'
    ws = wb['Questionnaire']
    hdr = [_text(c.value) for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(hdr) if h}
    notes_col = next((i + 1 for i, h in enumerate(hdr) if h.lower().startswith('notes')), None)
    if 'Details' not in col or notes_col is None:
        return [], ('report predates the Notes → Wonka column (its Notes column holds the '
                    "script's own text) — notes not read")

    def get(r, h):
        c = col.get(h)
        return ws.cell(row=r, column=c).value if c else None

    out = []
    for r in range(2, ws.max_row + 1):
        note = _text(ws.cell(row=r, column=notes_col).value)
        if not note:
            continue
        out.append({'row': r, 'question_id': _text(get(r, 'Question ID')),
                    'category': _text(get(r, 'Category')), 'dish_group': _text(get(r, 'Dish Group')),
                    'dish': _text(get(r, 'Dish Name')), 'current_price': get(r, 'Current Price'),
                    'expected_price': get(r, 'Expected Price'), 'menu_price': get(r, 'Menu Price'),
                    'details': _text(get(r, 'Details')), 'note': note})
    return out, None


FS_TAB = 'Fried Specialty'
FS_TABLES = ('PRICE DIFFERENCE — ', 'ABSOLUTE PRICE — ')     # the editable tables (modifier_rule_pricing)
FS_NOTES_BOX = 'NOTES → WONKA'
FS_HEADINGS = ('HOW THE PAPER WAS READ', 'FULL PRICE — ', 'PRICE DIFFERENCE', 'ABSOLUTE PRICE — ',
               'FRIED SPECIALTY — ', 'COLUMN → MODIFIER', 'DISH NAME MAPPING', 'PROPAGATION DETAIL')


def _fs_notes(wb):
    """Notes → Wonka on the Fried Specialty tab: the last column of its price
    tables (one note per dish and size) and the notes box under them
    (instructions or formulas for the whole tab)."""
    if FS_TAB not in wb.sheetnames:
        return []
    ws = wb[FS_TAB]
    rows = [[c.value for c in row] for row in ws.iter_rows()]
    out = []
    i = 0
    while i < len(rows):
        title = _text(rows[i][0]) if rows[i] else ''
        if title.startswith(FS_TABLES) and i + 1 < len(rows):
            hdr = [_text(v) for v in rows[i + 1]]
            col = {h: j for j, h in enumerate(hdr) if h}
            nc = next((j for h, j in col.items() if h.lower().startswith('notes')), None)
            j = i + 2
            # to the next section of the tab (a blank row inside the table is skipped)
            while j < len(rows) and not (rows[j] and _text(rows[j][0]).startswith(FS_HEADINGS + (FS_NOTES_BOX,))):
                r = rows[j]
                if not (r and any(v not in (None, '') for v in r)):
                    j += 1
                    continue
                note = _text(r[nc]) if nc is not None and nc < len(r) else ''
                if note:
                    dish, size = _text(r[col['Dish']]) if 'Dish' in col else '', \
                        _text(r[col['Size']]) if 'Size' in col else ''
                    out.append({'tab': FS_TAB, 'row': j + 1, 'question_id': '', 'category': 'Fried Specialty note',
                                'dish_group': _text(r[col['Dish Group']]) if 'Dish Group' in col else '',
                                'dish': f'{dish} [{size}]' if size else dish, 'current_price': None,
                                'expected_price': None, 'menu_price': None,
                                'details': title.split(' — ')[0].title() + ' table', 'note': note})
                j += 1
            i = j
            continue
        if title.startswith(FS_NOTES_BOX):
            j = i + 1
            # the box ends at the next section heading of the tab
            while j < len(rows) and not (rows[j] and _text(rows[j][0]).startswith(FS_HEADINGS)):
                note = ' '.join(_text(v) for v in rows[j] if _text(v))
                if note:
                    out.append({'tab': FS_TAB, 'row': j + 1, 'question_id': '', 'category': 'Fried Specialty note',
                                'dish_group': '', 'dish': '', 'current_price': None, 'expected_price': None,
                                'menu_price': None, 'details': 'notes box (the whole tab)', 'note': note})
                j += 1
            i = j
            continue
        i += 1
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('report', help='the specialist-edited Price Comparison Report (.xlsx)')
    ap.add_argument('--dir', default='.',
                    help='Part 2 content dir holding part2_summary.json (default: current dir)')
    args = ap.parse_args()

    wb = load_workbook(args.report, rich_text=True)
    unavailable = _unavailable(wb)
    notes, skipped = _notes(wb)
    notes += _fs_notes(wb)
    wb.close()

    path = os.path.join(args.dir, 'part2_summary.json')
    summary = {}
    if os.path.isfile(path):
        try:
            with open(path, encoding='utf-8') as f:
                summary = json.load(f)
        except (OSError, ValueError):
            summary = {}
    summary['specialist_requests'] = {
        'report': args.report,
        'make_unavailable': {'total': len(unavailable), 'dishes': unavailable},
        'notes_for_wonka': {'total': len(notes), 'notes': notes,
                            'skipped_reason': skipped},
    }
    os.makedirs(args.dir, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(summary, f, indent=2, default=str)

    print('=' * 60)
    print('  SPECIALIST REQUESTS — for Wonka (from the edited report)')
    print(f'  Make permanently unavailable (strikethrough on the dish name): {len(unavailable)}')
    for u in unavailable:
        extra = f' — only "{u["struck_text"]}" is struck' if u['partial'] else ''
        print(f'    - Dish Item row {u["row"]}: {u["dish"]} ({u["dish_group"]}){extra}')
    if skipped:
        print(f'  Notes → Wonka: not read — {skipped}')
    else:
        print(f'  Notes → Wonka (Questionnaire column I, Fried Specialty tab): {len(notes)}')
        for n in notes:
            where = ' · '.join(x for x in (n['question_id'], n['dish'], n['details'] if n.get('tab') else '') if x)
            print(f'    - {n.get("tab", "Questionnaire")} row {n["row"]} [{where}]: "{n["note"]}"')
    print('  The scripts apply none of these — Wonka carries them out (part2.md §0b).')
    print(f'  Summary JSON: {path}')
    print('=' * 60)
    return 0


if __name__ == '__main__':
    sys.exit(main())
