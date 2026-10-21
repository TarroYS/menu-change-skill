"""Fried Specialty / Modifier Rule tabs (maintainer, 2026-10-16).

    python3 check_fried_specialty.py <bundle_dir>

Pins: both tabs are built whenever Fried Specialty dishes are priced per dish
on their Select Rice items — also when the paper prints no Fried Specialty
table (scenario A) or only some of the dishes (scenario B); the Fried
Specialty tab shows FULL PRICE (paper; grey = not printed, today's price),
PRICE DIFFERENCE and ABSOLUTE PRICE (the specialist types there — plain
numbers, no + sign), Notes → Wonka; Part 2 applies those cells, refuses what
is not a price, never mixes up a sorted table, reports the same rule edited
differently on both tabs, and hands the notes to Wonka. A restaurant that
prices Select Rice once for every dish (on modifier_item) gets no tables.
Exit 1 if any check fails."""
import json
import os
import shutil
import sys
import tempfile

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'modifier_rule'))
sys.path.insert(0, os.path.join(HERE, '..', 'edits'))
sys.path.insert(0, os.path.join(HERE, '..', 'lib'))
from build_fixture import make_raw  # noqa: E402
import check_edits as ce  # noqa: E402
from xlformula import cf_kind  # noqa: E402

RESULTS = []
FS = 'Fried Specialty'
TITLES = {'FULL': 'FULL PRICE — ', 'DIFF': 'PRICE DIFFERENCE — ', 'ABS': 'ABSOLUTE PRICE — '}


def check(desc, ok, detail=''):
    RESULTS.append(bool(ok))
    print(f"  {'PASS' if ok else 'FAIL'}  {desc}" + ('' if ok or not detail else f'  -- {detail}'))


def kind(rule):
    d = rule.dxf
    fill = (d.fill.fgColor.rgb or '') + (d.fill.bgColor.rgb or '') if d is not None and d.fill is not None else ''
    font = (d.font.color.rgb or '') if d is not None and d.font is not None and d.font.color is not None else ''
    if 'C00000' in font:
        return 'INVALID'
    if 'FFC000' in fill:
        return 'EDIT'
    return '?'


S = 'Select Rice'


def build(paper):
    """R99937 Side Wok: Fried Specialties priced per dish on Select Rice items —
    four sizes by difference (one on 'Order', not 'Specialty'), two absolute;
    a Honey Sauce add-on and a group-level Select Rice attachment."""
    variants = ['Order', 'Specialty', 'Small', 'Large']
    dishes = [('Chicken Wings', 'Fried Specialties', {'Specialty': 7.99}),
              ('Chicken Fingers', 'Fried Specialties', {'Specialty': 8.50}),
              ('Fried Shrimp', 'Fried Specialties', {'Specialty': 9.50}),
              ('Teriyaki Beef', 'Fried Specialties', {'Order': 8.50}),
              ('Fried Fish', 'Fried Specialties', {'Specialty': 0.00}),
              ('Fried Scallop', 'Fried Specialties', {'Specialty': 0.00}),
              ('Egg Roll Platter', 'Fried Specialties', {'Specialty': 6.00}),
              ('Fried Wonton', 'Fried Specialties', {'Specialty': 5.00}),      # side rules with no price
              ('Lemon Chicken', 'Fried Specialties', {'Small': 7.00, 'Large': 10.00}),   # a rule for every size
              ('Egg Roll', 'Appetizers', {'Order': 1.75})]
    mod_groups = {S: (1, ['No Rice', 'By Itself', 'French Fries', 'Pork Fried Rice', 'Plain Lo Mein', 'White Rice'],
                      ['Fried Specialties']),
                  'Sauce': (0, ['Honey Sauce'], ['Fried Specialties']),
                  'Select Sauce': (0, ['Honey Mustard'], ['Fried Specialties'])}   # an optional add-on
    mod_items = {m: {'Order': 0.0} for m in ('No Rice', 'By Itself', 'French Fries', 'Pork Fried Rice',
                                             'Plain Lo Mein', 'White Rice', 'Honey Mustard')}
    mod_items['Honey Sauce'] = {'Order': 0.50}
    rules = [('Fried Specialties', '', 'Specialty', S, '', None)]           # attaches Select Rice (a gate)
    for dish, var in (('Chicken Wings', 'Specialty'), ('Chicken Fingers', 'Specialty'),
                      ('Fried Shrimp', 'Specialty'), ('Teriyaki Beef', 'Order')):
        rules += [('', dish, var, '', 'French Fries', 1.60), ('', dish, var, '', 'Pork Fried Rice', 2.40)]
    rules += [('', 'Chicken Wings', 'Specialty', '', 'Plain Lo Mein', 2.90),
              ('', 'Fried Fish', 'Specialty', '', 'By Itself', 8.25),
              ('', 'Fried Fish', 'Specialty', '', 'French Fries', 10.95),
              ('', 'Fried Fish', 'Specialty', '', 'Pork Fried Rice', 10.95),
              ('', 'Fried Fish', '', 'Sauce', 'Honey Sauce', 0.50),                # an add-on, not a side
              ('', 'Fried Fish', 'Specialty', '', 'Honey Mustard', 0.75),          # a "Select" add-on, not a side
              ('', 'Fried Scallop', 'Specialty', '', 'No Rice', 0.00),             # an alias before By Itself
              ('', 'Fried Scallop', 'Specialty', '', 'By Itself', 6.95),
              ('', 'Fried Scallop', 'Specialty', '', 'French Fries', 9.95),     # no Pork Fried Rice rule
              ('', 'Fried Wonton', 'Specialty', '', 'French Fries', None),
              ('', 'Lemon Chicken', '', '', 'French Fries', 1.50)]
    make_raw('raw.xlsx', 'R99937', 'Side Wok', variants, dishes, mod_groups, mod_items, rules)
    secs = [{'menu_section': 'Appetizers', 'price_columns': [], 'default_price': None,
             'items': [{'name': 'Egg Roll', 'prices': {'Order': 1.95}}]}]
    if paper:
        cols = ['w. French Fries', 'w. Pork Fried Rice']
        secs.append({'menu_section': 'Fried Specialties', 'price_columns': ['Specialty'], 'default_price': None,
                     'side_pricing': {'columns': cols}, 'items': [
                         {'name': 'Chicken Wings', 'prices': {'Specialty': 7.99},
                          'side_prices': {cols[0]: 9.79, cols[1]: 10.39}},          # fries +1.80 (was 1.60)
                         {'name': 'Chicken Fingers', 'prices': {'Specialty': 8.50},
                          'side_prices': {cols[0]: 10.10, cols[1]: 10.90}},         # unchanged
                         {'name': 'Fried Fish', 'prices': {'Specialty': 8.50},
                          'side_prices': {cols[0]: 11.25, cols[1]: 11.25}}]})       # absolute, all up
    json.dump({'restaurant': {'name': 'Side Wok'}, 'sections': secs}, open('ocr.json', 'w'))


def fs_tables(path='report.xlsx'):
    """{'FULL' / 'DIFF' / 'ABS': {(dish, size): {header: (value, grey font, fill rgb)}}} + rows / headers."""
    ws = load_workbook(path)[FS]
    out, where = {}, {}
    for r in range(1, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or '')
        t = next((k for k, v in TITLES.items() if a.startswith(v)), None)
        if not t:
            continue
        hdr = {str(ws.cell(r + 1, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(r + 1, c).value}
        out[t], where[t] = {}, {}
        rr = r + 2
        while ws.cell(rr, 2).value:
            key = (ws.cell(rr, 2).value, ws.cell(rr, 3).value or '')
            cells = {}
            for h, c in hdr.items():
                cell = ws.cell(rr, c)
                cells[h] = (cell.value, bool(cell.font and cell.font.color is not None and cell.font.italic),
                            cell.fill.fgColor.rgb[-6:] if cell.fill and cell.fill.fill_type else '')
            out[t][key] = cells
            where[t][key] = (rr, hdr)
            rr += 1
    return out, where


def final_rules(path='final_raw.xlsx'):
    ws = load_workbook(path)['modifier_rule']
    return {(r[9] or '', r[10] or '', r[12] or ''): r[2] for r in ws.iter_rows(min_row=4, values_only=True)
            if r and str(r[0] or '').startswith('{')}


def p_fs(path='report.xlsx'):
    return [' '.join(str(c) for c in row if c not in (None, ''))
            for row in load_workbook(path)['Questionnaire'].iter_rows(values_only=True)
            if str(row[0] or '').startswith('P-FS')]


def scenario_a(scripts, work):
    print('\nScenario A — the paper prints no Fried Specialty table')
    d = os.path.join(work, 'a')
    os.makedirs(d)
    os.chdir(d)
    build(paper=False)
    out = ce.part1(scripts, compare=True)
    check('Part 1 runs (compare included)', all(v[0] == 0 for v in out.values()),
          {k: v[1][-300:] for k, v in out.items() if v[0]})
    wb = load_workbook('report.xlsx')
    check('both tabs are built although the paper prints no side table', 'Modifier Rule' in wb.sheetnames
          and FS in wb.sheetnames, wb.sheetnames)
    _c, mrows = ce.mrv(wb['Modifier Rule'])
    mkeys = {(d or '', v or '', m or '') for d, v, m in mrows}
    check('the Modifier Rule tab lists every Fried Specialty rule — a dish on "Order" too, the Honey Sauce add-on',
          ('Teriyaki Beef', 'Order', 'French Fries') in mkeys and ('Fried Fish', '', 'Honey Sauce') in mkeys
          and ('Fried Scallop', 'Specialty', 'By Itself') in mkeys, sorted(mkeys)[:6])
    t, _w = fs_tables()
    check('PRICE DIFFERENCE lists the difference-priced sizes, ABSOLUTE PRICE the absolute ones',
          sorted(t.get('DIFF', {})) == [('Chicken Fingers', 'Specialty'), ('Chicken Wings', 'Specialty'),
                                        ('Fried Shrimp', 'Specialty'), ('Lemon Chicken', ''), ('Teriyaki Beef', 'Order')]
          and sorted(t.get('ABS', {})) == [('Fried Fish', 'Specialty'), ('Fried Scallop', 'Specialty')],
          (sorted(t.get('DIFF', {})), sorted(t.get('ABS', {}))))
    hdr = list(t['DIFF'][('Chicken Wings', 'Specialty')])
    check('one column per Select Rice side the dishes price (not the add-ons, even one in a "Select" group; not '
          'White Rice; not By Itself / No Rice); a size whose side rules have no price is left out (Fried Wonton)',
          [h for h in hdr if h in ('French Fries', 'Pork Fried Rice', 'Plain Lo Mein', 'Honey Sauce', 'Honey Mustard',
                                   'White Rice', 'By Itself', 'No Rice')] == ['French Fries', 'Pork Fried Rice',
                                                                             'Plain Lo Mein']
          and not any(k[0] == 'Fried Wonton' for tb in t.values() for k in tb), hdr)
    full = t['FULL']
    check("FULL PRICE shows today's prices in grey when nothing is printed (Chicken Wings fries 7.99 + 1.60 = "
          "9.59; Fried Fish fries 10.95, Plain 8.25)",
          full[('Chicken Wings', 'Specialty')]['French Fries'][:2] == (9.59, True)
          and full[('Fried Fish', 'Specialty')]['French Fries'][:2] == (10.95, True)
          and full[('Fried Fish', 'Specialty')]['Plain'][:2] == (8.25, True),
          (full[('Chicken Wings', 'Specialty')]['French Fries'], full[('Fried Fish', 'Specialty')]['Plain']))
    ab = t['ABS'][('Fried Scallop', 'Specialty')]
    check('a side the dish has no rule for is a grey, empty cell (Fried Scallop has no Pork Fried Rice rule); '
          'Plain is its By Itself 6.95, not the No Rice $0.00 listed first',
          ab['Pork Fried Rice'][0] is None and ab['Pork Fried Rice'][2] == 'EDEDED'
          and ab['Plain (By Itself)'][0] == 6.95, ab)
    check('a rule for every size of a two-size dish has no single full price (Lemon Chicken: blank, not 10.00 + 1.50)',
          full[('Lemon Chicken', '')]['French Fries'][0] is None
          and t['DIFF'][('Lemon Chicken', '')]['French Fries'][0] == 1.5, full[('Lemon Chicken', '')])
    pfs = p_fs()
    check('the Questionnaire asks about them (P-FS): no Fried Specialty table on the paper, 7 of 7 dish sizes',
          any('7 of 7' in r and 'prints no Fried Specialty table' in r for r in pfs), pfs)
    shutil.copy('report.xlsx', 'edited.xlsx')
    p2 = ce.part2(scripts, with_rules=True)
    mr = p2['summary'].get('modifier_rule', {})
    check('an unedited report writes no modifier_rule cell', p2['apply'][0] == 0
          and mr.get('out_unit_price_cells_written') == 0, (p2['apply'][1][-400:], mr.get('out_unit_price_cells_written')))
    rc, out = ce.run(scripts, 'modifier_rule_pricing.py', 'compare', 'ocr.json', 'raw.xlsx', '--report', 'report.xlsx')
    wb2 = load_workbook('report.xlsx')
    titles = [str(c.value) for c in wb2[FS]['A'] if str(c.value or '').startswith(tuple(TITLES.values()))]
    check('re-running compare rebuilds the tab and the P-FS row, never doubles them',
          rc == 0 and len(titles) == 3 and len(p_fs()) == 2, (rc, titles, p_fs()))

    # a blank row inside a table: the rows below it are still read
    base = os.getcwd()
    for tag, drop_mr in (('blank', False), ('no_mr', True)):
        dd = os.path.join(work, 'a_' + tag)
        shutil.copytree(base, dd)
        os.chdir(dd)
        for f_ in ('final_raw.xlsx', 'part2_summary.json'):
            if os.path.exists(f_):
                os.remove(f_)
        wb = load_workbook('report.xlsx')
        _t, where = fs_tables()
        ws = wb[FS]
        rows_ = sorted(r for r, _h in where['DIFF'].values())
        hdr = where['DIFF'][('Teriyaki Beef', 'Order')][1]
        if not drop_mr:
            for c in range(1, ws.max_column + 1):                   # empty a whole row of the table
                ws.cell(rows_[0], c).value = None
        else:
            del wb['Modifier Rule']
        rr = where['DIFF'][('Teriyaki Beef', 'Order')][0]
        ws.cell(rr, hdr['French Fries']).value = 1.95
        ws.cell(rr, hdr['Notes → Wonka']).value = 'teriyaki fries per the client'
        wb.save('edited.xlsx')
        p2 = ce.part2(scripts, with_rules=True)
        f = final_rules() if os.path.exists('final_raw.xlsx') else {}
        notes = [n.get('note') for n in p2['summary'].get('specialist_requests', {}).get('notes_for_wonka', {})
                 .get('notes', [])]
        if not drop_mr:
            check('a blank row inside a table does not end it: the edit and the note below it are read',
                  f.get(('Teriyaki Beef', 'Order', 'French Fries')) == '$1.95' and 'teriyaki fries per the client' in notes,
                  (f.get(('Teriyaki Beef', 'Order', 'French Fries')), notes))
        else:
            check('a report whose Modifier Rule tab was deleted still applies the Fried Specialty tables',
                  p2['apply'][0] == 0 and f.get(('Teriyaki Beef', 'Order', 'French Fries')) == '$1.95',
                  (p2['apply'][1][-300:], f.get(('Teriyaki Beef', 'Order', 'French Fries'))))
        os.chdir(base)


def scenario_b(scripts, work):
    print('\nScenario B — the paper prints 3 of the 6, and the specialist edits the tables')
    base = os.path.join(work, 'b')
    os.makedirs(base)
    os.chdir(base)
    build(paper=True)
    out = ce.part1(scripts, compare=True)
    check('Part 1 runs', all(v[0] == 0 for v in out.values()), {k: v[1][-300:] for k, v in out.items() if v[0]})
    t, where = fs_tables()
    full, diff, ab = t['FULL'], t['DIFF'], t['ABS']
    check('FULL PRICE: printed prices in black, the rest in grey (Fried Shrimp, Teriyaki Beef, Fried Scallop)',
          full[('Chicken Wings', 'Specialty')]['French Fries'][:2] == (9.79, False)
          and full[('Fried Fish', 'Specialty')]['Plain'][:2] == (8.5, False)
          and full[('Fried Shrimp', 'Specialty')]['French Fries'][:2] == (11.1, True)
          and full[('Teriyaki Beef', 'Order')]['Pork Fried Rice'][:2] == (10.9, True),
          {k: v.get('French Fries') for k, v in full.items()})
    check('PRICE DIFFERENCE / ABSOLUTE PRICE hold what Part 2 would write: Chicken Wings fries 1.80 (blue), '
          'Chicken Fingers fries 1.60 (green), Fried Fish By Itself 8.50 and fries 11.25 (blue)',
          diff[('Chicken Wings', 'Specialty')]['French Fries'][::2] == (1.8, 'ADD8E6')
          and diff[('Chicken Fingers', 'Specialty')]['French Fries'][::2] == (1.6, 'C6EFCE')
          and ab[('Fried Fish', 'Specialty')]['Plain (By Itself)'][::2] == (8.5, 'ADD8E6')
          and ab[('Fried Fish', 'Specialty')]['French Fries'][::2] == (11.25, 'ADD8E6'),
          (diff[('Chicken Wings', 'Specialty')]['French Fries'], ab[('Fried Fish', 'Specialty')]))
    pfs = p_fs()
    check('P-FS lists only the 4 not printed', any('4 of 7' in r and 'Fried Shrimp' in r and 'Teriyaki Beef [Order]' in r
                                                   and 'Fried Scallop' in r and 'Lemon Chicken' in r for r in pfs), pfs)

    # the specialist's edits
    wb = load_workbook('report.xlsx')
    ws = wb[FS]

    def put(tb, dish, size, col, v):
        rr, hdr = where[tb][(dish, size)]
        ws.cell(rr, hdr[col]).value = v
        return ws.cell(rr, hdr[col]).coordinate
    typed = {
        'diff 1.75': put('DIFF', 'Fried Shrimp', 'Specialty', 'French Fries', 1.75),          # an edit, no + sign
        'diff +2.50 text': put('DIFF', 'Teriyaki Beef', 'Order', 'Pork Fried Rice', '+2.50'),  # a + sign is fine too
        'abs 11.5 text': put('ABS', 'Fried Scallop', 'Specialty', 'French Fries', '11.5'),
        'by itself 7.25': put('ABS', 'Fried Scallop', 'Specialty', 'Plain (By Itself)', 7.25),
        'over part 1': put('ABS', 'Fried Fish', 'Specialty', 'Plain (By Itself)', 8.75),     # Part 1 said 8.50
        'letters': put('DIFF', 'Chicken Fingers', 'Specialty', 'Pork Fried Rice', 'abc'),
        '1395': put('DIFF', 'Chicken Fingers', 'Specialty', 'French Fries', 1395),
        'misplaced': put('DIFF', 'Fried Shrimp', 'Specialty', 'Pork Fried Rice', 24.0),       # 10× the live 2.40
        'x': put('ABS', 'Fried Fish', 'Specialty', 'Pork Fried Rice', 'x'),
        'no rule': put('ABS', 'Fried Scallop', 'Specialty', 'Pork Fried Rice', 10.95),
        'plain over': put('DIFF', 'Chicken Wings', 'Specialty', 'Plain (Dish Item price)', 8.49),
        'conflict': put('DIFF', 'Chicken Wings', 'Specialty', 'Plain Lo Mein', 3.10),
    }
    put('DIFF', 'Fried Shrimp', 'Specialty', 'Notes → Wonka', 'client: shrimp + lo mein is 12.95')
    # the Modifier Rule tab: one rule edited differently there, one edited only there
    _c, mrows = ce.mrv(wb['Modifier Rule'])
    mrows = {(d or '', v or '', m or ''): r for (d, v, m), r in mrows.items()}
    wb['Modifier Rule'].cell(mrows[('Chicken Wings', 'Specialty', 'Plain Lo Mein')], 2).value = 3.20
    wb['Modifier Rule'].cell(mrows[('Fried Fish', 'Specialty', 'French Fries')], 2).value = 11.45   # FS untouched (11.25)
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, 1).value or '').startswith('NOTES → WONKA'):
            ws.cell(r + 2, 2).value = 'Every Lo Mein side = Pork Fried Rice side + 0.50'
            break
    wb.save('edited.xlsx')
    wbv = load_workbook('edited.xlsx')
    seen = {k: cf_kind(wbv, wbv[FS], c, kind) for k, c in typed.items()}
    want = {'diff 1.75': 'EDIT', 'diff +2.50 text': 'EDIT', 'abs 11.5 text': 'EDIT', 'by itself 7.25': 'EDIT',
            'over part 1': 'EDIT',
            'letters': 'INVALID', '1395': 'INVALID', 'misplaced': 'INVALID', 'x': 'INVALID', 'no rule': 'INVALID',
            'conflict': 'EDIT', 'plain over': 'INVALID'}
    check('the sheet shows each typed cell the way Part 2 treats it (orange = applied edit, red = refused)',
          seen == want, {k: (seen[k], want[k]) for k in want if seen[k] != want[k]})

    # a quick sort of the PRICE DIFFERENCE block (visible columns only, descending dish name)
    rows_ = sorted(where['DIFF'].values(), key=lambda x: x[0])
    first, last = rows_[0][0], rows_[-1][0]
    hdr = rows_[0][1]
    notes_c = hdr['Notes → Wonka']
    block = [[ws.cell(r, c).value for c in range(1, notes_c + 1)] for r in range(first, last + 1)]
    block.sort(key=lambda row: str(row[1]), reverse=True)
    wb2 = load_workbook('edited.xlsx')
    ws2 = wb2[FS]
    for i, row in enumerate(block):
        for c, v in enumerate(row, 1):
            ws2.cell(first + i, c).value = v
    wb2.save('edited.xlsx')
    # after the sort the colours still belong to each dish (the rules look its hidden values up by key)
    wbs = load_workbook('edited.xlsx')
    wss = wbs[FS]
    at = {wss.cell(r, 2).value: r for r in range(first, last + 1)}
    cols = {'diff 1.75': ('Fried Shrimp', 'French Fries'), 'misplaced': ('Fried Shrimp', 'Pork Fried Rice'),
            'letters': ('Chicken Fingers', 'Pork Fried Rice'), 'conflict': ('Chicken Wings', 'Plain Lo Mein'),
            'untouched Part 1 change': ('Chicken Wings', 'French Fries'), 'untouched': ('Teriyaki Beef', 'French Fries')}
    seen2 = {k: cf_kind(wbs, wss, wss.cell(at[d], hdr[c]).coordinate, kind) for k, (d, c) in cols.items()}
    check('after a quick sort of the table, each cell keeps its own colour (untouched cells none)',
          seen2 == {'diff 1.75': 'EDIT', 'misplaced': 'INVALID', 'letters': 'INVALID', 'conflict': 'EDIT',
                    'untouched Part 1 change': None, 'untouched': None}, seen2)
    rc, out = ce.run(scripts, 'modifier_rule_pricing.py', 'compare', 'ocr.json', 'raw.xlsx', '--report',
                     'edited.xlsx', '-o', 'rerun.xlsx')
    check('compare refuses to rebuild tabs that hold the specialist\'s edits and notes (nothing written)',
          rc == 1 and 'already holds the specialist' in out and not os.path.exists('rerun.xlsx'), out[-300:])

    p2 = ce.part2(scripts, with_rules=True)
    check('Part 2 runs', all(v[0] == 0 for k, v in p2.items() if isinstance(v, tuple) and k != 'tsv'),
          {k: v[1][-300:] for k, v in p2.items() if isinstance(v, tuple) and v[0]})
    f = final_rules()
    want = {('Fried Shrimp', 'Specialty', 'French Fries'): '$1.75',
            ('Teriyaki Beef', 'Order', 'Pork Fried Rice'): '$2.50',
            ('Fried Scallop', 'Specialty', 'French Fries'): '$11.50',
            ('Fried Scallop', 'Specialty', 'By Itself'): '$7.25',
            ('Fried Fish', 'Specialty', 'By Itself'): '$8.75',                     # typed over Part 1's 8.50
            ('Chicken Fingers', 'Specialty', 'Pork Fried Rice'): '$2.40',          # letters — kept
            ('Chicken Fingers', 'Specialty', 'French Fries'): '$1.60',             # 1395 — kept
            ('Fried Shrimp', 'Specialty', 'Pork Fried Rice'): '$2.40',             # misplaced — kept
            ('Fried Fish', 'Specialty', 'Pork Fried Rice'): '$10.95',              # x refused: the rule is held
            ('Chicken Wings', 'Specialty', 'Plain Lo Mein'): '$2.90',              # conflict — not applied
            ('Fried Fish', 'Specialty', 'French Fries'): '$11.45',                 # edited on Modifier Rule only
            ('Chicken Wings', 'Specialty', 'French Fries'): '$1.80'}               # Part 1's paper change
    got = {k: f.get(k) for k in want}
    check('Part 2 writes the table edits (no + sign needed, a + sign fine), keeps every refused cell, and '
          'a sorted table still lands on the right dishes', got == want,
          {k: (got[k], want[k]) for k in want if got[k] != want[k]})
    mr = p2['summary'].get('modifier_rule', {})
    whys = {(x.get('dish'), x.get('modifier')): x.get('why', '') for x in mr.get('ignored_input', [])
            if x.get('tab', '').startswith('Fried Specialty')}
    check('each refusal is reported with its reason (not a price, misplaced decimal, x, no rule, Plain of a '
          'difference-priced dish)',
          'not a valid price' in whys.get(('Chicken Fingers', 'Pork Fried Rice'), '')
          and 'not a valid price' in whys.get(('Chicken Fingers', 'French Fries'), '')
          and 'misplaced decimal' in whys.get(('Fried Shrimp', 'Pork Fried Rice'), '')
          and 'x is not used' in whys.get(('Fried Fish', 'Pork Fried Rice'), '')
          and 'no rule for Pork Fried Rice' in whys.get(('Fried Scallop', 'Pork Fried Rice'), '')
          and 'Dish Item tab' in whys.get(('Chicken Wings', 'Plain'), ''), whys)
    tc = mr.get('tab_conflicts', [])
    check('the rule edited differently on both tabs is reported, not applied',
          len(tc) == 1 and tc[0].get('dish') == 'Chicken Wings' and tc[0].get('modifier_rule_tab') == '$3.20'
          and tc[0].get('fried_specialty_tab') == '$3.10', tc)
    notes = [(n.get('tab'), n.get('dish'), n.get('note'))
             for n in p2['summary'].get('specialist_requests', {}).get('notes_for_wonka', {}).get('notes', [])]
    check('Notes → Wonka on the tab reach Wonka: the row note (with its dish) and the box note',
          ('Fried Specialty', 'Fried Shrimp [Specialty]', 'client: shrimp + lo mein is 12.95') in notes
          and ('Fried Specialty', '', 'Every Lo Mein side = Pork Fried Rice side + 0.50') in notes, notes)
    rep = p2['rep'][1]
    check('part2_report shows the conflict, the refusals and the notes',
          'Modifier Rule / Fried Specialty' in rep and 'has no rule for Pork Fried Rice' in rep
          and 'Every Lo Mein side' in rep, rep[-1500:])


def twin(scripts, work):
    print('\nA dish the paper does not print takes the printed dishes\' same side change (price twin) — said so')
    d = os.path.join(work, 'twin')
    os.makedirs(d)
    os.chdir(d)
    dishes = [('Chicken Wings', 'Fried Specialties', {'Specialty': 7.99}),
              ('Chicken Fingers', 'Fried Specialties', {'Specialty': 8.50}),
              ('Teriyaki Beef', 'Fried Specialties', {'Order': 8.50}),
              ('Egg Roll', 'Appetizers', {'Order': 1.75})]
    mod_groups = {S: (1, ['By Itself', 'French Fries', 'Pork Fried Rice'], ['Fried Specialties'])}
    mod_items = {m: {'Order': 0.0} for m in ('By Itself', 'French Fries', 'Pork Fried Rice')}
    rules = [('Fried Specialties', '', 'Specialty', S, '', None)]
    for dish, var in (('Chicken Wings', 'Specialty'), ('Chicken Fingers', 'Specialty'), ('Teriyaki Beef', 'Order')):
        rules += [('', dish, var, '', 'French Fries', 1.60), ('', dish, var, '', 'Pork Fried Rice', 2.40)]
    make_raw('raw.xlsx', 'R99939', 'Twin Wok', ['Order', 'Specialty'], dishes, mod_groups, mod_items, rules)
    cols = ['w. French Fries', 'w. Pork Fried Rice']
    json.dump({'restaurant': {'name': 'Twin Wok'}, 'sections': [
        {'menu_section': 'Fried Specialties', 'price_columns': ['Specialty'], 'default_price': None,
         'side_pricing': {'columns': cols}, 'items': [
             {'name': 'Chicken Wings', 'prices': {'Specialty': 7.99}, 'side_prices': {cols[0]: 9.79, cols[1]: 10.59}},
             {'name': 'Chicken Fingers', 'prices': {'Specialty': 8.50},
              'side_prices': {cols[0]: 10.30, cols[1]: 11.10}}]}]}, open('ocr.json', 'w'))
    out = ce.part1(scripts, compare=True)
    t, _w = fs_tables()
    tb = t['DIFF'][('Teriyaki Beef', 'Order')]
    pfs = p_fs()
    check('the filled sides of the unprinted dish are purple (not printed), and P-FS names them as filled — '
          'not "today\'s prices"', tb['French Fries'][::2] == (1.8, 'E4DFEC') and tb['Pork Fried Rice'][::2] == (2.6, 'E4DFEC')
          and any('filled from the printed dishes' in r and 'Teriyaki Beef [Order]' in r for r in pfs)
          and 'got side prices from the printed dishes' in out['cmp'][1], (tb, pfs))
    shutil.copy('report.xlsx', 'edited.xlsx')
    ce.part2(scripts, with_rules=True)
    f = final_rules()
    check('… and Part 2 writes them unless the specialist changes them',
          f.get(('Teriyaki Beef', 'Order', 'French Fries')) == '$1.80'
          and f.get(('Teriyaki Beef', 'Order', 'Pork Fried Rice')) == '$2.60', f)


def no_side_columns(scripts, work):
    print('\nA dish priced per dish only on By Itself (its sides priced once, on the items): no table, no false DRIFT')
    d = os.path.join(work, 'noside')
    os.makedirs(d)
    os.chdir(d)
    make_raw('raw.xlsx', 'R99940', 'Plain Wok', ['Order', 'Specialty'],
             [('Fried Fish', 'Fried Specialties', {'Specialty': 0.00}), ('Egg Roll', 'Appetizers', {'Order': 1.75})],
             {S: (1, ['By Itself', 'French Fries'], ['Fried Specialties'])},
             {'By Itself': {'Order': 0.0}, 'French Fries': {'Order': 2.00}},
             [('Fried Specialties', '', 'Specialty', S, '', None), ('', 'Fried Fish', 'Specialty', '', 'By Itself', 8.25)])
    json.dump({'restaurant': {'name': 'Plain Wok'}, 'sections': [
        {'menu_section': 'Appetizers', 'price_columns': [], 'default_price': None,
         'items': [{'name': 'Egg Roll', 'prices': {'Order': 1.95}}]}]}, open('ocr.json', 'w'))
    out = ce.part1(scripts, compare=True)
    wb = load_workbook('report.xlsx')
    check('compare builds the Modifier Rule tab, no Fried Specialty tab, and no DRIFT warning',
          out['cmp'][0] == 0 and 'DRIFT' not in out['cmp'][1] and 'Modifier Rule' in wb.sheetnames
          and FS not in wb.sheetnames, (out['cmp'][1][-300:], wb.sheetnames))


def shared_prices(scripts, work):
    print('\nA restaurant that prices Select Rice once for every dish (Modifier Summary) — unchanged')
    d = os.path.join(work, 'shared')
    os.makedirs(d)
    os.chdir(d)
    dishes = [('Chicken Wings', 'Fried Specialties', {'Small': 7.75, 'Large': 12.95}),
              ('Chicken Finger', 'Fried Specialties', {'Order': 13.00})]
    mod_groups = {S: (1, ['White Rice', 'Pork Fried Rice', 'French Fries'], ['Fried Specialties'])}
    mod_items = {'White Rice': {'Order': 0.0}, 'Pork Fried Rice': {'Order': 1.00}, 'French Fries': {'Order': 2.00}}
    rules = [('Fried Specialties', '', 'Order', S, '', None), ('Fried Specialties', '', '', 'Sauce', 'Sauce', 1.00)]
    make_raw('raw.xlsx', 'R99938', 'Shared Wok', ['Order', 'Small', 'Large'], dishes, mod_groups, mod_items, rules)
    json.dump({'restaurant': {'name': 'Shared Wok'}, 'sections': [
        {'menu_section': 'Fried Specialties', 'price_columns': [], 'default_price': None,
         'items': [{'name': 'Chicken Wings', 'prices': {'Small': 7.95, 'Large': 12.95}}]}]}, open('ocr.json', 'w'))
    out = ce.part1(scripts, compare=True)
    wb = load_workbook('report.xlsx')
    check('no per-dish side prices: compare has nothing to do — no Modifier Rule / Fried Specialty tab, no P-FS',
          out['cmp'][0] == 0 and 'Nothing to do' in out['cmp'][1] and 'Modifier Rule' not in wb.sheetnames
          and FS not in wb.sheetnames and not p_fs(), (out['cmp'][1][-300:], wb.sheetnames))


def main():
    bundle = os.path.abspath(sys.argv[1])
    scripts = os.path.join(bundle, 'menu-change-payload', 'scripts')
    work = tempfile.mkdtemp(prefix='fs_checks_')
    for fn in (scenario_a, scenario_b, twin, no_side_columns, shared_prices):
        try:
            fn(scripts, work)
        except Exception as e:                         # a crash fails the check, never hides the rest
            import traceback
            traceback.print_exc()
            check(f'{fn.__name__} ran', False, repr(e))
    os.chdir(HERE)
    shutil.rmtree(work, ignore_errors=True)
    print(f'\nfried_specialty: {sum(RESULTS)}/{len(RESULTS)} PASS')
    sys.exit(0 if all(RESULTS) else 1)


if __name__ == '__main__':
    main()
