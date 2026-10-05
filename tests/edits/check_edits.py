"""Pass/fail checks for the edit-sheet batch of 2026-09-30 (maintainer's requests a–f).

    python3 check_edits.py <bundle_dir>

R99906 "Edit Wok" — DELTA, dish codes, dishes under two dish groups, a modifier
under two modifier groups, a same-name unpriced modifier twin:
  (a) a price typed into an EMPTY Dish Item / Modifier Summary cell is written
      (and reaches the TSV); paste + x = a move; letters / comma / negative /
      1000+ / 0-in-an-empty-cell are ignored and reported; the report shows the
      entry rules (x → not a price → new price → edit), each formula < 255 chars.
  (c) a dish under two groups: the edited copy wins over the untouched one;
      copies edited to different prices → conflict, nothing applied; untouched
      copies carrying different new prices → the change wins. Modifier Summary:
      edits in two rows of one item are merged per size; a conflict is held.
  codes: a size whose price was removed / moved keeps its dish code → listed.
R99907 "Abs Wok" — ABSOLUTE, group-level rules, the Fried Specialties rule:
  (f) a $0.00 dish priced by GROUP-level rules, and a $0.00 Fried Specialties dish
      with no rules in a restaurant that prices on modifier rules, never take a
      price (absolute_pricing, P-ABS, action5 guard); a new size on an ABSOLUTE
      dish and removing / zeroing an option of a required group that prices a
      $0.00 dish are held.
  (e) Modifier Rule 'by itself' is a live link — ABSOLUTE: the By Itself row's
      Price; DELTA (R99901): the dish's Dish Item cell, so an edit there moves
      Expected.
  (d) action2 warns when side-rule dishes are read without a multi-column table;
      compare builds a missing Fried Specialty tab and says DRIFT.
Exit 1 if any check fails."""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'modifier_rule'))
sys.path.insert(0, os.path.join(HERE, '..', 'lib'))
from build_fixture import make_raw  # noqa: E402
from xlformula import evaluate  # noqa: E402

RESULTS = []


def check(desc, ok, detail=''):
    RESULTS.append(bool(ok))
    print(f"  {'PASS' if ok else 'FAIL'}  {desc}" + ('' if ok or not detail else f'  -- {detail}'))


def num(v):
    try:
        return round(float(str(v).replace('$', '').strip()), 2)
    except (TypeError, ValueError):
        return None


def run(scripts, script, *args):
    p = subprocess.run([sys.executable, os.path.join(scripts, script), *args],
                       capture_output=True, text=True)
    return p.returncode, p.stdout + p.stderr


def part1(scripts, ocr='ocr.json', compare=False):
    out = {}
    out['a1'] = run(scripts, 'action1_raw_to_2a.py', 'raw.xlsx', '-o', 'preview.xlsx')
    out['a2'] = run(scripts, 'action2_ocr_to_2b_0625v2.py', ocr, 'preview.xlsx', '-o', 'ocr_out.xlsx')
    out['a3'] = run(scripts, 'action3_compare_and_merge.py', 'preview.xlsx', 'ocr_out.xlsx', '-o', 'report.xlsx')
    out['pi'] = run(scripts, 'price_inference.py', 'preview.xlsx', 'ocr_out.xlsx', '--report', 'report.xlsx',
                    '--mode', 'apply')
    if compare:
        out['pre'] = run(scripts, 'modifier_rule_pricing.py', 'preflight', 'raw.xlsx')
        out['cmp'] = run(scripts, 'modifier_rule_pricing.py', 'compare', ocr, 'raw.xlsx', '--report', 'report.xlsx')
    return out


def raw_to_tsv(raw, tsv_dir):
    os.makedirs(tsv_dir, exist_ok=True)
    wb = load_workbook(raw)
    for ws in wb.worksheets:
        lines = ['\t'.join('' if v is None else str(v) for v in row) for row in ws.iter_rows(values_only=True)]
        with open(os.path.join(tsv_dir, f'{ws.title}.tsv'), 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')


def tsv_cells(path):
    return [r.split('\t') for r in open(path, encoding='utf-8').read().split('\n') if r != '']


def part2(scripts, with_rules=False):
    raw_to_tsv('raw.xlsx', 'tsv')
    run(scripts, 'wsheet_xlsx_bridge.py', 'to-xlsx', 'tsv', '-o', 'raw_b.xlsx')
    tabs = ('dish_item', 'modifier_item') + (('modifier_rule',) if with_rules else ())
    before = {t: tsv_cells(os.path.join('tsv', f'{t}.tsv')) for t in tabs}
    out = {'req': run(scripts, 'specialist_requests.py', 'edited.xlsx', '--dir', '.'),
           'a4': run(scripts, 'action4_apply_modifier_changes.py', 'edited.xlsx', '-o', 'final_review.xlsx'),
           'a5': run(scripts, 'action5_to_raw_gsheet.py', 'final_review.xlsx', 'raw_b.xlsx', '-o', 'final_raw.xlsx')}
    if with_rules:
        out['pre'] = run(scripts, 'modifier_rule_pricing.py', 'preflight', 'final_raw.xlsx')
        out['apply'] = run(scripts, 'modifier_rule_pricing.py', 'apply', 'edited.xlsx', 'final_raw.xlsx',
                           '-o', 'final_raw.xlsx')
    out['tsv'] = run(scripts, 'wsheet_xlsx_bridge.py', 'to-tsv', 'final_raw.xlsx', 'tsv', '--tabs', ','.join(tabs))
    out['rep'] = run(scripts, 'part2_report.py', '.')
    after = {t: tsv_cells(os.path.join('tsv', f'{t}.tsv')) for t in tabs}
    diff = []
    for t in tabs:
        for ri, (o, n) in enumerate(zip(before[t], after[t])):
            for ci in range(max(len(o), len(n))):
                ov, nv = (o[ci] if ci < len(o) else ''), (n[ci] if ci < len(n) else '')
                if ov != nv:
                    diff.append((t, None, ri + 1, ci + 1, ov, nv))
    out['tsv_diff'] = diff
    out['summary'] = json.load(open('part2_summary.json')) if os.path.isfile('part2_summary.json') else {}
    return out


def live(path, tab):
    """{name: [ {variant: value} per root row ]} from a raw / final raw tab."""
    ws = load_workbook(path)[tab]
    h2 = [str(c.value or '') for c in ws[2]]
    h3 = [c.value for c in ws[3]]
    nc = h3.index(tab) + 1
    pcols = {h3[i]: i + 1 for i, m in enumerate(h2) if m.endswith('unit_price')}
    out = {}
    for r in range(4, ws.max_row + 1):
        if str(ws.cell(r, 1).value or '').startswith('{'):
            out.setdefault(ws.cell(r, nc).value, []).append(
                {v: ws.cell(r, c).value for v, c in pcols.items()})
    return out


def report_rows(ws):
    hdr = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(hdr) if h}
    ref = {str(h)[5:]: i + 1 for i, h in enumerate(hdr) if h and str(h).startswith('_ref_')}
    rows = {(ws.cell(r, 2).value, ws.cell(r, 1).value): r for r in range(2, ws.max_row + 1) if ws.cell(r, 2).value}
    return col, ref, rows


def rules_by_range(ws):
    out = {}
    for cf in ws.conditional_formatting:
        for rule in cf.rules:
            out.setdefault(str(cf.sqref), []).append(rule)
    return out


def rule_kind(rule):
    d = rule.dxf
    fill = (d.fill.fgColor.rgb or '') + (d.fill.bgColor.rgb or '') if d is not None and d.fill is not None else ''
    font = (d.font.color.rgb or '') if d is not None and d.font is not None and d.font.color is not None else ''
    if '595959' in fill:
        return 'X'
    if 'C00000' in font:
        return 'INVALID'
    if 'FFFF00' in fill:
        return 'NEW'
    if 'FBCEB1' in fill:
        return 'EDIT'
    return '?'


def entry_rules_ok(ws):
    """Every price column that has the apricot edit rule has, in priority order,
    x → 10 × not-a-price (6 on a Dish Item tray column: no new-price 10× / tenth
    rules) → new → edit; every formula under 255 characters."""
    from openpyxl.utils import column_index_from_string
    bad = []
    n = 0
    for sqref, rules in rules_by_range(ws).items():
        rules = sorted(rules, key=lambda r: r.priority)
        kinds = [rule_kind(r) for r in rules]
        if 'EDIT' not in kinds:
            continue
        n += 1
        head = ws.cell(1, column_index_from_string(re.match(r'[A-Z]+', sqref).group(0))).value
        k = 6 if (ws.title == 'Dish Item' and head == 'Party Tray') else 10
        if kinds != ['X'] + ['INVALID'] * k + ['NEW', 'EDIT']:
            bad.append((sqref, kinds))
        long_ = [len(f) for r in rules for f in r.formula if len(f) >= 255]
        if long_:
            bad.append((sqref, 'formula length', long_))
    return n > 0 and not bad, bad or n


def mrv(ws):
    hdr_row = next(r for r in range(1, 12)
                   if any(ws.cell(r, c).value == 'out_unit_price' for c in range(1, 12)))
    cols = {ws.cell(hdr_row, c).value: c for c in range(1, 16) if ws.cell(hdr_row, c).value}
    rows = {}
    for r in range(hdr_row + 1, ws.max_row + 1):
        if ws.cell(r, cols['modifier_item']).value:
            rows[(ws.cell(r, cols['dish_item']).value or '', ws.cell(r, cols['variant']).value,
                  ws.cell(r, cols['modifier_item']).value)] = r
    return cols, rows


# ══ R99906 Edit Wok ═════════════════════════════════════════════════════════

def build_r99906():
    variants = ['Order', 'Small', 'Large', 'Lunch']
    dishes = [
        ('Vegetable Lo Mein', 'Lo Mein', {'Small': 8.95, 'Large': 12.95, 'Lunch': 9.95}),
        ('Kung Pao Chicken', 'Chicken', {'Small': 9.25, 'Large': 13.25}),
        ('Sesame Chicken', 'Chicken', {'Small': 9.25, 'Large': 13.25}),
        ('Beef with Broccoli', 'Beef', {'Small': 9.50, 'Large': 13.95}),
        ('Egg Roll', 'Appetizers', {'Order': 1.75}),
        ('Wonton Soup', 'Soup', {'Small': 3.25, 'Large': 5.25}),
        ('Hot and Sour Soup', 'Soup', {'Small': 3.50, 'Large': 5.50}),
        ('Fried Rice', 'Rice', {'Small': 7.00}),
        ('Egg Drop Soup', 'Soup', {'Small': 3.00, 'Large': 5.00}),     # not on the paper
    ]
    also_in = [('Vegetarian', 'Vegetable Lo Mein'), ('Chef Specials', 'Kung Pao Chicken'),
               ('Chef Specials', 'Beef with Broccoli')]
    mod_groups = {'Extras': (0, ['Extra Meat', 'Extra Shrimp', 'Extra Sauce', 'Chili Oil', 'Hot Mustard'],
                             ['Lo Mein', 'Chicken']),
                  'Add Ons': (0, ['Extra Meat', 'Extra Shrimp'], ['Beef']),
                  'Select Side': (0, ['French Fries', 'White Rice'], ['Chicken', 'Beef'])}
    mod_items = {'Extra Meat': {'Small': 2.00, 'Large': 3.00}, 'Extra Shrimp': {'Small': 3.00, 'Large': 4.00},
                 'Extra Sauce': {'Small': 0.50}, 'French Fries': {'Order': 2.50}, 'White Rice': {'Order': 1.00},
                 'Chili Oil': {'Small': 0.25}, 'Hot Mustard': {'Small': 0.0, 'Large': 0.0}}
    codes = {'Egg Roll': {'Order': 'A1'}, 'Wonton Soup': {'Small': 'S1', 'Large': 'S2'}}
    make_raw('raw.xlsx', 'R99906', 'Edit Wok', variants, dishes, mod_groups, mod_items, [],
             codes=codes, also_in=also_in, mod_twins=['Extra Sauce'])
    paper = [('Lo Mein', 'Vegetable Lo Mein', {'Small': 9.25, 'Large': 13.25}),
             ('Chicken', 'Kung Pao Chicken', {'Small': 9.45, 'Large': 13.45}),
             ('Chicken', 'Sesame Chicken', {'Small': 9.50, 'Large': 13.50}),
             ('Beef', 'Beef with Broccoli', {'Small': 9.75, 'Large': 14.25}),
             ('Appetizers', 'Egg Roll', {'Order': 1.95}),
             ('Soup', 'Wonton Soup', {'Small': 3.50, 'Large': 5.50}),
             ('Soup', 'Hot and Sour Soup', {'Small': 3.75, 'Large': 5.75}),
             ('Rice', 'Fried Rice', {'Small': 7.25})]
    secs = {}
    for sec, dish, pr in paper:
        secs.setdefault(sec, []).append({'name': dish, 'prices': pr})
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Edit Wok'}, 'sections': [
            {'menu_section': s, 'price_columns': [], 'default_price': None, 'items': it}
            for s, it in secs.items()]}, f)


def edit_r99906():
    wb = load_workbook('edited.xlsx')
    di = wb['Dish Item']
    col, ref, rows = report_rows(di)

    def setv(dish, group, v, val, also_ref=False):
        r = rows[(dish, group)]
        di.cell(r, col[v]).value = val
        if also_ref:
            di.cell(r, ref[v]).value = val
    setv('Kung Pao Chicken', 'Chicken', 'Small', 9.95)              # upper copy edited
    setv('Beef with Broccoli', 'Beef', 'Large', 14.50)               # two copies edited …
    setv('Beef with Broccoli', 'Chef Specials', 'Large', 14.75)      # … differently
    setv('Vegetable Lo Mein', 'Vegetarian', 'Lunch', 10.45, also_ref=True)   # as if Part 1 inferred it …
    di.cell(rows[('Vegetable Lo Mein', 'Vegetarian')], col['Lunch']).fill = PatternFill('solid', fgColor='E4DFEC')
    setv('Vegetable Lo Mein', 'Vegetarian', 'Large', 'x')            # x on the lower copy
    setv('Egg Roll', 'Appetizers', 'Small', 1.95)                    # move: paste into the empty cell …
    setv('Egg Roll', 'Appetizers', 'Order', 'x')                     # … and x the old one
    setv('Wonton Soup', 'Soup', 'Large', 'X')
    setv('Fried Rice', 'Rice', 'Large', 11.5)                        # new size
    setv('Hot and Sour Soup', 'Soup', 'Small', -3)
    setv('Hot and Sour Soup', 'Soup', 'Large', 5500)
    setv('Hot and Sour Soup', 'Soup', 'Lunch', 0)                    # 0 into an empty cell
    setv('Sesame Chicken', 'Chicken', 'Small', '9.5o')
    setv('Sesame Chicken', 'Chicken', 'Large', '13,50')
    ms = wb['Modifier Summary']
    mcol, _mref, _rows = report_rows(ms)
    mrows = {(ms.cell(r, 1).value, ms.cell(r, 2).value): r for r in range(2, ms.max_row + 1) if ms.cell(r, 2).value}

    def setm(group, item, v, val):
        ms.cell(mrows[(group, item)], mcol[v]).value = val
    setm('Extras', 'Extra Meat', 'Small', 2.25)                      # merged per size …
    setm('Add Ons', 'Extra Meat', 'Large', 3.5)                      # … across two rows
    setm('Extras', 'Extra Shrimp', 'Small', 3.10)                    # conflict
    setm('Add Ons', 'Extra Shrimp', 'Small', 3.25)
    setm('Extras', 'Extra Sauce', 'Large', 0.75)                     # new, empty cell
    setm('Extras', 'Extra Sauce', 'Small', 'abc')
    setm('Select Side', 'White Rice', 'Order', 'x')
    setm('Select Side', 'French Fries', 'Small', 2.5)                # move Order → Small
    setm('Select Side', 'French Fries', 'Order', 'x')
    wb.save('edited.xlsx')


def set_live(path, tab, name, variant, value):
    """Someone edits wsheet between Part 1 and Part 2."""
    wb = load_workbook(path)
    ws = wb[tab]
    h2 = [str(c.value or '') for c in ws[2]]
    h3 = [c.value for c in ws[3]]
    nc = h3.index(tab) + 1
    ci = next(i + 1 for i, m in enumerate(h2) if m.endswith('unit_price') and h3[i] == variant)
    r = next(r for r in range(4, ws.max_row + 1)
             if ws.cell(r, nc).value == name and str(ws.cell(r, 1).value or '').startswith('{'))
    ws.cell(r, ci).value = value
    wb.save(path)


def rule_cell(path, group, dish, variant, mod, value=None):
    """Read (or, with value, set) one live modifier_rule out_unit_price."""
    wb = load_workbook(path)
    ws = wb['modifier_rule']
    for r in range(4, ws.max_row + 1):
        if [str(ws.cell(r, c).value or '') for c in (9, 10, 11, 13)] == [group, dish, variant, mod]:
            if value is None:
                return ws.cell(r, 3).value
            ws.cell(r, 3).value = value
            wb.save(path)
            return value
    return 'missing'


def check_r99906(scripts):
    build_r99906()
    p1 = part1(scripts)
    check('R99906 Part 1 runs', all(p1[k][0] == 0 for k in p1), {k: v[1][-200:] for k, v in p1.items() if v[0]})
    wb = load_workbook('report.xlsx')
    ok, detail = entry_rules_ok(wb['Dish Item'])
    check('(a) Dish Item: x → not-a-price ×5 → NEW (yellow) → edit rules on every price column, '
          'each formula < 255 chars', ok, detail)
    ok, detail = entry_rules_ok(wb['Modifier Summary'])
    check('(a) Modifier Summary: the same rules (x now works there too)', ok, detail)
    col, ref, rows = report_rows(wb['Dish Item'])
    check('(c) fixture: the dishes listed under two groups are two report rows',
          all((d, g) in rows for d, g in (('Vegetable Lo Mein', 'Lo Mein'), ('Vegetable Lo Mein', 'Vegetarian'),
                                          ('Kung Pao Chicken', 'Chef Specials'))), sorted(rows)[:12])

    shutil.copy('report.xlsx', 'edited.xlsx')
    edit_r99906()
    set_live('raw.xlsx', 'dish_item', 'Egg Drop Soup', 'Small', '$3.25')     # unprinted, untouched
    set_live('raw.xlsx', 'modifier_item', 'Hot Mustard', 'Large', '$0.75')   # the snapshot says $0.00
    set_live('raw.xlsx', 'modifier_item', 'Chili Oil', 'Small', '$0.35')     # untouched in the summary
    p2 = part2(scripts)
    check('R99906 Part 2 runs (requests, action4, action5, to-tsv, report)',
          all(p2[k][0] == 0 for k in ('req', 'a4', 'a5', 'tsv', 'rep')),
          {k: p2[k][1][-300:] for k in ('a4', 'a5', 'tsv', 'rep') if p2[k][0]})
    d = {k: v[0] for k, v in live('final_raw.xlsx', 'dish_item').items()}
    mods = live('final_raw.xlsx', 'modifier_item')
    a5 = p2['summary'].get('action5', {})

    print('\n  (c) a dish listed under two dish groups')
    check('edited upper copy wins over the untouched lower copy (Kung Pao Chicken Small $9.95)',
          d['Kung Pao Chicken']['Small'] == '$9.95', d['Kung Pao Chicken'])
    check('… and the untouched sizes still take the paper price (Kung Pao Chicken Large $13.45)',
          d['Kung Pao Chicken']['Large'] == '$13.45', d['Kung Pao Chicken'])
    check('two copies edited to different prices → conflict, nothing applied (Beef with Broccoli Large $13.95)',
          d['Beef with Broccoli']['Large'] == '$13.95'
          and any(c['name'] == 'Beef with Broccoli' and c['variant'] == 'Large'
                  for c in a5.get('duplicates', {}).get('conflicts', [])), d['Beef with Broccoli'])
    check('untouched copies with different new prices → the change wins (Vegetable Lo Mein Lunch $10.45)',
          d['Vegetable Lo Mein']['Lunch'] == '$10.45', d['Vegetable Lo Mein'])
    check('x on the lower copy wins over the upper copy\'s untouched paper price (Vegetable Lo Mein Large removed)',
          d['Vegetable Lo Mein']['Large'] is None and d['Vegetable Lo Mein']['Small'] == '$9.25',
          d['Vegetable Lo Mein'])

    print('\n  untouched prices never overwrite a live change made since Part 1')
    check('an unprinted, untouched dish price keeps its new live value (Egg Drop Soup Small $3.25, not $3.00)',
          d['Egg Drop Soup']['Small'] == '$3.25' and d['Egg Drop Soup']['Large'] == '$5.00', d['Egg Drop Soup'])
    check("a modifier nobody edited is not written from the report's snapshot "
          "(Hot Mustard Large stays $0.75 — the snapshot's $0.00 is not written; Chili Oil Small $0.35)",
          mods['Hot Mustard'][0]['Large'] == '$0.75' and mods['Chili Oil'][0]['Small'] == '$0.35',
          (mods['Hot Mustard'][0], mods['Chili Oil'][0]))
    check('… counted as "live changed since Part 1, not reverted"',
          a5.get('not_reverted', {}).get('total') == 3, a5.get('not_reverted'))

    print('\n  (a) new prices, moves, ignored input')
    check('a price typed into an empty cell is written (Fried Rice Large $11.50)',
          d['Fried Rice']['Large'] == '$11.50' and d['Fried Rice']['Small'] == '$7.25', d['Fried Rice'])
    check('paste + x = a move (Egg Roll Order removed, Small $1.95)',
          d['Egg Roll']['Order'] is None and d['Egg Roll']['Small'] == '$1.95', d['Egg Roll'])
    check('x removes a price (Wonton Soup Large)', d['Wonton Soup']['Large'] is None and
          d['Wonton Soup']['Small'] == '$3.50', d['Wonton Soup'])
    check('letters / comma / negative / 1000+ / 0 in an empty cell → nothing written (live price kept)',
          d['Sesame Chicken'] == {'Order': None, 'Small': '$9.25', 'Large': '$13.25', 'Lunch': None}
          and d['Hot and Sour Soup'] == {'Order': None, 'Small': '$3.50', 'Large': '$5.50', 'Lunch': None},
          (d['Sesame Chicken'], d['Hot and Sour Soup']))
    ign = a5.get('ignored_input', {}).get('cells', [])
    got = sorted((c.get('name'), c.get('variant'), str(c.get('value'))) for c in ign)
    want = sorted([('Hot and Sour Soup', 'Small', '-3'), ('Hot and Sour Soup', 'Large', '5500'),
                   ('Hot and Sour Soup', 'Lunch', '0'), ('Sesame Chicken', 'Small', '9.5o'),
                   ('Sesame Chicken', 'Large', '13,50'), ('Extra Sauce', 'Small', 'abc')])
    check('every ignored input is reported (6: 5 Dish Item + 1 Modifier Summary)', got == want, got)

    print('\n  Modifier Summary')
    em = mods['Extra Meat'][0]
    check('edits in two rows of one modifier are merged per size (Extra Meat Small $2.25 AND Large $3.50)',
          em['Small'] == '$2.25' and em['Large'] == '$3.50', em)
    check('two rows edited to different prices → conflict, held (Extra Shrimp Small $3.00)',
          mods['Extra Shrimp'][0]['Small'] == '$3.00'
          and any(c['name'] == 'Extra Shrimp' for c in a5.get('duplicates', {}).get('conflicts', [])),
          mods['Extra Shrimp'][0])
    es = mods['Extra Sauce']
    check('a new price in an empty cell reaches the item the report shows, never its unpriced twin '
          '(Extra Sauce Large $0.75; twin stays empty)',
          es[0]['Large'] == '$0.75' and es[0]['Small'] == '$0.50' and not any(v for v in es[1].values()), es)
    ff = mods['French Fries'][0]
    check('Modifier Summary move (French Fries Order → Small) and x (White Rice Order removed)',
          ff['Order'] is None and ff['Small'] == '$2.50' and mods['White Rice'][0]['Order'] is None,
          (ff, mods['White Rice'][0]))

    print('\n  what reaches wsheet (the TSV)')
    diff = p2['tsv_diff']
    filled = sorted((t, nv) for t, _n, _r, _c, ov, nv in diff if ov == '' and nv != '')
    blanked = sorted((t, ov) for t, _n, _r, _c, ov, nv in diff if nv == '')
    check('the TSV gets exactly the 4 new prices (2 dish_item, 2 modifier_item)',
          filled == sorted([('dish_item', '$1.95'), ('dish_item', '$11.50'),
                            ('modifier_item', '$0.75'), ('modifier_item', '$2.50')]), filled)
    check('… and exactly the 5 removals', blanked == sorted([
        ('dish_item', '$12.95'), ('dish_item', '$1.75'), ('dish_item', '$5.25'),
        ('modifier_item', '$1.00'), ('modifier_item', '$2.50')]), blanked)
    check('every price the TSV received is $X.XX',
          all(re.match(r'^\$\d+\.\d\d$', nv) for *_x, nv in diff if nv), [x for x in diff if x[5]][:6])

    print('\n  dish codes')
    cf = sorted((c['dish'], c['action'], c['code'], c['from'], c['to'])
                for c in a5.get('code_followups', {}).get('items', []))
    check('a moved / removed size keeps its code — listed for Wonka, codes never written '
          '(Egg Roll A1 Order→Small; Wonton Soup S2 removed)',
          cf == [('Egg Roll', 'move', 'A1', 'Order', 'Small'), ('Wonton Soup', 'remove', 'S2', 'Large', '')],
          cf)
    wbf = load_workbook('final_raw.xlsx')['dish_item']
    h2 = [str(c.value or '') for c in wbf[2]]
    h3 = [c.value for c in wbf[3]]
    codes = {}
    for r in range(4, wbf.max_row + 1):
        for i, m in enumerate(h2):
            if m.endswith('unit_code') and wbf.cell(r, i + 1).value:
                codes[(wbf.cell(r, 2).value, h3[i])] = wbf.cell(r, i + 1).value
    check('… and the raw\'s codes are untouched', codes == {('Egg Roll', 'Order'): 'A1',
                                                           ('Wonton Soup', 'Small'): 'S1',
                                                           ('Wonton Soup', 'Large'): 'S2'}, codes)
    check('x on a dish\'s last price is applied but flagged (Egg Roll kept its moved price → not flagged; '
          'Wonton Soup keeps Small → not flagged)', a5.get('no_price_left') == [], a5.get('no_price_left'))
    rep = p2['rep'][1]
    check('part2_report lists new prices, removals, conflicts, ignored input and dish codes',
          all(s in rep for s in ('new prices in empty cells', 'prices removed', 'CONFLICT',
                                 'Not a valid price / code — ignored', 'Dish codes left on a moved / removed price')), rep[-1200:])


# ══ R99907 Abs Wok ══════════════════════════════════════════════════════════

def build_r99907(side=True):
    variants = ['Order', 'Specialty', 'Small', 'Large']
    dishes = [
        ('Chicken Wings', 'Fried Specialties', {'Specialty': 0.00}),        # dish-level ABSOLUTE rules
        ('Fried Shrimp', 'Fried Specialties', {'Specialty': 0.00}),         # group-level rules only
        ('Fried Crab Stick', 'Fried Specialties::RPOS', {'Specialty': 0.00}),  # no rules at all
        ('Seafood Boil', 'Seafood Boil', {'Order': 0.00}),                  # required priced group
        ('Egg Roll', 'Appetizers', {'Order': 1.75}),
        ('Egg Roll Platter', 'Appetizers', {'Order': 6.00}),
    ]
    mod_groups = {'Select Side': (0, ['By Itself', 'French Fries'], ['Fried Specialties']),
                  'Select Seafood': (1, ['1 lb Shrimp', '1 lb Crawfish'], ['Seafood Boil'])}
    mod_items = {'By Itself': {'Specialty': 0.0}, 'French Fries': {'Specialty': 0.0},
                 '1 lb Shrimp': {'Order': 19.00}, '1 lb Crawfish': {'Order': 27.00}}
    S = 'Select Side'
    rules = [('', 'Chicken Wings', 'Specialty', S, 'By Itself', 7.50),
             ('', 'Chicken Wings', 'Specialty', S, 'French Fries', 8.90),
             ('Fried Specialties', '', 'Specialty', S, 'By Itself', 8.00),
             ('Fried Specialties', '', 'Specialty', S, 'French Fries', 9.40)]
    make_raw('raw.xlsx', 'R99907', 'Abs Wok', variants, dishes, mod_groups, mod_items, rules)
    fs = {'menu_section': 'Fried Specialties', 'price_columns': ['Specialty'], 'default_price': None,
          'items': [{'code': 'F1', 'name': 'Chicken Wings', 'prices': {'Specialty': 7.99}},
                    {'code': 'F2', 'name': 'Fried Shrimp', 'prices': {'Specialty': 8.25}},
                    {'code': 'F3', 'name': 'Fried Crab Stick', 'prices': {'Specialty': 6.50}}]}
    if side:
        fs['side_pricing'] = {'columns': ['w. French Fries']}
        fs['items'][0]['side_prices'] = {'w. French Fries': 9.39}
        fs['items'][1]['side_prices'] = {'w. French Fries': 9.65}
    ocr = {'restaurant': {'name': 'Abs Wok'}, 'sections': [
        fs,
        {'menu_section': 'Seafood', 'price_columns': ['Order'], 'default_price': None,
         'items': [{'name': 'Seafood Boil', 'prices': {'Order': 25.00}}]},
        {'menu_section': 'Appetizers', 'price_columns': ['Order'], 'default_price': None,
         'items': [{'name': 'Egg Roll', 'prices': {'Order': 1.95}}]}]}
    with open('ocr.json' if side else 'ocr_noside.json', 'w') as f:
        json.dump(ocr, f)


def check_r99907(scripts):
    build_r99907(side=True)
    build_r99907(side=False)
    p1 = part1(scripts, compare=True)
    check('R99907 Part 1 runs (incl. preflight + compare)', all(p1[k][0] == 0 for k in p1),
          {k: v[1][-300:] for k, v in p1.items() if v[0]})
    abs_rows = {(r[1], r[2]): (r[3], str(r[5] or '')) for r in
                load_workbook('preview.xlsx')['absolute_pricing'].iter_rows(min_row=2, values_only=True) if r[1]}
    print('\n  (f) $0.00 dishes that must never take a dish price')
    check('group-level rules put Fried Shrimp [Specialty] on absolute_pricing',
          abs_rows.get(('Fried Shrimp', 'Specialty'), ('', ''))[0] == 'modifier_rule'
          and 'group-level' in abs_rows[('Fried Shrimp', 'Specialty')][1], abs_rows)
    check('the Fried Specialties rule puts Fried Crab Stick [Specialty] (no rules) on it too',
          abs_rows.get(('Fried Crab Stick', 'Specialty'), ('',))[0] == 'fried_specialty', abs_rows)
    wb = load_workbook('report.xlsx')
    col, _ref, rows = report_rows(wb['Dish Item'])
    zero = {d: wb['Dish Item'].cell(rows[(d, g)], col[v]).value
            for d, g, v in (('Fried Shrimp', 'Fried Specialties', 'Specialty'),
                            ('Fried Crab Stick', 'Fried Specialties::RPOS', 'Specialty'),
                            ('Chicken Wings', 'Fried Specialties', 'Specialty'))}
    check('Part 1: their paper prices are NOT written to Dish Item (all stay 0)',
          all(num(v) == 0.0 for v in zero.values()), zero)
    q = [str(r[3]) for r in wb['Questionnaire'].iter_rows(min_row=2, values_only=True)
         if str(r[0] or '').startswith('P-ABS-')]
    check('… and each is asked in P-ABS', all(x in q for x in ('Fried Shrimp [Specialty]',
                                                                'Fried Crab Stick [Specialty]')), q)

    print('\n  (e) Modifier Rule \'by itself\' — ABSOLUTE')
    ws = wb['Modifier Rule']
    mc, mr = mrv(ws)
    ff, bi = mr.get(('Chicken Wings', 'Specialty', 'French Fries')), mr.get(('Chicken Wings', 'Specialty', 'By Itself'))
    h = ws.cell(ff, 8).value if ff else None
    shown = num(evaluate(wb, ws, h)) if ff else None
    check("ABSOLUTE row: 'by itself' is a live link to the By Itself row's Price (evaluates to it)",
          ff and bi and f'B{bi}' in str(h) and shown == num(ws.cell(bi, 2).value), (h, shown, bi))
    keep = ws.cell(bi, 2).value
    ws.cell(bi, 2).value = 8.25                              # edit the By Itself price …
    moved = num(evaluate(wb, ws, h))
    ws.cell(bi, 2).value = keep
    check('… and follows an edit of that Price (8.25)', moved == 8.25, moved)

    shutil.copy('report.xlsx', 'edited.xlsx')
    e = load_workbook('edited.xlsx')
    di = e['Dish Item']
    col, _ref, rows = report_rows(di)
    di.cell(rows[('Fried Crab Stick', 'Fried Specialties::RPOS')], col['Specialty']).value = 6.50
    di.cell(rows[('Fried Shrimp', 'Fried Specialties')], col['Specialty']).value = 8.25
    di.cell(rows[('Chicken Wings', 'Fried Specialties')], col['Small']).value = 9.99   # new size
    di.cell(rows[('Egg Roll', 'Appetizers')], col['Order']).value = 2.05                # control
    ms = e['Modifier Summary']
    mcol, _r, _x = report_rows(ms)
    mrow = {ms.cell(r, 2).value: r for r in range(2, ms.max_row + 1) if ms.cell(r, 2).value}
    ms.cell(mrow['1 lb Shrimp'], mcol['Order']).value = 'x'
    ms.cell(mrow['1 lb Crawfish'], mcol['Order']).value = 0
    e.save('edited.xlsx')
    rule_cell('raw.xlsx', 'Fried Specialties', '', 'Specialty', 'French Fries', '$9.60')  # wsheet edit
    p2 = part2(scripts, with_rules=True)
    check('R99907 Part 2 runs (action5, preflight on the result, apply, to-tsv)',
          all(p2[k][0] == 0 for k in ('a4', 'a5', 'pre', 'apply', 'tsv', 'rep')),
          {k: p2[k][1][-300:] for k in ('a5', 'pre', 'apply', 'tsv') if p2[k][0]})
    d = {k: v[0] for k, v in live('final_raw.xlsx', 'dish_item').items()}
    mods = {k: v[0] for k, v in live('final_raw.xlsx', 'modifier_item').items()}
    check('action5 keeps Fried Crab Stick and Fried Shrimp at $0.00 (ABSOLUTE GUARD)',
          d['Fried Crab Stick']['Specialty'] == '$0.00' and d['Fried Shrimp']['Specialty'] == '$0.00',
          (d['Fried Crab Stick'], d['Fried Shrimp']))
    check('a new size on an ABSOLUTE dish is held (Chicken Wings Small stays empty)',
          d['Chicken Wings']['Small'] is None, d['Chicken Wings'])
    check('removing / zeroing an option of the required group that prices Seafood Boil is held',
          mods['1 lb Shrimp']['Order'] == '$19.00' and mods['1 lb Crawfish']['Order'] == '$27.00',
          (mods['1 lb Shrimp'], mods['1 lb Crawfish']))
    check('the ordinary edit goes through (Egg Roll $2.05)', d['Egg Roll']['Order'] == '$2.05', d['Egg Roll'])
    check('apply writes the rule compare changed (Chicken Wings + French Fries $9.39) but keeps an '
          'untouched rule someone changed in wsheet since Part 1 (group-level French Fries $9.60)',
          rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'French Fries') == '$9.39'
          and rule_cell('final_raw.xlsx', 'Fried Specialties', '', 'Specialty', 'French Fries') == '$9.60',
          (rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'French Fries'),
           rule_cell('final_raw.xlsx', 'Fried Specialties', '', 'Specialty', 'French Fries'),
           p2['summary'].get('modifier_rule', {}).get('not_reverted')))
    g = p2['summary'].get('action5', {}).get('absolute_guard', {}).get('cells', [])
    check('all 5 held edits are reported (ABSOLUTE GUARD)', len(g) == 5, [(x['dish'], x['report_price']) for x in g])

    print('\n  (d) Fried Specialty tab triggers')
    rc, out = run(scripts, 'action2_ocr_to_2b_0625v2.py', 'ocr_noside.json', 'preview.xlsx', '-o', 'ocr_noside.xlsx')
    check('action2 warns: side-rule dishes read from the paper, but no multi-column table in the OCR JSON',
          rc == 0 and 'NO MULTI-COLUMN TABLE' in out and 'Chicken Wings' in out, out[-600:])
    run(scripts, 'action3_compare_and_merge.py', 'preview.xlsx', 'ocr_noside.xlsx', '-o', 'report_drift.xlsx')
    had = 'Fried Specialty' in load_workbook('report_drift.xlsx').sheetnames
    rc, out = run(scripts, 'modifier_rule_pricing.py', 'compare', 'ocr.json', 'raw.xlsx', '--report',
                  'report_drift.xlsx')
    wbd = load_workbook('report_drift.xlsx')
    check('compare on a report without a Fried Specialty tab builds it and says DRIFT',
          not had and rc == 0 and 'Fried Specialty' in wbd.sheetnames and '!! DRIFT' in out
          and 'different OCR JSON' in str(wbd['Fried Specialty']['A1'].value), (had, rc, out[-300:]))


# ══ R99901 Delta Wok — the DELTA link ═══════════════════════════════════════

def check_delta_link(scripts):
    subprocess.run([sys.executable, os.path.join(HERE, '..', 'modifier_rule', 'build_fixture.py')],
                   check=True, capture_output=True)
    shutil.copy('R99901_raw.xlsx', 'raw.xlsx')
    shutil.copy('R99901_menu_ocr.json', 'ocr.json')
    p1 = part1(scripts, compare=True)
    check('R99901 Part 1 runs', all(p1[k][0] == 0 for k in p1), {k: v[1][-300:] for k, v in p1.items() if v[0]})
    wb = load_workbook('report.xlsx')
    ws = wb['Modifier Rule']
    mc, mr = mrv(ws)
    r = mr.get(('Chicken Wings', 'Specialty', 'French Fries'))
    h = str(ws.cell(r, 8).value) if r else ''
    m = re.search(r"'Dish Item'!\$([A-Z]+)\$(\d+)", h)
    di = wb['Dish Item']
    target = (di.cell(int(m.group(2)), 2).value, di[f'{m.group(1)}1'].value) if m else None
    check("DELTA row: 'by itself' links to the dish's own Dish Item cell (Chicken Wings · Specialty)",
          target == ('Chicken Wings', 'Specialty'), (h, target))
    cell = f'{m.group(1)}{m.group(2)}' if m else 'A1'
    b = num(ws.cell(r, 2).value) if r else None
    di[cell].value = 8.49                                  # the specialist corrects the dish price
    h1, j1 = (num(evaluate(wb, ws, ws.cell(r, c).value)) for c in (8, 10))
    check('… so an edit on the Dish Item tab moves it (8.49) and Expected = 8.49 + the upcharge',
          h1 == 8.49 and j1 == round(8.49 + b, 2), (h1, j1, b))
    for typed in ('x', None):
        di[cell].value = typed
        hv, jv, kv = (evaluate(wb, ws, ws.cell(r, c).value) for c in (8, 10, 11))
        check(f"… with {typed!r} in the Dish Item cell: 'by itself' and Expected blank, Match? N/A",
              hv == '' and jv == '' and kv == 'N/A', (hv, jv, kv))


def build_r99908():
    """DELTA, a multi-column dish listed under two dish groups."""
    variants = ['Order', 'Specialty', 'Small', 'Large']
    dishes = [('Chicken Wings', 'Fried Specialties', {'Specialty': 7.99}),
              ('Egg Roll', 'Appetizers', {'Order': 1.75})]
    mod_groups = {'Select Side': (0, ['French Fries', 'Pork Fried Rice'], ['Fried Specialties'])}
    mod_items = {'French Fries': {'Order': 0.0}, 'Pork Fried Rice': {'Order': 0.0}}
    rules = [('', 'Chicken Wings', 'Specialty', 'Select Side', 'French Fries', 1.40),
             ('', 'Chicken Wings', 'Specialty', 'Select Side', 'Pork Fried Rice', 1.90)]
    make_raw('raw.xlsx', 'R99908', 'Two Group Wok', variants, dishes, mod_groups, mod_items, rules,
             also_in=[('Chef Specials', 'Chicken Wings')])
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Two Group Wok'}, 'sections': [
            {'menu_section': 'Specialties', 'price_columns': ['Specialty'], 'default_price': None,
             'side_pricing': {'columns': ['w. French Fries', 'w. Pork Fried Rice']},
             'items': [{'code': 'S1', 'name': 'Chicken Wings', 'prices': {'Specialty': 7.99},
                        'side_prices': {'w. French Fries': 9.59, 'w. Pork Fried Rice': 10.19}}]}]}, f)


def check_two_group_link(scripts):
    build_r99908()
    p1 = part1(scripts, compare=True)
    check('R99908 Part 1 runs', all(p1[k][0] == 0 for k in p1), {k: v[1][-300:] for k, v in p1.items() if v[0]})
    wb = load_workbook('report.xlsx')
    ws, di = wb['Modifier Rule'], wb['Dish Item']
    _c, mr = mrv(ws)
    r = mr.get(('Chicken Wings', 'Specialty', 'French Fries'))
    col, _ref, rows = report_rows(di)
    lower = rows[('Chicken Wings', 'Chef Specials')]
    di.cell(lower, col['Specialty']).value = 8.25          # only the OTHER copy is edited
    h = num(evaluate(wb, ws, ws.cell(r, 8).value))
    check("DELTA dish in two groups: 'by itself' follows the edited copy (8.25), like Part 2",
          h == 8.25, (ws.cell(r, 8).value, h))
    upper = rows[('Chicken Wings', 'Fried Specialties')]
    di.cell(upper, col['Specialty']).value = 8.10           # both copies edited, differently
    h2 = evaluate(wb, ws, ws.cell(r, 8).value)
    k2 = evaluate(wb, ws, ws.cell(r, 11).value)
    check('… and shows one of the edited copies (Part 2 holds such a conflict)', num(h2) in (8.10, 8.25),
          (h2, k2))


# ══ Robustness: recoloured / sorted sheets, formulas, the price parser ═══════

def set_cached(path, sheet, coord, value):
    """Give a formula cell a saved result, as Excel / Google Sheets would."""
    import zipfile
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    wbx = items['xl/workbook.xml'].decode()
    tag = re.search(r'<sheet\b[^>]*\bname="%s"[^>]*/>' % re.escape(sheet), wbx).group(0)
    rid = re.search(r'r:id="([^"]+)"', tag).group(1)
    rels = items['xl/_rels/workbook.xml.rels'].decode()
    rel = re.search(r'<Relationship\b[^>]*\bId="%s"[^>]*/>' % re.escape(rid), rels).group(0)
    target = re.search(r'Target="([^"]+)"', rel).group(1).lstrip('/')
    target = target if target.startswith('xl/') else 'xl/' + target
    x = items[target].decode()
    x2 = re.sub(r'(<c r="%s"[^>]*>\s*<f>[^<]*</f>)\s*<v\s*/>' % coord, r'\g<1><v>%s</v>' % value, x)
    assert x2 != x, coord
    items[target] = x2.encode()
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, d in items.items():
            z.writestr(n, d)


def recolour_and_sort(ws, key_col=2):
    """What a specialist may do: recolour every visible cell, then a quick sort of
    the visible block by dish name — Excel leaves the hidden columns behind."""
    hdr = [c.value for c in ws[1]]
    vis = [i + 1 for i, h in enumerate(hdr) if h not in (None, '') and not str(h).startswith('_ref_')]
    rows = []
    for r in range(2, ws.max_row + 1):
        rows.append([ws.cell(r, c).value for c in vis])
    rows.sort(key=lambda x: str(x[vis.index(key_col)] or ''), reverse=True)
    for r, vals in enumerate(rows, 2):
        for c, v in zip(vis, vals):
            cell = ws.cell(r, c)
            cell.value = v
            cell.fill = PatternFill('solid', fgColor='FFF2CC')


def check_robustness(scripts):
    build_r99906()
    p1 = part1(scripts)
    check('R99906 Part 1 runs', all(p1[k][0] == 0 for k in p1))
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    di, ms = wb['Dish Item'], wb['Modifier Summary']
    recolour_and_sort(di)
    recolour_and_sort(ms)
    col, _ref, rows = report_rows(di)
    di.cell(rows[('Kung Pao Chicken', 'Chicken')], col['Small']).value = 9.95
    di.cell(rows[('Fried Rice', 'Rice')], col['Large']).value = 11.5
    di.cell(rows[('Wonton Soup', 'Soup')], col['Large']).value = 'x'
    di.cell(rows[('Egg Drop Soup', 'Soup')], col['Small']).value = 'x'      # both prices x'd …
    di.cell(rows[('Egg Drop Soup', 'Soup')], col['Large']).value = 'X'      # … → no price left
    mcol, _r, _x = report_rows(ms)
    mrows = {(ms.cell(r, 1).value, ms.cell(r, 2).value): r for r in range(2, ms.max_row + 1) if ms.cell(r, 2).value}
    ms.cell(mrows[('Extras', 'Extra Meat')], mcol['Small']).value = 2.25
    chili = ms.cell(mrows[('Extras', 'Chili Oil')], mcol['Small'])
    chili.value = '=0.25+0.15'                        # a formula, saved with its result
    rice = ms.cell(mrows[('Select Side', 'White Rice')], mcol['Order'])
    rice.value = '=1+1'                               # a formula with no saved result
    wb.save('edited.xlsx')
    set_cached('edited.xlsx', 'Modifier Summary', chili.coordinate, '0.4')
    p2 = part2(scripts)
    check('R99906 Part 2 runs on a recoloured + sorted report', all(p2[k][0] == 0 for k in ('a4', 'a5', 'tsv')),
          {k: p2[k][1][-300:] for k in ('a4', 'a5') if p2[k][0]})
    d = {k: v[0] for k, v in live('final_raw.xlsx', 'dish_item').items()}
    mods = {k: v[0] for k, v in live('final_raw.xlsx', 'modifier_item').items()}
    a5 = p2['summary'].get('action5', {})
    check('recoloured + sorted Dish Item: Part 1 paper prices still land (Kung Pao Large $13.45, Wonton Soup '
          'Small $3.50, Fried Rice Small $7.25, Sesame Chicken $9.50 / $13.50)',
          d['Kung Pao Chicken']['Large'] == '$13.45' and d['Wonton Soup']['Small'] == '$3.50'
          and d['Fried Rice']['Small'] == '$7.25' and d['Sesame Chicken']['Small'] == '$9.50'
          and d['Sesame Chicken']['Large'] == '$13.50',
          (d['Kung Pao Chicken'], d['Wonton Soup'], d['Fried Rice'], d['Sesame Chicken']))
    check('… and the edits made after sorting are read against the right baselines '
          '(Kung Pao Small $9.95, Fried Rice Large new $11.50, Wonton Soup Large removed)',
          d['Kung Pao Chicken']['Small'] == '$9.95' and d['Fried Rice']['Large'] == '$11.50'
          and d['Wonton Soup']['Large'] is None, (d['Kung Pao Chicken'], d['Fried Rice'], d['Wonton Soup']))
    check('… nothing untouched is written (Hot and Sour Soup keeps its paper $3.75 / $5.75 only; no '
          'conflicts, no drift)',
          d['Hot and Sour Soup']['Small'] == '$3.75' and d['Hot and Sour Soup']['Lunch'] is None
          and not a5.get('duplicates', {}).get('conflicts') and not a5.get('new_cell_drift'),
          (d['Hot and Sour Soup'], a5.get('duplicates'), a5.get('new_cell_drift')))
    check('x on every price of a dish removes them but flags it: strike the name to take a dish off the menu '
          '(Egg Drop Soup)',
          d['Egg Drop Soup']['Small'] is None and d['Egg Drop Soup']['Large'] is None
          and [c['dish'] for c in a5.get('no_price_left', [])] == ['Egg Drop Soup']
          and 'NO price left' in p2['rep'][1], (d['Egg Drop Soup'], a5.get('no_price_left')))
    check('sorted Modifier Summary: the edit lands on the right item (Extra Meat Small $2.25, Large $3.00)',
          mods['Extra Meat']['Small'] == '$2.25' and mods['Extra Meat']['Large'] == '$3.00', mods['Extra Meat'])
    ign = a5.get('ignored_input', {}).get('cells', [])
    check('a Modifier Summary formula counts by its saved result (Chili Oil Small $0.40); one with no saved '
          'result is reported, not used (White Rice Order stays $1.00)',
          mods['Chili Oil']['Small'] == '$0.40' and mods['White Rice']['Order'] == '$1.00'
          and any(c.get('name') == 'White Rice' and 'formula' in c.get('why', '') for c in ign),
          (mods['Chili Oil'], mods['White Rice'], ign))


def check_rule_recolour(scripts):
    build_r99907(side=True)
    p1 = part1(scripts, compare=True)
    check('R99907 Part 1 runs', all(p1[k][0] == 0 for k in p1))
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    ws = wb['Modifier Rule']
    for r in range(1, ws.max_row + 1):
        ws.cell(r, 2).fill = PatternFill(fill_type=None)          # the specialist clears the colours
    # … and quick-sorts the visible A–F block by modifier (Excel stops at the blank column G)
    _cols, mr = mrv(ws)
    data = sorted([[ws.cell(r, c).value for c in range(1, 7)] for r in sorted(mr.values())],
                  key=lambda x: str(x[5] or ''), reverse=True)
    for r, vals in zip(sorted(mr.values()), data):
        for c, v in enumerate(vals, 1):
            ws.cell(r, c).value = v
    wb.save('edited.xlsx')
    rule_cell('raw.xlsx', 'Fried Specialties', '', 'Specialty', 'French Fries', '$9.60')  # wsheet edit
    p2 = part2(scripts, with_rules=True)
    check("recoloured + sorted Modifier Rule tab: compare's changes still applied (Chicken Wings By Itself "
          "$7.99, French Fries $9.39)",
          rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'By Itself') == '$7.99'
          and rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'French Fries') == '$9.39',
          (rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'By Itself'),
           rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'French Fries'), p2['apply'][1][-300:]))
    check('… and an untouched rule changed in wsheet since Part 1 is kept, not reverted ($9.60)',
          rule_cell('final_raw.xlsx', 'Fried Specialties', '', 'Specialty', 'By Itself') == '$8.00'
          and rule_cell('final_raw.xlsx', 'Fried Specialties', '', 'Specialty', 'French Fries') == '$9.60',
          (rule_cell('final_raw.xlsx', 'Fried Specialties', '', 'Specialty', 'French Fries'),
           p2['summary'].get('modifier_rule', {}).get('not_reverted')))


def check_parser(scripts):
    import importlib.util
    mods = []
    for f in ('action5_to_raw_gsheet.py', 'action4_apply_modifier_changes.py'):
        spec = importlib.util.spec_from_file_location(f[:-3], os.path.join(scripts, f))
        m = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(m)
        mods.append(m)
    cases = {'9.45': ('price', 9.45), ' $9.45 ': ('price', 9.45), '9.45$': ('price', 9.45),
             '$ 9.45': ('price', 9.45), '+1.50': ('price', 1.5), '.5': ('price', 0.5), '5.': ('price', 5.0),
             '-0': ('price', 0.0), 12.3: ('price', 12.3), 0: ('price', 0.0), 'x': ('x', None),
             ' X ': ('x', None), '': ('blank', None), None: ('blank', None)}
    bad = ['9,45', '1,000', '9.5o', '1e2', '٩', '10%', '12:30', '(9.45)', 1000, 1395, 999.995, True, 'abc',
           'x ', '2.25\n', '\t2.25', '9 1/2']
    got_ok = {k: [m._entry(k) for m in mods] for k in cases}
    wrong = {k: v for k, v in got_ok.items() if any(x != cases[k] for x in v)}
    wrong_bad = {k: [m._entry(k) for m in mods] for k in bad if any(m._entry(k)[0] != 'invalid' for m in mods)}
    check('the price parser (action4 = action5) accepts what the sheet shows as a price and nothing else '
          '(no-break space, tab, newline, fraction, 999.995 → invalid)',
          not wrong and not wrong_bad, (wrong, wrong_bad))
    a5, a4 = mods
    check('a negative is never a dish price but is a modifier price ("Less Whole Wing" −$1.70)',
          a5._entry('-3')[0] == 'invalid' and a5._entry(-1)[0] == 'invalid'
          and a4._entry('-1.70') == ('price', -1.7) and a4._entry(-1.7) == ('price', -1.7),
          (a5._entry('-3'), a4._entry('-1.70')))
    check('a misplaced decimal point is caught both ways (195 for 1.95, 0.95 for 9.50) — not 2× / half',
          all(m._misplaced(195, 1.95) and m._misplaced(0.95, 9.5) and not m._misplaced(3.9, 1.95)
              and not m._misplaced(4.75, 9.5) and not m._misplaced(5, 0) for m in mods))


def main():
    bundle = os.path.abspath(sys.argv[1])
    scripts = os.path.join(bundle, 'menu-change-payload', 'scripts')
    work = tempfile.mkdtemp(prefix='edits_checks_')
    for name, fn in (('R99906 Edit Wok', check_r99906), ('R99907 Abs Wok', check_r99907),
                     ('R99901 Delta Wok (DELTA link)', check_delta_link),
                     ('R99906b recoloured + sorted report', check_robustness),
                     ('R99907b recoloured Modifier Rule tab', check_rule_recolour),
                     ('price parser', check_parser),
                     ('R99908 DELTA dish in two groups', check_two_group_link)):
        d = os.path.join(work, name.split()[0])
        os.makedirs(d, exist_ok=True)
        os.chdir(d)
        print(f'\n{name}')
        try:
            fn(scripts)
        except Exception as e:          # a crash is a failure, never a silent pass
            check(f'{name}: checks ran to the end', False, f'{type(e).__name__}: {e}')
    os.chdir(HERE)
    shutil.rmtree(work, ignore_errors=True)
    print(f'\nedits: {sum(RESULTS)}/{len(RESULTS)} PASS')
    sys.exit(0 if all(RESULTS) else 1)


if __name__ == '__main__':
    main()
