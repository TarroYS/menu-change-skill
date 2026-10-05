"""Regression checks for the self-audit of 2026-09-30 (three independent reviews +
LibreOffice evaluation of the sheet rules). Each check pins one confirmed bug.

    python3 check_audit.py <bundle_dir>

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
sys.path.insert(0, os.path.join(HERE, '..', 'edits'))
sys.path.insert(0, os.path.join(HERE, '..', 'lib'))
from build_fixture import make_raw  # noqa: E402
import check_edits as ce  # noqa: E402  (fixtures + pipeline helpers)
from xlformula import evaluate, cf_kind  # noqa: E402

RESULTS = []


def check(desc, ok, detail=''):
    RESULTS.append(bool(ok))
    print(f"  {'PASS' if ok else 'FAIL'}  {desc}" + ('' if ok or not detail else f'  -- {detail}'))


num = ce.num


def run(scripts, script, *args):
    return ce.run(scripts, script, *args)


def fresh(work, name):
    d = os.path.join(work, name)
    os.makedirs(d, exist_ok=True)
    os.chdir(d)
    return d


def delta_wok():
    subprocess.run([sys.executable, os.path.join(HERE, '..', 'modifier_rule', 'build_fixture.py')],
                   check=True, capture_output=True)
    shutil.copy('R99901_raw.xlsx', 'raw.xlsx')
    shutil.copy('R99901_menu_ocr.json', 'ocr.json')


# ── Part 2: Modifier Rule input (apply) ──────────────────────────────────────

def mrv_input(scripts, work):
    print('\nModifier Rule tab — what the specialist types (apply)')
    fresh(work, 'mrv_input')
    delta_wok()
    ce.part1(scripts, compare=True)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    ws = wb['Modifier Rule']
    _c, mr = ce.mrv(ws)
    typed = {('Chicken Wings', 'Specialty', 'French Fries'): '1,50',          # decimal comma
             ('Chicken Fingers', 'Specialty', 'French Fries'): 1395,         # missing decimal point
             ('Fried Shrimp', 'Specialty', 'French Fries'): '1.5%',
             ('Fried Shrimp', 'Specialty', 'Pork Fried Rice'): 'x',
             ('Buffalo Wings', 'Order', 'French Fries'): 16.0,               # 10× its 1.60
             ('Fried Scallops', 'Specialty', 'French Fries'): 0.15}          # typed as 15%
    for k, v in typed.items():
        ws.cell(mr[k], 2).value = v
    ws.cell(mr[('Fried Scallops', 'Specialty', 'French Fries')], 2).number_format = '0%'
    formula_cell = ws.cell(mr[('Chicken Wings', 'Specialty', 'Chicken Fried Rice')], 2)
    formula_cell.value = '=1.5+0.5'
    wb.save('edited.xlsx')
    red = lambda rule: 'INVALID' if (rule.dxf and rule.dxf.font and rule.dxf.font.color is not None
                                     and 'C00000' in (rule.dxf.font.color.rgb or '')) else 'ORANGE'
    shown = {k: cf_kind(wb, ws, ws.cell(mr[k], 2).coordinate, red) for k in typed}
    eq = [(t.title, str(cf.sqref)) for t in wb.worksheets for cf in t.conditional_formatting
          for rule in cf.rules if any(str(f).startswith('=') for f in (rule.formula or []))]
    check('no conditional-format formula is stored with a leading "=" (Excel drops such a rule)', not eq, eq)
    check('the Modifier Rule tab shows each of those red (not a price) before Part 2 runs — except the '
          'percentage, which the sheet shows as "15%"',
          all(v == 'INVALID' for k, v in shown.items() if k[0] != 'Fried Scallops'), shown)
    ce.set_cached('edited.xlsx', 'Modifier Rule', formula_cell.coordinate, '2')
    p2 = ce.part2(scripts, with_rules=True)
    mr_s = p2['summary'].get('modifier_rule', {})
    got = {(t['dish'], t['modifier']): t['value'] for t in mr_s.get('ignored_input', [])}
    live = {k: ce.rule_cell('final_raw.xlsx', '', *k) for k in typed}
    check('invalid prices never reach the raw: 1,50 / 1395 / 1.5% / x / 10× / a percentage keep their '
          'live rule price',
          all(v not in ('$150.00', '$1395.00', '$16.00', '$0.15', '$1.50') for v in live.values())
          and live[('Fried Shrimp', 'Specialty', 'Pork Fried Rice')] == '$1.90', live)
    check('… and each is reported as "not a price" (apply JSON + part2_report)',
          len(got) == 6 and 'Not a valid price' in p2['rep'][1], (got, p2['rep'][1][-600:]))
    check('a Modifier Rule formula counts by its saved result (Chicken Wings + Chicken Fried Rice $2.00)',
          ce.rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'Chicken Fried Rice') == '$2.00',
          ce.rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'Chicken Fried Rice'))


def mrv_absolute_zero(scripts, work):
    print('\nABSOLUTE: the By Itself base never goes to $0.00')
    fresh(work, 'mrv_abs0')
    ce.build_r99907(side=True)
    ce.part1(scripts, compare=True)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    ws = wb['Modifier Rule']
    _c, mr = ce.mrv(ws)
    ws.cell(mr[('Chicken Wings', 'Specialty', 'By Itself')], 2).value = 0
    wb.save('edited.xlsx')
    p2 = ce.part2(scripts, with_rules=True)
    check('By Itself of a $0.00 dish set to 0 → held (the plain dish would be free)',
          ce.rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'By Itself') == '$7.50'
          and p2['summary'].get('modifier_rule', {}).get('held'),
          (ce.rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'By Itself'),
           p2['summary'].get('modifier_rule', {}).get('held')))


# ── Part 2: action4 / action5 refusals and warnings ─────────────────────────

def refusals(scripts, work):
    print('\naction4 / action5 refusals and warnings')
    fresh(work, 'refuse')
    ce.build_r99906()
    ce.part1(scripts)
    ce.raw_to_tsv('raw.xlsx', 'tsv')
    run(scripts, 'wsheet_xlsx_bridge.py', 'to-xlsx', 'tsv', '-o', 'raw_b.xlsx')
    rc, out = run(scripts, 'action5_to_raw_gsheet.py', 'report.xlsx', 'raw_b.xlsx', '-o', 'f.xlsx')
    check('action5 refuses a report that did not go through action4 (no Part 2 Intent sheet)',
          rc == 1 and 'not action4' in out and not os.path.exists('f.xlsx'), out[-300:])
    run(scripts, 'action4_apply_modifier_changes.py', 'report.xlsx', '-o', 'fr.xlsx')
    rc, out = run(scripts, 'action4_apply_modifier_changes.py', 'fr.xlsx', '-o', 'fr2.xlsx')
    check('action4 refuses its own output (a second pass would drop the edits)',
          rc == 1 and 'already a Final Review' in out, out[-300:])
    rc, out = run(scripts, 'part2_report.py', '.')
    check('part2_report exits 2 when action5 has not written its result (never renders zeros)', rc == 2, out[-200:])
    run(scripts, 'action5_to_raw_gsheet.py', 'fr.xlsx', 'raw_b.xlsx', '-o', 'f.xlsx')
    rc, out = run(scripts, 'action5_to_raw_gsheet.py', 'fr.xlsx', 'raw_b.xlsx', '-o', 'f.xlsx')
    check('a second action5 run in the same directory warns (patched base → counts are relative)',
          'already ran in this directory' in out, out[-400:])
    # an older report layout: no hidden baselines
    wb = load_workbook('fr.xlsx')
    di = wb['Dish Item']
    for c in range(di.max_column, 0, -1):
        if str(di.cell(1, c).value or '').startswith('_ref_'):
            di.delete_cols(c)
    wb.save('fr_old.xlsx')
    rc, out = run(scripts, 'action5_to_raw_gsheet.py', 'fr_old.xlsx', 'raw_b.xlsx', '-o', 'f_old.xlsx')
    check('a report without hidden baselines runs with a loud warning', rc == 0 and 'no hidden baselines' in out,
          out[-300:])
    # a report re-saved by another app: the hidden Modifier tab and every _ref_ column gone (5 real ones)
    wb = load_workbook('report.xlsx')
    del wb['Modifier']
    ms = wb['Modifier Summary']
    for c in range(ms.max_column, 0, -1):
        if str(ms.cell(1, c).value or '').startswith('_ref_'):
            ms.delete_cols(c)
    wb.save('resaved.xlsx')
    rc, out = run(scripts, 'action4_apply_modifier_changes.py', 'resaved.xlsx', '-o', 'fr_resaved.xlsx')
    check('a report that lost its hidden data stops (kept as a stop, maintainer 2026-10-05) and says to re-make it '
          'with Part 1', rc == 1 and "lost the report's hidden data" in out and 'Re-make the report with Part 1' in out
          and not os.path.exists('fr_resaved.xlsx'), out[-300:])


# ── Part 2: Dish Item input ──────────────────────────────────────────────────

def dish_input(scripts, work):
    print('\nDish Item — what the specialist types (action5)')
    fresh(work, 'dish_input')
    ce.build_r99906()
    ce.part1(scripts)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    di = wb['Dish Item']
    col, _ref, rows = ce.report_rows(di)
    di.cell(rows[('Egg Roll', 'Appetizers')], col['Order']).value = 195          # for 1.95
    di.cell(rows[('Wonton Soup', 'Soup')], col['Small']).value = 0.35            # a tenth
    c = di.cell(rows[('Hot and Sour Soup', 'Soup')], col['Large'])
    c.value, c.number_format = 0.1, '0%'                                          # typed 10%
    di.cell(rows[('Fried Rice', 'Rice')], col['Large']).value = 0                 # 0 into empty
    di.cell(rows[('Fried Rice', 'Rice')], col['Lunch']).value = 'x'               # x into empty
    di.cell(rows[('Egg Drop Soup', 'Soup')], col['Lunch']).value = 6.5            # new size …
    di.cell(rows[('Sesame Chicken', 'Chicken')], col['Lunch']).value = 150.0      # 10× its prices
    di.cell(rows[('Kung Pao Chicken', 'Chicken')], col['Combo']).value = 12.0     # no such live size
    di.cell(rows[('Hot and Sour Soup', 'Soup')], col['Small']).value = 3.675      # shown as 3.68
    wb.save('edited.xlsx')
    # what the specialist SEES (the report's own rules, evaluated) before Part 2 runs
    want = {('Egg Roll', 'Appetizers', 'Order'): 'INVALID', ('Wonton Soup', 'Soup', 'Small'): 'INVALID',
            ('Fried Rice', 'Rice', 'Large'): 'INVALID', ('Fried Rice', 'Rice', 'Lunch'): 'X',   # empty cells
            ('Egg Drop Soup', 'Soup', 'Lunch'): 'NEW', ('Kung Pao Chicken', 'Chicken', 'Combo'): 'NEW',
            ('Sesame Chicken', 'Chicken', 'Lunch'): 'INVALID'}                 # a new size at 10× its sizes
    seen = {k: cf_kind(wb, di, di.cell(rows[(k[0], k[1])], col[k[2]]).coordinate, ce.rule_kind) for k in want}
    check('the sheet shows each typed cell the way Part 2 treats it (red / grey / yellow)', seen == want,
          {k: v for k, v in seen.items() if v != want[k]})
    ce.set_live('raw.xlsx', 'dish_item', 'Fried Rice', 'Large', '$9.00')          # wsheet added …
    ce.set_live('raw.xlsx', 'dish_item', 'Fried Rice', 'Lunch', '$8.00')          # … these since Part 1
    ce.set_live('raw.xlsx', 'dish_item', 'Egg Drop Soup', 'Lunch', '$4.00')
    p2 = ce.part2(scripts)
    d = {k: v[0] for k, v in ce.live('final_raw.xlsx', 'dish_item').items()}
    a5 = p2['summary'].get('action5', {})
    ign = {(c['name'], c['variant']): c['why'] for c in a5.get('ignored_input', {}).get('cells', [])}
    check('a misplaced decimal point never ships (195 for 1.95; 0.35 for 3.50) — held, the live price '
          'kept, and reported',
          d['Egg Roll']['Order'] == '$1.75' and d['Wonton Soup']['Small'] == '$3.25'
          and 'misplaced' in ign.get(('Egg Roll', 'Order'), '') and 'misplaced' in ign.get(('Wonton Soup', 'Small'), ''),
          (d['Egg Roll'], d['Wonton Soup'], ign))
    check('a cell typed as a percentage (Excel stores 10% as 0.1) is not a price (live kept)',
          d['Hot and Sour Soup']['Large'] == '$5.50' and ('Hot and Sour Soup', 'Large') in ign,
          (d['Hot and Sour Soup'], ign.get(('Hot and Sour Soup', 'Large'))))
    check('0 / x typed into a cell that was EMPTY in the report never touch a price wsheet added since',
          d['Fried Rice']['Large'] == '$9.00' and d['Fried Rice']['Lunch'] == '$8.00', d['Fried Rice'])
    check('a new price typed where wsheet has since added one is reported, not written over it',
          d['Egg Drop Soup']['Lunch'] == '$4.00'
          and any(c['dish'] == 'Egg Drop Soup' for c in a5.get('new_cell_drift', [])), a5.get('new_cell_drift'))
    check('… and the report says why: live now has $4.00 (not "the live cell has no price")',
          'Egg Drop Soup [Lunch] report $6.50 — typed into an empty cell, but live now has $4.00' in p2['rep'][1],
          [l for l in p2['rep'][1].splitlines() if 'Egg Drop Soup' in l])
    check('a new size at 10× the dish\'s other prices is held as a misplaced decimal point',
          d['Sesame Chicken']['Lunch'] is None and 'misplaced' in ign.get(('Sesame Chicken', 'Lunch'), ''),
          (d['Sesame Chicken'], ign.get(('Sesame Chicken', 'Lunch'))))
    check('3.675 (the sheet shows 3.68) is written $3.68 — rounded like the spreadsheet; Python\'s round() '
          'gives 3.67', d['Hot and Sour Soup']['Small'] == '$3.68', d['Hot and Sour Soup'])
    check('an edit in a size the live menu does not have is reported (Combo)',
          'no \'Combo\' size' in ign.get(('Kung Pao Chicken', 'Combo'), ''), ign.get(('Kung Pao Chicken', 'Combo')))


def twins(scripts, work):
    print('\nSame-name live rows (availability rows, twins) are never written')
    fresh(work, 'twins')
    ce.build_r99906()
    ce.part1(scripts)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    di = wb['Dish Item']
    col, _ref, rows = ce.report_rows(di)
    di.cell(rows[('Fried Rice', 'Rice')], col['Large']).value = 11.5
    ms = wb['Modifier Summary']
    mcol, _r, _x = ce.report_rows(ms)
    mrow = {(ms.cell(r, 1).value, ms.cell(r, 2).value): r for r in range(2, ms.max_row + 1)}
    ms.cell(mrow[('Extras', 'Extra Meat')], mcol['Small']).value = 2.25
    wb.save('edited.xlsx')
    # wsheet since Part 1: Fried Rice got an availability row (repeated root id); a priced twin
    raw = load_workbook('raw.xlsx')
    ws = raw['dish_item']
    r = next(r for r in range(4, ws.max_row + 1) if ws.cell(r, 2).value == 'Fried Rice')
    ws.insert_rows(r + 1)
    ws.cell(r + 1, 1).value, ws.cell(r + 1, 2).value = ws.cell(r, 1).value, 'Fried Rice'
    mi = raw['modifier_item']
    mt = next(r for r in range(4, mi.max_row + 1) if mi.cell(r, 2).value == 'Extra Sauce'
              and str(mi.cell(r, 1).value).startswith('{"id":"mt'))
    mi.cell(mt, 2).value = 'Extra Meat'
    mi.cell(mt, 4).value = '$1.50'
    raw.save('raw.xlsx')
    p2 = ce.part2(scripts)
    fw = load_workbook('final_raw.xlsx')
    rows_fr = [[c.value for c in row] for row in fw['dish_item'].iter_rows(min_row=4) if row[1].value == 'Fried Rice']
    twin = [[c.value for c in row] for row in fw['modifier_item'].iter_rows(min_row=4)
            if row[1].value == 'Extra Meat']
    check('a second root row of a dish (availability row) is not written; the first gets the edit',
          rows_fr[0][4] == '$11.50' and all(v in (None, '') for v in rows_fr[1][2:]), rows_fr)
    check('a priced same-name modifier twin keeps its own price ($1.50); the item the report shows gets 2.25',
          twin[0][3] == '$2.25' and twin[1][3] == '$1.50', twin)


# ── bridge ───────────────────────────────────────────────────────────────────

def bridge(scripts, work):
    print('\nTSV bridge edge cases')
    fresh(work, 'bridge')
    ce.build_r99906()
    ce.part1(scripts)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    di = wb['Dish Item']
    col, _ref, rows = ce.report_rows(di)
    di.cell(rows[('Wonton Soup', 'Soup')], col['Large']).value = 'x'
    wb.save('edited.xlsx')
    ce.raw_to_tsv('raw.xlsx', 'tsv')
    with open('tsv/dish_item.tsv', 'a', encoding='utf-8') as f:
        f.write('\n')                                       # a trailing blank line
    t = open('tsv/restaurant_location.tsv', encoding='utf-8').read().replace('1 Test St', '1 Test\x0bSt')
    open('tsv/restaurant_location.tsv', 'w', encoding='utf-8').write(t)     # a control character
    rc, out = run(scripts, 'wsheet_xlsx_bridge.py', 'to-xlsx', 'tsv', '-o', 'raw_b.xlsx')
    check('to-xlsx survives a control character in a TSV', rc == 0, out[-200:])
    run(scripts, 'action4_apply_modifier_changes.py', 'edited.xlsx', '-o', 'fr.xlsx')
    run(scripts, 'action5_to_raw_gsheet.py', 'fr.xlsx', 'raw_b.xlsx', '-o', 'final_raw.xlsx')
    before_loc = open('tsv/restaurant_location.tsv', encoding='utf-8').read()
    rc, out = run(scripts, 'wsheet_xlsx_bridge.py', 'to-tsv', 'final_raw.xlsx', 'tsv')
    after = ce.tsv_cells('tsv/dish_item.tsv')
    check('to-tsv: a trailing blank line no longer skips the tab; the default tabs include modifier_rule',
          rc == 0 and re.search(r'dish_item: [1-9]\d* cells patched', out) and 'modifier_rule' in out
          and 'NOT patched' not in out, out[-400:])
    check('… and a control character elsewhere is not "patched" away',
          open('tsv/restaurant_location.tsv', encoding='utf-8').read() == before_loc)
    # a TSV whose row count really differs → the run fails (never a silent skip)
    rows_ = open('tsv/modifier_item.tsv', encoding='utf-8').read().split('\n')
    open('tsv/modifier_item.tsv', 'w', encoding='utf-8').write('\n'.join(rows_[:-3]) + '\n')
    rc, out = run(scripts, 'wsheet_xlsx_bridge.py', 'to-tsv', 'final_raw.xlsx', 'tsv')
    check('to-tsv exits 1 when a tab cannot be patched (row count differs)', rc == 1 and 'NOT patched' in out,
          out[-300:])
    _ = after


def main():
    bundle = os.path.abspath(sys.argv[1])
    scripts = os.path.join(bundle, 'menu-change-payload', 'scripts')
    work = tempfile.mkdtemp(prefix='audit_checks_')
    for fn in (mrv_input, mrv_absolute_zero, refusals, dish_input, twins, bridge) + tuple(EXTRA):
        try:
            fn(scripts, work)
        except Exception as e:          # a crash is a failure, never a silent pass
            check(f'{fn.__name__}: checks ran to the end', False, f'{type(e).__name__}: {e}')
    os.chdir(HERE)
    shutil.rmtree(work, ignore_errors=True)
    print(f'\naudit: {sum(RESULTS)}/{len(RESULTS)} PASS')
    sys.exit(0 if all(RESULTS) else 1)


# ── Part 1: the $0 guard, G11, compare, side rules ───────────────────────────

S = 'Select Side'


def guard_addon(scripts, work):
    print('\nThe $0.00 guard — a group-level add-on is not ABSOLUTE pricing (R85325 / R83807 pattern)')
    fresh(work, 'guard_addon')
    variants = ['Order', 'Specialty', 'Small', 'Large', 'Party Tray']
    dishes = [('Honey Wings', 'Fried Specialties', {'Specialty': 8.95, 'Party Tray': 0.00}),
              ('Egg Roll', 'Appetizers', {'Order': 1.75})]
    mod_groups = {'Sauce': (0, ['With'], ['Fried Specialties'])}
    make_raw('raw.xlsx', 'R99910', 'Addon Wok', variants, dishes, mod_groups, {'With': {'Order': 0.0}},
             [('Fried Specialties', '', '', 'Sauce', 'With', 1.00)])
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Addon Wok'}, 'sections': [
            {'menu_section': 'Fried Specialties', 'price_columns': [], 'default_price': None,
             'items': [{'name': 'Honey Wings', 'prices': {'Specialty': 9.25, 'Party Tray': 45.00}}]}]}, f)
    ce.part1(scripts)
    ab = [r for r in load_workbook('preview.xlsx')['absolute_pricing'].iter_rows(min_row=2, values_only=True) if r[1]]
    wb = load_workbook('report.xlsx')
    col, _ref, rows = ce.report_rows(wb['Dish Item'])
    tray = wb['Dish Item'].cell(rows[('Honey Wings', 'Fried Specialties')], col['Party Tray']).value
    check('a $0.00 tray placeholder under a group-level "Sauce / With +$1.00" add-on is NOT held: the paper '
          'tray price is compared like any price', not ab and num(tray) == 45.0, (ab, tray))


def guard_platter(scripts, work):
    print('\nThe Fried Specialties rule: action1 and action5 agree (a priced $0 dish outside any group)')
    fresh(work, 'guard_platter')
    variants = ['Order', 'Specialty']
    dishes = [('Fried Crab Stick', 'Fried Specialties', {'Specialty': 0.00}),
              ('Egg Roll', 'Appetizers', {'Order': 1.75})]
    make_raw('raw.xlsx', 'R99911', 'Platter Wok', variants, dishes,
             {S: (0, ['By Itself', 'French Fries'], [])}, {'By Itself': {'Order': 0.0}, 'French Fries': {'Order': 0.0}},
             [('', 'Party Platter', 'Order', S, 'By Itself', 30.0)])
    raw = load_workbook('raw.xlsx')
    ws = raw['dish_item']
    ws.append(['{"id":"pp"}', 'Party Platter', '$0.00', ''])         # a live dish in no dish group
    raw.save('raw.xlsx')
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Platter Wok'}, 'sections': [
            {'menu_section': 'Fried Specialties', 'price_columns': [], 'default_price': None,
             'items': [{'name': 'Fried Crab Stick', 'prices': {'Specialty': 6.50}}]}]}, f)
    ce.part1(scripts)
    ab = {(r[1], r[2]): r[3] for r in load_workbook('preview.xlsx')['absolute_pricing'].iter_rows(min_row=2, values_only=True) if r[1]}
    check('action1 sees the restaurant prices on modifier rules (Party Platter, in no group) and holds the '
          '$0.00 Fried Specialties dish, as action5 would', ab.get(('Fried Crab Stick', 'Specialty')) == 'fried_specialty', ab)


def g11_group_anchor(scripts, work):
    print('\nG11: a priced GROUP-level By Itself rule on a dish that also has a dish price')
    fresh(work, 'g11_group')
    variants = ['Order', 'Specialty']
    dishes = [('Chicken Wings', 'Fried Specialties', {'Specialty': 7.99})]
    make_raw('raw.xlsx', 'R99912', 'Group Base Wok', variants, dishes,
             {S: (0, ['By Itself', 'French Fries'], ['Fried Specialties'])},
             {'By Itself': {'Order': 0.0}, 'French Fries': {'Order': 0.0}},
             [('Fried Specialties', '', 'Specialty', S, 'By Itself', 7.99),
              ('Fried Specialties', '', 'Specialty', S, 'French Fries', 9.39)])
    rc, out = run(scripts, 'modifier_rule_pricing.py', 'preflight', 'raw.xlsx')
    check('preflight blocks it as a DOUBLE CHARGE (dish_item $7.99 AND group-level By Itself $7.99)',
          rc != 0 and 'group-level' in out and 'DOUBLE CHARGE' in out, out[-400:])


def compare_flags(scripts, work):
    print('\ncompare: paper side prices for a dish with no rules of its own; re-runs; size-less rules')
    fresh(work, 'compare_flags')
    ce.build_r99907(side=True)
    ce.part1(scripts, compare=True)
    _rc, out = run(scripts, 'modifier_rule_pricing.py', 'compare', 'ocr.json', 'raw.xlsx', '--report', 'report.xlsx')
    fs = load_workbook('report.xlsx')['Fried Specialty']
    text = ' | '.join(str(c) for row in fs.iter_rows(values_only=True) for c in row if c is not None)
    check('a printed side price on a dish without its own rules is flagged (Fried Shrimp — NO DISH RULES), '
          'not mislabelled "⚠ size?"', 'NO DISH RULES' in out and 'Fried Shrimp' in out and '⚠ size?' not in text,
          out[-500:])
    rows1 = [r for r in range(1, fs.max_row + 1) if str(fs.cell(r, 1).value or '').startswith('PRICE DIFFERENCE')
             or str(fs.cell(r, 1).value or '').startswith('DELTA')]
    run(scripts, 'modifier_rule_pricing.py', 'compare', 'ocr.json', 'raw.xlsx', '--report', 'report.xlsx')
    fs2 = load_workbook('report.xlsx')['Fried Specialty']
    rows2 = [r for r in range(1, fs2.max_row + 1) if str(fs2.cell(r, 1).value or '').startswith('PRICE DIFFERENCE')
             or str(fs2.cell(r, 1).value or '').startswith('DELTA')]
    check('re-running compare does not push the Fried Specialty block down', rows1 == rows2 and fs.max_row == fs2.max_row,
          (rows1, rows2, fs.max_row, fs2.max_row))


def compare_variantless(scripts, work):
    print('\ncompare: a rule with no size links to the dish\'s only priced size')
    fresh(work, 'compare_vless')
    variants = ['Order', 'Specialty']
    dishes = [('Chicken Wings', 'Fried Specialties', {'Specialty': 7.99})]
    make_raw('raw.xlsx', 'R99913', 'Sizeless Wok', variants, dishes,
             {S: (0, ['French Fries'], ['Fried Specialties'])}, {'French Fries': {'Order': 0.0}},
             [('', 'Chicken Wings', '', S, 'French Fries', 1.40)])
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Sizeless Wok'}, 'sections': [
            {'menu_section': 'Specialties', 'price_columns': ['Specialty'], 'default_price': None,
             'side_pricing': {'columns': ['w. French Fries']},
             'items': [{'name': 'Chicken Wings', 'prices': {'Specialty': 7.99},
                        'side_prices': {'w. French Fries': 9.59}}]}]}, f)
    ce.part1(scripts, compare=True)
    wb = load_workbook('report.xlsx')
    ws = wb['Modifier Rule']
    _c, mr = ce.mrv(ws)
    r = next(iter(mr.values()))
    h = ws.cell(r, 8).value
    check("'by itself' of a size-less DELTA rule is a live link to the dish's price",
          "'Dish Item'!" in str(h) and num(evaluate(wb, ws, h)) == 7.99, h)


def side_rules_warning(scripts, work):
    print('\naction2 "no multi-column table" warning: only for real side tables')
    fresh(work, 'side_rules')
    variants = ['Order', 'Specialty']
    dishes = [('Edamame', 'Appetizers', {'Order': 4.95}),
              ('Chicken Wings', 'Fried Specialties', {'Specialty': 7.99}),
              ('Fried Shrimp', 'Fried Specialties', {'Specialty': 8.25})]
    dishes.append(('Yellowtail Roll', 'Sushi Rolls', {'Order': 6.50}))
    make_raw('raw.xlsx', 'R99914', 'Side Wok', variants, dishes,
             {'Spicy': (0, ['Make it Spicy'], ['Appetizers']), S: (0, ['French Fries', 'Fried Rice'], ['Fried Specialties']),
              'Toppings': (0, ['Masago', 'Quail Egg'], ['Sushi Rolls'])},
             {'Make it Spicy': {'Order': 0.0}, 'French Fries': {'Order': 0.0}, 'Fried Rice': {'Order': 0.0},
              'Masago': {'Order': 0.0}, 'Quail Egg': {'Order': 0.0}},
             [('', 'Edamame', 'Order', 'Spicy', 'Make it Spicy', 1.50),
              ('Fried Specialties', '', 'Specialty', S, 'French Fries', 1.40),
              ('Fried Specialties', '', 'Specialty', S, 'Fried Rice', 1.90),
              ('Sushi Rolls', '', 'Order', 'Toppings', 'Masago', 1.00),           # sushi toppings (real ER87669)
              ('Sushi Rolls', '', 'Order', 'Toppings', 'Quail Egg', 1.50)])
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Side Wok'}, 'sections': [
            {'menu_section': 'Menu', 'price_columns': [], 'default_price': None,
             'items': [{'name': 'Edamame', 'prices': {'Order': 4.95}},
                       {'name': 'Chicken Wings', 'prices': {'Specialty': 7.99}}]}]}, f)
    run(scripts, 'action1_raw_to_2a.py', 'raw.xlsx', '-o', 'preview.xlsx')
    rc, out = run(scripts, 'action2_ocr_to_2b_0625v2.py', 'ocr.json', 'preview.xlsx', '-o', 'ocr_out.xlsx')
    listed = [r[1] for r in load_workbook('preview.xlsx')['side_rules'].iter_rows(min_row=2, values_only=True) if r[1]]
    check('a single add-on (Edamame "Make it Spicy" +$1.50) and sushi toppings (Masago / Quail Egg) are not a '
          'side table; group-level side rules are',
          'Edamame' not in listed and 'Yellowtail Roll' not in listed
          and {'Chicken Wings', 'Fried Shrimp'} <= set(listed), listed)
    check('… so the warning names Chicken Wings (printed with no table), not Edamame',
          'NO MULTI-COLUMN TABLE' in out and 'Chicken Wings' in out and 'Edamame' not in out.split('NO MULTI')[1],
          out[-400:])


# ── price_inference ──────────────────────────────────────────────────────────

def pi_case(scripts, name, dishes, paper, variants=('Order', 'Small', 'Large', 'Party Tray'),
            also_in=(), extra=(), unmatched=()):
    """Part 1 on a made-up menu → (records by (group, dish, variant), report cells, summary)."""
    make_raw('raw.xlsx', 'R99920', name, list(variants), dishes, {}, {}, [], also_in=list(also_in))
    secs = {}
    for sec, dish, pr in paper:
        secs.setdefault(sec, []).append({'name': dish, 'prices': pr})
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': name}, 'sections': [
            {'menu_section': s_, 'price_columns': [], 'default_price': None, 'items': it}
            for s_, it in secs.items()]}, f)
    out = ce.part1(scripts)
    if extra:
        out['pi2'] = run(scripts, 'price_inference.py', 'preview.xlsx', 'ocr_out.xlsx', '--report',
                         'report.xlsx', *extra)
    js = json.load(open('price_inference_summary.json'))
    recs = {(r['group'], r['dish'], r['variant']): r for r in js['records']}
    wb = load_workbook('report.xlsx')
    di = wb['Dish Item']
    col, _ref, rows = ce.report_rows(di)
    cells = {}
    for (dish, grp), r in rows.items():
        for v in variants:
            if v in col:
                c = di.cell(r, col[v])
                if c.value not in (None, ''):
                    f = c.fill.fgColor.rgb[-6:] if c.fill and c.fill.fill_type else ''
                    cells[(grp, dish, v)] = (num(c.value), f == 'E4DFEC')
    return recs, cells, js, out


def pi_qualifier(scripts, work):
    print('\nprice_inference: "(White Meat)" is its own same-price family')
    fresh(work, 'pi_qual')
    dishes = [('General Tso Chicken', 'Chicken', {'Large': 14.25}), ('Orange Chicken', 'Chicken', {'Large': 14.25}),
              ('General Tso Chicken (White Meat)', 'Chicken', {'Large': 16.50}),
              ('Orange Chicken (White Meat)', 'Chicken', {'Large': 16.50})]
    paper = [('Chicken', 'General Tso Chicken', {'Large': 14.95}),
             ('Chicken', 'General Tso Chicken (White Meat)', {'Large': 16.95})]
    _r, cells, _j, _o = pi_case(scripts, 'Qual Wok', dishes, paper)
    check('Orange Chicken takes the plain 14.95, Orange Chicken (White Meat) the White Meat 16.95',
          cells.get(('Chicken', 'Orange Chicken', 'Large')) == (14.95, True)
          and cells.get(('Chicken', 'Orange Chicken (White Meat)', 'Large')) == (16.95, True), cells)


def pi_two_groups(scripts, work):
    print('\nprice_inference: a dish under two dish groups gets ONE answer')
    fresh(work, 'pi_two')
    dishes = [('Beef with Scallops', 'Beef', {'Large': 12.95}), ('Pepper Steak', 'Beef', {'Large': 12.95}),
              ('Shrimp with Garlic Sauce', 'Seafood', {'Large': 12.95}),
              ('Beef with Peppers', 'Beef', {'Large': 13.50}), ('Shrimp with Peppers', 'Seafood', {'Large': 13.50})]
    paper = [('Beef', 'Pepper Steak', {'Large': 13.95}), ('Seafood', 'Shrimp with Garlic Sauce', {'Large': 14.50}),
             ('Beef', 'Beef with Peppers', {'Large': 14.25})]
    recs, cells, _j, _o = pi_case(scripts, 'Two Wok', dishes, paper,
                                  also_in=[('Seafood', 'Beef with Scallops'), ('Seafood', 'Beef with Peppers')])
    check('two groups inferring different prices → neither row is inferred, both say why',
          cells.get(('Beef', 'Beef with Scallops', 'Large')) == (12.95, False)
          and cells.get(('Seafood', 'Beef with Scallops', 'Large')) == (12.95, False)
          and 'SEVERAL GROUPS' in recs[('Seafood', 'Beef with Scallops', 'Large')]['reason'],
          (cells.get(('Beef', 'Beef with Scallops', 'Large')), cells.get(('Seafood', 'Beef with Scallops', 'Large'))))


def pi_ordering(scripts, work):
    print('\nprice_inference: the Small < Large check uses the NEW printed price')
    fresh(work, 'pi_order')
    dishes = [('Orange Chicken', 'Chicken', {'Small': 8.95, 'Large': 9.95}),
              ("General Tso's Chicken", 'Chicken', {'Small': 8.95, 'Large': 9.95})]
    paper = [('Chicken', 'Orange Chicken', {'Small': 10.45}), ('Chicken', "General Tso's Chicken", {'Large': 10.25})]
    recs, cells, _j, _o = pi_case(scripts, 'Order Wok', dishes, paper)
    check('Orange Chicken Large is not inferred to 10.25 below its printed Small 10.45 (suggest-only)',
          cells.get(('Chicken', 'Orange Chicken', 'Large')) == (9.95, False)
          and recs[('Chicken', 'Orange Chicken', 'Large')]['status'] == 'SUGGESTED', recs[('Chicken', 'Orange Chicken', 'Large')])


def pi_rerun_and_screens(scripts, work):
    print('\nprice_inference: re-runs undo their own purple cells; the safety screens')
    fresh(work, 'pi_rerun')
    dishes = [('Kung Pao Chicken', 'Chicken', {'Small': 8.00, 'Large': 12.00}),
              ('Chicken with Broccoli', 'Chicken', {'Small': 8.00, 'Large': 12.00}),
              ('Chicken with Cashew Nuts', 'Chicken', {'Small': 8.00, 'Large': 12.00}),
              ('Hunan Beef', 'Beef', {'Large': 13.00}), ('Beef with Broccoli', 'Beef', {'Large': 13.00}),
              ('Pepper Steak', 'Beef', {'Large': 13.00}),
              ('Soda', 'Drink', {'Order': 1.50}), ('Lemonade', 'Drink', {'Order': 1.50}), ('Iced Tea', 'Drink', {'Order': 1.50}),
              ('Pork Fried Rice', 'Fried Rice', {'Large': 9.00}), ('Shrimp Fried Rice', 'Fried Rice', {'Large': 9.00}),
              ('Beef Fried Rice', 'Fried Rice', {'Large': 9.00}),
              ('Wonton Soup', 'Soup', {'Small': 3.00}), ('Egg Drop Soup', 'Soup', {'Small': 3.00}),
              ('Hot Sour Soup', 'Soup', {'Small': 3.00})]
    paper = [('Chicken', 'Kung Pao Chicken', {'Small': 8.50, 'Large': 12.75}),
             ('Chicken', 'Chicken with Broccoli', {'Small': 8.50, 'Large': 12.75}),
             ('Beef', 'Hunan Beef', {'Large': 12.50}), ('Beef', 'Beef with Broccoli', {'Large': 12.50}),     # decrease
             ('Drink', 'Soda', {'Order': 1.75}), ('Drink', 'Lemonade', {'Order': 1.75}),                    # drinks
             ('Fried Rice', 'Pork Fried Rice', {'Large': 15.00}), ('Fried Rice', 'Shrimp Fried Rice', {'Large': 15.00}),  # +6
             ('Soup', 'Wonton Soup', {'Small': 3.25}), ('Soup', 'Egg Drop Soup', {'Small': 3.50})]           # conflict
    recs, cells, js, _o = pi_case(scripts, 'Screen Wok', dishes, paper)
    st = lambda g, d, v: recs[(g, d, v)]['status']
    check('a cohort inference is applied purple (Chicken with Cashew Nuts 8.50 / 12.75)',
          cells.get(('Chicken', 'Chicken with Cashew Nuts', 'Small')) == (8.50, True))
    check('screens: an inferred decrease, a +$6 jump and a drink are suggest-only, never purple',
          st('Beef', 'Pepper Steak', 'Large') == 'SUGGESTED' and st('Fried Rice', 'Beef Fried Rice', 'Large') == 'SUGGESTED'
          and st('Drink', 'Iced Tea', 'Order') == 'SUGGESTED'
          and not cells[('Beef', 'Pepper Steak', 'Large')][1] and not cells[('Drink', 'Iced Tea', 'Order')][1],
          (st('Beef', 'Pepper Steak', 'Large'), st('Fried Rice', 'Beef Fried Rice', 'Large'), st('Drink', 'Iced Tea', 'Order')))
    check('a cohort whose donors disagree is never guessed (Hot Sour Soup)',
          recs[('Soup', 'Hot Sour Soup', 'Small')]['tier'] == 'T1 conflict'
          and cells[('Soup', 'Hot Sour Soup', 'Small')] == (3.0, False), recs[('Soup', 'Hot Sour Soup', 'Small')]['tier'])
    rc, _out = run(scripts, 'price_inference.py', 'preview.xlsx', 'ocr_out.xlsx', '--report', 'report.xlsx',
                   '--mode', 'suggest')
    wb = load_workbook('report.xlsx')
    di = wb['Dish Item']
    col, ref, rows = ce.report_rows(di)
    r = rows[('Chicken with Cashew Nuts', 'Chicken')]
    c = di.cell(r, col['Small'])
    js2 = json.load(open('price_inference_summary.json'))
    check('a --mode suggest re-run undoes the earlier purple cells (back to the system 8.00, _ref_ too)',
          num(c.value) == 8.0 and not (c.fill and c.fill.fill_type and c.fill.fgColor.rgb.endswith('E4DFEC'))
          and num(di.cell(r, ref['Small']).value) == 8.0, (c.value, di.cell(r, ref['Small']).value))
    check('… and its counts add up (eligible 0, applied 0 in suggest mode)',
          js2['stats']['eligible'] == 0 and js2['stats']['applied'] == 0, js2['stats'])
    # the specialist has typed over two cells; price_inference is re-run on that report
    run(scripts, 'price_inference.py', 'preview.xlsx', 'ocr_out.xlsx', '--report', 'report.xlsx', '--mode', 'apply')
    wb = load_workbook('report.xlsx')
    di = wb['Dish Item']
    di.cell(r, col['Small']).value = 8.95                     # over a purple 8.50
    di.cell(r, col['Large']).value = 13.25                    # over a purple 12.75
    wb.save('report.xlsx')
    run(scripts, 'price_inference.py', 'preview.xlsx', 'ocr_out.xlsx', '--report', 'report.xlsx', '--mode', 'apply')
    di = load_workbook('report.xlsx')['Dish Item']
    check('a re-run never undoes or overwrites a cell the specialist typed over (8.95 / 13.25 stay)',
          (num(di.cell(r, col['Small']).value), num(di.cell(r, col['Large']).value)) == (8.95, 13.25),
          (di.cell(r, col['Small']).value, di.cell(r, col['Large']).value))


def pi_tray_dupes(scripts, work):
    print('\nprice_inference: T-TRAY agreement counts dishes, not rows')
    fresh(work, 'pi_tray')
    dishes = [('Dish A', 'Chicken', {'Large': 10.00, 'Party Tray': 48.00}),
              ('Dish B', 'Chicken', {'Large': 12.00, 'Party Tray': 58.00}),
              ('Dish C', 'Chicken', {'Large': 14.00, 'Party Tray': 68.00}),
              ('Dish D', 'Chicken', {'Large': 11.00, 'Party Tray': 60.00}),
              ('Dish E', 'Chicken', {'Large': 13.00, 'Party Tray': 70.00})]
    paper = [('Chicken', 'Dish A', {'Large': 10.50})]
    _r, _c, js, _o = pi_case(scripts, 'Tray Wok', dishes, paper,
                             also_in=[('Chef Special', 'Dish A'), ('Chef Special', 'Dish B')])
    t = js['tray_formulas'][0]
    check('3 of 5 dishes fit the formula (60%) — listing two of them in a second group does not make it 5/7',
          t['status'] == 'NOT DETECTED' and t['population'] == 5, (t['status'], t['population'], t['reason']))
    q = [' '.join(str(v) for v in row if v not in (None, ''))
         for row in load_workbook('report.xlsx')['Questionnaire'].iter_rows(values_only=True)]
    p0 = [r for r in q if r.startswith('P-TRAY-0') and ' Chicken ' in r]
    check('P-TRAY-0 says there is no reliable formula for Chicken AND why (the agreement found)',
          len(p0) == 1 and 'No reliable tray formula for Chicken' in p0[0] and t['reason'] in p0[0], p0)
    run(scripts, 'price_inference.py', 'preview.xlsx', 'ocr_out.xlsx', '--report', 'report.xlsx', '--mode', 'apply')
    wb = load_workbook('report.xlsx')
    n_sum = sum(1 for row in wb['Summary'].iter_rows(values_only=True) if row[0] == 'AUTO-INFERENCE (price_inference.py)')
    n_tray = sum(1 for row in wb['Questionnaire'].iter_rows(values_only=True)
                 if str(row[0] or '').startswith('P-TRAY — '))
    n_p0 = sum(1 for row in wb['Questionnaire'].iter_rows(values_only=True) if row[0] == 'P-TRAY-0')
    check('a second run replaces its Summary and P-TRAY blocks (one of each — one formula row per dish '
          'group, Chicken and Chef Special — not two)', (n_sum, n_tray, n_p0) == (1, 1, 2), (n_sum, n_tray, n_p0))
    run(scripts, 'price_inference.py', 'preview.xlsx', 'ocr_out.xlsx', '--report', 'report.xlsx', '--mode', 'apply',
        '--no-tray')
    left = [row[0] for row in load_workbook('report.xlsx')['Questionnaire'].iter_rows(values_only=True)
            if str(row[0] or '').startswith('P-TRAY')]
    check('a --no-tray re-run removes the earlier P-TRAY block (no stale tray answers)', not left, left)


def pi_healthy(scripts, work):
    print('\nprice_inference: Healthy rule — "on the paper" includes a printed tray and a P1 name')
    fresh(work, 'pi_healthy')
    dishes = [('Chicken with Broccoli', 'Chicken', {'Small': 8.00, 'Large': 12.00, 'Party Tray': 55.00}),
              ('Shrimp with Broccoli', 'Seafood', {'Small': 9.00, 'Large': 13.00}),
              ('Steam Chicken with Broccoli', 'Healthy Dish', {'Small': 8.00, 'Large': 12.00, 'Party Tray': 55.00}),
              ('Steam Shrimp with Broccoli', 'Healthy Dish', {'Small': 9.00, 'Large': 13.00}),
              ('Beef with Broccoli', 'Beef', {'Small': 9.50, 'Large': 13.50}),
              ('Steam Beef with Broccoli', 'Healthy Dish', {'Small': 9.50, 'Large': 13.50}),
              ('Pork with Broccoli', 'Pork', {'Small': 8.25, 'Large': 12.25}),
              ('Steam Pork with Broccoli', 'Healthy Dish', {'Small': 8.25, 'Large': 12.25})]
    paper = [('Chicken', 'Chicken with Broccoli', {'Small': 8.50, 'Large': 12.75}),
             ('Seafood', 'Shrimp with Broccoli', {'Small': 9.50, 'Large': 13.75}),
             ('Beef', 'Beef with Broccoli', {'Small': 9.95, 'Large': 14.25}),
             ('Pork', 'Pork with Broccoli', {'Small': 8.75, 'Large': 12.95}),
             ('Healthy Dish', 'Steam Pork with Broccoli', {'Lunch': 9.45}),       # a size it does not have
             ('Healthy Dish', 'Steam Chicken with Broccoli', {'Party Tray': 60.00}),
             ('Healthy Dish', 'Diet Shrimp w/ Broccoli (Steamed)', {'Small': 10.25, 'Large': 14.95}),
             ('Healthy Dish', 'Steam Beef with Broccoli Combo Plate', {'Small': 10.95, 'Large': 15.25})]
    recs, cells, _j, _o = pi_case(scripts, 'Healthy Wok', dishes, paper)
    check('printed only as a tray → on the paper → not given the regular dish\'s price',
          cells.get(('Healthy Dish', 'Steam Chicken with Broccoli', 'Small')) == (8.0, False),
          recs.get(('Healthy Dish', 'Steam Chicken with Broccoli', 'Small'), {}).get('reason'))
    check('printed under another name (a P1 row "Diet Shrimp w/ Broccoli (Steamed)") → not inferred',
          cells.get(('Healthy Dish', 'Steam Shrimp with Broccoli', 'Small')) == (9.0, False),
          recs.get(('Healthy Dish', 'Steam Shrimp with Broccoli', 'Small'), {}).get('reason'))
    hdr, q = questionnaire_rows()
    psize = [r for r in q if str(r[0] or '').startswith('P-SIZE-')]
    check('a size the dish does not have (paper "Steam Pork with Broccoli" Lunch, system Small / Large) is '
          'asked as P-SIZE — never dropped silently',
          any('Steam Pork with Broccoli [Lunch]' == str(r[3]) for r in psize), [r[3] for r in psize])
    check('… and that dish counts as on the paper (not given its regular dish\'s price)',
          cells.get(('Healthy Dish', 'Steam Pork with Broccoli', 'Large')) == (12.25, False),
          recs.get(('Healthy Dish', 'Steam Pork with Broccoli', 'Large'), {}).get('reason'))
    check('… or a P1 row action2 suggests for it ("Maybe: Steam Beef with Broccoli (substring)") → not inferred',
          cells.get(('Healthy Dish', 'Steam Beef with Broccoli', 'Small')) == (9.5, False),
          recs.get(('Healthy Dish', 'Steam Beef with Broccoli', 'Small'), {}).get('reason'))


def pi_names_and_labels(scripts, work):
    print('\nprice_inference: name variants, donor counting, the merged-cohort label')
    fresh(work, 'pi_names')
    dishes = [("Gen. Tso's Chicken", 'Chicken', {'Large': 12.95}), ('Orange Flavored Chicken', 'Chicken', {'Large': 12.95}),
              ('Steamed Chicken w. Broccoli', 'Healthy Dish', {'Large': 12.00}),
              ('Chicken with Broccoli', 'Chicken', {'Large': 12.00}),
              ('Chicken Chow Mein', 'Chow Mein', {'Large': 12.25}), ('Pork Chow Mein', 'Chow Mein', {'Large': 12.25})]
    paper = [('Chicken', "Gen. Tso's Chicken", {'Large': 13.95}), ('Chicken', 'Chicken with Broccoli', {'Large': 12.75}),
             ('Chow Mein', 'Chicken Chow Mein', {'Large': 12.50})]
    recs, cells, _j, _o = pi_case(scripts, 'Name Wok', dishes, paper, also_in=[('Chow Mein::RPOS', 'Chicken Chow Mein')])
    check('"Gen. Tso\'s" and "Orange Flavored" are one family; "Steamed … w." is the Healthy twin',
          cells.get(('Chicken', 'Orange Flavored Chicken', 'Large')) == (13.95, True)
          and cells.get(('Healthy Dish', 'Steamed Chicken w. Broccoli', 'Large')) == (12.75, True), cells)
    r = recs[('Chow Mein', 'Pork Chow Mein', 'Large')]
    check('a dish listed twice counts as ONE donor; no "(Chow Mein + Chop Suey)" label between Chow Mein groups',
          r['donor_n'] == 1 and 'Chop Suey' not in r['tier'], (r['donor_n'], r['tier']))


def pi_stale_report(scripts, work):
    print('\nprice_inference: never writes into a report row that is a different dish')
    fresh(work, 'pi_stale')
    dishes = [('Chicken with Broccoli', 'Chicken', {'Small': 8.0, 'Large': 12.0}),
              ('Chicken with Cashew Nuts', 'Chicken', {'Small': 8.0, 'Large': 12.0}),
              ('Kung Pao Chicken', 'Chicken', {'Small': 8.0, 'Large': 12.0})]
    paper = [('Chicken', 'Chicken with Broccoli', {'Small': 8.5, 'Large': 12.75}),
             ('Chicken', 'Kung Pao Chicken', {'Small': 8.5, 'Large': 12.75})]
    pi_case(scripts, 'Stale Wok', dishes, paper)
    shutil.copy('report.xlsx', 'old_report.xlsx')

    def printed(path):
        ws = load_workbook(path)['Dish Item']
        return {(r, c): ws.cell(r, c).value for r in range(2, ws.max_row + 1) for c in range(1, ws.max_column + 1)
                if ws.cell(r, c).fill.fill_type and ws.cell(r, c).fill.fgColor.rgb.endswith('ADD8E6')}
    blue = printed('old_report.xlsx')
    # wsheet since Part 1: one more Chicken dish — the Menu Preview rows shift by one
    make_raw('raw.xlsx', 'R99920', 'Stale Wok', ['Order', 'Small', 'Large', 'Party Tray'],
             dishes + [('Sesame Chicken', 'Chicken', {'Small': 8.0, 'Large': 12.0})], {}, {}, [])
    run(scripts, 'action1_raw_to_2a.py', 'raw.xlsx', '-o', 'preview.xlsx')
    run(scripts, 'action2_ocr_to_2b_0625v2.py', 'ocr.json', 'preview.xlsx', '-o', 'ocr_out.xlsx')
    shifted = load_workbook('preview.xlsx').worksheets[0].cell(3, 2).value == 'Sesame Chicken'
    rc, out = run(scripts, 'price_inference.py', 'preview.xlsx', 'ocr_out.xlsx', '--report', 'old_report.xlsx',
                  '--mode', 'apply')
    after = printed('old_report.xlsx')
    check('a stale report (rows shifted) keeps every printed blue cell; nothing is written onto the wrong dish',
          shifted and after == blue and 'not this dish' in json.dumps(json.load(open('price_inference_summary.json'))),
          (shifted, set(blue) - set(after)))


def pi_caps_t2_labels(scripts, work):
    print('\nprice_inference: T2 group delta, the $ and % caps one at a time, labels and colours')
    fresh(work, 'pi_caps')
    dishes = [('Chicken Lo Mein', 'Lo Mein', {'Small': 7.00, 'Large': 10.00}),
              ('Pork Lo Mein', 'Lo Mein', {'Large': 11.00}), ('Beef Lo Mein', 'Lo Mein', {'Large': 12.00}),
              ('House Lo Mein', 'Lo Mein', {'Large': 13.00}),
              ('Wonton Soup', 'Soup', {'Small': 3.00}), ('Egg Drop Soup', 'Soup', {'Small': 3.00}),
              ('Shrimp Lobster Sauce', 'Seafood', {'Large': 20.00}), ('Shrimp Delight', 'Seafood', {'Large': 20.00}),
              ('General Tso Chicken', 'Chicken', {'Large': 14.25}), ('Orange Chicken', 'Chicken', {'Large': 14.25}),
              ('Sesame Chicken', 'Chicken', {'Large': 14.25})]
    paper = [('Lo Mein', 'Chicken Lo Mein', {'Large': 10.50}), ('Lo Mein', 'Pork Lo Mein', {'Large': 11.50}),
             ('Lo Mein', 'Beef Lo Mein', {'Large': 12.50}),
             ('Soup', 'Wonton Soup', {'Small': 4.50}),                        # +1.50 = 50 % (cap 40 %)
             ('Seafood', 'Shrimp Lobster Sauce', {'Large': 25.75}),           # +5.75 (cap $5), 29 %
             ('Chicken', 'General Tso Chicken', {'Large': 14.95}), ('Chicken', 'Orange Chicken', {'Large': 15.25})]
    recs, cells, _j, _o = pi_case(scripts, 'Caps Wok', dishes, paper)
    rec = lambda g, d, v: recs.get((g, d, v), {})
    check('T2: every priced Lo Mein moved +0.50 (3 donors) → House Lo Mein Large 13.50, purple',
          cells.get(('Lo Mein', 'House Lo Mein', 'Large')) == (13.5, True)
          and rec('Lo Mein', 'House Lo Mein', 'Large').get('tier') == 'T2 uniform group delta',
          (cells.get(('Lo Mein', 'House Lo Mein', 'Large')), rec('Lo Mein', 'House Lo Mein', 'Large').get('tier')))
    r1, r2 = rec('Soup', 'Egg Drop Soup', 'Small'), rec('Seafood', 'Shrimp Delight', 'Large')
    check('the % cap alone holds a +50 % inference (Egg Drop Soup 3.00 → 4.50)',
          r1.get('status') == 'SUGGESTED' and 'max-delta-pct' in r1.get('reason', '')
          and 'max-delta-abs' not in r1.get('reason', ''), r1.get('reason'))
    check('the $ cap alone holds a +$5.75 inference (Shrimp Delight 20.00 → 25.75)',
          r2.get('status') == 'SUGGESTED' and 'max-delta-abs' in r2.get('reason', '')
          and 'max-delta-pct' not in r2.get('reason', ''), r2.get('reason'))
    check('V-GAP = the dish is printed, this size is not; D-GAP = the dish is not printed',
          rec('Lo Mein', 'Chicken Lo Mein', 'Small').get('scenario') == 'V-GAP'
          and rec('Lo Mein', 'House Lo Mein', 'Large').get('scenario') == 'D-GAP',
          (rec('Lo Mein', 'Chicken Lo Mein', 'Small').get('scenario'), rec('Lo Mein', 'House Lo Mein', 'Large').get('scenario')))
    wb = load_workbook('report.xlsx')
    tab = {(row[4].value, row[5].value): row for row in wb['Price Inference'].iter_rows(min_row=6)
           if str(row[0].value or '').startswith('INF-')}
    fam = tab.get(('Sesame Chicken', 'Large'))
    check('a same-price family conflict (Gen Tso 14.95 vs Orange 15.25) is red on the Price Inference tab',
          fam is not None and fam[9].value == 'T0 conflict' and fam[0].fill.fgColor.rgb.endswith('FFC7CE'),
          fam and (fam[9].value, fam[0].fill.fgColor.rgb))
    hl = tab.get(('House Lo Mein', 'Large'))
    check('… and its columns line up: Scenario / Status / Current / Inferred',
          hl is not None and (hl[1].value, hl[2].value, num(hl[6].value), num(hl[7].value)) == ('D-GAP', 'APPLIED', 13.0, 13.5),
          hl and [c.value for c in hl[:8]])
    di = wb['Dish Item']
    col, ref, rows = ce.report_rows(di)
    r = rows[('House Lo Mein', 'Lo Mein')]
    check('the hidden _ref_ copy is refreshed to the written price (no false "specialist edit" colour)',
          num(di.cell(r, ref['Large']).value) == 13.5, di.cell(r, ref['Large']).value)


def abs_x_held(scripts, work):
    print('\naction5: x never removes an ABSOLUTE size')
    fresh(work, 'abs_x')
    ce.build_r99907(side=True)
    ce.part1(scripts, compare=True)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    di = wb['Dish Item']
    col, _ref, rows = ce.report_rows(di)
    di.cell(rows[('Chicken Wings', 'Fried Specialties')], col['Specialty']).value = 'x'
    wb.save('edited.xlsx')
    p2 = ce.part2(scripts)
    d = ce.live('final_raw.xlsx', 'dish_item')['Chicken Wings'][0]
    g = p2['summary'].get('action5', {}).get('absolute_guard', {}).get('cells', [])
    check('x on Chicken Wings [Specialty] ($0.00, priced on its modifier rules) keeps the $0.00 and is held',
          d['Specialty'] == '$0.00' and any(c['dish'] == 'Chicken Wings' and 'x' in c['report_price'] for c in g),
          (d, g))


def bridge_stale_tsv(scripts, work):
    print('\nbridge: a TSV that changed since to-xlsx is never overwritten')
    fresh(work, 'bridge_stale')
    ce.build_r99906()
    ce.part1(scripts)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    di = wb['Dish Item']
    col, _ref, rows = ce.report_rows(di)
    di.cell(rows[('Wonton Soup', 'Soup')], col['Large']).value = 'x'              # remove $5.25
    di.cell(rows[('Fried Rice', 'Rice')], col['Large']).value = 9.5               # new price, empty cell
    wb.save('edited.xlsx')
    ce.raw_to_tsv('raw.xlsx', 'tsv')
    run(scripts, 'wsheet_xlsx_bridge.py', 'to-xlsx', 'tsv', '-o', 'raw_b.xlsx')
    run(scripts, 'action4_apply_modifier_changes.py', 'edited.xlsx', '-o', 'fr.xlsx')
    run(scripts, 'action5_to_raw_gsheet.py', 'fr.xlsx', 'raw_b.xlsx', '-o', 'final_raw.xlsx')
    rows_ = ce.tsv_cells('tsv/dish_item.tsv')
    h2, h3 = rows_[1], rows_[2]
    nc = h3.index('dish_item')
    pc = {h3[i]: i for i, m in enumerate(h2) if m.endswith('unit_price')}
    at = {r[nc]: i for i, r in enumerate(rows_) if i >= 3 and r[0].startswith('{')}
    clean = [list(r) for r in rows_]
    # someone changed both cells in wsheet and the TSVs were re-fetched into the same dir
    for name, v, val in (('Wonton Soup', 'Large', '$5.95'), ('Fried Rice', 'Large', '$9.00')):
        row = rows_[at[name]]
        row += [''] * (pc[v] + 1 - len(row))
        row[pc[v]] = val
    open('tsv/dish_item.tsv', 'w', encoding='utf-8').write('\n'.join('\t'.join(r) for r in rows_) + '\n')
    rc, out = run(scripts, 'wsheet_xlsx_bridge.py', 'to-tsv', 'final_raw.xlsx', 'tsv')
    after = ce.tsv_cells('tsv/dish_item.tsv')
    check('to-tsv keeps the newer $5.95 (the x was for $5.25) and the newer $9.00 (the cell was empty), '
          'says so, and exits 1',
          after[at['Wonton Soup']][pc['Large']] == '$5.95' and after[at['Fried Rice']][pc['Large']] == '$9.00'
          and rc == 1 and 'NOT applied' in out, out[-500:])
    # the TSV row of Fried Rice ends early (its trailing empty fields left out) — the fill must still land
    ragged = [list(r) for r in clean]
    fr = ragged[at['Fried Rice']]
    while fr and fr[-1] == '':
        fr.pop()
    ragged[at['Fried Rice']] = fr[:pc['Large']]
    open('tsv/dish_item.tsv', 'w', encoding='utf-8').write('\n'.join('\t'.join(r) for r in ragged) + '\n')
    rc, out = run(scripts, 'wsheet_xlsx_bridge.py', 'to-tsv', 'final_raw.xlsx', 'tsv')
    after = ce.tsv_cells('tsv/dish_item.tsv')
    check('… while on the unchanged TSV both requests apply (x clears $5.25; $9.50 fills the empty cell, even '
          'past the end of a short TSV row)',
          rc == 0 and after[at['Wonton Soup']][pc['Large']] == '' and after[at['Fried Rice']][pc['Large']] == '$9.50',
          out[-300:])


def questionnaire_rows(path='report.xlsx'):
    ws = load_workbook(path)['Questionnaire']
    hdr = [c.value for c in ws[1]]
    return hdr, [[c.value for c in row] for row in ws.iter_rows(min_row=2)]


def p1_unmatched(scripts, work):
    print('\nQuestionnaire P1: an unmatched paper dish shows what the paper printed')
    fresh(work, 'p1')
    variants = ['Order', 'Small', 'Large']
    dishes = [('Egg Roll', 'Appetizers', {'Order': 1.75}), ('Wonton Soup', 'Soup', {'Small': 3.25, 'Large': 5.25})]
    make_raw('raw.xlsx', 'R99915', 'P1 Wok', variants, dishes, {}, {}, [])
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'P1 Wok'}, 'sections': [
            {'menu_section': 'Chef Specials', 'price_columns': [], 'default_price': None,
             'items': [{'name': 'Dragon and Phoenix', 'prices': {'Large': 16.95}},
                       {'name': 'Seafood Delight Soup', 'prices': {'Small': 4.25, 'Large': 7.95}}]},
            {'menu_section': 'Appetizers', 'price_columns': [], 'default_price': None,
             'items': [{'name': 'Egg Roll', 'prices': {'Order': 1.95}}]}]}, f)
    ce.part1(scripts)
    hdr, q = questionnaire_rows()
    mp, det = hdr.index('Menu Price'), hdr.index('Details')
    by = {r[3]: r for r in q if str(r[0] or '').startswith('P1-') and r[3]}
    dp, sd = by.get('Dragon and Phoenix'), by.get('Seafood Delight Soup')
    check('one paper price → Menu Price is that number (16.95)', dp is not None and num(dp[mp]) == 16.95, dp)
    check('several sizes → every printed size is shown (Small 4.25 / Large 7.95)',
          sd is not None and all(x in f'{sd[mp]} {sd[det]}' for x in ('4.25', '7.95')), sd)


def pabs_notes_fallback(scripts, work):
    print('\nP-ABS: with nowhere to edit, the row says "write it in Notes → Wonka"')
    fresh(work, 'pabs_notes')
    ce.build_r99907(side=True)
    ce.part1(scripts)                                        # no compare → no Modifier Rule tab
    hdr, q = questionnaire_rows()
    det = hdr.index('Details')
    rules = [r for r in q if str(r[0] or '').startswith('P-ABS-') and 'Modifier Rule' in str(r[det])]
    check('without a Modifier Rule tab the P-ABS rule rows say to write the price in Notes → Wonka',
          rules and all('No Modifier Rule tab in this report? Write the price you want in Notes → Wonka'
                        in str(r[det]) for r in rules), [r[det] for r in rules][:2])
    run(scripts, 'modifier_rule_pricing.py', 'compare', 'ocr.json', 'raw.xlsx', '--report', 'report.xlsx')
    hdr, q = questionnaire_rows()
    rows = {str(r[3]): str(r[det]) for r in q if str(r[0] or '').startswith('P-ABS-')}
    cw = rows.get('Chicken Wings [Specialty]', '')
    cs = rows.get('Fried Crab Stick [Specialty]', '')
    check('after compare: Chicken Wings points at its By Itself row; Fried Crab Stick (no By Itself row) '
          'says Notes → Wonka', re.search(r"'Modifier Rule' tab · row \d+ · column B", cw)
          and 'write the plain price you want in Notes → Wonka' in cs, (cw[:120], cs[:160]))


def compare_no_side_rules(scripts, work):
    print('\ncompare: printed sides on a dish no rule prices sides for → a note, not a callback item')
    fresh(work, 'compare_noside')
    variants = ['Order', 'Specialty']
    dishes = [('Chicken Wings', 'Fried Specialties', {'Specialty': 7.99}),
              ('Fried Crab Stick', 'Fried Specialties', {'Specialty': 6.50})]
    make_raw('raw.xlsx', 'R99916', 'Note Wok', variants, dishes,
             {S: (0, ['French Fries'], []), 'Sauce': (0, ['With'], [])},
             {'French Fries': {'Order': 0.0}, 'With': {'Order': 0.0}},
             [('', 'Chicken Wings', 'Specialty', S, 'French Fries', 1.40),
              ('Fried Specialties', '', '', 'Sauce', 'With', 1.00)])       # a group ADD-ON, not a side
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Note Wok'}, 'sections': [
            {'menu_section': 'Specialties', 'price_columns': ['Specialty'], 'default_price': None,
             'side_pricing': {'columns': ['w. French Fries']},
             'items': [{'name': 'Chicken Wings', 'prices': {'Specialty': 7.99}, 'side_prices': {'w. French Fries': 9.39}},
                       {'name': 'Fried Crab Stick', 'prices': {'Specialty': 6.50},
                        'side_prices': {'w. French Fries': 7.90}}]}]}, f)
    ce.part1(scripts, compare=True)
    _rc, out = run(scripts, 'modifier_rule_pricing.py', 'compare', 'ocr.json', 'raw.xlsx', '--report', 'report.xlsx')
    conflicts = out.split('Conflicts NOT auto-resolved', 1)[1].split('Notes:', 1)[0] if 'Conflicts NOT' in out else ''
    check('Fried Crab Stick (only a "Sauce / With +$1.00" group add-on) is a NO SIDE RULES note, not a '
          'callback conflict', 'NO SIDE RULES: Fried Crab Stick' in out and 'Fried Crab Stick' not in conflicts,
          out[-600:])


def mrv_absolute_zero_sizeless(scripts, work):
    print('\nABSOLUTE: a By Itself rule for every size (no size on the rule) never goes to $0.00 either')
    fresh(work, 'mrv_abs0_nosize')
    variants = ['Order', 'Specialty']
    make_raw('raw.xlsx', 'R99917', 'Sizeless Abs Wok', variants,
             [('Chicken Wings', 'Fried Specialties', {'Specialty': 0.00}), ('Egg Roll', 'Appetizers', {'Order': 1.75})],
             {S: (0, ['By Itself', 'French Fries'], ['Fried Specialties'])},
             {'By Itself': {'Order': 0.0}, 'French Fries': {'Order': 0.0}},
             [('', 'Chicken Wings', '', S, 'By Itself', 7.50), ('', 'Chicken Wings', '', S, 'French Fries', 8.90)])
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Sizeless Abs Wok'}, 'sections': [
            {'menu_section': 'Fried Specialties', 'price_columns': ['Specialty'], 'default_price': None,
             'side_pricing': {'columns': ['w. French Fries']},
             'items': [{'name': 'Chicken Wings', 'prices': {'Specialty': 7.50}, 'side_prices': {'w. French Fries': 8.90}}]}]}, f)
    ce.part1(scripts, compare=True)
    hdr, q = questionnaire_rows()
    pabs = [str(r[hdr.index('Details')]) for r in q if str(r[0] or '').startswith('P-ABS-')
            and str(r[3]).startswith('Chicken Wings')]
    check('its P-ABS row points at that By Itself row (a rule with no size counts)',
          pabs and re.search(r"'Modifier Rule' tab · row \d+ · column B", pabs[0]), pabs[:1])
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    ws = wb['Modifier Rule']
    _c, mr = ce.mrv(ws)
    k = next(k for k in mr if k[0] == 'Chicken Wings' and k[2] == 'By Itself')
    ws.cell(mr[k], 2).value = 0
    wb.save('edited.xlsx')
    p2 = ce.part2(scripts, with_rules=True)
    check('By Itself (no size) of a $0.00 dish set to 0 → held, $7.50 kept',
          ce.rule_cell('final_raw.xlsx', '', 'Chicken Wings', '', 'By Itself') == '$7.50'
          and p2['summary'].get('modifier_rule', {}).get('held'),
          (ce.rule_cell('final_raw.xlsx', '', 'Chicken Wings', '', 'By Itself'), p2['summary'].get('modifier_rule', {}).get('held')))


def guard_required_choice(scripts, work):
    print('\nThe $0.00 guard — group-level rules on a REQUIRED choice hold the price (full totals, no By Itself)')
    fresh(work, 'guard_required')
    variants = ['Order', 'Specialty']
    dishes = [('Wings Combo', 'Fried Specialties', {'Order': 0.00}), ('Egg Roll', 'Appetizers', {'Order': 1.75})]
    make_raw('raw.xlsx', 'R99918', 'Combo Wok', variants, dishes,
             {S: (1, ['French Fries', 'Fried Rice'], ['Fried Specialties'])},          # required
             {'French Fries': {'Order': 0.0}, 'Fried Rice': {'Order': 0.0}},
             [('Fried Specialties', '', 'Order', S, 'French Fries', 9.95),
              ('Fried Specialties', '', 'Order', S, 'Fried Rice', 10.45)])
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Combo Wok'}, 'sections': [
            {'menu_section': 'Fried Specialties', 'price_columns': ['Order'], 'default_price': None,
             'items': [{'name': 'Wings Combo', 'prices': {'Order': 9.95}}]}]}, f)
    ce.part1(scripts)
    ab = {(r[1], r[2]) for r in load_workbook('preview.xlsx')['absolute_pricing'].iter_rows(min_row=2, values_only=True) if r[1]}
    check('Wings Combo ($0.00, its price on the required "Select Side" rules) is held — the paper $9.95 is '
          'not written onto the dish', ('Wings Combo', 'Order') in ab, ab)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    di = wb['Dish Item']
    col, _ref, rows = ce.report_rows(di)
    di.cell(rows[('Wings Combo', 'Fried Specialties')], col['Order']).value = 9.95    # typed anyway
    wb.save('edited.xlsx')
    p2 = ce.part2(scripts)
    d = ce.live('final_raw.xlsx', 'dish_item')['Wings Combo'][0]
    g = p2['summary'].get('action5', {}).get('absolute_guard', {}).get('cells', [])
    check('… and action5 holds a price typed onto it (no $9.95 dish + $9.95 rule double charge)',
          d['Order'] == '$0.00' and any(c['dish'] == 'Wings Combo' for c in g), (d, g))


def apply_g11_full_price(scripts, work):
    print('\napply: an ABSOLUTE rule (the full price of a $0.00 dish) never goes to $0.00')
    fresh(work, 'apply_g11_full')
    ce.build_r99907(side=True)
    ce.part1(scripts, compare=True)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    ws = wb['Modifier Rule']
    _c, mr = ce.mrv(ws)
    ws.cell(mr[('', 'Specialty', 'By Itself')], 2).value = 0              # group-level By Itself
    ws.cell(mr[('Chicken Wings', 'Specialty', 'French Fries')], 2).value = 0   # an ABSOLUTE side rule
    wb.save('edited.xlsx')
    p2 = ce.part2(scripts, with_rules=True)
    held = p2['summary'].get('modifier_rule', {}).get('held', [])
    check('the group-level By Itself ($8.00) and Chicken Wings + French Fries ($8.90, the full combo) set to 0 '
          'are both held — they would make the plain dish / the combo free',
          ce.rule_cell('final_raw.xlsx', 'Fried Specialties', '', 'Specialty', 'By Itself') == '$8.00'
          and ce.rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'French Fries') == '$8.90'
          and len(held) == 2, (ce.rule_cell('final_raw.xlsx', 'Fried Specialties', '', 'Specialty', 'By Itself'),
                               ce.rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'French Fries'), held))
    # DELTA: an upcharge set to 0 is a free side — applied
    fresh(work, 'apply_delta_zero')
    delta_wok()
    ce.part1(scripts, compare=True)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    ws = wb['Modifier Rule']
    _c, mr = ce.mrv(ws)
    ws.cell(mr[('Chicken Wings', 'Specialty', 'French Fries')], 2).value = 0
    wb.save('edited.xlsx')
    p2 = ce.part2(scripts, with_rules=True)
    check('… while a DELTA upcharge set to 0 (a free side on a priced dish) is applied',
          ce.rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'French Fries') == '$0.00'
          and not p2['summary'].get('modifier_rule', {}).get('held'),
          (ce.rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'French Fries'),
           p2['summary'].get('modifier_rule', {}).get('held')))


def new_size_group_anchor(scripts, work):
    print('\naction5: a new size on a dish whose group prices that size on a By Itself rule is held')
    fresh(work, 'new_size_group')
    variants = ['Order', 'Specialty']
    dishes = [('Fried Fish', 'Fried Specialties', {'Order': 6.00}), ('Egg Roll', 'Appetizers', {'Order': 1.75})]
    make_raw('raw.xlsx', 'R99919', 'Group Base Wok', variants, dishes,
             {S: (0, ['By Itself', 'French Fries'], ['Fried Specialties'])},
             {'By Itself': {'Order': 0.0}, 'French Fries': {'Order': 0.0}},
             [('Fried Specialties', '', 'Specialty', S, 'By Itself', 8.00),
              ('Fried Specialties', '', 'Specialty', S, 'French Fries', 9.40)])
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Group Base Wok'}, 'sections': [
            {'menu_section': 'Appetizers', 'price_columns': ['Order'], 'default_price': None,
             'items': [{'name': 'Egg Roll', 'prices': {'Order': 1.95}}]}]}, f)
    ce.part1(scripts)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    di = wb['Dish Item']
    col, _ref, rows = ce.report_rows(di)
    di.cell(rows[('Fried Fish', 'Fried Specialties')], col['Specialty']).value = 8.0
    wb.save('edited.xlsx')
    p2 = ce.part2(scripts)
    d = ce.live('final_raw.xlsx', 'dish_item')['Fried Fish'][0]
    g = p2['summary'].get('action5', {}).get('absolute_guard', {}).get('cells', [])
    check('typing 8.00 into Fried Fish\'s empty Specialty cell is held: the group\'s By Itself $8.00 already '
          'charges it (no Modifier Rule tab, so no preflight would catch the double charge)',
          d['Specialty'] in (None, '') and any(c['dish'] == 'Fried Fish' for c in g), (d, g))


def ms_new_cell_drift(scripts, work):
    print('\naction5: a new Modifier Summary price never overwrites a price wsheet added since Part 1')
    fresh(work, 'ms_new_drift')
    ce.build_r99906()
    ce.part1(scripts)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    ms = wb['Modifier Summary']
    mcol, _r, _x = ce.report_rows(ms)
    mrow = {(ms.cell(r, 1).value, ms.cell(r, 2).value): r for r in range(2, ms.max_row + 1)}
    ms.cell(mrow[('Extras', 'Extra Sauce')], mcol['Large']).value = 0.75      # empty in the report
    ms.cell(mrow[('Extras', 'Extra Shrimp')], mcol['Order']).value = 0.3       # a tenth of its 3.00 / 4.00
    di = wb['Dish Item']
    col, _ref, rows = ce.report_rows(di)
    di.cell(rows[('Hot and Sour Soup', 'Soup')], col['Lunch']).value = 0.3     # a tenth of its 3.75 / 5.75
    wb.save('edited.xlsx')
    ce.set_live('raw.xlsx', 'modifier_item', 'Extra Sauce', 'Large', '$0.60')  # wsheet added it since
    p2 = ce.part2(scripts)
    live_ = ce.live('final_raw.xlsx', 'modifier_item')['Extra Sauce'][0]
    drift = p2['summary'].get('action5', {}).get('new_cell_drift', [])
    ign = {(c['name'], c['variant']) for c in p2['summary'].get('action5', {}).get('ignored_input', {}).get('cells', [])}
    check('a new price at a tenth of the item\'s other prices is held — Dish Item (Hot and Sour Soup Lunch 0.30) '
          'and Modifier Summary (Extra Shrimp Order 0.30)',
          ce.live('final_raw.xlsx', 'dish_item')['Hot and Sour Soup'][0]['Lunch'] in (None, '')
          and ce.live('final_raw.xlsx', 'modifier_item')['Extra Shrimp'][0]['Order'] in (None, '')
          and ('Hot and Sour Soup', 'Lunch') in ign and ('Extra Shrimp', 'Order') in ign, ign)
    check('Extra Sauce Large stays the live $0.60 and is reported (new_cell_drift, "live now has $0.60")',
          live_['Large'] == '$0.60' and any(c['dish'] == 'Extra Sauce (modifier)' and c.get('live') == '$0.60'
                                           for c in drift), (live_, drift))


def misread_and_typos(scripts, work):
    print('\nA misread paper price is corrected, never shipped; a slipped decimal is caught either way')
    for case in ('corrected', 'untouched'):
        fresh(work, 'misread_' + case)
        ce.build_r99906()
        o = json.load(open('ocr.json'))
        for sec in o['sections']:
            for it in sec['items']:
                if it['name'] == 'Egg Roll':
                    it['prices'] = {'Order': 19.50}          # the OCR read 19.50 for 1.95 (system $1.75)
        json.dump(o, open('ocr.json', 'w'))
        ce.part1(scripts)
        shutil.copy('report.xlsx', 'edited.xlsx')
        wb = load_workbook('edited.xlsx')
        di = wb['Dish Item']
        col, _ref, rows = ce.report_rows(di)
        egg = di.cell(rows[('Egg Roll', 'Appetizers')], col['Order'])
        if case == 'corrected':
            egg.value = 1.95
            di.cell(rows[('Sesame Chicken', 'Chicken')], col['Small']).value = 0.95   # 9.50 on the paper, $9.25 now
        wb.save('edited.xlsx')
        shown = cf_kind(wb, di, egg.coordinate, ce.rule_kind)
        p2 = ce.part2(scripts)
        d = {k: v[0] for k, v in ce.live('final_raw.xlsx', 'dish_item').items()}
        ign = {(c['name'], c['variant']): c['why'] for c in p2['summary'].get('action5', {}).get('ignored_input', {}).get('cells', [])}
        if case == 'corrected':
            sc = cf_kind(wb, di, di.cell(rows[('Sesame Chicken', 'Chicken')], col['Small']).coordinate, ce.rule_kind)
            check('the specialist corrects the misread to 1.95: the sheet shows an edit (not red) and $1.95 is written',
                  shown == 'EDIT' and d['Egg Roll']['Order'] == '$1.95', (shown, d['Egg Roll']))
            check('0.95 typed for the new 9.50 (system $9.25) is red in the sheet and held — a fifth of the system price',
                  sc == 'INVALID' and d['Sesame Chicken']['Small'] == '$9.25'
                  and 'misplaced' in ign.get(('Sesame Chicken', 'Small'), ''), (sc, d['Sesame Chicken'], ign))
        else:
            check('left untouched, the misread 19.50 (11× the system $1.75) is red in the sheet and never shipped',
                  shown == 'INVALID' and d['Egg Roll']['Order'] == '$1.75'
                  and 'misread' in ign.get(('Egg Roll', 'Order'), ''), (shown, d['Egg Roll'], ign.get(('Egg Roll', 'Order'))))

    # the same on the Modifier Rule tab: compare took "w. French Fries 93.90" (for 9.39; live $8.90)
    for case in ('corrected', 'untouched'):
        fresh(work, 'mr_misread_' + case)
        ce.build_r99907(side=True)
        o = json.load(open('ocr.json'))
        o['sections'][0]['items'][0]['side_prices'] = {'w. French Fries': 93.9}
        json.dump(o, open('ocr.json', 'w'))
        ce.part1(scripts, compare=True)
        shutil.copy('report.xlsx', 'edited.xlsx')
        wb = load_workbook('edited.xlsx')
        ws = wb['Modifier Rule']
        _c, mr = ce.mrv(ws)
        cell = ws.cell(mr[('Chicken Wings', 'Specialty', 'French Fries')], 2)
        if case == 'corrected':
            cell.value = 9.39
        wb.save('edited.xlsx')
        ce.part2(scripts, with_rules=True)
        got = ce.rule_cell('final_raw.xlsx', '', 'Chicken Wings', 'Specialty', 'French Fries')
        check(f'Modifier Rule, misread {case}: ' + ('the correction 9.39 is written' if case == 'corrected'
                                                    else 'compare\'s 93.90 (10× the live $8.90) is held'),
              got == ('$9.39' if case == 'corrected' else '$8.90'), got)


def negative_format(scripts, work):
    print('\nNegative modifier prices are written the way wsheet holds them (-$X.XX)')
    fresh(work, 'neg_format')
    variants = ['Order', 'Specialty']
    dishes = [('Chicken Wings', 'Fried Specialties', {'Specialty': 7.99}),
              ('Whole Wings', 'Fried Specialties', {'Specialty': 9.50})]
    make_raw('raw.xlsx', 'R99921', 'Neg Wok', variants, dishes,
             {S: (0, ['French Fries', 'Less Whole Wing'], ['Fried Specialties'])},
             {'French Fries': {'Specialty': 0.0}, 'Less Whole Wing': {'Order': -1.70}},
             [('', 'Chicken Wings', 'Specialty', S, 'French Fries', 1.40),
              ('', 'Whole Wings', 'Specialty', S, 'Less Whole Wing', -2.25)])
    ce.set_live('raw.xlsx', 'modifier_item', 'Less Whole Wing', 'Order', '-$1.70')     # as real exports
    ce.rule_cell('raw.xlsx', '', 'Whole Wings', 'Specialty', 'Less Whole Wing', '-$2.25')
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Neg Wok'}, 'sections': [
            {'menu_section': 'Fried Specialties', 'price_columns': ['Specialty'], 'default_price': None,
             'side_pricing': {'columns': ['w. French Fries']},
             'items': [{'name': 'Chicken Wings', 'prices': {'Specialty': 7.99}, 'side_prices': {'w. French Fries': 9.39}},
                       {'name': 'Whole Wings', 'prices': {'Specialty': 9.50}}]}]}, f)
    ce.part1(scripts, compare=True)
    shutil.copy('report.xlsx', 'edited.xlsx')
    wb = load_workbook('edited.xlsx')
    ms = wb['Modifier Summary']
    mcol, _r, _x = ce.report_rows(ms)
    mrow = {(ms.cell(r, 1).value, ms.cell(r, 2).value): r for r in range(2, ms.max_row + 1)}
    ms.cell(mrow[(S, 'Less Whole Wing')], mcol['Order']).value = -2.00
    ws = wb['Modifier Rule']
    _c, mr = ce.mrv(ws)
    ws.cell(mr[('Whole Wings', 'Specialty', 'Less Whole Wing')], 2).value = -2.50
    wb.save('edited.xlsx')
    ce.part2(scripts, with_rules=True)
    mi = ce.live('final_raw.xlsx', 'modifier_item')['Less Whole Wing'][0]['Order']
    rule = ce.rule_cell('final_raw.xlsx', '', 'Whole Wings', 'Specialty', 'Less Whole Wing')
    check('Less Whole Wing −2.00 / rule −2.50 are written "-$2.00" / "-$2.50" (not "-2.00" / "$-2.50")',
          (mi, rule) == ('-$2.00', '-$2.50'), (mi, rule))


def ocr_text_prices(scripts, work):
    print('\nA side / plain price written as text in the OCR JSON ("$9.59", "MP") never crashes Part 1')
    fresh(work, 'ocr_text')
    ce.build_r99907(side=True)
    o = json.load(open('ocr.json'))
    items = o['sections'][0]['items']
    items[0]['side_prices'] = {'w. French Fries': '$9.39'}       # a number as text
    items[1]['side_prices'] = {'w. French Fries': 'MP'}          # no number at all
    items[2]['prices'] = {'Specialty': '6.50'}
    json.dump(o, open('ocr.json', 'w'))
    out = ce.part1(scripts, compare=True)
    rcs = {k: v[0] for k, v in out.items()}
    check('action2 / action3 / compare all run (exit 0), "$9.39" is read as 9.39 and "MP" is reported',
          all(v == 0 for v in rcs.values()) and 'NOT A PRICE' in out['a2'][1] + out['cmp'][1]
          and ce.rule_cell('raw.xlsx', '', 'Chicken Wings', 'Specialty', 'French Fries') == '$8.90'
          and num(load_workbook('report.xlsx')['Modifier Rule'].cell(
              ce.mrv(load_workbook('report.xlsx')['Modifier Rule'])[1][('Chicken Wings', 'Specialty', 'French Fries')], 2).value) == 9.39,
          (rcs, [l for l in (out['a2'][1] + out['cmp'][1]).splitlines() if 'NOT A PRICE' in l][:2]))


def compare_no_rule_for_side(scripts, work):
    print('\ncompare: a printed side the dish has no rule for is a callback item (never dropped)')
    fresh(work, 'compare_norule_side')
    variants = ['Order', 'Specialty']
    dishes = [('Chicken Wings', 'Fried Specialties', {'Specialty': 7.99})]
    make_raw('raw.xlsx', 'R99922', 'Partial Wok', variants, dishes,
             {S: (0, ['French Fries', 'Pork Fried Rice'], ['Fried Specialties'])},
             {'French Fries': {'Order': 0.0}, 'Pork Fried Rice': {'Order': 0.0}},
             [('', 'Chicken Wings', 'Specialty', S, 'French Fries', 1.40),
              ('Fried Specialties', '', 'Specialty', S, 'Pork Fried Rice', 1.90)])     # group-level side
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Partial Wok'}, 'sections': [
            {'menu_section': 'Fried Specialties', 'price_columns': ['Specialty'], 'default_price': None,
             'side_pricing': {'columns': ['w. French Fries', 'w. Pork Fried Rice', 'w. Lo Mein']},
             'items': [{'name': 'Chicken Wings', 'prices': {'Specialty': 7.99},
                        'side_prices': {'w. French Fries': 9.59, 'w. Pork Fried Rice': 10.39, 'w. Lo Mein': 10.89}}]}]}, f)
    ce.part1(scripts, compare=True)
    _rc, out = run(scripts, 'modifier_rule_pricing.py', 'compare', 'ocr.json', 'raw.xlsx', '--report', 'report.xlsx')
    conflicts = out.split('Conflicts NOT auto-resolved', 1)[1].split('Notes:', 1)[0] if 'Conflicts NOT' in out else ''
    check('"w. Pork Fried Rice" (priced only by a group-level rule) and "w. Lo Mein" (no rule) are both flagged '
          'NO RULE FOR SIDE', "'w. Pork Fried Rice'" in conflicts and 'group-level rule on Fried Specialties' in conflicts
          and "'w. Lo Mein'" in conflicts, conflicts[-500:])


def psize_absolute(scripts, work):
    print("\nP-SIZE: a paper price that is the By Itself price of the dish's $0.00 size is not asked")
    fresh(work, 'psize_abs')
    variants = ['Order', 'Specialty', 'Small', 'Large']
    dishes = [('Rib Tips', 'Fried Specialties', {'Specialty': 0.00, 'Large': 11.95}),     # China Village shape
              ('Pork Ribs', 'Fried Specialties', {'Specialty': 0.00, 'Large': 12.95})]
    make_raw('raw.xlsx', 'R99923', 'Rib Wok', variants, dishes,
             {S: (0, ['By Itself', 'French Fries'], [])}, {'By Itself': {'Order': 0.0}, 'French Fries': {'Order': 0.0}},
             [('', 'Rib Tips', 'Specialty', S, 'By Itself', 7.55), ('', 'Rib Tips', 'Specialty', S, 'French Fries', 9.55),
              ('', 'Pork Ribs', 'Specialty', S, 'By Itself', 8.25), ('', 'Pork Ribs', 'Specialty', S, 'French Fries', 10.25)])
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Rib Wok'}, 'sections': [
            {'menu_section': 'Appetizer', 'price_columns': [], 'default_price': None,
             'items': [{'name': 'Rib Tips', 'prices': {'Small': 7.55, 'Large': 11.95}},       # Small = By Itself
                       {'name': 'Pork Ribs', 'prices': {'Small': 8.95, 'Large': 12.95}}]}]}, f)  # 8.95 ≠ 8.25
    ce.part1(scripts)
    hdr, q = questionnaire_rows()
    rows = {str(r[3]): str(r[hdr.index('Details')]) for r in q if str(r[0] or '').startswith('P-SIZE-')}
    pr = rows.get('Pork Ribs [Small]', '')
    check('Rib Tips "Small 7.55" (= its Specialty By Itself $7.55) is not asked; Pork Ribs "Small 8.95" is, with its '
          'sizes and the By Itself it probably is', 'Rib Tips [Small]' not in rows and 'its sizes: ' in pr
          and 'By Itself rule $8.25' in pr, rows)


def tray_per_group(scripts, work):
    print('\nT-TRAY per dish group (maintainer, 2026-10-05): each group its own formula')
    fresh(work, 'tray_groups')
    C = lambda n, l: (n, 'Chicken', {'Large': l, 'Party Tray': round(l * 5 - 2, 2)})        # × 5 − $2.00
    F = lambda n, l: (n, 'Seafood', {'Large': l, 'Party Tray': round(l * 4 + 5.5, 2)})      # × 4 + $5.50
    dishes = [C('Chicken A', 10.00), C('Chicken B', 11.00), C('Chicken C', 12.00), C('Chicken D', 13.00),
              F('Seafood A', 14.00), F('Seafood B', 15.00), F('Seafood C', 16.00),
              C('Combo X', 7.50)]                      # 35.50 — fits BOTH formulas at its old price
    paper = [('Chicken', 'Chicken A', {'Large': 10.50}), ('Seafood', 'Seafood A', {'Large': 14.50}),
             ('Chicken', 'Combo X', {'Large': 8.00})]
    recs, cells, js, _o = pi_case(scripts, 'Group Tray Wok', dishes, paper, also_in=[('Seafood', 'Combo X')])
    check('Chicken A takes the Chicken formula (10.50 × 5 − 2 = 50.50), Seafood A the Seafood one '
          '(14.50 × 4 + 5.50 = 63.50) — no single menu formula would fit both',
          cells.get(('Chicken', 'Chicken A', 'Party Tray')) == (50.5, True)
          and cells.get(('Seafood', 'Seafood A', 'Party Tray')) == (63.5, True),
          (cells.get(('Chicken', 'Chicken A', 'Party Tray')), cells.get(('Seafood', 'Seafood A', 'Party Tray'))))
    cx = [r for k, r in recs.items() if k[1] == 'Combo X' and k[2] == 'Party Tray']
    check('Combo X, in both groups, would get 38.00 (Chicken) or 37.50 (Seafood) — not derived, asked',
          len(cx) == 2 and all(r['inferred'] is None and 'SEVERAL GROUPS' in r['reason'] for r in cx)
          and cells.get(('Chicken', 'Combo X', 'Party Tray'), (0, True))[1] is False, [(r['group'], r['reason'][:80]) for r in cx])


def tray_small_group(scripts, work):
    print('\nT-TRAY: a group too small for its own formula uses the whole-menu one only if its trays fit it')
    fresh(work, 'tray_small')
    T = lambda n, g, l, t=None: (n, g, {'Large': l, 'Party Tray': t if t is not None else round(l * 5 - 2, 2)})
    dishes = [T('Chicken A', 'Chicken', 10.00), T('Chicken B', 'Chicken', 11.00), T('Chicken C', 'Chicken', 12.00),
              T('Chicken D', 'Chicken', 13.00), T('Chicken E', 'Chicken', 14.00),
              T('Beef A', 'Beef', 15.00), T('Beef B', 'Beef', 16.00),                    # 2 trays, both × 5 − 2
              T('Pork A', 'Pork', 12.50), T('Pork B', 'Pork', 13.50, 70.00)]             # one fits, one does not
    paper = [('Chicken', 'Chicken A', {'Large': 10.50}), ('Beef', 'Beef A', {'Large': 15.50}),
             ('Pork', 'Pork A', {'Large': 13.00})]
    recs, cells, js, _o = pi_case(scripts, 'Small Group Wok', dishes, paper)
    pork = recs.get(('Pork', 'Pork A', 'Party Tray'), {})
    check('Beef (2 trays, both fit the whole-menu × 5 − 2) uses it: Beef A 15.50 × 5 − 2 = 75.50 purple',
          cells.get(('Beef', 'Beef A', 'Party Tray')) == (75.5, True), cells.get(('Beef', 'Beef A', 'Party Tray')))
    check('Pork (1 of its 2 trays fits) has no formula: Pork A tray not derived, asked on the callback',
          cells.get(('Pork', 'Pork A', 'Party Tray')) == (60.5, False) and pork.get('ask'), (cells.get(('Pork', 'Pork A', 'Party Tray')), pork.get('reason')))
    hdr, q = questionnaire_rows()
    p0 = {str(r[2]): str(r[hdr.index('Details')]) for r in q if r[0] == 'P-TRAY-0'}
    check('P-TRAY-0 has one row per group: Beef says it uses the whole-menu formula, Pork why it has none',
          'whole-menu formula' in p0.get('Beef', '') and 'No reliable tray formula for Pork' in p0.get('Pork', '')
          and 'Large × 5' in p0.get('Chicken', ''), p0)


def tray_two_groups(scripts, work):
    print('\nT-TRAY: a dish in two groups gets a derived tray only when every group derives the same one')
    fresh(work, 'tray_two')
    C = lambda n, l, g='Chicken': (n, g, {'Large': l, 'Party Tray': round(l * 5 - 2, 2)})      # × 5 − $2.00
    S = lambda n, l: (n, 'Seafood', {'Large': l, 'Party Tray': round(l * 4 + 5.5, 2)})         # × 4 + $5.50
    dishes = [C('Chicken A', 10.00), C('Chicken B', 11.00), C('Chicken C', 12.00), C('Chicken D', 13.00),
              S('Seafood A', 14.00), S('Seafood B', 15.00), S('Seafood C', 16.00),
              C('Beef A', 15.00, 'Beef'), C('Beef B', 16.00, 'Beef'), C('Beef C', 17.00, 'Beef'),
              ('Pork A', 'Pork', {'Large': 12.00, 'Party Tray': 60.00}),               # no formula fits Pork
              ('Pork B', 'Pork', {'Large': 13.00, 'Party Tray': 71.00}),
              ('Pork C', 'Pork', {'Large': 14.00, 'Party Tray': 59.00}),
              C('Combo X', 10.00),    # 48.00: Chicken's formula; Seafood's would say 45.50 — an outlier there
              C('Combo Y', 11.50),    # Pork has no formula
              C('Combo W', 9.00),     # Beef's formula is Chicken's: both derive the same tray
              C('Side A', 6.00, 'Side Orders'), C('Side B', 7.00, 'Side Orders'), C('Side C', 8.00, 'Side Orders'),
              C('Combo V', 8.50)]     # also a Side Order: the same tray, held back there (operational group)
    paper = [('Chicken', 'Combo X', {'Large': 10.50}), ('Chicken', 'Combo Y', {'Large': 12.00}),
             ('Chicken', 'Combo W', {'Large': 9.50}), ('Chicken', 'Combo V', {'Large': 9.00})]
    recs, cells, _js, _o = pi_case(scripts, 'Two Group Wok', dishes, paper,
                                   also_in=[('Seafood', 'Combo X'), ('Pork', 'Combo Y'), ('Beef', 'Combo W'),
                                            ('Side Orders', 'Combo V')])
    for dish, other, why in (('Combo X', 'Seafood', "does not follow this group's formula"),
                             ('Combo Y', 'Pork', 'no reliable formula')):
        rows = [recs.get((g, dish, 'Party Tray'), {}) for g in ('Chicken', other)]
        check(f'{dish}: Chicken derives a tray but {other} does not ({why}) — derived on neither row, asked',
              all(r.get('inferred') is None and r.get('ask') and why in r.get('reason', '') for r in rows)
              and not any(cells.get((g, dish, 'Party Tray'), (0, True))[1] for g in ('Chicken', other)),
              [(r.get('group'), r.get('status'), r.get('reason', '')[:90]) for r in rows])
    check('Combo W: Chicken and Beef both derive 9.50 × 5 − 2 = 45.50 — purple on both rows',
          cells.get(('Chicken', 'Combo W', 'Party Tray')) == (45.5, True)
          and cells.get(('Beef', 'Combo W', 'Party Tray')) == (45.5, True),
          (cells.get(('Chicken', 'Combo W', 'Party Tray')), cells.get(('Beef', 'Combo W', 'Party Tray'))))
    cv = [recs.get((g, 'Combo V', 'Party Tray'), {}) for g in ('Chicken', 'Side Orders')]
    check('Combo V: both groups say 43.00, but its Side Orders row is held back (suggest-only) — so is its '
          'Chicken row: purple on neither', all(r.get('status') == 'SUGGESTED' and r.get('inferred') == 43.0 for r in cv)
          and not any(cells.get((g, 'Combo V', 'Party Tray'), (0, True))[1] for g in ('Chicken', 'Side Orders')),
          [(r.get('group'), r.get('status'), r.get('reason', '')[:90]) for r in cv])
    # both groups' formulas fit the old tray, but on the new paper one says "no change"
    fresh(work, 'tray_two_b')
    C2 = lambda n, s, l: (n, 'Chicken', {'Small': s, 'Large': l, 'Party Tray': round(l * 5 - 2, 2)})
    S2 = lambda n, s: (n, 'Seafood', {'Small': s, 'Party Tray': round(s * 6, 2)})               # Small × 6
    dishes = [C2('Chicken A', 7.00, 10.00), C2('Chicken B', 7.50, 11.00), C2('Chicken C', 8.25, 12.00),
              S2('Seafood A', 9.00), S2('Seafood B', 9.50), S2('Seafood C', 10.00),
              ('Combo Z', 'Chicken', {'Small': 8.00, 'Large': 10.00, 'Party Tray': 48.00})]   # 10 × 5 − 2 = 8 × 6
    recs, cells, _js, _o = pi_case(scripts, 'Two Group Wok B', dishes,
                                   [('Chicken', 'Combo Z', {'Small': 8.00, 'Large': 10.50})],
                                   also_in=[('Seafood', 'Combo Z')])
    rows = [recs.get((g, 'Combo Z', 'Party Tray'), {}) for g in ('Chicken', 'Seafood')]
    check('Combo Z: Chicken says 50.50 (Large went up), Seafood says the tray stays (Small did not) — not '
          'derived, asked', all(r.get('inferred') is None and r.get('ask') and 'tray stays' in r.get('reason', '')
                                for r in rows)
          and cells.get(('Chicken', 'Combo Z', 'Party Tray')) == (48.0, False),
          [(r.get('group'), r.get('status'), r.get('reason', '')[:90]) for r in rows])


def paper_two_prices(scripts, work):
    print('\nTwo different prices printed for one size: neither written, asked (P-PRICE; maintainer 2026-10-05)')
    fresh(work, 'two_prices')
    T = lambda n, l: (n, 'Beef', {'Large': l, 'Party Tray': round(l * 5 - 2, 2)})           # Beef trays: × 5 − 2
    dishes = [('Sesame Chicken', 'Chicken', {'Small': 9.25, 'Large': 13.25}),
              ('Kung Pao Chicken', 'Chicken', {'Small': 9.25, 'Large': 13.25}),
              ('Mongolian Beef', 'Beef', {'Small': 9.75, 'Large': 13.75}),
              T('Pepper Steak', 13.00), T('Beef A', 12.00), T('Beef B', 14.00), T('Beef C', 15.00),
              ('General Tso Chicken', 'Chicken', {'Lunch': 8.95, 'Dinner': 11.95}),   # Lunch-based menu:
              ('Hunan Chicken', 'Chicken', {'Lunch': 8.95, 'Dinner': 11.95}),         # paper Small → Lunch
              ('Garlic Chicken', 'Chicken', {'Lunch': 8.95, 'Dinner': 11.95}),
              ('Curry Chicken', 'Chicken', {'Lunch': 8.95, 'Dinner': 11.95})]
    make_raw('raw.xlsx', 'R99924', 'Twice Wok', ['Order', 'Small', 'Large', 'Lunch', 'Dinner', 'Party Tray'],
             dishes, {}, {}, [])
    sec = lambda name, items: {'menu_section': name, 'price_columns': [], 'default_price': None, 'items': items}
    with open('ocr.json', 'w') as f:
        json.dump({'restaurant': {'name': 'Twice Wok'}, 'sections': [
            sec('Chicken', [{'name': 'Sesame Chicken', 'prices': {'Small': 9.50, 'Large': 13.45}},
                            {'name': 'Kung Pao Chicken', 'prices': {'Large': 13.45}},
                            {'name': 'General Tso Chicken', 'prices': {'Small': 9.25}}]),
            sec("Chef's Specials", [{'name': 'Sesame Chicken', 'prices': {'Large': 13.50}},
                                    {'name': 'Kung Pao Chicken', 'prices': {'Large': 13.45}},   # same price twice
                                    {'name': 'Mongolian Beef', 'prices': {'Large': 14.25}},
                                    {'name': 'Beef C', 'prices': {'Large': 15.75}}]),
            sec('Beef', [{'name': 'Mongolian Beef', 'prices': {'Order': 14.00}},               # Order → Large
                         {'name': 'Pepper Steak', 'prices': {'Large': 13.75, 'Party Tray': 65.50}},
                         {'name': 'Beef B', 'prices': {'Large': 14.50}},
                         {'name': 'Beef C', 'prices': {'Large': 15.50}}]),
            sec('Party Trays', [{'name': 'Pepper Steak', 'prices': {'Party Tray': 66.00}}])],
            'lunch_specials': {'items': [{'name': 'General Tso Chicken', 'price': 9.50}]}}, f)
    out = ce.part1(scripts)
    check('Part 1 runs', all(v[0] == 0 for v in out.values()), {k: v[1][-300:] for k, v in out.items() if v[0]})
    di = load_workbook('report.xlsx')['Dish Item']
    col, _ref, rows = ce.report_rows(di)

    def cell(dish, grp, size):
        c = di.cell(rows[(dish, grp)], col[size])
        return num(c.value), (c.fill.fgColor.rgb[-6:] if c.fill and c.fill.fill_type else '')
    held = {k: cell(*k) for k in (('Sesame Chicken', 'Chicken', 'Large'), ('Mongolian Beef', 'Beef', 'Large'),
                                  ('General Tso Chicken', 'Chicken', 'Lunch'), ('Pepper Steak', 'Beef', 'Party Tray'),
                                  ('Beef C', 'Beef', 'Large'))}
    check('a size printed twice at different prices keeps the system price, not coloured, not inferred: Sesame '
          'Chicken Large (13.45 / 13.50 in two sections), Mongolian Beef Large (14.00 as Order, 14.25), General Tso '
          'Lunch (9.25 as Small, lunch special 9.50), Pepper Steak tray (65.50 / 66.00), Beef C Large',
          held == {('Sesame Chicken', 'Chicken', 'Large'): (13.25, ''), ('Mongolian Beef', 'Beef', 'Large'): (13.75, ''),
                   ('General Tso Chicken', 'Chicken', 'Lunch'): (8.95, ''), ('Pepper Steak', 'Beef', 'Party Tray'): (63.0, ''),
                   ('Beef C', 'Beef', 'Large'): (15.0, '')}, held)
    check('the same price printed twice is no conflict (Kung Pao Chicken Large 13.45, blue); the other printed '
          'sizes are compared as usual (Sesame Chicken Small 9.50, Pepper Steak Large 13.75)',
          cell('Kung Pao Chicken', 'Chicken', 'Large') == (13.45, 'ADD8E6')
          and cell('Sesame Chicken', 'Chicken', 'Small') == (9.5, 'ADD8E6')
          and cell('Pepper Steak', 'Beef', 'Large') == (13.75, 'ADD8E6'))
    hdr, q = questionnaire_rows()
    pp = {str(r[3]): str(r[hdr.index('Details')]) for r in q if str(r[0] or '').startswith('P-PRICE-')}
    check('each is asked once in P-PRICE, naming both prices and where they were printed',
          sorted(pp) == sorted(['Sesame Chicken [Large]', 'Mongolian Beef [Large]', 'General Tso Chicken [Lunch]',
                                'Pepper Steak [Party Tray]', 'Beef C [Large]'])
          and '$13.45' in pp.get('Sesame Chicken [Large]', '') and "Chef's Specials" in pp.get('Sesame Chicken [Large]', '')
          and 'Lunch Specials' in pp.get('General Tso Chicken [Lunch]', ''), pp)
    js = json.load(open('price_inference_summary.json'))
    recs = {(r['group'], r['dish'], r['variant']): r for r in js['records']}
    bc = recs.get(('Beef', 'Beef C', 'Party Tray'), {})
    check('a tray whose base was printed twice is not derived, and asked (Beef C); one with a clear base still is '
          '(Beef B 14.50 × 5 − 2 = 70.50 purple)', bc.get('inferred') is None and bc.get('ask')
          and 'two different Large prices' in bc.get('reason', '') and cell('Beef B', 'Beef', 'Party Tray') == (70.5, 'E4DFEC'),
          (bc.get('reason'), cell('Beef B', 'Beef', 'Party Tray')))


EXTRA = [guard_addon, guard_required_choice, guard_platter, g11_group_anchor, compare_flags, compare_variantless, side_rules_warning,
         pi_qualifier, pi_two_groups, pi_ordering, pi_rerun_and_screens, pi_tray_dupes, pi_healthy,
         pi_names_and_labels, pi_stale_report, pi_caps_t2_labels, abs_x_held, bridge_stale_tsv, p1_unmatched,
         pabs_notes_fallback, compare_no_side_rules, mrv_absolute_zero_sizeless, apply_g11_full_price,
         new_size_group_anchor, ms_new_cell_drift, misread_and_typos,
         negative_format, ocr_text_prices, compare_no_rule_for_side, psize_absolute,
         tray_per_group, tray_small_group, tray_two_groups, paper_two_prices]

if __name__ == '__main__':
    main()
