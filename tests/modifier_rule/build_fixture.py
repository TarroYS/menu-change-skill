"""Synthetic ChAm restaurants for the Fried Specialty / modifier_rule issues.

  R99901 "Delta Wok"     DELTA model: base on dish_item, rules carry upcharges.
                         Covers #2 (printed price-group conflict), #3 ("Plain
                         Fried Rice" is an upcharge, not a By Itself anchor),
                         #4 (group-level rule edit), #5 (second multi-column
                         section + same dish in two sections), #6 (tab order).
  R99902 "Absolute Wok"  ABSOLUTE model: dish_item $0.00, By Itself rule carries
                         the base, side rules carry the full total. Covers #1,
                         plus the R85325 pattern (a $0.00 dish whose price lives
                         on a required, fully-priced modifier group).

    python3 build_fixture.py        # writes R99901_* and R99902_* into cwd

Raw layout mirrors a real wsheet export (3-row header, JSON root ids in col A,
modifier_rule columns in the same positions as R85325)."""
import json
from openpyxl import Workbook


def m(v):
    return '' if v is None else f'${v:.2f}'


# modifier_rule header, copied from a real export (R85325)
MR_ROW1 = ['id'] + ['rule'] * 24
MR_ROW2 = ['', 'wprop', 'wprop', 'wprop', 'wprop', 'parent', 'parent', 'parent', 'parent', 'parent',
           'parent', 'parent', 'parent', 'parent', 'parent', 'wprop', 'wprop.times[]',
           'wprop.times[].availability', 'wprop.times[].availability', 'wprop.times[].availability',
           'wprop.times[].date_range', 'wprop.times[].date_range', 'wprop.times[]', 'wprop', 'wprop']
MR_ROW3 = ['basic', 'out_is_enabled', 'out_unit_price', 'out_start_price', 'out_is_open_price',
           'restaurant_group', 'restaurant_location', 'menu_item', 'dish_group', 'dish_item', 'variant',
           'modifier_group', 'modifier_item', 'premod_group', 'premod_item', 'start_quantity',
           'day_of_week', 'kind', 'start_time', 'end_time', 'start_date', 'end_date', 'reason',
           'order_sources[]', 'service_sources[]']


def make_raw(path, r_number, name, variants, dishes, mod_groups, mod_items, rules,
             codes=None, also_in=None, mod_twins=None):
    """dishes: [(dish, group, {variant: price})]; mod_groups: {group: (min_sel, [items])};
    mod_items: {item: {variant: price}}; rules: [(group, dish, variant, modgrp, mod, price)].
    Optional, like a real export: codes {dish: {variant: unit_code}} adds the
    unit_code columns; also_in [(group, dish)] lists a dish under a second dish
    group (still ONE dish_item row); mod_twins [item] adds a second, unpriced
    modifier_item root row with the same name."""
    wb = Workbook()
    wb.remove(wb.active)

    def tab(title, row1, row2, row3, data):
        ws = wb.create_sheet(title)
        for r in (row1, row2, row3):
            ws.append(r)
        for r in data:
            ws.append(r)

    tab('restaurant_location', ['id', '', ''], ['', '', ''], ['id', 'r_number', 'address1'],
        [['{"id":1}', r_number, '1 Test St']])
    tab('restaurant_enterprise', ['id', '', ''], ['', '', ''], ['id', 'e_number', 'restaurant_enterprise'],
        [['{"id":1}', 'E' + r_number[1:], name]])

    # dish_group: root row per group, then one row per dish / modifier group
    also_in = also_in or []
    dg_groups = []
    for grp in [g for _d, g, _p in dishes] + [g for g, _d in also_in]:
        if grp not in dg_groups:
            dg_groups.append(grp)
    dg_rows = []
    for gi, grp in enumerate(dg_groups):
        dg_rows.append([f'{{"id":"g{gi}"}}', grp, '', ''])
        for dish, g, _ in dishes:
            if g == grp:
                dg_rows.append(['', grp, dish, ''])
        for g, dish in also_in:
            if g == grp:
                dg_rows.append(['', grp, dish, ''])
        for mg, (_min, _items, groups) in mod_groups.items():
            if grp in groups:
                dg_rows.append(['', grp, '', mg])
    tab('dish_group', ['id', 'from', 'comprise', 'comprise'], ['', '', 'to', 'to'],
        ['id', 'dish_group', 'dish_item', 'modifier_group'], dg_rows)

    nc = len(variants) if codes is not None else 0
    tab('dish_item', ['id', 'from'] + ['variant'] * (len(variants) + nc),
        ['', ''] + ['wprop.unit_price'] * len(variants) + ['wprop.unit_code'] * nc,
        ['id', 'dish_item'] + variants + (variants if nc else []),
        [[f'{{"id":"d{i}"}}', dish] + [m(pr.get(v)) for v in variants]
         + ([(codes.get(dish) or {}).get(v, '') for v in variants] if nc else [])
         for i, (dish, _g, pr) in enumerate(dishes)])

    mg_rows = []
    for gi, (mg, (min_sel, items, _groups)) in enumerate(mod_groups.items()):
        mg_rows.append([f'{{"id":"mg{gi}"}}', mg, '', str(min_sel)])
        for it in items:
            mg_rows.append(['', mg, it, ''])
    tab('modifier_group', ['id', 'from', 'comprise', 'object'], ['', '', 'to', 'wprop'],
        ['id', 'modifier_group', 'modifier_item', 'min_number_of_selections'], mg_rows)

    tab('modifier_item', ['id', 'from'] + ['variant'] * len(variants),
        ['', ''] + ['wprop.unit_price'] * len(variants),
        ['id', 'modifier_item'] + variants,
        [[f'{{"id":"mi{i}"}}', it] + [m(pr.get(v)) for v in variants]
         for i, (it, pr) in enumerate(mod_items.items())]
        + [[f'{{"id":"mt{i}"}}', it] + [''] * len(variants) for i, it in enumerate(mod_twins or [])])

    mr_rows = []
    for i, (grp, dish, var, mg, mod, price) in enumerate(rules):
        row = [''] * len(MR_ROW3)
        row[0] = f'{{"prod":{{"stable_wid":"Wid(0,1,1,0,{i})"}}}}'
        row[1] = 'TRUE'
        row[2] = m(price)
        row[8], row[9], row[10], row[11], row[12] = grp, dish, var, mg, mod
        mr_rows.append(row)
    tab('modifier_rule', MR_ROW1, MR_ROW2, MR_ROW3, mr_rows)
    wb.save(path)


VARIANTS = ['Order', 'Specialty', 'Small', 'Large']
SIDES = ['French Fries', 'Pork Fried Rice', 'Chicken Fried Rice', 'Beef Fried Rice', 'Plain Fried Rice']


def build_delta():
    dishes = [
        ('Chicken Wings',   'Fried Specialties', {'Specialty': 7.99}),
        ('Chicken Fingers', 'Fried Specialties', {'Specialty': 8.50}),
        ('Fried Shrimp',    'Fried Specialties', {'Specialty': 9.50}),
        ('Fried Scallops',  'Fried Specialties', {'Specialty': 10.50}),   # not printed -> twin
        ('Buffalo Wings',   'Wings',             {'Order': 8.50}),        # 2nd multi-col section (#5)
        ('Egg Roll',        'Appetizers',        {'Order': 1.75}),
    ]
    mod_groups = {'Select Side': (0, SIDES, ['Fried Specialties', 'Wings']),
                  'Sauce': (0, ['Extra Sauce'], ['Fried Specialties'])}
    mod_items = {s: {'Order': 0.0} for s in SIDES}
    mod_items['Extra Sauce'] = {'Order': 0.50}
    S = 'Select Side'
    rules = [
        ('', 'Chicken Wings',   'Specialty', S, 'French Fries',       1.40),
        ('', 'Chicken Wings',   'Specialty', S, 'Pork Fried Rice',    1.90),  # #2 printed +2.20
        ('', 'Chicken Wings',   'Specialty', S, 'Chicken Fried Rice', 1.90),  # #2 printed +1.90
        ('', 'Chicken Wings',   'Specialty', S, 'Beef Fried Rice',    1.90),  # #2 unprinted member
        ('', 'Chicken Wings',   'Specialty', S, 'Plain Fried Rice',   1.50),  # #3 upcharge, not anchor
        ('', 'Chicken Fingers', 'Specialty', S, 'French Fries',       1.40),  # #5 sections disagree
        ('', 'Chicken Fingers', 'Specialty', S, 'Pork Fried Rice',    1.90),
        ('', 'Chicken Fingers', 'Specialty', S, 'Chicken Fried Rice', 1.90),
        ('', 'Fried Shrimp',    'Specialty', S, 'French Fries',       1.40),
        ('', 'Fried Shrimp',    'Specialty', S, 'Pork Fried Rice',    1.90),
        ('', 'Fried Scallops',  'Specialty', S, 'French Fries',       1.40),
        ('', 'Buffalo Wings',   'Order',     S, 'French Fries',       1.40),  # #5
        ('', 'Buffalo Wings',   'Order',     S, 'Pork Fried Rice',    1.90),  # #5
        ('Fried Specialties', '', 'Specialty', 'Sauce', 'Extra Sauce', 0.50),  # #4 group-level
    ]
    make_raw('R99901_raw.xlsx', 'R99901', 'Delta Wok', VARIANTS, dishes, mod_groups, mod_items, rules)

    ocr = {'restaurant': {'name': 'Delta Wok', 'r_number': 'R99901'}, 'sections': [
        {'menu_section': 'Specialties', 'price_columns': ['Specialty'], 'default_price': None,
         'side_pricing': {'columns': ['w. French Fries', 'w. Pork Fried Rice',
                                      'w. Chicken Fried Rice', 'w. Fried Rice']},
         'items': [
             {'code': 'S1', 'name': 'Chicken Wings', 'prices': {'Specialty': 7.99},
              'side_prices': {'w. French Fries': 9.59, 'w. Pork Fried Rice': 10.19,
                              'w. Chicken Fried Rice': 9.89, 'w. Fried Rice': 9.69}},
             {'code': 'S2', 'name': 'Chicken Fingers', 'prices': {'Specialty': 8.75},
              'side_prices': {'w. French Fries': 10.35, 'w. Pork Fried Rice': 10.95,
                              'w. Chicken Fried Rice': 10.95, 'w. Fried Rice': None}},
             {'code': 'S3', 'name': 'Fried Shrimp', 'prices': {'Specialty': 9.95},
              'side_prices': {'w. French Fries': 11.55, 'w. Pork Fried Rice': 12.15,
                              'w. Chicken Fried Rice': None, 'w. Fried Rice': None}}]},
        {'menu_section': 'Wings', 'price_columns': ['Order'], 'default_price': None,
         'side_pricing': {'columns': ['with French Fries', 'with Pork Fried Rice']},
         'items': [
             {'code': 'W1', 'name': 'Buffalo Wings', 'prices': {'Order': 8.95},
              'side_prices': {'with French Fries': 10.55, 'with Pork Fried Rice': 11.15}},
             # same dish as S2, fries total disagrees (+1.70 here vs +1.60 there)
             {'code': 'W2', 'name': 'Chicken Fingers', 'prices': {'Specialty': 8.75},
              'side_prices': {'with French Fries': 10.45, 'with Pork Fried Rice': 10.95}}]},
        {'menu_section': 'Appetizers', 'price_columns': ['Order'], 'default_price': None,
         'items': [{'code': 'A1', 'name': 'Egg Roll', 'prices': {'Order': 1.95}}]}]}
    with open('R99901_menu_ocr.json', 'w') as f:
        json.dump(ocr, f, indent=2)


def build_absolute():
    dishes = [
        ('Chicken Wings',       'Fried Specialties', {'Specialty': 0.00}),  # ABSOLUTE, printed
        ('Chicken Fingers',     'Fried Specialties', {'Specialty': 0.00}),  # ABSOLUTE, not printed
        ('Egg Roll Platter',    'Fried Specialties', {'Specialty': 6.00}),  # ordinary dish (donor)
        ('Spring Roll Platter', 'Fried Specialties', {'Specialty': 6.00}),  # ordinary dish (donor)
        ('Seafood Boil',        'Seafood Boil',      {'Order': 0.00}),      # R85325 pattern
        ('Egg Roll',            'Appetizers',        {'Order': 1.75}),
        ('Spring Roll',         'Appetizers',        {'Order': 1.75}),
        ('Crab Rangoon',        'Appetizers',        {'Order': 6.25}),
        ('Free Egg Roll',       'Appetizers',        {'Order': 0.00}),      # genuine freebie
    ]
    mod_groups = {'Select Side': (0, ['By Itself', 'French Fries', 'Pork Fried Rice'], ['Fried Specialties']),
                  'Select Seafood': (1, ['1 lb Shrimp', '1 lb Crawfish'], ['Seafood Boil'])}
    mod_items = {'By Itself': {'Order': 0.0}, 'French Fries': {'Order': 0.0},
                 'Pork Fried Rice': {'Order': 0.0},
                 '1 lb Shrimp': {'Order': 19.00}, '1 lb Crawfish': {'Order': 27.00}}
    S = 'Select Side'
    rules = [
        ('', 'Chicken Wings',   'Specialty', S, 'By Itself',       7.50),
        ('', 'Chicken Wings',   'Specialty', S, 'French Fries',    8.90),   # full total
        ('', 'Chicken Wings',   'Specialty', S, 'Pork Fried Rice', 9.40),   # full total
        ('', 'Chicken Fingers', 'Specialty', S, 'By Itself',       8.00),
        ('', 'Chicken Fingers', 'Specialty', S, 'French Fries',    9.40),
    ]
    make_raw('R99902_raw.xlsx', 'R99902', 'Absolute Wok', VARIANTS, dishes, mod_groups, mod_items, rules)

    ocr = {'restaurant': {'name': 'Absolute Wok', 'r_number': 'R99902'}, 'sections': [
        {'menu_section': 'Specialties', 'price_columns': ['Specialty'], 'default_price': None,
         'side_pricing': {'columns': ['w. French Fries', 'w. Pork Fried Rice']},
         'items': [
             {'code': 'S1', 'name': 'Chicken Wings', 'prices': {'Specialty': 7.99},
              'side_prices': {'w. French Fries': 9.39, 'w. Pork Fried Rice': 9.89}},
             {'code': 'S5', 'name': 'Egg Roll Platter', 'prices': {'Specialty': 6.50}, 'side_prices': None},
             {'code': 'S6', 'name': 'Spring Roll Platter', 'prices': {'Specialty': 6.50}, 'side_prices': None}]},
        {'menu_section': 'Seafood Boil', 'price_columns': ['Order'], 'default_price': None,
         'items': [{'code': 'B1', 'name': 'Seafood Boil', 'prices': {'Order': 21.00}}]},
        # every printed appetizer moves +$0.20 — a uniform group delta (T2)
        {'menu_section': 'Appetizers', 'price_columns': ['Order'], 'default_price': None,
         'items': [{'code': 'A1', 'name': 'Egg Roll', 'prices': {'Order': 1.95}},
                   {'code': 'A2', 'name': 'Spring Roll', 'prices': {'Order': 1.95}},
                   {'code': 'A3', 'name': 'Crab Rangoon', 'prices': {'Order': 6.45}}]}]}
    with open('R99902_menu_ocr.json', 'w') as f:
        json.dump(ocr, f, indent=2)


def build_village():
    """Shape of a real ABSOLUTE restaurant (ER81988 China Village): the model
    differs per dish AND per variant, and a restaurant-wide 'Plain' $0.00 rule
    sits in front of every dish anchor."""
    variants = ['Order', 'Specialty', 'Small', 'Large', '3 Pieces', '4 Pieces', '5 Pieces', '10 Pieces',
                '50 Pieces']
    dishes = [
        ('Chicken Nuggets', 'Fried Specialties', {'Specialty': 0.00}),          # ABSOLUTE
        ('Fried Fish',      'Fried Specialties', {'Specialty': 0.00}),          # ABSOLUTE
        ('Cut Wings',       'Fried Specialties', {'5 Pieces': 0.00,             # ABSOLUTE
                                                  '10 Pieces': 0.00,            # ABSOLUTE, not printed
                                                  '50 Pieces': 50.00}),         # DELTA (party size)
        ('Whole Wing',      'Fried Specialties', {'3 Pieces': 0.00, '4 Pieces': 0.00}),
        ('Bento Box',       'Bento',             {'Order': 13.95}),             # DELTA, variant-less rules
        ('Fried Scallop',   'Fried Specialties', {'Specialty': 0.00}),          # ABSOLUTE, combo box only
        ('Spare Rib Tips',  'Fried Specialties', {'Order': 7.55}),              # DELTA, combo box only
        ('Barbecue Boneless Spare Ribs', 'Fried Specialties', {'Specialty': 0.00}),   # ABSOLUTE
        ('Barbecue Spare Ribs', 'Fried Specialties', {'Small': 8.95, 'Large': 16.95}),  # no rules
    ]
    mod_groups = {'Select Side': (0, ['By Itself', 'French Fries', 'Fried Rice', 'Plain Lo Mein'],
                                  ['Fried Specialties'])}
    mod_items = {m: {'Order': 0.0} for m in ('By Itself', 'French Fries', 'Fried Rice', 'Plain Lo Mein',
                                             'Pork Fried Rice', 'Pork Lo Mein', 'Plain')}
    rules = [
        ('', '', '', '', 'Plain', 0.00),                         # restaurant-wide — not a dish anchor
        ('', '', 'Specialty', '', 'By Itself', 0.00),            # dish-less anchor, also not one
    ]
    for dish, base, fries, rice, lomein in (('Chicken Nuggets', 6.95, 10.95, 10.95, 11.95),
                                            ('Fried Fish', 8.25, 10.95, 10.95, 11.95)):
        rules += [('', dish, 'Specialty', '', 'By Itself', base), ('', dish, 'Specialty', '', 'French Fries', fries),
                  ('', dish, 'Specialty', '', 'Fried Rice', rice), ('', dish, 'Specialty', '', 'Plain Lo Mein', lomein)]
    for var, base, side in (('5 Pieces', 6.25, 9.95), ('10 Pieces', 11.25, 13.95)):
        rules += [('', 'Cut Wings', var, '', 'By Itself', base), ('', 'Cut Wings', var, '', 'French Fries', side),
                  ('', 'Cut Wings', var, '', 'Fried Rice', side)]
    rules += [('', 'Cut Wings', '50 Pieces', '', 'By Itself', 0.00),       # DELTA: base on dish_item
              ('', 'Cut Wings', '50 Pieces', '', 'French Fries', 3.00),    # upcharge
              ('', 'Cut Wings', '50 Pieces', '', 'Fried Rice', 3.00),      # upcharge
              ('', 'Cut Wings', '50 Pieces', 'Wing Sauce', '', 0.50)]      # sauce add-on
    for var, base, side in (('3 Pieces', 7.25, 10.95), ('4 Pieces', 9.25, 12.95)):
        rules += [('', 'Whole Wing', var, '', 'By Itself', base), ('', 'Whole Wing', var, '', 'French Fries', side)]
    rules += [('', 'Bento Box', '', '', 'Pork Fried Rice', 1.75), ('', 'Bento Box', '', '', 'Pork Lo Mein', 2.00)]
    rules += [('', 'Fried Scallop', 'Specialty', '', 'By Itself', 6.95),
              ('', 'Fried Scallop', 'Specialty', '', 'French Fries', 10.95),
              ('', 'Fried Scallop', 'Specialty', '', 'Fried Rice', 10.95),
              ('', 'Fried Scallop', 'Specialty', '', 'Plain Lo Mein', 11.95),
              ('', 'Spare Rib Tips', 'Order', '', 'French Fries', 3.40),    # DELTA upcharge
              ('', 'Barbecue Boneless Spare Ribs', 'Specialty', '', 'By Itself', 7.55),
              ('', 'Barbecue Boneless Spare Ribs', 'Specialty', '', 'French Fries', 10.95)]
    make_raw('R99904_raw.xlsx', 'R99904', 'Village Wok', variants, dishes, mod_groups, mod_items, rules)

    cols = ['w. French Fries', 'w. Fried Rice', 'Lo Mein']
    ocr = {'restaurant': {'name': 'Village Wok', 'r_number': 'R99904'}, 'sections': [
        {'menu_section': 'Fried Specialties', 'price_columns': ['Specialty'], 'default_price': None,
         'side_pricing': {'columns': cols},
         'items': [
             {'code': 'F1', 'name': 'Chicken Nuggets', 'prices': {'Specialty': 7.25},
              'side_prices': {cols[0]: 11.25, cols[1]: 11.25, cols[2]: 12.25}},   # bare "Lo Mein"
             {'code': 'F2', 'name': 'Fried Fish', 'prices': {'Specialty': 8.75},
              'side_prices': {cols[0]: 11.25, cols[1]: 11.25, cols[2]: 12.25}},
             # sized rows: the Plain is keyed by the size, one item per printed size
             {'code': 'W5', 'name': 'Cut Wings', 'prices': {'5 Pieces': 6.75},
              'side_prices': {cols[0]: 10.45, cols[1]: 10.45}},
             {'code': 'W50', 'name': 'Cut Wings', 'prices': {'50 Pieces': 52.00},
              'side_prices': {cols[0]: 55.50, cols[1]: 55.50}},
             # size only in the name
             {'code': 'H4', 'name': '4pcs Whole Wing', 'prices': {'Specialty': 9.75},
              'side_prices': {cols[0]: 13.45, cols[1]: None}},
             # no size at all — cannot tell 3 from 4 pieces
             {'code': 'H?', 'name': 'Whole Wing', 'prices': {'Specialty': 7.75},
              'side_prices': {cols[0]: 11.45, cols[1]: None}}]},
        {'menu_section': 'Bento', 'price_columns': ['Order'], 'default_price': None,
         'items': [{'code': 'B1', 'name': 'Bento Box', 'prices': {'Order': 14.50}}]},
        # China Village layout: own prices in APPETIZER, the with-side price in a combo box
        {'menu_section': 'APPETIZER', 'price_columns': ['Order'], 'default_price': None,
         'items': [{'code': 'A1', 'name': 'Fried Scallops', 'prices': {'Order': 7.25}},
                   {'code': 'A2', 'name': 'Spare Rib Tips', 'prices': {'Small': 7.55, 'Large': 11.95}}]},
        {'menu_section': 'Appetizer Combos', 'price_columns': [], 'default_price': 11.25,
         'note': 'All served w. french fries or fried rice',
         'side_pricing': {'columns': ['w. french fries or fried rice']},
         'items': [{'code': 'C1', 'name': 'Fried Scallops', 'prices': {},
                    'side_prices': {'w. french fries or fried rice': 11.25}},
                   {'code': 'C2', 'name': 'Spare Rib Tips', 'prices': {},
                    'side_prices': {'w. french fries or fried rice': 11.45}},
                   {'code': 'C3', 'name': 'Boneless Spare Ribs', 'prices': {},
                    'side_prices': {'w. french fries or fried rice': 11.45}},
                   # a different dish (no modifier rules) — must not land on Boneless
                   {'code': 'C4', 'name': 'BBQ Spare Ribs', 'prices': {},
                    'side_prices': {'w. french fries or fried rice': 12.45}}]}]}
    with open('R99904_menu_ocr.json', 'w') as f:
        json.dump(ocr, f, indent=2)


if __name__ == '__main__':
    build_delta()
    build_absolute()
    build_village()
    print('wrote R99901_*, R99902_*, R99904_* (raw.xlsx + menu_ocr.json)')
