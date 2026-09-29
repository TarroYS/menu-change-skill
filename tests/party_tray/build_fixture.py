"""Synthetic ChAm restaurant for Party Tray testing.
Tray formula baked into the system data: Party Tray = Large x 5 - $2.00
(except Beef with Broccoli, a deliberate outlier)."""
import json, sys
from openpyxl import Workbook

MIXED = len(sys.argv) > 1 and sys.argv[1] == 'mixed'   # negative fixture: no dominant ratio

# name, group, Small, Large, Tray
DISHES = [
    ("General Tso's Chicken", 'Chicken', 8.95, 12.95, 62.75),
    ('Sesame Chicken',        'Chicken', 8.95, 12.95, 62.75),
    ('Chicken with Broccoli', 'Chicken', 8.25, 11.95, 57.75),
    ('Kung Pao Chicken',      'Chicken', 8.25, 11.95, 57.75),
    ('Hunan Chicken',         'Chicken', 8.95, 12.95, 62.75),
    ('Lemon Chicken',         'Chicken', 8.50, 12.50, 60.50),
    ('Beef with Broccoli',    'Beef',    9.95, 14.95, 80.00),   # outlier (formula says 72.75)
    ('Pepper Steak',          'Beef',    9.95, 14.95, 72.75),
    ('Mongolian Beef',        'Beef',   10.50, 15.50, 75.50),
    ('Fried Rice Tray',       'Party Trays', None, None, 45.00),  # tray-only
    ('Lo Mein Tray',          'Party Trays', None, None, 45.00),  # tray-only
]
if MIXED:  # every tray on a different ratio -> no dominant formula
    trays = [62.75, 55.00, 57.75, 49.00, 70.00, 60.50, 80.00, 66.00, 58.00]
    DISHES = [(n, g, s, l, trays[i] if i < 9 else t) for i, (n, g, s, l, t) in enumerate(DISHES)]

def m(v): return '' if v is None else f'${v:.2f}'

wb = Workbook(); wb.remove(wb.active)
def tab(name, row2, row3, data):
    ws = wb.create_sheet(name)
    ws.append([name] + [''] * (len(row3) - 1)); ws.append(row2); ws.append(row3)
    for r in data: ws.append(r)

tab('restaurant_location', ['', '', ''], ['id', 'r_number', 'address1'], [['{"id":1}', 'R99999', '1 Test St']])
tab('restaurant_enterprise', ['', '', ''], ['id', 'e_number', 'restaurant_enterprise'], [['{"id":1}', 'E99999', 'Tray Test Kitchen']])

dg, cur = [], None
for i, (n, g, *_ ) in enumerate(DISHES):
    if g != cur:
        dg.append([f'{{"id":"g{i}"}}', g, n, 'Add Ons' if g != 'Party Trays' else '']); cur = g
    else:
        dg.append(['', '', n, ''])
tab('dish_group', ['', 'root_wprop.name', 'dish_item', 'modifier_group'], ['id', 'dish_group', 'dish_item', 'modifier_group'], dg)
tab('dish_item', ['', 'root_wprop.name', 'wprop.unit_price', 'wprop.unit_price', 'wprop.unit_price'],
    ['id', 'dish_item', 'Small', 'Large', 'Party Tray'],
    [[f'{{"id":"d{i}"}}', n, m(s), m(l), m(t)] for i, (n, g, s, l, t) in enumerate(DISHES)])
tab('modifier_group', ['', 'root_wprop.name', 'modifier_item'], ['id', 'modifier_group', 'modifier_item'],
    [['{"id":"mg1"}', 'Add Ons', 'Extra Sauce'], ['', '', 'Extra Rice']])
tab('modifier_item', ['', 'root_wprop.name', 'wprop.unit_price'], ['id', 'modifier_item', 'Order'],
    [['{"id":"m1"}', 'Extra Sauce', '$0.50'], ['{"id":"m2"}', 'Extra Rice', '$1.50']])
out = 'R99999_raw_mixed.xlsx' if MIXED else 'R99999_raw.xlsx'
wb.save(out); print('wrote', out)

# Paper menu OCR — what the vision pass would emit
ocr = {"restaurant": {"name": "Tray Test Kitchen", "r_number": "R99999"},
  "sections": [
    {"menu_section": "Chicken", "price_columns": ["Small", "Large", "Party Tray"], "default_price": None, "items": [
        {"name": "General Tso's Chicken", "prices": {"Small": 9.45, "Large": 13.95}},             # tray NOT printed
        {"name": "Sesame Chicken",        "prices": {"Small": 9.45, "Large": 13.95, "Party Tray": 68.00}},  # tray printed
        {"name": "Chicken with Broccoli", "prices": {"Small": 8.75, "Large": 12.95}},
        {"name": "Kung Pao Chicken",      "prices": {"Small": 8.25, "Large": 11.95}},             # unchanged
        {"name": "Lemon Chicken",         "prices": {"Small": 8.50, "Large": 12.00}}]},           # decrease
    {"menu_section": "Beef", "price_columns": ["Small", "Large"], "default_price": None, "items": [
        {"name": "Beef with Broccoli", "prices": {"Small": 10.45, "Large": 15.95}},
        {"name": "Pepper Steak",       "prices": {"Small": 10.45, "Large": 15.95}},
        {"name": "Mongolian Beef",     "prices": {"Small": 10.95, "Large": 16.50}}]},
    {"menu_section": "Party Trays", "price_columns": ["Party Tray"], "default_price": None, "items": [
        {"name": "Lo Mein Tray", "prices": {"Party Tray": 48.00}}]}]}          # Fried Rice Tray not printed; Hunan Chicken not printed at all
with open('R99999_menu_ocr.json', 'w') as f: json.dump(ocr, f, indent=2)
print('wrote R99999_menu_ocr.json')
