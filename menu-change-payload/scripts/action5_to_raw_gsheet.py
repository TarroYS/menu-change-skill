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
    G2. NEVER write to blank/empty cells in Raw Excel 1 — except a NEW price the
        specialist typed into an EMPTY report cell (Dish Item: the cell's hidden
        _ref_ baseline is empty too; Modifier Summary: resolved by action4 in
        the 'Part 2 Intent' sheet, and only on the live row the report shows).
        Recorded in part2_summary.json (action5.added) for bridge to-tsv.
    G3. Only touches dish_item and modifier_item tabs — all other tabs untouched
    G4. Only overwrites cells where the price actually changed
    G11. NEVER writes a price onto a $0.00 dish whose price lives on modifiers
         (priced modifier_rule rows for the dish or its dish group, a required
         fully-priced modifier group, or — in a restaurant that prices dishes on
         modifier rules — any $0.00 Fried Specialties dish), never adds a size
         such a dish would be double-charged in, never removes an ABSOLUTE size,
         and never removes / zeroes an option of a required group that prices a
         $0.00 dish. Printed as "!! ABSOLUTE GUARD", recorded in part2_summary.json.
    X.   A Dish Item / Modifier Summary price the specialist replaced with x / X =
         remove that price: the live cell is cleared (the ONLY blank write, and
         only onto a cell that has a price). Recorded in part2_summary.json
         (action5.removed), which bridge to-tsv reads. An x on an empty cell is
         ignored and reported. A removed size's dish code (unit_code) is never
         moved automatically — unless the specialist moved / removed it on the
         Dish Code tab, it is listed in action5.code_followups for Wonka.
    DUP. A dish listed under several dish groups is ONE live dish: its report
         rows are resolved per size — an edited row (≠ its _ref_) wins over
         untouched rows (so an untouched lower copy never restores the old
         price); rows edited to different prices, or untouched rows carrying
         different new prices, are a CONFLICT — nothing applied, reported.
    Report prices may be plain numbers (12.95) or '$12.95' strings; they are
    written in the live cell's own format ($X.XX), and 0 is a real price.
    Anything that is not a valid price (letters, comma, %, negative, 1000+) is
    ignored and reported — a stray keystroke never ships.
"""

import argparse
import re
import sys
from typing import Optional, Dict, Tuple, List
from collections import OrderedDict, defaultdict
from decimal import Decimal, ROUND_HALF_UP

try:
    from openpyxl import load_workbook
except ImportError:
    print("ERROR: openpyxl required. Install: pip install openpyxl")
    sys.exit(1)


# Dish codes (maintainer, 2026-10-05): Part 1's Dish Code tab shows each size's
# code next to the printed one; a code is written ONLY where the specialist put
# one in a code cell (fills and corrections alike are confirmed by a person —
# Wonka's dish-code brief). x removes a code. The old automatic overwrite (off
# since it made mistakes) is gone.
CODE_TAB = 'Dish Code'
CODE_MAX = 128          # unit_code length limit
PAPER_SUFFIX = ' · paper'   # the tab's read-only paper block ('Order · paper')


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
    """Format price as $X.XX (matching Raw Excel 1 format; a negative as
    -$X.XX, like wsheet)."""
    if val is None: return ''
    try:
        f = float(val)
        return f'-${-f:.2f}' if f < 0 else f'${f:.2f}'
    except: return ''

def _is_x(val) -> bool:
    """x / X typed over a Dish Item price = remove that price."""
    return isinstance(val, str) and val.strip().lower() == 'x'


def _has_value(val) -> bool:
    """A report cell carries something (0 is a real price, not blank)."""
    return val is not None and str(val).strip() != ''


def _money(new_p, dollar: bool) -> str:
    """'$X.XX' / '-$X.XX' (a negative modifier, as wsheet writes it), or a
    bare 'X.XX' / '-X.XX'."""
    if not dollar:
        return f'{new_p:.2f}'
    return f'-${-new_p:.2f}' if new_p < 0 else f'${new_p:.2f}'


def _price_str(new_p, live_val) -> str:
    """The new price in the live cell's own format: '$X.XX' when the live cell
    uses '$' (every wsheet price does — '-$1.70' too), else plain 'X.XX' (G6)."""
    return _money(new_p, '$' in str(live_val or ''))


# A price entry after every '$' is removed (like the sheet's SUBSTITUTE): an
# optional sign, ASCII digits, optional decimals. Kept in step with the report's
# red-text rule (action3), so the sheet and Part 2 agree on what is a price.
_ENTRY_RE = re.compile(r'^([+-]?)([0-9]+(?:\.[0-9]*)?|\.[0-9]+)$')
PRICE_CAP = 1000.0      # no dish / modifier costs this — a missing decimal point (1395)
TRAY_VARIANTS = {'Party Tray'}   # keep in sync with action2 / price_inference
INTENT_TAB = 'Part 2 Intent'    # hidden sheet action4 writes into the Final Review
# Dish Item fills that mean "Part 1 changed this price": blue (increase), pink
# (decrease), purple (inferred). An UNTOUCHED cell is written only with one of
# these — an unchanged snapshot price never overwrites the live one.
PART1_CHANGE_FILLS = {'ADD8E6', 'FFB6C1', 'E4DFEC'}
BY_ITSELF_ALIASES = ['by itself', 'plain', 'no rice', 'no side']


def _round2(x) -> float:
    """Round to cents the way a spreadsheet does — half away from zero, on the
    decimal value as typed (2.675 → 2.68, 15.995 → 16.00; Python's round()
    works on the binary value and gives 2.67 / 15.99, so Part 2 would write a
    cent less than the report shows)."""
    f = float(x)
    if f != f or abs(f) >= 1e12:          # NaN / huge: no cents to round (fails the cap anyway)
        return f
    return float(Decimal(x if isinstance(x, str) else repr(f)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))


def _entry(val, allow_negative=False):
    """(kind, number) for a report price cell: 'blank', 'x', 'price' or
    'invalid' — never applied, so a stray keystroke never ships. Matches the
    report's red-text rules (action3): a number cell is checked as a number;
    text is trimmed of plain spaces only (like TRIM — a tab or no-break space is
    not a price), '$' removed, then must be a plain decimal. Invalid: letters,
    comma, %, a negative (dish prices; allow_negative for modifiers), 1000 or
    more after rounding."""
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
    return 'price', f + 0.0          # -0.0 → 0.0


def _sys_of(cells, v):
    """The SYSTEM price of size v on one report row: the hidden _ref_sys_
    mirror (a report made before it existed: the _ref_ baseline)."""
    cv = cells.get(v)
    if not cv:
        return None
    return _parse(cv[3]) if cv[4] else _parse(cv[1])


def _new_size_misplaced(new_p, others) -> bool:
    """A new price (typed into an empty cell) at 10× the item's highest other
    price, or a tenth of its lowest — a misplaced decimal point. `others` = the
    item's other prices as the report shows them; a $0.00 one switches the tenth
    test off. The report's red-text rule is the same."""
    pos = [p for p in others if p > 0.001]
    if not pos:
        return False
    return (abs(new_p) >= 10 * max(pos) - 1e-9
            or (min(others) > 0.001 and abs(new_p) * 10 <= min(others) + 1e-9))


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
    """Report cell equals its _ref_ baseline (numbers compared as numbers)."""
    pa, pb = _parse(a), _parse(b)
    if pa is not None and pb is not None:
        return abs(pa - pb) < 0.001
    return str(a if a is not None else '').strip().lower() == \
        str(b if b is not None else '').strip().lower()


def _is_blank(val) -> bool:
    return val is None or str(val).strip() == ''


def _is_by_itself(name: str) -> bool:
    """A modifier_item name that IS a base-price anchor (whole name, normalized
    — same rule as action1 / modifier_rule_pricing)."""
    s = str(name or '').lower().strip()
    s = re.sub(r'\s*\([^)]*\)\s*', ' ', s)
    s = s.replace('w.', 'with').replace('&', 'and')
    s = re.sub(r'[^a-z0-9\s]', '', s)
    return re.sub(r'\s+', ' ', s).strip() in BY_ITSELF_ALIASES


def _is_fried_specialty_group(group: str) -> bool:
    return bool(group) and 'fried special' in group.lower()


def _money_prefix(ws, row_idx, cols) -> str:
    """'$' when the row's prices are '$X.XX' (every wsheet price is), for a
    price written into a cell that had none."""
    for ci in cols:
        v = ws.cell(row=row_idx, column=ci + 1).value
        if not _is_blank(v):
            return '$' if '$' in str(v) else ''
    return '$'


def _hidden_by_key(rows, headers):
    """{(group, name): row} from the hidden '_ref_key:group' / '_ref_key:name'
    columns — each row's baselines, wherever the visible row is now. None for
    a report without the key columns (then rows are read by position)."""
    if '_ref_key:group' not in headers or '_ref_key:name' not in headers:
        return None
    gi, ni = headers.index('_ref_key:group'), headers.index('_ref_key:name')
    out = {}
    for r in rows[1:]:
        n = str(r[ni]).strip() if ni < len(r) and r[ni] is not None else ''
        if n:
            out[(str(r[gi]).strip() if gi < len(r) and r[gi] is not None else '', n)] = r
    return out


def _dish_item_styles(path):
    """(changed, percent) — Dish Item price cells Part 1 coloured as a change
    (blue / pink / purple: the fallback for reports without _ref_sys_), and
    cells formatted as a percentage (Excel stores a typed "10%" as 0.1 — never a
    price). Needs the styles, so the Final Review is read again."""
    wb = load_workbook(path)
    tab = 'Dish Item' if 'Dish Item' in wb.sheetnames else ('price' if 'price' in wb.sheetnames else None)
    changed, percent = set(), set()
    if tab:
        ws = wb[tab]
        hdr = [str(c.value).strip() if c.value is not None else '' for c in ws[1]]
        for ci, h in enumerate(hdr, 1):
            if not _is_variant_header(h):
                continue
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=ci)
                f = cell.fill
                if f is not None and f.fill_type == 'solid' and \
                        str(f.fgColor.rgb or '')[-6:].upper() in PART1_CHANGE_FILLS:
                    changed.add((r, h))
                if '%' in str(cell.number_format or ''):
                    percent.add((r, h))
    wb.close()
    return changed, percent


def _split(insts, v, changed, percent=frozenset()):
    """The rows of one dish (a dish listed under two dish groups has two rows
    but is ONE live dish) for variant v → (edits, script, invalid).
    edits  — rows whose value differs from its _ref_ baseline (the specialist
             typed it), as (row, group, kind, num, ref_was_blank, True);
    script — untouched rows (Part 1's paper / inferred / system value), as
             (row, group, kind, num, False, part1_changed) — part1_changed:
             Part 1 changed the price, i.e. _ref_ ≠ the hidden _ref_sys_ system
             price (older reports: the blue / pink / purple fill);
    invalid— rows the specialist typed something that is not a price into.
    A report without _ref_ columns: only an x counts as an edit, and every
    untouched value counts as a change (the pre-2026-09-30 behaviour)."""
    edits, script, invalid = [], [], []
    for ri, grp, cells in insts:
        raw, ref, has_ref, sys_, has_sys = cells.get(v, (None, None, False, None, False))
        kind, num = _entry(raw)
        if kind == 'price' and (ri, v) in percent:
            kind = 'invalid'            # typed as a percentage
        if kind == 'blank':
            continue                    # a cleared cell is no instruction (G1)
        edited = (not _same(raw, ref)) if has_ref else (kind == 'x')
        if kind == 'invalid':
            if edited or not has_ref:
                invalid.append((ri, grp, raw))
            continue
        if edited:
            edits.append((ri, grp, kind, num, has_ref and _is_blank(ref), True))
        else:
            # Part 1 changed it ⇔ what Part 1 wrote (_ref_) ≠ the system price
            # (_ref_sys_). Older reports without the mirror: the blue / pink /
            # purple fill; without _ref_ columns at all: always.
            chg = (not _same(ref, sys_)) if has_sys else ((ri, v) in changed if has_ref else True)
            script.append((ri, grp, kind, num, False, chg))
    return edits, script, invalid


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


def _code_entry(val):
    """(kind, code) for a Dish Code cell: 'blank', 'x', 'code' or 'invalid' —
    the same verdict as the tab's red text: a line break / tab, more than 128
    characters, a leading = (a formula — wsheet would run it), or nothing but
    # - . and spaces is not a code."""
    if val is None:
        return 'blank', None
    if isinstance(val, bool):
        return 'invalid', None
    if isinstance(val, float) and val.is_integer():
        val = int(val)
    s = str(val)
    if any(ch in s for ch in '\n\r\t') or len(s) > CODE_MAX:
        return 'invalid', None
    t = _norm_code(s)
    if not t:
        return 'blank', None
    if t.lower() == 'x':
        return 'x', None
    if t.startswith('=') or not re.sub(r'[#\-. ]', '', t):
        return 'invalid', None
    return 'code', t


def _norm_code(val) -> str:
    """A dish code as text, read the way the Dish Code tab's rules read it
    (Excel TRIM, no-break spaces as spaces): a number cell 38 / 38.0 → '38';
    text stays as typed ('1.0' stays '1.0'); every run of spaces → one space,
    none at the ends. The bridge (to-tsv) compares codes the same way."""
    if val is None:
        return ''
    if isinstance(val, float) and val.is_integer():
        val = int(val)
    return ' '.join(str(val).split())


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


def _pricing_structure(wb) -> dict:
    """How the raw being patched prices its $0.00 dishes (G11). Columns are
    found by header text (G4). Returns
      'absolute'    {(dish, variant): evidence} — $0.00 dish prices whose real
                    price lives on modifiers: priced modifier_rule rows for the
                    dish (variant matches or is blank) or for one of its dish
                    groups (group-level rows), or a REQUIRED modifier group on
                    the dish / its dish group whose every item is priced for
                    that variant; and, in a restaurant that prices dishes on
                    modifier rules, every $0.00 Fried Specialties dish.
      'new_size'    {(dish, variant): evidence} — a NEW price in this empty
                    size would be charged on top of the modifiers (a priced By
                    Itself rule for that size, or a dish priced on its
                    modifiers in its other sizes).
      'mod_options' {(modifier_item, variant): evidence} — options of a
                    required group that prices a $0.00 dish: removing one, or
                    setting it to 0, would make that dish free."""
    def tab(name):
        if name not in wb.sheetnames:
            return [], [], []
        rows = list(wb[name].iter_rows(values_only=True))
        if len(rows) < 3:
            return rows, [], []
        return (rows, [str(c).strip() if c else '' for c in rows[2]],
                [str(c).strip() if c else '' for c in rows[1]])

    def cell(r, ci):
        return str(r[ci]).strip() if ci is not None and ci < len(r) and r[ci] is not None else ''

    empty = {'absolute': {}, 'new_size': {}, 'mod_options': {}}
    rows, h3, h2 = tab('dish_item')
    if 'dish_item' not in h3:
        return empty
    nc = h3.index('dish_item')
    pcols = _find_unit_price_cols(h3, h2)
    mg_cols = [i for i, h in enumerate(h3) if h == 'modifier_group']
    mgc = next((i for i in mg_cols if i < len(rows[0]) and cell(rows[0], i) == 'comprise'),
               mg_cols[0] if mg_cols else None)
    zero, dish_mgs, dishes, cur = [], defaultdict(set), [], ''
    for r in rows[3:]:
        if _is_root_row(r):
            cur = cell(r, nc)
            if cur:
                dishes.append(cur)
            for v, ci in pcols.items():
                p = _parse(r[ci]) if ci < len(r) else None
                if cur and p is not None and abs(p) <= 0.001:
                    zero.append((cur, v))
        owner = cell(r, nc) or cur
        if owner and cell(r, mgc):
            dish_mgs[owner].add(cell(r, mgc))

    rules = defaultdict(list)                 # dish → [(variant, modifier, price, modifier group)]
    grules = defaultdict(list)                # dish_group → [(variant, modifier, price, modifier group)]
    rows, h3, _ = tab('modifier_rule')
    if all(k in h3 for k in ('out_unit_price', 'dish_item', 'variant')):
        pi, di, vi = h3.index('out_unit_price'), h3.index('dish_item'), h3.index('variant')
        mi = h3.index('modifier_item') if 'modifier_item' in h3 else None
        gi = h3.index('dish_group') if 'dish_group' in h3 else None
        mgi = h3.index('modifier_group') if 'modifier_group' in h3 else None
        for r in rows[3:]:
            p = _parse(r[pi]) if pi < len(r) else None
            if p is None or p <= 0.001:
                continue
            if cell(r, di):
                rules[cell(r, di)].append((cell(r, vi), cell(r, mi), p, cell(r, mgi)))
            elif cell(r, gi):
                grules[cell(r, gi)].append((cell(r, vi), cell(r, mi), p, cell(r, mgi)))

    dish_groups, group_mgs = defaultdict(set), defaultdict(set)
    rows, h3, _ = tab('dish_group')
    if 'dish_group' in h3:
        gi = h3.index('dish_group')
        dci = h3.index('dish_item') if 'dish_item' in h3 else None
        mci = h3.index('modifier_group') if 'modifier_group' in h3 else None
        g = ''
        for r in rows[3:]:
            if _is_root_row(r) and cell(r, gi):
                g = cell(r, gi)
            if g and cell(r, dci):
                dish_groups[cell(r, dci)].add(g)
            if g and cell(r, mci):
                group_mgs[g].add(cell(r, mci))

    required, mg_items = set(), defaultdict(list)
    rows, h3, _ = tab('modifier_group')
    if 'modifier_group' in h3:
        gi = h3.index('modifier_group')
        ii = h3.index('modifier_item') if 'modifier_item' in h3 else None
        mn = h3.index('min_number_of_selections') if 'min_number_of_selections' in h3 else None
        st = h3.index('subtype') if 'subtype' in h3 else None
        g = ''
        for r in rows[3:]:
            if _is_root_row(r) and cell(r, gi):
                g = cell(r, gi)
                n = _parse(cell(r, mn))
                if (n is not None and n >= 1) or cell(r, st).lower() == 'required':
                    required.add(g)
            if g and cell(r, ii) and cell(r, ii) not in mg_items[g]:
                mg_items[g].append(cell(r, ii))

    mod_prices = {}
    rows, h3, h2 = tab('modifier_item')
    if 'modifier_item' in h3:
        ni = h3.index('modifier_item')
        mcols = _find_unit_price_cols(h3, h2)
        for r in rows[3:]:
            if _is_root_row(r) and cell(r, ni) and cell(r, ni) not in mod_prices:
                mod_prices[cell(r, ni)] = {v: _parse(r[ci]) for v, ci in mcols.items() if ci < len(r)}

    def shown(rr):
        return ', '.join(f'{x[1] or "?"} {_fmt(x[2])}' for x in rr[:4])

    def holds_price(rr):
        """Priced rules hold a $0.00 dish's price when a By Itself is among
        them, or they price a REQUIRED choice (the customer must pick one — its
        rule price is the dish's price). An optional add-on ("Sauce / With
        +$1.00") is only an upcharge."""
        return any(_is_by_itself(x[1]) or (x[3] and x[3] in required) for x in rr)

    def rule_evidence(dish, v):
        rr = [x for x in rules.get(dish, []) if x[0] in ('', v)]
        if rr:
            return f'{len(rr)} priced modifier_rule row(s): ' + shown(rr)
        for g in sorted(dish_groups.get(dish, ())):
            gr = [x for x in grules.get(g, []) if x[0] in ('', v)]
            # group-level rules hold the price only with a priced By Itself or
            # on a required choice — a group add-on ("Sauce / With +$1.00") is
            # an upcharge
            if gr and holds_price(gr):
                return f"{len(gr)} priced group-level modifier_rule row(s) on dish group '{g}': " + shown(gr)
        return None

    def has_anchor(dish, v):
        """Priced rules that hold this size's price (holds_price) — dish-level
        or group-level."""
        if holds_price([x for x in rules.get(dish, []) if x[0] in ('', v)]):
            return True
        return any(holds_price([x for x in grules.get(g, []) if x[0] in ('', v)])
                   for g in dish_groups.get(dish, ()))

    absolute, mod_options = {}, {}
    # "prices on modifier rules" = some live $0.00 dish has priced rules that
    # hold its price — a By Itself, or a required choice (the same test action1
    # makes)
    rule_based = any(has_anchor(dish, v) for dish, v in zero)
    for dish, v in zero:
        ev = rule_evidence(dish, v)
        if ev:
            absolute[(dish, v)] = ev
            continue
        mgs = set(dish_mgs.get(dish, ())) | {m for g in dish_groups.get(dish, ())
                                              for m in group_mgs.get(g, ())}
        for mg in sorted(mgs):
            items = mg_items.get(mg, [])
            ip = [mod_prices.get(it, {}).get(v) for it in items]
            if mg in required and ip and all(x is not None and x > 0.001 for x in ip):
                absolute[(dish, v)] = f"required modifier group '{mg}': all {len(ip)} items priced"
                for it in items:
                    mod_options.setdefault((it, v), f"an option of required group '{mg}', which "
                                                    f"prices {dish} [{v}] (a $0.00 dish)")
                break
    # Maintainer (2026-09-30): in a restaurant that prices dishes on modifier
    # rules (ABSOLUTE), a $0.00 Fried Specialties dish is never given a price.
    if rule_based:
        for dish, v in zero:
            if (dish, v) not in absolute and any(_is_fried_specialty_group(g)
                                                 for g in dish_groups.get(dish, ())):
                absolute[(dish, v)] = ('Fried Specialties dish at $0.00 in a restaurant that '
                                       'prices dishes on modifier rules (ABSOLUTE)')
    new_size = {}
    # dishes whose price lives on modifiers (rules / a required group) — not
    # one held only by the Fried Specialties name rule, which may be priced
    # on its dish_item in its other sizes
    abs_dishes = {d for (d, _v), ev in absolute.items() if not ev.startswith('Fried Specialties dish')}
    for dish in dishes:
        for v in pcols:
            bi = [x for x in rules.get(dish, []) if x[0] in ('', v) and _is_by_itself(x[1])]
            if bi:
                new_size[(dish, v)] = f'a priced By Itself rule for this size ({_fmt(bi[0][2])})'
            elif has_anchor(dish, v):
                # group-level By Itself, or rules on a required choice, price this size
                new_size[(dish, v)] = (rule_evidence(dish, v)
                                       or 'priced modifier rules hold this size\'s price') + ' — a dish price would be charged on top'
            elif dish in abs_dishes:
                new_size[(dish, v)] = 'the dish is priced on its modifiers ($0.00 dish price)'
    return {'absolute': absolute, 'new_size': new_size, 'mod_options': mod_options}


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

    # Read Excel 4 price tab → every report row of each dish (a dish listed
    # under two dish groups has two rows but is ONE live dish), with each price
    # cell's hidden _ref_ baseline (what Part 1 wrote) so an EDITED copy wins
    # over an untouched one — the lower copy never puts the old price back.
    e4_price_rows = list(wb_4[_pick_tab(wb_4, 'Dish Item', 'price')].iter_rows(values_only=True))
    e4_price_headers = [str(c).strip() if c else '' for c in e4_price_rows[0]]
    e4_price_name_col = e4_price_headers.index('Dish Name:Primary') if 'Dish Name:Primary' in e4_price_headers else 1
    e4_price_group_col = e4_price_headers.index('Dish Group') if 'Dish Group' in e4_price_headers else 0
    # Auto-detect variant columns from headers (header-driven)
    e4_price_variants = {h.strip() for h in e4_price_headers if _is_variant_header(h)}
    e4_ref_cols = {h[len('_ref_'):]: i for i, h in enumerate(e4_price_headers)
                   if h.startswith('_ref_') and not h.startswith(('_ref_sys_', '_ref_key:'))}
    # hidden SYSTEM-price mirror (Part 1 changed a cell ⇔ its _ref_ ≠ its _ref_sys_)
    e4_sys_cols = {h[len('_ref_sys_'):]: i for i, h in enumerate(e4_price_headers)
                   if h.startswith('_ref_sys_')}
    if e4_ref_cols:
        # the price columns Part 1 made carry a _ref_ column; one a specialist
        # added (notes …) is not a size
        e4_price_variants = {v for v in e4_price_variants if v in e4_ref_cols}
    else:
        print("  !! This report has no hidden baselines (_ref_ columns — made before they "
              "existed): every report price that differs from live is treated as an "
              "instruction, as before 2026-09-30. Prefer a report made with the current Part 1.")
    # The hidden block (baselines) is found by each row's own key, not by
    # position — a quick sort in Excel moves the visible columns only.
    hidden_rows = _hidden_by_key(e4_price_rows, e4_price_headers)
    e4_prices = defaultdict(list)  # dish → [(row, group, {variant: (value, _ref_, has _ref_, _ref_sys_, has _ref_sys_)})]
    for ri, r in enumerate(e4_price_rows[1:], 2):
        name = str(r[e4_price_name_col]).strip() if e4_price_name_col < len(r) and r[e4_price_name_col] else ''
        if not name:
            continue
        group = str(r[e4_price_group_col]).strip() if e4_price_group_col < len(r) and r[e4_price_group_col] else ''
        hid = hidden_rows.get((group, name), r) if hidden_rows is not None else r
        cells = {}
        for v in e4_price_variants:
            ci = e4_price_headers.index(v)
            rc, sc = e4_ref_cols.get(v), e4_sys_cols.get(v)
            cells[v] = (r[ci] if ci < len(r) else None,
                        hid[rc] if rc is not None and rc < len(hid) else None, rc is not None,
                        hid[sc] if sc is not None and sc < len(hid) else None, sc is not None)
        e4_prices[name].append((ri, group, cells))

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
                if ci < len(r) and _has_value(r[ci]):
                    prices[v] = r[ci]
        e4_mod_prices[mn] = prices

    # Hidden 'Part 2 Intent' sheet (action4): the Modifier Summary new cells /
    # removals action5 may write, plus conflicts and ignored input to report.
    mod_intent = {}               # (item, variant) → 'price' | 'new' | 'x'
    has_intent = INTENT_TAB in wb_4.sheetnames
    if not has_intent and 'Modifier Summary' in wb_4.sheetnames:
        print(f"ERROR: {args.excel_4} has no '{INTENT_TAB}' sheet — it is not action4's Final Review. "
              f"Run action4_apply_modifier_changes.py on the edited report first, then action5 on "
              f"its output (otherwise Modifier Summary edits are lost and live prices reverted).")
        sys.exit(1)
    a4_notes = defaultdict(list)  # kind → [{'item','variant','value','detail'}]
    if INTENT_TAB in wb_4.sheetnames:
        for r in list(wb_4[INTENT_TAB].iter_rows(values_only=True))[1:]:
            if not r or not r[0]:
                continue
            kind, item, var, val, det = (list(r) + [None] * 5)[:5]
            item, var = str(item or '').strip(), str(var or '').strip()
            if kind in ('new', 'x', 'price'):
                mod_intent[(item, var)] = kind
            a4_notes[str(kind)].append({'item': item, 'variant': var,
                                        'value': '' if val is None else str(val),
                                        'detail': str(det or '')})

    # Read the Dish Code tab → {dish: [(row, group, {size: (typed, system code at Part 1)})]}.
    # The code columns are those with a hidden _ref_ (the paper block is reference
    # only); each row's baselines are found by its hidden key (sort-safe).
    e4_codes = defaultdict(list)
    has_code_tab = CODE_TAB in wb_4.sheetnames
    if has_code_tab:
        dc_rows = list(wb_4[CODE_TAB].iter_rows(values_only=True))
        if dc_rows:
            dc_headers = [str(c).strip() if c else '' for c in dc_rows[0]]
            dc_refs = {h[len('_ref_'):]: i for i, h in enumerate(dc_headers)
                       if h.startswith('_ref_') and not h.startswith(('_ref_sys_', '_ref_key:'))}
            # each size's code column by its header — or, when the specialist
            # pasted its whole paper column over it (header included, so the
            # header now reads 'Order · paper'), by its place: the code block is
            # the k-th column after the 3 name columns, in _ref_ order. A code
            # column found neither way stops Part 2 (its typed codes would be lost).
            hidden_from = min(dc_refs.values()) if dc_refs else len(dc_headers)
            dc_sizes, lost = {}, []
            for k, v in enumerate(dc_refs):
                pos = 3 + k
                if v in dc_headers[:hidden_from]:
                    dc_sizes[v] = dc_headers.index(v)
                elif pos < hidden_from and dc_headers[pos] == v + PAPER_SUFFIX:
                    dc_sizes[v] = pos
                else:
                    lost.append(v)
            if lost:
                print(f"ERROR: the {CODE_TAB} tab has no code column for {', '.join(repr(v) for v in lost)} "
                      f"any more (its header in row 1 was changed or the column deleted), so Part 2 cannot "
                      f"read the codes typed there. Put the column / header back as Part 1 made it "
                      f"(the code columns follow 'Dish Name:Secondary', one per size) and run Part 2 again. "
                      f"Nothing was written.")
                sys.exit(1)
            dc_name_col = dc_headers.index('Dish Name:Primary') if 'Dish Name:Primary' in dc_headers else 1
            dc_group_col = dc_headers.index('Dish Group') if 'Dish Group' in dc_headers else 0
            dc_hidden = _hidden_by_key(dc_rows, dc_headers)
            for ri, r in enumerate(dc_rows[1:], 2):
                nm = str(r[dc_name_col]).strip() if dc_name_col < len(r) and r[dc_name_col] else ''
                if not nm:
                    continue
                grp = str(r[dc_group_col]).strip() if dc_group_col < len(r) and r[dc_group_col] else ''
                hid = dc_hidden.get((grp, nm), r) if dc_hidden is not None else r
                e4_codes[nm].append((ri, grp, {
                    v: (r[ci] if ci < len(r) else None,
                        hid[dc_refs[v]] if dc_refs[v] < len(hid) else None)
                    for v, ci in dc_sizes.items()}))

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
    # G11: $0.00 dishes whose price lives on modifiers — never given a price.
    struct = _pricing_structure(wb_raw)
    absolute_cells = struct['absolute']
    absolute_guard = []
    removed = []        # x / X: prices cleared on request
    x_ignored = []      # x / X typed over a cell that had no price
    added = []          # new prices typed into empty cells
    ignored_input = []  # not a valid price — never applied
    dup_conflicts = []  # one dish / modifier edited to different prices
    dup_resolved = []   # an edited copy used over untouched copies
    new_cell_drift = [] # typed over a price that is gone from the live cell
    code_followups = [] # dish codes left on a size whose price was removed / moved
    not_reverted = []   # untouched report prices that differ from live (live moved since Part 1)
    no_price_left = []  # dishes whose every price was removed with x (take off the menu = strike)
    twin_rows = []      # same-name live root rows after the first — never written
    live_di_variants = set(di_price_cols)
    # Fills are only the fallback for reports without the _ref_sys_ mirror.
    _fills, percent_cells = _dish_item_styles(args.excel_4)
    part1_changed = set() if e4_sys_cols else _fills

    dish_changes = 0
    dish_cells = 0
    code_cells = 0      # codes corrected (written over an existing code)
    code_items = 0
    codes_done = {'filled': [], 'corrected': [], 'removed': [], 'drift': []}

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
        pending_followups = []      # codes left on a removed / moved price (see below)
        if name in e4_prices and name in di_price_seen:
            # a second live root row with the same name (an availability /
            # mutant row, or a twin the report never showed) — never written
            twin_rows.append({'tab': 'dish_item', 'name': name, 'row': row_idx})
            continue
        if name in e4_prices:
            di_price_seen.add(name)
            insts = e4_prices[name]
            item_changed = False
            row_removed, row_added = [], []     # for the dish-code follow-up
            for v, ci in di_price_cols.items():
                if v not in e4_price_variants:
                    continue
                old_val = ws_di.cell(row=row_idx, column=ci+1).value
                old_p = _parse(old_val)
                edits, script, invalid = _split(insts, v, part1_changed, percent_cells)
                for ri, grp, raw in invalid:
                    ignored_input.append({'tab': 'Dish Item', 'row': ri, 'name': name, 'group': grp,
                                          'variant': v, 'value': str(raw), 'why': 'not a valid price'})
                if edits:
                    distinct = {(k, n) for _r, _g, k, n, _b, _c in edits}
                    if len(distinct) > 1:
                        dup_conflicts.append({'tab': 'Dish Item', 'name': name, 'variant': v, 'rows': [
                            {'row': r_, 'group': g_, 'value': 'x' if k_ == 'x' else n_}
                            for r_, g_, k_, n_, _b, _c in edits]})
                        continue
                    kind, new_p = next(iter(distinct))
                    from_edit, typed_blank = True, all(b for *_x, b, _c in edits)
                    if len(insts) > 1 and len(edits) < len(insts):
                        dup_resolved.append({'dish': name, 'variant': v,
                                             'used_rows': [r_ for r_, *_x in edits]})
                else:
                    # Untouched rows: only a price Part 1 CHANGED (blue / pink /
                    # purple) is an instruction; an unchanged snapshot that now
                    # differs from live = the live price moved since Part 1 —
                    # kept, never reverted.
                    changes = sorted({n for _r, _g, k, n, _b, chg in script if k == 'price' and chg
                                      and (old_p is None or abs(n - old_p) > 0.001)})
                    if not changes:
                        if old_p is not None and any(k == 'price' and not chg and abs(n - old_p) > 0.001
                                                     for _r, _g, k, n, _b, chg in script):
                            not_reverted.append({'tab': 'dish_item', 'name': name, 'variant': v,
                                                 'live': _fmt(old_p)})
                        continue
                    if len(changes) > 1:
                        dup_conflicts.append({'tab': 'Dish Item', 'name': name, 'variant': v, 'rows': [
                            {'row': r_, 'group': g_, 'value': n_}
                            for r_, g_, _k, n_, _b, chg in script if chg]})
                        continue
                    kind, new_p, from_edit, typed_blank = 'price', changes[0], False, False
                    sys_p = next((_sys_of(c_, v) for _r, _g, c_ in insts if _sys_of(c_, v) is not None), None)
                    if any(c_.get(v, (None,) * 5)[4] for _r, _g, c_ in insts) and _misplaced(new_p, sys_p):
                        r0 = next((r_, g_) for r_, g_, _k, _n, _b, chg in script if chg)
                        ignored_input.append({'tab': 'Dish Item', 'row': r0[0], 'name': name, 'group': r0[1],
                                              'variant': v, 'value': str(new_p),
                                              'why': f"Part 1's price {_fmt(new_p)} is {MISPLACED_RATIO}× / a fifth of the "
                                                     f"system's {_fmt(sys_p)} — probably a misread paper price; "
                                                     f"not applied. Type the right price in the report"})
                        continue

                if from_edit and typed_blank:
                    # Typed into a cell that was EMPTY in the report: judged by
                    # what the specialist saw, never by a price wsheet has now.
                    if kind == 'x':
                        x_ignored.append({'dish': name, 'variant': v})
                        continue
                    if new_p <= 0.001:
                        ignored_input.append({'tab': 'Dish Item', 'row': edits[0][0], 'name': name,
                                              'group': edits[0][1], 'variant': v, 'value': '0',
                                              'why': '0 in an empty cell — a new price must be above 0'})
                        continue
                    if old_p is not None:
                        new_cell_drift.append({'dish': name, 'variant': v, 'report': _fmt(new_p),
                                               'live': _fmt(old_p),
                                               'note': 'the live menu now has a price the report did not show'})
                        continue
                    # the dish's other prices as the report shows them (what the
                    # specialist saw), like the report's red-text rule
                    others = [p_ for (r_, _g, c_) in insts for vv, cv in c_.items()
                              if vv != v and vv not in TRAY_VARIANTS
                              for p_ in [_parse(cv[1])] if p_ is not None]
                    if v not in TRAY_VARIANTS and _new_size_misplaced(new_p, others):
                        pos = [p for p in others if p > 0.001]
                        ignored_input.append({'tab': 'Dish Item', 'row': edits[0][0], 'name': name,
                                              'group': edits[0][1], 'variant': v, 'value': str(new_p),
                                              'why': f'looks like a misplaced decimal point (10× or a tenth of '
                                                     f'this dish\'s other prices, {_fmt(min(pos))}–'
                                                     f'{_fmt(max(pos))}) — not applied; retype it, or confirm '
                                                     f'in Notes → Wonka'})
                        continue
                if from_edit and kind == 'price' and not typed_blank:
                    rows_ = {r_ for r_, *_x in edits}
                    # judged against the SYSTEM price — the report's own price may
                    # be the misread being corrected (19.50 printed for 1.95)
                    base = next((_sys_of(c_, v) for r_, _g, c_ in insts
                                 if r_ in rows_ and _sys_of(c_, v) is not None), None)
                    if _misplaced(new_p, base):
                        ignored_input.append({'tab': 'Dish Item', 'row': edits[0][0], 'name': name,
                                              'group': edits[0][1], 'variant': v, 'value': str(new_p),
                                              'why': f'looks like a misplaced decimal point ({_fmt(new_p)} vs '
                                                     f'the system\'s {_fmt(base)}) — not applied; retype '
                                                     f'it, or confirm in Notes → Wonka'})
                        continue

                if kind == 'x':
                    # The specialist typed x over this price = remove it: the one
                    # sanctioned blank write (G1 exception), only onto a cell that
                    # has a price.
                    if old_p is None:
                        x_ignored.append({'dish': name, 'variant': v})
                        continue
                    if (name, v) in absolute_cells:
                        absolute_guard.append({'dish': name, 'variant': v, 'report_price': 'x (remove)',
                                               'evidence': absolute_cells[(name, v)] + ' — removing this '
                                               'size would leave its modifier pricing behind'})
                        continue
                    # (ws.cell(..., value=None) would leave the value as is)
                    ws_di.cell(row=row_idx, column=ci+1).value = None
                    removed.append({'tab': 'dish_item', 'row': row_idx, 'col': ci + 1,
                                    'dish': name, 'variant': v, 'old': _fmt(old_p)})
                    code_ci = di_code_cols.get(v)
                    code = (_norm_code(ws_di.cell(row=row_idx, column=code_ci+1).value)
                            if code_ci is not None else '')
                    row_removed.append((v, code))
                    item_changed = True
                    continue

                if old_p is None:
                    # An empty live cell takes a price only when the specialist
                    # typed it into a cell that was empty in the report too.
                    if not (from_edit and typed_blank):
                        new_cell_drift.append({'dish': name, 'variant': v, 'report': _fmt(new_p)})
                        continue
                    if (name, v) in struct['new_size']:
                        absolute_guard.append({'dish': name, 'variant': v,
                                               'report_price': f'{_fmt(new_p)} (new size)',
                                               'evidence': struct['new_size'][(name, v)] + ' — a dish '
                                               'price here would be charged on top'})
                        continue
                    new_str = _money(new_p, _money_prefix(ws_di, row_idx, di_price_cols.values()) == '$')
                    ws_di.cell(row=row_idx, column=ci+1).value = new_str
                    added.append({'tab': 'dish_item', 'row': row_idx, 'col': ci + 1,
                                  'dish': name, 'variant': v, 'new': new_str})
                    row_added.append(v)
                    item_changed = True
                    continue
                if abs(old_p - new_p) < 0.001:
                    continue  # no change
                if abs(old_p) <= 0.001 and (name, v) in absolute_cells:
                    # G11: this $0.00 is by design — the price lives on
                    # modifiers; a dish price would be charged on top of it.
                    absolute_guard.append({'dish': name, 'variant': v, 'report_price': _fmt(new_p),
                                           'evidence': absolute_cells[(name, v)]})
                    continue

                ws_di.cell(row=row_idx, column=ci+1, value=_price_str(new_p, old_val))
                dish_cells += 1
                item_changed = True

            # an edit in a report size the live menu does not have (a template
            # column) cannot land — report it, never drop it silently
            for v in sorted(e4_price_variants - live_di_variants):
                edits_, _s, invalid_ = _split(insts, v, part1_changed, percent_cells)
                for r_, g_, k_, n_, _b, _c in edits_:
                    ignored_input.append({'tab': 'Dish Item', 'row': r_, 'name': name, 'group': g_,
                                          'variant': v, 'value': 'x' if k_ == 'x' else str(n_),
                                          'why': f"the live menu has no '{v}' size — the scripts "
                                                 f"cannot add a size column (ask Wonka)"})
                for r_, g_, raw_ in invalid_:
                    ignored_input.append({'tab': 'Dish Item', 'row': r_, 'name': name, 'group': g_,
                                          'variant': v, 'value': str(raw_), 'why': 'not a valid price'})

            # x removed this dish's last price: valid in wsheet, but taking a
            # dish off the menu is a strikethrough — warn, Wonka confirms.
            if row_removed and not any(not _is_blank(ws_di.cell(row=row_idx, column=c + 1).value)
                                       for c in di_price_cols.values()):
                no_price_left.append({'dish': name, 'removed': [v_ for v_, _c in row_removed]})

            # A removed size keeps its dish code (unit_code) unless the
            # specialist moved / removed it on the Dish Code tab (checked after
            # the code block below). Moved with x + a new price on the same row
            # → the code probably belongs on the new size. Wonka settles each.
            coded = [(v_, c_) for v_, c_ in row_removed if c_]
            if coded and len(row_removed) == 1 and len(row_added) == 1:
                pending_followups.append({'dish': name, 'code': coded[0][1], 'from': coded[0][0],
                                          'to': row_added[0], 'action': 'move'})
            else:
                for v_, c_ in coded:
                    pending_followups.append({'dish': name, 'code': c_, 'from': v_, 'to': '',
                                              'action': 'remove'})

            if item_changed:
                dish_changes += 1

        # ── Dish codes (the Dish Code tab) ──
        # Only a cell the specialist changed (≠ its hidden system-code baseline)
        # is an instruction; an untouched printed code is never written. A dish
        # under two groups is ONE dish: an edited copy wins, two copies edited
        # differently = conflict. G1's code exceptions: a code typed into an
        # empty code cell (recorded in `added` for to-tsv) and x (in `removed`).
        if name in e4_codes and name not in di_code_seen:
            di_code_seen.add(name)
            code_changed = False
            for v in sorted({v_ for _r, _g, c_ in e4_codes[name] for v_ in c_}):
                edits = []
                for r_, g_, c_ in e4_codes[name]:
                    raw, ref = c_.get(v, (None, None))
                    kind, code = _code_entry(raw)
                    if kind == 'blank' or (kind != 'invalid' and _norm_code(raw) == _norm_code(ref)):
                        continue                          # cleared or untouched: no instruction
                    if kind == 'invalid':
                        ignored_input.append({'tab': CODE_TAB, 'row': r_, 'name': name, 'group': g_,
                                              'variant': v, 'value': str(raw),
                                              'why': 'not a dish code (a line break, more than 128 '
                                                     'characters, a leading = (a formula), or only '
                                                     '# - . and spaces)'})
                        continue
                    edits.append((r_, g_, kind, code, _norm_code(ref)))
                if not edits:
                    continue
                if len({(k_, c_) for _r, _g, k_, c_, _f in edits}) > 1:
                    dup_conflicts.append({'tab': CODE_TAB, 'name': name, 'variant': v, 'rows': [
                        {'row': r_, 'group': g_, 'value': 'x' if k_ == 'x' else c_}
                        for r_, g_, k_, c_, _f in edits]})
                    continue
                r0, g0, kind, code, was = edits[0]
                ci = di_code_cols.get(v)
                if ci is None:
                    ignored_input.append({'tab': CODE_TAB, 'row': r0, 'name': name, 'group': g0, 'variant': v,
                                          'value': 'x' if kind == 'x' else code,
                                          'why': f"the live menu has no code column for '{v}'"})
                    continue
                live = _norm_code(ws_di.cell(row=row_idx, column=ci + 1).value)
                if live != was:
                    # the code changed in wsheet since Part 1 — never written over
                    codes_done['drift'].append({'dish': name, 'variant': v, 'report': 'x' if kind == 'x' else code,
                                                'live': live, 'was': was})
                    continue
                if kind == 'x':
                    if not live:
                        x_ignored.append({'dish': name, 'variant': v, 'what': 'code'})
                        continue
                    ws_di.cell(row=row_idx, column=ci + 1).value = None
                    removed.append({'tab': 'dish_item', 'row': row_idx, 'col': ci + 1, 'dish': name,
                                    'variant': v, 'old': live, 'kind': 'code'})
                    codes_done['removed'].append({'dish': name, 'variant': v, 'old': live})
                    code_changed = True
                    continue
                pci = di_price_cols.get(v)
                if pci is None or _is_blank(ws_di.cell(row=row_idx, column=pci + 1).value):
                    ignored_input.append({'tab': CODE_TAB, 'row': r0, 'name': name, 'group': g0, 'variant': v,
                                          'value': code, 'why': 'this size has no price — a dish code needs one'})
                    continue
                if code == live:
                    continue
                ws_di.cell(row=row_idx, column=ci + 1).value = code
                if not live:
                    added.append({'tab': 'dish_item', 'row': row_idx, 'col': ci + 1, 'dish': name,
                                  'variant': v, 'new': code, 'kind': 'code'})
                    codes_done['filled'].append({'dish': name, 'variant': v, 'new': code})
                else:
                    code_cells += 1
                    codes_done['corrected'].append({'dish': name, 'variant': v, 'old': live, 'new': code})
                code_changed = True
            if code_changed:
                code_items += 1

        # a code left on a removed / moved price is asked about only while it
        # is still there — the specialist may have moved / removed it above
        for f_ in pending_followups:
            cc_ = di_code_cols.get(f_['from'])
            if cc_ is not None and _norm_code(ws_di.cell(row=row_idx, column=cc_ + 1).value) == f_['code']:
                code_followups.append(f_)

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
    mi_first = {}       # item → its first live root row (the one the report shows)

    for row_idx in range(4, ws_mi.max_row + 1):
        row = [ws_mi.cell(row=row_idx, column=c+1).value for c in range(ws_mi.max_column)]

        if not _is_root_row(row):
            continue

        name = str(row[mi_name_col]).strip() if mi_name_col is not None and mi_name_col < len(row) and row[mi_name_col] else ''
        if not name or name not in e4_mod_prices:
            continue
        mi_seen.add(name)
        is_first = mi_first.setdefault(name, row_idx) == row_idx
        if has_intent and not is_first:
            # the report shows the FIRST live row of this name only; a same-name
            # twin is never written (its own price stays)
            if any((name, v) in mod_intent for v in mi_price_cols):
                twin_rows.append({'tab': 'modifier_item', 'name': name, 'row': row_idx})
            continue

        e4 = e4_mod_prices[name]
        item_changed = False

        for v, ci in mi_price_cols.items():
            new_raw = e4.get(v)
            if new_raw is None:
                continue
            old_val = ws_mi.cell(row=row_idx, column=ci+1).value
            old_p = _parse(old_val)
            intent = mod_intent.get((name, v))
            if has_intent and intent is None:
                # Not edited in the Modifier Summary: the Final Review holds Part
                # 1's snapshot — never written over the live price (a change
                # made in wsheet since Part 1 is kept, not reverted).
                snap = _parse(new_raw)
                if snap is not None and old_p is not None and abs(snap - old_p) > 0.001:
                    not_reverted.append({'tab': 'modifier_item', 'name': name, 'variant': v,
                                         'live': _fmt(old_p)})
                continue
            if _is_x(new_raw):
                # x in the Modifier Summary (action4 resolved it) = remove it
                if intent != 'x':
                    continue
                if old_p is None:
                    new_cell_drift.append({'dish': f'{name} (modifier)', 'variant': v, 'report': 'x'})
                    continue
                if (name, v) in struct['mod_options']:
                    absolute_guard.append({'dish': f'{name} (modifier)', 'variant': v,
                                           'report_price': 'x (remove)',
                                           'evidence': struct['mod_options'][(name, v)]
                                           + ' — removing it would make that choice free'})
                    continue
                ws_mi.cell(row=row_idx, column=ci+1).value = None
                removed.append({'tab': 'modifier_item', 'row': row_idx, 'col': ci + 1,
                                'dish': f'{name} (modifier)', 'item': name, 'variant': v,
                                'old': _fmt(old_p)})
                item_changed = True
                continue
            new_p = _parse(new_raw)
            if new_p is None:
                continue
            if old_p is None:
                # A new price only where action4 resolved one (typed into an
                # empty Modifier Summary cell), and only on the live row the
                # report shows — never into a same-name twin's empty cells (G1).
                if intent == 'price':
                    # edited in the report, but the live cell has lost its price
                    new_cell_drift.append({'dish': f'{name} (modifier)', 'variant': v,
                                           'report': _fmt(new_p)})
                if intent != 'new' or not is_first:
                    continue
                new_str = _money(new_p, _money_prefix(ws_mi, row_idx, mi_price_cols.values()) == '$')
                ws_mi.cell(row=row_idx, column=ci+1).value = new_str
                added.append({'tab': 'modifier_item', 'row': row_idx, 'col': ci + 1,
                              'dish': f'{name} (modifier)', 'item': name, 'variant': v,
                              'new': new_str})
                item_changed = True
                continue
            if abs(old_p - new_p) < 0.001:
                continue
            if intent == 'new':
                # typed into a cell that was EMPTY in the report, but wsheet has a
                # price there now — judged by what the specialist saw: not written
                new_cell_drift.append({'dish': f'{name} (modifier)', 'variant': v, 'report': _fmt(new_p),
                                       'live': _fmt(old_p),
                                       'note': 'the live menu now has a price the report did not show'})
                continue
            if new_p <= 0.001 and (name, v) in struct['mod_options']:
                absolute_guard.append({'dish': f'{name} (modifier)', 'variant': v,
                                       'report_price': _fmt(new_p),
                                       'evidence': struct['mod_options'][(name, v)]
                                       + ' — at $0.00 that choice would make the dish free'})
                continue

            ws_mi.cell(row=row_idx, column=ci+1, value=_price_str(new_p, old_val))
            mod_cells += 1
            item_changed = True

        if item_changed:
            mod_changes += 1

    for (item, v), kind in sorted(mod_intent.items()):
        if v not in mi_price_cols:
            ignored_input.append({'tab': 'Modifier Summary', 'name': item, 'variant': v,
                                  'value': kind, 'why': f"the live menu has no '{v}' size for "
                                                        f"modifiers — the scripts cannot add one"})

    # ── Save ──────────────────────────────────────────────────────
    wb_raw.save(out)
    wb_raw.close()

    # ── Drift detection ───────────────────────────────────────────
    # Report names present in the report (Excel 4) but NOT matched to any live
    # row are content drift: the dish/modifier was renamed or removed in wsheet
    # between Part 1 and Part 2, so its edit silently never landed.
    drift_price = sorted(set(e4_prices) - di_price_seen)
    drift_code = []     # a dish on the Dish Code tab is on Dish Item too — counted under drift_price
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
    rerun = 'action5' in _existing
    _existing['r_number'] = r_number
    # A new Part 2 result: drop a previous run's modifier_rule block (apply,
    # which always runs after action5, writes its own) so part2_report never
    # shows stale counts.
    _existing.pop('modifier_rule', None)
    _existing['action5'] = {
        'dish_item_items_changed': dish_changes,
        'dish_item_cells': dish_cells,
        'dish_code_enabled': has_code_tab,
        'dish_code_items_changed': code_items,
        'dish_code_cells': code_cells,
        'codes': {k: {'total': len(v_), 'items': v_[:200]} for k, v_ in codes_done.items()},
        'modifier_item_items_changed': mod_changes,
        'modifier_item_cells': mod_cells,
        'total_cells_written': dish_cells + code_cells + mod_cells + len(removed) + len(added),
        'drift': {
            'total': drift_total,
            'dish_price': drift_price,
            'dish_code': drift_code,
            'modifier': drift_mod,
        },
        'absolute_guard': {
            'total': len(absolute_guard),
            'cells': absolute_guard,
        },
        'removed': {
            'total': len(removed),
            'cells': removed,
            'x_on_empty_cell_ignored': x_ignored
            + [{'dish': f"{c['item']} (modifier)", 'variant': c['variant']}
               for c in a4_notes.get('x_empty', [])],
        },
        'added': {
            'total': len(added),
            'cells': added,
        },
        'ignored_input': {
            'total': len(ignored_input) + len(a4_notes.get('ignored', [])),
            'cells': ignored_input + [{'tab': 'Modifier Summary', 'name': c['item'],
                                       'variant': c['variant'], 'value': c['value'],
                                       'why': c['detail']} for c in a4_notes.get('ignored', [])],
        },
        'duplicates': {
            'resolved_from_edited_copy': dup_resolved,
            'conflicts': dup_conflicts + [{'tab': 'Modifier Summary', 'name': c['item'],
                                           'variant': c['variant'], 'detail': c['detail']}
                                          for c in a4_notes.get('conflict', [])],
        },
        'new_cell_drift': new_cell_drift,
        'no_price_left': no_price_left,
        'twin_rows_skipped': twin_rows,
        'rerun_in_same_dir': rerun,
        'not_reverted': {
            'total': len(not_reverted),
            'cells': not_reverted[:200],
        },
        'code_followups': {
            'total': len(code_followups),
            'items': code_followups,
        },
        'output_xlsx': out,
    }
    with open(_summary_path, 'w', encoding='utf-8') as _f:
        _json.dump(_existing, _f, indent=2)

    # ── Print summary ─────────────────────────────────────────────
    print(f"\n{'='*50}")
    if rerun:
        print("  !! Part 2 already ran in this directory (part2_summary.json had a result). If its "
              "TSVs were patched (to-tsv), this raw already carries those changes — counts below are "
              "relative to it, and an x you undid shows as drift. Re-fetch into a FRESH directory and "
              "re-run from step 1 instead (part2.md 4b).")
    print(f"  dish_item tab: {dish_changes} items changed, {dish_cells} cells updated")
    if added:
        print(f"  NEW prices in empty cells (typed by the specialist): {len(added)}")
        for x in added:
            print(f"       + {x['dish']} [{x['variant']}] = {x['new']}")
    if removed:
        print(f"  prices REMOVED (x in the report): {len(removed)}")
        for x in removed:
            print(f"       - {x['dish']} [{x['variant']}] was {x['old']} → cleared")
    for x in x_ignored:
        print(f"  NOTE: x on {x['dish']} [{x['variant']}] ignored — that cell has no "
              f"{'code' if x.get('what') == 'code' else 'price'} to remove")
    for c in a4_notes.get('x_empty', []):
        print(f"  NOTE: x on {c['item']} (modifier) [{c['variant']}] ignored — no price to remove")
    if dup_resolved:
        print(f"  Dish listed under several groups — the EDITED copy was used "
              f"(untouched copies ignored): {len(dup_resolved)} cell(s)")
    conflicts_all = dup_conflicts + a4_notes.get('conflict', [])
    if conflicts_all:
        print(f"\n  !! CONFLICT (NOT applied): one dish / modifier edited to different prices in "
              f"two of its rows — ask the specialist which is right: {len(conflicts_all)}")
        for c in dup_conflicts:
            print(f"       - {c['name']} [{c['variant']}]: "
                  + '; '.join(f"row {x['row']} ({x['group']}) {x['value']}" for x in c['rows']))
        for c in a4_notes.get('conflict', []):
            print(f"       - {c['item']} (modifier) [{c['variant']}]: {c['detail']}")
    ign_all = ignored_input + a4_notes.get('ignored', [])
    if ign_all:
        print(f"\n  !! NOT A PRICE — ignored (never applied): {len(ign_all)}")
        for c in ignored_input:
            print(f"       - Dish Item row {c['row']} {c['name']} [{c['variant']}]: "
                  f"{c['value']!r} — {c['why']}")
        for c in a4_notes.get('ignored', []):
            print(f"       - {c['item']} [{c['variant']}]: {c['value']!r} — {c['detail']}")
    if not_reverted:
        print(f"  Live prices that changed since Part 1 (untouched in the report — kept, NOT "
              f"reverted): {len(not_reverted)}")
    for t in twin_rows:
        print(f"  NOTE: {t['tab']} row {t['row']} '{t['name']}' is a second live row with that name "
              f"(availability / twin) — not written; the report's edits went to the first row")
    for c in no_price_left:
        print(f"  !! {c['dish']}: x removed its LAST price ({', '.join(c['removed'])}) — it is left with "
              f"no price. To take a dish off the menu, strike its name instead; confirm with the specialist.")
    if new_cell_drift:
        print(f"\n  !! DRIFT: {len(new_cell_drift)} report price(s) NOT applied — the live cell changed "
              f"in wsheet since Part 1:")
        for c in new_cell_drift:
            why = (f"typed into an empty cell, but live now has {c['live']} — not overwritten" if c.get('live')
                   else 'the live cell has no price any more')
            print(f"       - {c['dish']} [{c['variant']}] report {c['report']} — {why}")
    if code_followups:
        print(f"\n  !! DISH CODES to settle (the scripts never write codes): {len(code_followups)}")
        for c in code_followups:
            if c['action'] == 'move':
                print(f"       - {c['dish']}: code '{c['code']}' stays on [{c['from']}] (price moved to "
                      f"[{c['to']}]) — move the code to [{c['to']}]?")
            else:
                print(f"       - {c['dish']}: code '{c['code']}' stays on [{c['from']}], whose price was "
                      f"removed — remove the code too?")
    if has_code_tab:
        print(f"  dish_item codes (Dish Code tab): {len(codes_done['filled'])} filled, "
              f"{len(codes_done['corrected'])} corrected, {len(codes_done['removed'])} removed "
              f"({code_items} dishes)")
        for c in codes_done['drift']:
            print(f"  !! CODE DRIFT: {c['dish']} [{c['variant']}] — the code changed in wsheet since Part 1 "
                  f"('{c['was']}' → '{c['live']}'); the report's '{c['report']}' was not written")
    else:
        print(f"  dish_item codes: no Dish Code tab in this report — none written")
    print(f"  modifier_item tab: {mod_changes} items changed, {mod_cells} cells updated")
    print(f"  Total cells written: {dish_cells + code_cells + mod_cells + len(removed) + len(added)}")
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

    if absolute_guard:
        print(f"\n  !! ABSOLUTE GUARD (G11): {len(absolute_guard)} edit(s) NOT written — the dish's "
              f"price lives on its modifiers; a dish price would be charged on top (double "
              f"charge), or the change would make it free. Kept as it was:")
        for g in absolute_guard:
            print(f"       - {g['dish']} [{g['variant']}]: report {g['report_price']} — {g['evidence']}")
        print(f"     Put the base on the By Itself rule (Modifier Rule tab) or the option prices "
              f"(Modifier Summary) instead.")

    print(f"\n  Output: {out}")
    print(f"  Summary JSON: {_summary_path}  (source of truth for part2_report.py)")
    print(f"{'='*50}")
    print("Done.")


if __name__ == '__main__':
    main()