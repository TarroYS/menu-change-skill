#!/usr/bin/env python3
"""
action1_raw_to_2a.py — Convert Raw Excel 1 → Excel 2a
======================================================

Standalone script. No external dependencies beyond openpyxl.
Works for any restaurant cuisine in the Wonders system — price columns are
discovered from the Raw Gsheet's row-2 `unit_price` markers (not from a
hardcoded variant whitelist).

Usage:
    python action1_raw_to_2a.py <raw_excel_1.xlsx> [--output <output.xlsx>]

If --output is omitted, auto-names from the R-number (e.g., R88312_Excel_2a.xlsx).

What it does:
    Reads the complex multi-tab Raw Excel 1 system spreadsheet and produces
    a simplified Excel 2a with:
      - price tab: Dish Group, Dish Name:Primary, Dish Name:Secondary, + price variant columns
      - System Codes tab: same row layout as price, cells carry each size's dish
        code (unit_code) — action3 builds the report's Dish Code tab from it.
        (Deliberately not named 'dish_code': that name would switch on action2's
        dormant code tie-break in name matching.)
      - modifier tab: Dish Group, Modifier Group, Modifier Name:Primary/Secondary, + price variants
      - absolute_pricing tab (v4.4): $0.00 dish cells whose price lives on
        modifiers (priced modifier_rule rows for the dish or for its dish group,
        or a required, fully priced modifier group — and, in a restaurant that
        prices dishes on modifier rules, every $0.00 Fried Specialties dish).
        action2 never writes a paper price onto them and price_inference never
        infers one (G11 — it would double-charge).
      - side_rules tab: dishes with priced dish-level modifier_rule rows (side
        combos). action2 warns when the OCR JSON reads such a dish without a
        multi-column (side_pricing) table — no Fried Specialty tab otherwise.

Inclusion rules (discovered by reverse-engineering the Raw Excel 1 → Excel 2a relationship):
    Price tab items must:
      1. Exist in the dish_group tab (have a group assignment)
      2. Have exactly 1 root row in dish_item tab (no duplicates)
      3. Have at least 1 non-empty price in any variant column

    Modifier tab:
      - Side choice groups (any group starting with "Select ") are replicated per dish_group
      - Other modifier groups appear once as standalone sections
      - Standalone groups must be referenced in dish_group tab
      - Modifier items must have at least 1 non-empty price
      - Orphan modifiers (not in any group but have prices) appear at the end

    All items within a section are sorted reverse-alphabetically.

Raw Excel 1 format assumptions (constant across restaurants):
    - Row 1: entity types, Row 2: property paths, Row 3: column names, Row 4+: data
    - Root rows have JSON string starting with '{' in column 0
    - Column positions vary per restaurant — always search by header text (G4)
"""

import argparse
import re
import sys
from collections import OrderedDict, Counter, defaultdict
from typing import Dict, List, Optional, Tuple

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import PatternFill, Font
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl is required. Install with: pip install openpyxl")
    sys.exit(1)


# ── Constants ──────────────────────────────────────────────────────

PRICE_VARIANTS = [
    'Order', 'Specialty', 'Small', 'Large', 'Lunch', 'Combo',
    'Party Tray', '2 Ounce', '8 Ounce', '13 Ounce', '15 Ounce',
    '16 Ounce', 'Can', 'Bottle'
]

SIDE_CHOICE_PREFIX = 'Select '  # Modifier groups starting with "Select " are side choice groups

# Base-price anchor names on an ABSOLUTE dish's modifier_rule rows (whole name,
# normalized). Keep in sync with modifier_rule_pricing.py BY_ITSELF_ALIASES.
BY_ITSELF_ALIASES = {'by itself', 'plain', 'no rice', 'no side'}

HEADER_FILL = PatternFill('solid', fgColor='D9E1F2')
HEADER_FONT = Font(bold=True, size=11)


# ── Helpers ────────────────────────────────────────────────────────

def _parse_price(val) -> Optional[float]:
    """Parse '$7.50', 7.5, or similar into a float. Returns None for empty/invalid."""
    if val is None or str(val).strip() in ('', 'None', 'NULL'):
        return None
    s = str(val).strip().replace('$', '').replace(',', '')
    try:
        return float(s)
    except ValueError:
        return None


def _fmt(val: Optional[float]) -> str:
    """Format float as '$X.XX'. Returns '' for None."""
    return f'${val:.2f}' if val is not None else ''


def _is_root_row(row) -> bool:
    """Check if a data row is a root row (has JSON id in column 0)."""
    val = str(row[0]).strip() if row[0] else ''
    return val.startswith('{')


def _find_unit_price_cols(headers: list, row2: list) -> Dict[str, int]:
    """
    Find price variant column indices in the Raw Gsheet.

    A column is a price column iff row 2 contains 'unit_price' for that
    column. The header name is used as the variant key in the output.

    Older versions of this function gated discovery by a hardcoded
    PRICE_VARIANTS whitelist plus a narrow regex for "N Ounce/Pieces/
    Liter". That silently dropped sushi-specific columns ('Roll',
    'Regular Roll', 'Bowl'), bento ('Dinner'), combo sets ('Seapark'),
    etc. — even though the Raw Gsheet already marked them with
    'unit_price' in row 2. The 'unit_price' marker is the source of
    truth; trust it.

    Columns with empty/duplicate headers are skipped defensively.
    """
    cols = {}
    seen = set()
    for i, h in enumerate(headers):
        h_str = str(h).strip() if h else ''
        if not h_str:
            continue
        if i >= len(row2):
            continue
        marker = str(row2[i]) if row2[i] is not None else ''
        if 'unit_price' not in marker:
            continue
        if h_str in seen:
            continue
        seen.add(h_str)
        cols[h_str] = i
    return cols


def _find_unit_code_cols(headers: list, row2: list) -> Dict[str, int]:
    """
    Find dish-code variant column indices in the Raw Gsheet.

    Mirrors _find_unit_price_cols, but keys off row-2 'unit_code' markers
    instead of 'unit_price'. Each variant slot in the Raw Gsheet that carries
    a per-variant dish code (path 'variant.wprop.unit_code.{slot}') has
    'unit_code' in row 2 for that column; the row-3 header is the variant key.

    The 'unit_code' marker is the source of truth — no hardcoded whitelist.
    Columns with empty/duplicate headers are skipped defensively.
    """
    cols = {}
    seen = set()
    for i, h in enumerate(headers):
        h_str = str(h).strip() if h else ''
        if not h_str:
            continue
        if i >= len(row2):
            continue
        marker = str(row2[i]) if row2[i] is not None else ''
        if 'unit_code' not in marker:
            continue
        if h_str in seen:
            continue
        seen.add(h_str)
        cols[h_str] = i
    return cols


def _parse_code(val) -> Optional[str]:
    """Parse a dish-code cell into a clean string. Returns None for empty."""
    if val is None:
        return None
    if isinstance(val, float) and val.is_integer():
        val = int(val)                  # a code stored as a number: 48, not '48.0'
    s = str(val).strip()
    if s in ('', 'None', 'NULL'):
        return None
    return s


def _find_col(headers: list, name: str) -> Optional[int]:
    """Find the first column index matching a header name."""
    for i, h in enumerate(headers):
        if h == name:
            return i
    return None


def _find_col_with_path(headers: list, row2: list, name: str, path_contains: str) -> Optional[int]:
    """Find a column by header name AND property path substring."""
    for i, h in enumerate(headers):
        if h == name and i < len(row2) and path_contains in row2[i]:
            return i
    return None


def _is_fried_specialty_group(group: str) -> bool:
    """'Fried Specialties', 'Fried Specialty', 'Fried Specialties::RPOS' …
    (the same test modifier_rule_pricing.py uses)."""
    return bool(group) and 'fried special' in group.lower()


def _is_by_itself(name: str) -> bool:
    """True if a modifier_item name IS a base-price anchor alias (whole name)."""
    s = str(name).lower().strip()
    s = re.sub(r'\s*\([^)]*\)\s*', ' ', s)
    s = s.replace('w.', 'with').replace('&', 'and')
    s = re.sub(r'[^a-z0-9\s]', '', s)
    return re.sub(r'\s+', ' ', s).strip() in BY_ITSELF_ALIASES


# ── Raw Excel 1 Reader ────────────────────────────────────────────

class RawExcel1Reader:
    """Reads and parses all relevant tabs from Raw Excel 1."""

    def __init__(self, path: str):
        self.path = path
        self.wb = load_workbook(path, read_only=True)
        self._validate_tabs()
        self._read_identity()
        self._read_dish_group()
        self._read_dish_item()
        self._read_modifier_group()
        self._read_modifier_item()
        self._read_modifier_rule()

    def close(self):
        self.wb.close()

    def _validate_tabs(self):
        required = ['dish_group', 'dish_item', 'modifier_group', 'modifier_item']
        missing = [t for t in required if t not in self.wb.sheetnames]
        if missing:
            print(f"ERROR: Raw Excel 1 is missing required tabs: {missing}")
            print(f"  Available tabs: {self.wb.sheetnames[:10]}...")
            sys.exit(1)

    def _get_rows_and_headers(self, tab_name: str):
        """Read a tab and return (all_rows, header_row3, property_path_row2)."""
        rows = list(self.wb[tab_name].iter_rows(values_only=True))
        if len(rows) < 4:
            return rows, [], []
        headers = [str(c).strip() if c else '' for c in rows[2]]
        row2 = [str(c).strip() if c else '' for c in rows[1]]
        return rows, headers, row2

    # ── Identity (G7) ─────────────────────────────────────────────

    def _read_identity(self):
        self.r_number = None
        self.e_number = None
        self.restaurant_name = None
        self.address = None

        if 'restaurant_location' in self.wb.sheetnames:
            rows, headers, _ = self._get_rows_and_headers('restaurant_location')
            if len(rows) >= 4:
                data = rows[3]
                for i, h in enumerate(headers):
                    if i >= len(data):
                        break
                    val = str(data[i]).strip() if data[i] else None
                    if h == 'r_number' and val:
                        self.r_number = val
                    elif h == 'address1' and val:
                        self.address = val

        if 'restaurant_enterprise' in self.wb.sheetnames:
            rows, headers, _ = self._get_rows_and_headers('restaurant_enterprise')
            if len(rows) >= 4:
                data = rows[3]
                for i, h in enumerate(headers):
                    if i >= len(data):
                        break
                    val = str(data[i]).strip() if data[i] else None
                    if h == 'e_number' and val:
                        self.e_number = val
                    elif h == 'restaurant_enterprise' and val:
                        self.restaurant_name = val

    # ── dish_group tab ────────────────────────────────────────────

    def _read_dish_group(self):
        rows, headers, _ = self._get_rows_and_headers('dish_group')

        col_group = _find_col(headers, 'dish_group')
        col_dish = _find_col(headers, 'dish_item')
        col_mg = _find_col(headers, 'modifier_group')

        if col_group is None or col_dish is None:
            print(f"ERROR: dish_group tab missing 'dish_group' or 'dish_item' column")
            sys.exit(1)

        self.group_order = []                   # ordered group names
        self.group_dishes = OrderedDict()       # group → [ordered dish names]
        self.group_modifier_groups = defaultdict(list)  # group → [modifier_group names]
        self.dish_to_group = {}                 # dish_name → group_name

        current = None
        for r in rows[3:]:
            is_root = _is_root_row(r)
            group = str(r[col_group]).strip() if col_group < len(r) and r[col_group] else ''
            dish = str(r[col_dish]).strip() if col_dish < len(r) and r[col_dish] else ''
            mg = str(r[col_mg]).strip() if col_mg is not None and col_mg < len(r) and r[col_mg] else ''

            if is_root and group:
                current = group
                if current not in self.group_dishes:
                    self.group_order.append(current)
                    self.group_dishes[current] = []

            if dish and current:
                if dish not in self.group_dishes[current]:
                    self.group_dishes[current].append(dish)
                self.dish_to_group[dish] = current

            if mg and current and mg not in self.group_modifier_groups[current]:
                self.group_modifier_groups[current].append(mg)

    # ── dish_item tab ─────────────────────────────────────────────

    def _read_dish_item(self):
        rows, headers, row2 = self._get_rows_and_headers('dish_item')

        col_name = _find_col(headers, 'dish_item')
        col_zh = _find_col_with_path(headers, row2, 'i18n_names:zh_cn', 'root_wprop.name_complex')
        price_cols = _find_unit_price_cols(headers, row2)
        code_cols = _find_unit_code_cols(headers, row2)
        # The dish's own modifier groups ('comprise → modifier_group'; a second
        # 'modifier_group' header under sort_coordinate is only an ordering hint).
        mg_cols = [i for i, h in enumerate(headers) if h == 'modifier_group']
        col_mg = next((i for i in mg_cols if i < len(rows[0])
                       and str(rows[0][i] or '').strip() == 'comprise'),
                      mg_cols[0] if mg_cols else None)

        if col_name is None:
            print("ERROR: dish_item tab missing 'dish_item' column")
            sys.exit(1)

        self.dish_prices = {}       # name → {variant: float}
        self.dish_codes = {}        # name → {variant: str}
        self.dish_zh = {}           # name → chinese name
        self.dish_root_counts = Counter()
        self.dish_modifier_groups = defaultdict(list)  # name → [modifier_group]

        current = None
        for r in rows[3:]:
            if _is_root_row(r):
                current = str(r[col_name]).strip() if col_name < len(r) and r[col_name] else ''
            mg = str(r[col_mg]).strip() if col_mg is not None and col_mg < len(r) and r[col_mg] else ''
            owner = (str(r[col_name]).strip() if col_name < len(r) and r[col_name] else '') or current
            if mg and owner and mg not in self.dish_modifier_groups[owner]:
                self.dish_modifier_groups[owner].append(mg)
            if not _is_root_row(r):
                continue
            name = str(r[col_name]).strip() if col_name < len(r) and r[col_name] else ''
            if not name:
                continue

            self.dish_root_counts[name] += 1
            if name in self.dish_prices:
                continue  # take first occurrence

            # Chinese name
            zh = ''
            if col_zh is not None and col_zh < len(r) and r[col_zh]:
                zh = str(r[col_zh]).strip()
                if zh == 'NULL':
                    zh = ''
            self.dish_zh[name] = zh

            # Prices
            prices = {}
            for v, ci in price_cols.items():
                p = _parse_price(r[ci]) if ci < len(r) else None
                if p is not None:
                    prices[v] = p
            self.dish_prices[name] = prices

            # Dish codes (per variant slot)
            codes = {}
            for v, ci in code_cols.items():
                c = _parse_code(r[ci]) if ci < len(r) else None
                if c is not None:
                    codes[v] = c
            self.dish_codes[name] = codes

    # ── modifier_group tab ────────────────────────────────────────

    def _read_modifier_group(self):
        rows, headers, _ = self._get_rows_and_headers('modifier_group')

        col_mg = _find_col(headers, 'modifier_group')
        col_mi = _find_col(headers, 'modifier_item')
        col_min = _find_col(headers, 'min_number_of_selections')
        col_sub = _find_col(headers, 'subtype')

        self.mg_required = set()        # groups the customer MUST pick from
        if col_mg is None or col_mi is None:
            print("WARNING: modifier_group tab missing columns, modifier tab will be empty")
            self.mg_items = OrderedDict()
            self.mod_to_group = {}
            return

        self.mg_items = OrderedDict()   # modifier_group → [ordered modifier_item names]
        current = None
        for r in rows[3:]:
            is_root = _is_root_row(r)
            mg = str(r[col_mg]).strip() if col_mg < len(r) and r[col_mg] else ''
            mi = str(r[col_mi]).strip() if col_mi < len(r) and r[col_mi] else ''

            if is_root and mg:
                current = mg
                if current not in self.mg_items:
                    self.mg_items[current] = []
                mn = _parse_price(r[col_min]) if col_min is not None and col_min < len(r) else None
                sub = str(r[col_sub]).strip().lower() if col_sub is not None and col_sub < len(r) and r[col_sub] else ''
                if (mn is not None and mn >= 1) or sub == 'required':
                    self.mg_required.add(current)

            if mi and current and mi not in self.mg_items[current]:
                self.mg_items[current].append(mi)

        # Reverse lookup
        self.mod_to_group = {}
        for mg, items in self.mg_items.items():
            for mi in items:
                if mi not in self.mod_to_group:
                    self.mod_to_group[mi] = mg

    # ── modifier_item tab ─────────────────────────────────────────

    def _read_modifier_item(self):
        rows, headers, row2 = self._get_rows_and_headers('modifier_item')

        col_name = _find_col(headers, 'modifier_item')
        col_zh = _find_col_with_path(headers, row2, 'i18n_names:zh_cn', 'root_wprop.name_complex')
        price_cols = _find_unit_price_cols(headers, row2)

        if col_name is None:
            print("WARNING: modifier_item tab missing 'modifier_item' column, modifier tab will be empty")
            self.mod_prices = {}
            self.mod_zh = {}
            self.mod_root_counts = Counter()
            return

        self.mod_prices = {}
        self.mod_zh = {}
        self.mod_root_counts = Counter()
        self.mod_item_order = []  # Tracks root row order from modifier_item tab

        for r in rows[3:]:
            if not _is_root_row(r):
                continue
            name = str(r[col_name]).strip() if col_name < len(r) and r[col_name] else ''
            if not name:
                continue

            self.mod_root_counts[name] += 1
            if name not in self.mod_prices:
                self.mod_item_order.append(name)
            if name in self.mod_prices:
                continue

            zh = ''
            if col_zh is not None and col_zh < len(r) and r[col_zh]:
                zh = str(r[col_zh]).strip()
                if zh == 'NULL':
                    zh = ''
            self.mod_zh[name] = zh

            prices = {}
            for v, ci in price_cols.items():
                p = _parse_price(r[ci]) if ci < len(r) else None
                if p is not None:
                    prices[v] = p
            self.mod_prices[name] = prices

    # ── modifier_rule tab (optional) ──────────────────────────────

    def _read_modifier_rule(self):
        """Priced modifier_rule rows, by header text (G4): dish-level rows (they
        name a dish_item) and group-level rows (only a dish_group — they apply
        to every dish in that group). Used to find $0.00 dishes whose price
        lives on modifier rules (ABSOLUTE) and to list the side-rule dishes."""
        self.mr_rules = []              # [{'dish','variant','mod','mg','price'}]
        self.mr_group_rules = []        # [{'group','variant','mod','mg','price'}]
        if 'modifier_rule' not in self.wb.sheetnames:
            return
        rows, headers, _ = self._get_rows_and_headers('modifier_rule')
        cols = {k: _find_col(headers, k) for k in
                ('out_unit_price', 'dish_group', 'dish_item', 'variant', 'modifier_group',
                 'modifier_item')}
        if cols['out_unit_price'] is None or cols['dish_item'] is None:
            return

        def _v(r, k):
            ci = cols[k]
            return str(r[ci]).strip() if ci is not None and ci < len(r) and r[ci] else ''

        for r in rows[3:]:
            dish, group = _v(r, 'dish_item'), _v(r, 'dish_group')
            price = _parse_price(r[cols['out_unit_price']]) if cols['out_unit_price'] < len(r) else None
            if price is None:
                continue
            rule = {'variant': _v(r, 'variant'), 'mod': _v(r, 'modifier_item'),
                    'mg': _v(r, 'modifier_group'), 'price': price}
            if dish:
                self.mr_rules.append({'dish': dish, **rule})
            elif group:
                self.mr_group_rules.append({'group': group, **rule})


# ── Excel 2a Generator ────────────────────────────────────────────

class Excel2aGenerator:
    """Generates Excel 2a from a RawExcel1Reader."""

    def __init__(self, reader: RawExcel1Reader):
        self.r = reader

    def _get_price_tab_items(self) -> List[Tuple[str, str, str, Dict]]:
        """
        Returns ordered list of (group, name, zh_name, prices) for the price tab.

        Inclusion rules:
          1. Item exists in dish_group tab
          2. Item has exactly 1 root row in dish_item
          3. Item has at least 1 non-empty price
        Sort: reverse-alphabetical within each group.
        """
        result = []
        for group in self.r.group_order:
            candidates = []
            for dish in self.r.group_dishes[group]:
                if self.r.dish_root_counts.get(dish, 0) != 1:
                    continue
                prices = self.r.dish_prices.get(dish, {})
                if not prices:
                    continue
                zh = self.r.dish_zh.get(dish, '')
                candidates.append((group, dish, zh, prices))
            candidates.sort(key=lambda x: x[1], reverse=True)
            result.extend(candidates)
        return result

    def _get_dish_code_tab_items(self) -> List[Tuple[str, str, str, Dict]]:
        """
        Returns ordered list of (group, name, zh_name, codes) for the dish_code tab.

        Mirrors the price tab row-for-row so the two tabs align. Inclusion
        rules are identical to _get_price_tab_items (in dish_group tab,
        exactly 1 root row, at least 1 non-empty price); the cells carry the
        per-variant dish codes (possibly blank). This generates for every
        restaurant regardless of whether any codes are present.
        """
        result = []
        for group in self.r.group_order:
            candidates = []
            for dish in self.r.group_dishes[group]:
                if self.r.dish_root_counts.get(dish, 0) != 1:
                    continue
                prices = self.r.dish_prices.get(dish, {})
                if not prices:
                    continue
                zh = self.r.dish_zh.get(dish, '')
                codes = self.r.dish_codes.get(dish, {})
                candidates.append((group, dish, zh, codes))
            candidates.sort(key=lambda x: x[1], reverse=True)
            result.extend(candidates)
        return result

    def _get_filtered_mod_items(self, mg_name: str) -> List[Tuple[str, str, Dict]]:
        """Get modifier items for a group: must have prices, sorted by modifier_item tab order."""
        items = self.r.mg_items.get(mg_name, [])
        # Build order index from modifier_item tab root row order
        order_idx = {name: i for i, name in enumerate(self.r.mod_item_order)}
        filtered = []
        for mi in items:
            prices = self.r.mod_prices.get(mi, {})
            if not prices:
                continue
            zh = self.r.mod_zh.get(mi, '')
            filtered.append((mi, zh, prices))
        filtered.sort(key=lambda x: order_idx.get(x[0], 999999))
        return filtered

    def _get_modifier_tab_rows(self) -> list:
        """
        Returns ordered list of modifier row dicts for the modifier tab.

        Structure:
          Part 1: Side choice groups (Select *) interleaved per dish_group
                  For each dish_group: Select Rice, Select Sauce, Select Side (alphabetical)
          Part 2: Standalone modifier groups (reverse-alpha, must be in dish_group tab)
          Part 3: Orphan modifiers (not in any group, have prices)
        """
        rows = []

        # Collect all modifier groups referenced in dish_group tab
        all_dg_referenced_mgs = set()
        for mgs in self.r.group_modifier_groups.values():
            all_dg_referenced_mgs.update(mgs)

        # Part 1: Side choice groups interleaved per dish_group
        # Detect side choice groups dynamically: any modifier group starting with "Select "
        side_choice_set = set(mg for mg in self.r.mg_items if mg.startswith(SIDE_CHOICE_PREFIX))

        # Iterate dish groups first, then Select * groups in dish_group tab reference order
        for dg in self.r.group_order:
            dg_mgs = self.r.group_modifier_groups.get(dg, [])
            for sc_group in dg_mgs:
                if sc_group not in side_choice_set:
                    continue
                sc_items = self._get_filtered_mod_items(sc_group)
                if not sc_items:
                    continue
                # Separator row for this (dish_group, side_choice) section
                rows.append({'dg': dg, 'mg': '', 'name': '', 'zh': '', 'prices': {}, 'is_sep': True})
                for name, zh, prices in sc_items:
                    rows.append({'dg': dg, 'mg': sc_group, 'name': name, 'zh': zh,
                                 'prices': prices, 'is_sep': False})

        # Part 2: Standalone modifier groups
        # Order follows modifier_group tab root row order (self.r.mg_items is OrderedDict)
        # Filtered to: not side-choice, must be referenced in dish_group tab
        standalone = [
            mg for mg in self.r.mg_items
            if not mg.startswith(SIDE_CHOICE_PREFIX) and mg in all_dg_referenced_mgs
        ]
        first_standalone = True
        for mg_name in standalone:
            items = self._get_filtered_mod_items(mg_name)
            if not items:
                continue
            if first_standalone:
                # First standalone section gets an empty separator row
                rows.append({'dg': '', 'mg': '', 'name': '', 'zh': '', 'prices': {}, 'is_sep': True})
                first_standalone = False
            for name, zh, prices in items:
                rows.append({'dg': mg_name, 'mg': mg_name, 'name': name, 'zh': zh,
                             'prices': prices, 'is_sep': False})

        # Part 3: Orphan modifiers (not in any modifier_group, have prices)
        # No separator — flows directly after standalone groups
        all_grouped_items = set()
        for items in self.r.mg_items.values():
            all_grouped_items.update(items)
        orphans = []
        for mi_name, prices in self.r.mod_prices.items():
            if mi_name not in all_grouped_items and prices:
                zh = self.r.mod_zh.get(mi_name, '')
                orphans.append((mi_name, zh, prices))
        if orphans:
            orphans.sort(key=lambda x: x[0], reverse=True)
            for name, zh, prices in orphans:
                rows.append({'dg': '', 'mg': '', 'name': name, 'zh': zh,
                             'prices': prices, 'is_sep': False})

        return rows

    def _get_absolute_cells(self, dishes) -> List[Tuple[str, str, str, str, Optional[float], str]]:
        """$0.00 dish prices whose real price lives on modifiers (ABSOLUTE
        pricing). A paper price written onto one of these cells is charged ON
        TOP of the modifier price — the v4.3 double charge (wings $7.99 on the
        dish + the wings-with-fries rule's full $9.39 = $17.38) — so Part 1
        never writes one (G11). Two structures, both read from the raw:
          modifier_rule  — priced dish-level modifier_rule rows for this dish and
                           variant (or variant-less): By Itself carries the
                           base, each side rule the full total.
          modifier_group — a REQUIRED modifier group on the dish or its dish
                           group whose every item is priced for this variant
                           (e.g. R85325 Seafood Boil $0.00 → Select Seafood Boil).
        Returns [(group, dish, variant, source, by_itself_price, evidence,
        modifier_group)], one per (dish, variant), in Menu Preview order."""
        dish_groups = defaultdict(list)
        for g, ds in self.r.group_dishes.items():
            for d in ds:
                dish_groups[d].append(g)
        def _shown(rules):
            txt = ', '.join(f"{x['mod'] or x['mg'] or '?'} {_fmt(x['price'])}" for x in rules[:4])
            return txt + (f' (+{len(rules) - 4} more)' if len(rules) > 4 else '')

        out, seen = [], set()
        for group, dish, _zh, prices in dishes:
            for v, p in prices.items():
                if abs(p) > 0.001 or (dish, v) in seen:
                    continue
                rules = [x for x in self.r.mr_rules
                         if x['dish'] == dish and x['variant'] in ('', v) and x['price'] > 0.001]
                if rules:
                    by_itself = next((x['price'] for x in rules if _is_by_itself(x['mod'])), None)
                    out.append((group, dish, v, 'modifier_rule', by_itself,
                                f'{len(rules)} priced modifier_rule row(s): {_shown(rules)}', ''))
                    seen.add((dish, v))
                    continue
                # Group-level rules (no dish_item, only a dish_group) price every
                # dish of that group the same way — but only a priced By Itself,
                # or rules on a REQUIRED choice, mean the price lives there: a
                # group-level add-on ("Sauce / With +$1.00", real R85325) is
                # just an upcharge.
                grules = [x for x in getattr(self.r, 'mr_group_rules', [])
                          if x['group'] in dish_groups.get(dish, []) and x['variant'] in ('', v)
                          and x['price'] > 0.001]
                if grules and self._holds_price(grules):
                    by_itself = next((x['price'] for x in grules if _is_by_itself(x['mod'])), None)
                    out.append((group, dish, v, 'modifier_rule', by_itself,
                                f"{len(grules)} priced group-level modifier_rule row(s) on dish "
                                f"group '{grules[0]['group']}': {_shown(grules)}", ''))
                    seen.add((dish, v))
                    continue
                mgs = list(self.r.dish_modifier_groups.get(dish, []))
                for g in dish_groups.get(dish, []):
                    mgs += [m for m in self.r.group_modifier_groups.get(g, []) if m not in mgs]
                for mg in mgs:
                    items = self.r.mg_items.get(mg, [])
                    if mg not in self.r.mg_required or not items:
                        continue
                    ip = [self.r.mod_prices.get(it, {}).get(v) for it in items]
                    if all(x is not None and x > 0.001 for x in ip):
                        out.append((group, dish, v, 'modifier_group', None,
                                    f"required modifier group '{mg}': all {len(items)} items "
                                    f"priced ({_fmt(min(ip))}–{_fmt(max(ip))})", mg))
                        seen.add((dish, v))
                        break
        # Maintainer (2026-09-30): in a restaurant that prices dishes on
        # modifier rules (ABSOLUTE), a $0.00 Fried Specialties dish is never
        # given a dish price — even one whose own rules are missing or kept
        # somewhere this check cannot see. "Prices on modifier rules" = some
        # live $0.00 dish (any dish, in a dish group or not) has priced rules
        # that hold its price (a By Itself, or a required choice) — the same
        # test action5 makes on the live raw.
        if self._restaurant_absolute():
            for group, dish, _zh, prices in dishes:
                if not _is_fried_specialty_group(group):
                    continue
                for v, p in prices.items():
                    if abs(p) > 0.001 or (dish, v) in seen:
                        continue
                    out.append((group, dish, v, 'fried_specialty', None,
                                'Fried Specialties dish at $0.00 in a restaurant that prices dishes '
                                'on modifier rules (ABSOLUTE) — never given a dish price', ''))
                    seen.add((dish, v))
        return out

    def _holds_price(self, rules) -> bool:
        """Priced rules hold a $0.00 dish's price when a By Itself is among
        them, or they price a REQUIRED choice (the customer must pick one — its
        rule price is the dish's price). An optional add-on ("Sauce / With
        +$1.00") is only an upcharge. Same test as action5."""
        return any(_is_by_itself(x['mod']) or (x.get('mg') and x['mg'] in self.r.mg_required)
                   for x in rules)

    def _restaurant_absolute(self) -> bool:
        """Some live $0.00 dish price has priced rules that hold it for that
        size (_holds_price: a By Itself, or a required choice) — dish-level, or
        group-level on one of the dish's groups."""
        groups_of = defaultdict(set)
        for g, ds in self.r.group_dishes.items():
            for d in ds:
                groups_of[d].add(g)
        anchors = [x for x in self.r.mr_rules if x['price'] > 0.001 and self._holds_price([x])]
        ganchors = [x for x in getattr(self.r, 'mr_group_rules', [])
                    if x['price'] > 0.001 and self._holds_price([x])]
        for dish, prices in self.r.dish_prices.items():
            for v, p in prices.items():
                if abs(p) > 0.001:
                    continue
                if any(x['dish'] == dish and x['variant'] in ('', v) for x in anchors) or \
                        any(x['group'] in groups_of[dish] and x['variant'] in ('', v) for x in ganchors):
                    return True
        return False

    _SIDE_WORDS = re.compile(r'\b(fries|rice|lo mein|chow mein|mei fun|mai fun|noodles?|pancit|side)\b')

    def _get_side_rule_dishes(self, dishes) -> List[Tuple[str, str, int, str]]:
        """[(group, dish, n_priced_rules, sample)] — Menu Preview dishes priced
        with side-combo rules (a "Plain | w. Fries | w. Fried Rice" table): a
        priced By Itself rule, or a priced side (fries / rice / lo mein / chow
        mein / mei fun / noodles) — dish-level, or group-level on one of the
        dish's groups. An add-on ("Spicy +$1.50", a sushi topping) is not a
        side table."""
        def side_like(rules):
            # a priced By Itself, or a priced SIDE ("w. French Fries", "Pork
            # Fried Rice", "Lo Mein") — not an add-on such as "Make it Spicy" or
            # a sushi topping (Egg, Masago: real ER87669)
            return any(_is_by_itself(x['mod']) or (x['mod'] and self._SIDE_WORDS.search(x['mod'].lower()))
                       for x in rules)

        by_dish = defaultdict(list)
        for x in self.r.mr_rules:
            if x['price'] > 0.001:
                by_dish[x['dish']].append(x)
        by_group = defaultdict(list)
        for x in getattr(self.r, 'mr_group_rules', []):
            if x['price'] > 0.001:
                by_group[x['group']].append(x)
        out, seen = [], set()
        for group, dish, _zh, _prices in dishes:
            if dish in seen:
                continue
            rules = by_dish.get(dish, [])
            if not side_like(rules):
                rules = next((by_group[g] for g in self.r.group_dishes
                              if dish in self.r.group_dishes[g] and side_like(by_group.get(g, []))), [])
            if rules:
                seen.add(dish)
                sample = ', '.join(sorted({x['mod'] for x in rules if x['mod']})[:4])
                out.append((group, dish, len(rules), sample))
        return out

    def generate(self, output_path: str) -> dict:
        """Generate the Excel 2a file. Returns summary stats."""
        wb = Workbook()

        # Build the variant column order for the Menu Preview.
        # PRICE_VARIANTS provides the canonical column ORDER (Order first,
        # then Specialty, Small, Large, etc.) so common variants appear in
        # a predictable column position across restaurants. Any additional
        # variants found in the actual Raw Gsheet data (e.g., 'Roll',
        # 'Dinner', 'Seapark', 'Bowl', 'Regular Roll' for sushi
        # restaurants) are appended after the canonical ones. This is
        # ordering only — every variant in the data ends up as a column.
        output_variants = list(PRICE_VARIANTS)
        all_data_variants = set()
        for prices in self.r.dish_prices.values():
            all_data_variants.update(prices.keys())
        for prices in self.r.mod_prices.values():
            all_data_variants.update(prices.keys())
        for codes in self.r.dish_codes.values():
            all_data_variants.update(codes.keys())
        for v in sorted(all_data_variants):
            if v not in output_variants:
                output_variants.append(v)

        # ── Price tab ─────────────────────────────────────────────
        ws_price = wb.active
        ws_price.title = 'price'
        p_headers = ['Dish Group', 'Dish Name:Primary', 'Dish Name:Secondary'] + output_variants
        for ci, h in enumerate(p_headers, 1):
            cell = ws_price.cell(row=1, column=ci, value=h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT

        dishes = self._get_price_tab_items()
        row = 2
        current_group = None
        for group, name, zh, prices in dishes:
            if group != current_group:
                # Separator row (group name only, no dish name)
                ws_price.cell(row=row, column=1, value=group)
                row += 1
                current_group = group
            ws_price.cell(row=row, column=1, value=group)
            ws_price.cell(row=row, column=2, value=name)
            ws_price.cell(row=row, column=3, value=zh)
            for vi, v in enumerate(output_variants):
                if v in prices:
                    ws_price.cell(row=row, column=4 + vi, value=_fmt(prices[v]))
            row += 1

        # ── System Codes tab: each size's dish code, row-aligned with price ──
        ws_codes = wb.create_sheet('System Codes')
        for ci, h in enumerate(p_headers, 1):
            cell = ws_codes.cell(row=1, column=ci, value=h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
        row, current_group = 2, None
        for group, name, zh, _prices in dishes:
            if group != current_group:
                ws_codes.cell(row=row, column=1, value=group)
                row += 1
                current_group = group
            ws_codes.cell(row=row, column=1, value=group)
            ws_codes.cell(row=row, column=2, value=name)
            ws_codes.cell(row=row, column=3, value=zh)
            for vi, v in enumerate(output_variants):
                code = self.r.dish_codes.get(name, {}).get(v)
                if code:
                    ws_codes.cell(row=row, column=4 + vi, value=str(code)).number_format = '@'
            row += 1

        # ── Modifier tab ──────────────────────────────────────────
        ws_mod = wb.create_sheet('modifier')
        m_headers = ['Dish Group', 'Modifier Group', 'Modifier Name:Primary',
                     'Modifier Name:Secondary'] + output_variants
        for ci, h in enumerate(m_headers, 1):
            cell = ws_mod.cell(row=1, column=ci, value=h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT

        mod_rows = self._get_modifier_tab_rows()
        row = 2
        for mr in mod_rows:
            ws_mod.cell(row=row, column=1, value=mr['dg'])
            if not mr['is_sep']:
                ws_mod.cell(row=row, column=2, value=mr['mg'])
                ws_mod.cell(row=row, column=3, value=mr['name'])
                ws_mod.cell(row=row, column=4, value=mr['zh'])
                for vi, v in enumerate(output_variants):
                    if v in mr['prices']:
                        ws_mod.cell(row=row, column=5 + vi, value=_fmt(mr['prices'][v]))
            row += 1

        # ── absolute_pricing tab (v4.4) ───────────────────────────
        # $0.00 dish cells whose price lives on modifiers. action2 never writes
        # a paper price onto them, price_inference never infers one (G11).
        ws_abs = wb.create_sheet('absolute_pricing')
        abs_cells = self._get_absolute_cells(dishes)
        for ci, h in enumerate(['Dish Group', 'Dish Name', 'Variant', 'Source',
                                'By Itself Rule', 'Evidence', 'Modifier Group'], 1):
            cell = ws_abs.cell(row=1, column=ci, value=h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
        for ri, (group, name, v, source, by_itself, evidence, mg) in enumerate(abs_cells, 2):
            for ci, val in enumerate([group, name, v, source, _fmt(by_itself), evidence, mg], 1):
                if val:
                    ws_abs.cell(row=ri, column=ci, value=val)

        # ── Column widths ─────────────────────────────────────────
        _width_sheets = [ws_price, ws_mod]
        for sheet in _width_sheets:
            sheet.column_dimensions['A'].width = 22
            sheet.column_dimensions['B'].width = 38
            sheet.column_dimensions['C'].width = 20
            for i in range(4, 4 + len(output_variants) + 1):
                sheet.column_dimensions[get_column_letter(i)].width = 12
        for col, w in zip('ABCDEFG', (22, 38, 12, 16, 14, 70, 26)):
            ws_abs.column_dimensions[col].width = w

        # ── side_rules tab — dishes priced with side-combo rules ───
        ws_sr = wb.create_sheet('side_rules')
        side_dishes = self._get_side_rule_dishes(dishes)
        for ci, h in enumerate(['Dish Group', 'Dish Name', 'Priced Rules', 'Modifiers'], 1):
            cell = ws_sr.cell(row=1, column=ci, value=h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
        for ri, (group, name, n, sample) in enumerate(side_dishes, 2):
            for ci, val in enumerate([group, name, n, sample], 1):
                ws_sr.cell(row=ri, column=ci, value=val)
        for col, w in zip('ABCD', (22, 38, 12, 60)):
            ws_sr.column_dimensions[col].width = w

        wb.save(output_path)
        wb.close()

        data_mod_rows = len([r for r in mod_rows if not r['is_sep']])
        return {
            'price_items': len(dishes),
            'price_groups': len(set(d[0] for d in dishes)),
            'modifier_data_rows': data_mod_rows,
            'modifier_sections': len([r for r in mod_rows if r['is_sep']]),
            'absolute_cells': abs_cells,
            'side_rule_dishes': len(side_dishes),
        }


# ── Main ───────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='Convert Raw Excel 1 → Excel 2a (price + dish_code + modifier tabs)',
        epilog='Example: python action1_raw_to_2a.py R88312_NO_1-_Jane_Lexus.xlsx'
    )
    parser.add_argument('input', help='Path to Raw Excel 1 (.xlsx)')
    parser.add_argument('--output', '-o', help='Output path (default: auto-named from R-number)')
    args = parser.parse_args()

    print(f"Reading: {args.input}")
    reader = RawExcel1Reader(args.input)

    # Auto-name output from R-number if not specified
    if args.output:
        output_path = args.output
    else:
        r_num = reader.r_number or 'Unknown'
        output_path = f"{r_num}_Excel_2a.xlsx"

    # Print identity (G7)
    print(f"\n{'='*50}")
    print(f"  Restaurant: {reader.restaurant_name or '(not found)'}")
    print(f"  R-Number:   {reader.r_number or '(not found)'}")
    print(f"  E-Number:   {reader.e_number or '(not found)'}")
    print(f"  Address:    {reader.address or '(not found)'}")
    print(f"{'='*50}")

    # Generate Excel 2a
    generator = Excel2aGenerator(reader)
    stats = generator.generate(output_path)
    reader.close()

    # Print summary
    print(f"\n  Output: {output_path}")
    print(f"\n  Price tab:")
    print(f"    Dish items:  {stats['price_items']}")
    print(f"    Dish groups: {stats['price_groups']}")
    print(f"\n  Modifier tab:")
    print(f"    Data rows:   {stats['modifier_data_rows']}")
    print(f"    Sections:    {stats['modifier_sections']}")
    abs_cells = stats['absolute_cells']
    print(f"\n  absolute_pricing tab ($0.00 dish, price lives on modifiers — never "
          f"overwritten from paper, G11): {len(abs_cells)}")
    for group, name, v, source, _bi, evidence, _mg in abs_cells:
        print(f"    {name} [{v}] — {evidence}")
    print(f"  side_rules tab (dishes with priced side-combo rules): {stats['side_rule_dishes']}")

    # Warnings
    excluded_dupes = [name for name, count in reader.dish_root_counts.items()
                      if count > 1 and name in reader.dish_to_group]
    excluded_no_price = [name for name in reader.dish_to_group
                         if reader.dish_root_counts.get(name, 0) == 1
                         and not reader.dish_prices.get(name, {})]
    excluded_no_group = [name for name in reader.dish_prices
                         if name not in reader.dish_to_group
                         and reader.dish_prices[name]]

    if excluded_dupes or excluded_no_price or excluded_no_group:
        print(f"\n  Excluded items:")
        if excluded_dupes:
            print(f"    Duplicate root rows ({len(excluded_dupes)}): {', '.join(excluded_dupes)}")
        if excluded_no_price:
            print(f"    No prices ({len(excluded_no_price)}): {', '.join(excluded_no_price[:10])}"
                  + (f"... +{len(excluded_no_price)-10} more" if len(excluded_no_price) > 10 else ""))
        if excluded_no_group:
            print(f"    Not in dish_group ({len(excluded_no_group)}): {', '.join(excluded_no_group[:10])}"
                  + (f"... +{len(excluded_no_group)-10} more" if len(excluded_no_group) > 10 else ""))

    print(f"\nDone.")


if __name__ == '__main__':
    main()