#!/usr/bin/env python3
"""
modifier_rule_pricing.py  (compare + apply)
===========================================

Integrated multi-column-pricing → modifier_rule tool. Dependencies: openpyxl only.

Subcommands
-----------
  compare <ocr_json> <raw_gsheet.xlsx> [--report <price_comparison_report.xlsx>]
          [-o <output.xlsx>]

      Part-1 (review) half. Finds EVERY side_pricing section in the OCR JSON,
      computes deltas (combo_total - plain) per dish x column, matches OCR dish
      names + paper columns to the system modifier_rule rows, and propagates by
      price-group. A dish printed in two sections is merged; if the sections
      disagree for a modifier, that modifier is left at its current price and
      flagged (PAPER CONFLICT), never auto-picked (G14). Under ABSOLUTE the
      paper Plain price goes on the dish's By Itself rule (dish_item stays $0).
      Produces the analysis in exactly two places inside the report:
        - "Modifier Rule"  — a COMPACT edit view for matched dishes
          (Enabled | Price | Dish Group | Dish | Variant | Modifier) under a
          single friendly header row, colour coded (Blue=increase,
          Pink=decrease, Green=confirmed, Orange=specialist edit). To its right
          (after a blank spacer col) sits the OCR validation block (by-itself,
          Paper menu Total, live Expected, Match? TRUE/FALSE/N/A) — NOT
          imported by Part 2. Paper menu Total is the total actually PRINTED
          on the paper for that row (blank for PROPAGATED / PRICE TWIN rows,
          which have no printed total → Match? N/A). Expected = by-itself +
          edited price under DELTA, the edited price itself under ABSOLUTE
          (the rule stores the full total). Its own RED tab — the tab the
          specialist edits and `apply` reads (by header label, not fixed
          position).
        - "Fried Specialty"     — Action-3's section, appended below with the
          price-difference deltas, column→modifier + dish-name mappings, and
          per-dish price-group propagation detail.
      There are NO separate "Delta Analysis" / "Propagation Detail" tabs; if a
      prior run created them they are removed. With --report the content is
      folded INTO the existing Price Comparison Report, keeping one single
      workbook. Re-running replaces only compare's own content. compare NEVER
      modifies the Raw Gsheet.

  apply <price_comparison_report.xlsx> <raw_gsheet.xlsx> [-o <final_raw.xlsx>]

      Part-2 (write) half. Run it on action5's output (final_raw.xlsx) so the
      G11 check sees the dish_item prices production will actually get. Reads
      the specialist-approved out_unit_price values from the report's
      "Modifier Rule" tab, matches each (dish_group, dish_item, variant,
      modifier_item) row to the live modifier_rule rows — group-level rules
      (no dish_item) are keyed by their dish_group — and overwrites ONLY
      changed out_unit_price cells in a copy of the Raw Gsheet. Guardrails:
      refuses DOUBLE_CHARGE / BROKEN and any G11 violation in the resulting
      state (exit 1, nothing written), never writes a blank price (G1), never
      adds/removes rows (G2), touches only the modifier_rule tab (G3), only
      writes cells that actually changed (G4), prints "!! DRIFT WARNING" for
      any approved row not found live (G8) and "!! AMBIGUOUS" for a tab row
      that matches more than one live row — neither is applied.

Column offsets in modifier_rule (0-based):
  PRICE_COL=2 (wprop.out_unit_price), ENABLED_COL=1 (wprop.out_is_enabled),
  DISH_COL=9 (parent.dish_item), VARIANT_COL=10 (parent.variant),
  MODGRP_COL=11 (parent.modifier_group), MOD_COL=12 (parent.modifier_item)
"""

import argparse
import json
import os
import re
import sys
from typing import Optional, Dict, List
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import PatternFill, Font
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    from openpyxl.worksheet.hyperlink import Hyperlink
except ImportError:
    print("ERROR: openpyxl required. Install: pip install openpyxl")
    sys.exit(1)

# Optional: floating-image embedding needs openpyxl.drawing.image (which needs
# Pillow). Absence is non-fatal — the Modifier Rule tab just skips the picture.
try:
    from openpyxl.drawing.image import Image as XLImage
    _HAVE_XLIMAGE = True
except Exception:
    XLImage = None
    _HAVE_XLIMAGE = False

BLUE_FILL = PatternFill('solid', fgColor='ADD8E6')
PINK_FILL = PatternFill('solid', fgColor='FFB6C1')
GREEN_FILL = PatternFill('solid', fgColor='C6EFCE')
RED_FILL = PatternFill('solid', fgColor='FFC7CE')
GREY_FILL = PatternFill('solid', fgColor='D9D9D9')
ORANGE_FILL = PatternFill('solid', fgColor='FFC000')
# not a valid price (apply ignores it) — the same red text as Dish Item / Modifier Summary
INVALID_FILL = PatternFill('solid', fgColor='FFFFFF')
INVALID_FONT = Font(color='C00000', bold=True, italic=True)
DIVIDER_FILL = PatternFill('solid', fgColor='FFF2CC')
HEADER_FILL = PatternFill('solid', fgColor='D9E1F2')
GREEN_TAB = 'C6EFCE'
RED_TAB = 'FF0000'
HEADER_FONT = Font(bold=True, size=11)

# Tabs the specialist edits and Part 2 consumes → coloured red.
RED_TABS = {'Dish Item', 'Modifier Summary', 'Modifier Rule'}
# Tab hidden from the report view (read-forward baseline only).
HIDDEN_TABS = {'Modifier'}

# modifier_rule SOURCE column offsets (0-based) in the raw gsheet — used only to
# READ the live modifier_rule tab. Unchanged; the raw layout is fixed.
PRICE_COL = 2
ENABLED_COL = 1
DGROUP_COL = 8
DISH_COL = 9
VARIANT_COL = 10
MODGRP_COL = 11
MOD_COL = 12

# ── Pricing-model detection (ABSOLUTE vs DELTA) ─────────────────────────
# DELTA   : base price on dish_item; no "By Itself" modifier_rule row; each side
#           carries only its UPCHARGE  (out_unit_price = combo_total - plain).
# ABSOLUTE: dish_item == $0; a "By Itself"-type modifier_rule row carries the
#           base price; every side carries the FULL total (out_unit_price = total).
# Decided per (dish, variant) — a restaurant, even one dish, can mix both
# (_cell_models). The anchor row is not always literally "By Itself"; match a
# normalized alias, on the whole name.
BY_ITSELF_ALIASES = ['by itself', 'plain', 'no rice', 'no side']
MODEL_ABSOLUTE = 'ABSOLUTE'
MODEL_DELTA = 'DELTA'
MODEL_DOUBLE = 'DOUBLE_CHARGE'
MODEL_BROKEN = 'BROKEN'
MODEL_UNKNOWN = 'UNKNOWN'
MODEL_MIXED = 'MIXED'       # both kinds, decided per (dish, variant) — see _cell_models

# Compact Modifier Rule OUTPUT layout (1-based columns). Only the 5 fields
# the specialist needs, under a single friendly header row. Part 2's apply()
# locates these by HEADER LABEL (OUT_HEADERS), not by fixed position, so this
# block may be reordered without breaking the linkage.
OUT_ENABLED_COL = 1     # A  out_is_enabled
OUT_PRICE_COL = 2       # B  out_unit_price  (editable; orange-on-edit)
OUT_DGROUP_COL = 3      # C  dish_group (added per spec — one col after price)
OUT_DISH_COL = 4        # D  dish_item
OUT_VARIANT_COL = 5     # E  variant
OUT_MOD_COL = 6         # F  modifier_item
# (col 7 / G is a blank spacer separating data from validation)
VAL_PLAIN_COL = 8       # H  plain price (static)
VAL_OCRTOT_COL = 9      # I  OCR/anchor total (static reference)
VAL_EXPECT_COL = 10     # J  expected total = plain + edited delta (live formula)
VAL_MATCH_COL = 11      # K  Match? TRUE/FALSE/N/A (live formula)
VAL_BASELINE_COL = 12   # L  auto-generated price baseline (hidden; drives orange-on-edit)
# Paper-menu multi-column pricing photo is anchored here — one blank spacer col
# (M) past the validation block, top of the sheet. Embedded inside _render_mrv_tab
# so it survives the del/create-sheet regeneration cycle.
IMG_ANCHOR_COL = 14     # N  (col M is a spacer after the hidden baseline col L)
# The rule's live out_unit_price when Part 1 ran (hidden, far right, past the
# photo). apply: compare changed a row ⇔ baseline (L) ≠ this — never read from
# the cell colour, which a specialist may change.
VAL_LIVE_COL = 30       # AD
# The row's own key (dish_group|dish_item|variant|modifier_item), hidden: apply
# finds each row's L / AD by key, so a sort of the visible A–F block (Excel's
# quick sort stops at the blank spacer column G) cannot mix rows up.
VAL_KEY_COL = 31        # AE
# wsheet column-name header labels; the data columns Part 2's apply() reads are
# matched by substring, so the raw wsheet names are used verbatim. 'dish_item'
# is checked before the bare 'dish' token so 'dish_group' can't mis-bind.
OUT_HEADERS = {
    OUT_ENABLED_COL: 'out_is_enabled',
    OUT_PRICE_COL: 'out_unit_price',
    OUT_DGROUP_COL: 'dish_group',
    OUT_DISH_COL: 'dish_item',
    OUT_VARIANT_COL: 'variant',
    OUT_MOD_COL: 'modifier_item',
}

# Sentinel that marks compare's appended block in the Fried Specialty tab.
FS_SENTINEL = 'PRICE DIFFERENCE (column price minus Plain)'
MRV_TAB = 'Modifier Rule'
FS_DRIFT_NOTE = ("⚠ Built by compare — the report had no Fried Specialty tab although this OCR JSON "
                 "has a multi-column table, so action2 / action3 ran on a different OCR JSON. Re-run "
                 "action2 → action3 → price_inference → compare on the same JSON.")
DELTA_TAB = 'Delta Analysis'
PROP_TAB = 'Propagation Detail'
# canonical tab order used by the report — same order as part1.md, action3 and
# price_inference.py (compare runs last, so this is the order the specialist
# sees). Delta Analysis + Propagation Detail are NOT separate tabs — their
# content lives inside the Fried Specialty tab. 'Modifier' is hidden
# (read-forward baseline only) so it sorts to the end.
CANON_ORDER = ['Summary', 'Dish Item', 'Questionnaire', 'Price Inference',
               'Modifier Summary', MRV_TAB, 'Fried Specialty', 'Dish Code',
               'Modifier']


def _parse(val) -> Optional[float]:
    if val is None:
        return None
    s = str(val).strip()
    if not s or s in ('None', 'NULL', ''):
        return None
    s = s.replace('$', '').replace(',', '').replace('+', '')
    try:
        return float(s)
    except ValueError:
        return None


def _fmt(val) -> str:
    """$X.XX — a negative as -$X.XX, the way wsheet writes it."""
    if val is None:
        return ''
    try:
        f = float(val)
        return f'-${-f:.2f}' if f < 0 else f'${f:.2f}'
    except Exception:
        return ''


PRICE_CAP = 1000.0      # no rule costs this — a missing decimal point (1395)
_RULE_ENTRY_RE = re.compile(r'^([+-]?)([0-9]+(?:\.[0-9]*)?|\.[0-9]+)$')


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


def _round2(x) -> float:
    """Round to cents the way a spreadsheet does — half away from zero, on the
    decimal value as typed (2.675 → 2.68, 15.995 → 16.00; Python's round()
    works on the binary value and gives 2.67 / 15.99, so Part 2 would write a
    cent less than the report shows)."""
    f = float(x)
    if f != f or abs(f) >= 1e12:          # NaN / huge: no cents to round (fails the cap anyway)
        return f
    return float(Decimal(x if isinstance(x, str) else repr(f)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))


def _rule_entry(val):
    """(kind, number) for a Modifier Rule Price typed by the specialist:
    'blank', 'x', 'price' (|p| < PRICE_CAP — a rule may be negative, e.g. "Less
    Whole Wing" −$2.00) or 'invalid' (letters, comma, %, 1000 or more). Same
    rules as action5's Dish Item parser, signs allowed."""
    if val is None:
        return 'blank', None
    if isinstance(val, bool):
        return 'invalid', None
    if isinstance(val, (int, float)):
        f = _round2(val)
        return ('price', f + 0.0) if abs(f) < PRICE_CAP else ('invalid', None)
    s = str(val).strip(' ')              # plain spaces only, like the sheet's TRIM
    if not s:
        return 'blank', None
    if s.lower() == 'x':
        return 'x', None
    m = _RULE_ENTRY_RE.fullmatch(s.replace('$', '').strip(' '))
    if not m:
        return 'invalid', None
    f = _round2(m.group(2)) * (-1 if m.group(1) == '-' else 1)
    return ('price', f + 0.0) if abs(f) < PRICE_CAP else ('invalid', None)


def _is_by_itself(mod_name: str) -> bool:
    """True if a modifier_item name IS a base-price anchor alias — the whole
    name, after normalization. A substring test made 'Plain Fried Rice' (an
    ordinary DELTA upcharge) read as a By Itself anchor → false DOUBLE_CHARGE."""
    return _normalize(mod_name) in BY_ITSELF_ALIASES


def _load_raw(raw_gsheet_path):
    """Return (mr_data, dish_item_prices) from a Raw Gsheet .xlsx.
    mr_data = modifier_rule data rows (row 4+); dish_item_prices = {(dish,variant): float}.
    Self-contained (does not depend on action5 helpers)."""
    wb = load_workbook(raw_gsheet_path, read_only=True)
    mr_data = list(wb['modifier_rule'].iter_rows(values_only=True))[3:]
    di_rows = list(wb['dish_item'].iter_rows(values_only=True))
    wb.close()
    di_hdr = [str(c).strip() if c else '' for c in di_rows[2]]
    di_path = [str(c).strip().lower() if c else '' for c in di_rows[1]]
    di_name_col = di_hdr.index('dish_item') if 'dish_item' in di_hdr else 0
    _META = {'dish_group', 'dish_item', 'variant', 'modifier_group',
             'modifier_item', 'notes', 'id', ''}
    di_price_cols = {}
    for i in range(len(di_hdr)):
        h = di_hdr[i]
        path = di_path[i] if i < len(di_path) else ''
        if 'unit_price' in path and h and h.lower() not in _META and not h.startswith('_ref'):
            di_price_cols[h] = i
    dish_item_prices = {}
    for row in di_rows[3:]:
        if not row or not str(row[0]).strip().startswith('{'):
            continue
        nm = str(row[di_name_col]).strip() if di_name_col < len(row) and row[di_name_col] else ''
        if not nm:
            continue
        for variant, ci in di_price_cols.items():
            p = _parse(row[ci]) if ci < len(row) else None
            if p is not None:
                dish_item_prices[(nm, variant)] = p
    return mr_data, dish_item_prices


def _cell_models(mr_data, dish_item_prices):
    """Pricing model of every (dish, variant) that carries priced dish-level
    modifier_rule rows. The model belongs to the CELL, not the restaurant:
    ER81988 China Village prices Cut Wings 5–30 pcs ABSOLUTE ($0.00 dish, By
    Itself carries the base) and 40/50/100 pcs on dish_item with upcharge
    rules — inside one dish. Returns {(dish, variant): (model, by_itself)}:

        dish_item >$0, no priced By Itself   -> DELTA          (rules are upcharges)
        dish_item $0 (or none)               -> ABSOLUTE       (rules are full totals;
                                                               By Itself carries the base)
        dish_item >$0 AND priced By Itself   -> DOUBLE_CHARGE  (base in two places)

    Only rows that name a dish_item count: a restaurant-wide or group-level
    'Plain' / 'By Itself' rule is a modifier price, not a dish's base (China
    Village's restaurant-wide 'Plain $0.00' rule made v4.3 read UNKNOWN). A
    variant-less rule is judged against the dish's own dish_item prices."""
    priced = defaultdict(list)          # (dish, variant) -> [(modifier, price)]
    for row in mr_data:
        dish = str(row[DISH_COL]).strip() if len(row) > DISH_COL and row[DISH_COL] else ''
        price = _parse(row[PRICE_COL]) if len(row) > PRICE_COL else None
        if not dish or price is None or price <= 0.001:
            continue
        variant = str(row[VARIANT_COL]).strip() if len(row) > VARIANT_COL and row[VARIANT_COL] else ''
        mod = str(row[MOD_COL]).strip() if len(row) > MOD_COL and row[MOD_COL] else ''
        priced[(dish, variant)].append((mod, price))
    out = {}
    for (dish, variant), rules in priced.items():
        by_itself = max((p for m, p in rules if m and _is_by_itself(m)), default=None)
        dp = _dish_price(dish_item_prices, dish, variant)
        if dp is not None and dp > 0.001:
            model = MODEL_DOUBLE if by_itself is not None else MODEL_DELTA
        else:
            model = MODEL_ABSOLUTE
        out[(dish, variant)] = (model, by_itself)
    return out


def _dish_price(dish_item_prices, dish, variant):
    """dish_item price of (dish, variant); a variant-less rule applies to every
    variant of the dish, so it is judged against the dish's highest price."""
    if variant:
        return dish_item_prices.get((dish, variant))
    return max((p for (d, _v), p in dish_item_prices.items() if d == dish), default=None)


def _detect_pricing_model(mr_data, dish_item_prices):
    """Restaurant verdict from the per-(dish, variant) models (_cell_models).
    Returns (model, evidence_str).

        any DOUBLE_CHARGE cell               -> DOUBLE_CHARGE  (block)
        every cell ABSOLUTE / every DELTA    -> ABSOLUTE / DELTA
        both kinds                           -> MIXED   (each cell priced by its own model)
        no priced dish-level rule at all     -> DELTA if dish_item carries prices,
                                                else BROKEN (block)
    """
    cells = _cell_models(mr_data, dish_item_prices)
    double = sorted(k for k, (m, _b) in cells.items() if m == MODEL_DOUBLE)
    if double:
        return MODEL_DOUBLE, (f"{len(double)} dish/variant(s) priced on BOTH dish_item and a By "
                              f"Itself rule, e.g. {double[0]} - DOUBLE CHARGE")
    absolute = sorted(k for k, (m, _b) in cells.items() if m == MODEL_ABSOLUTE)
    delta = sorted(k for k, (m, _b) in cells.items() if m == MODEL_DELTA)
    if absolute and delta:
        return MODEL_MIXED, (f"{len(absolute)} ABSOLUTE dish/variant(s) (dish_item $0, e.g. "
                             f"{absolute[0]}) + {len(delta)} DELTA (dish_item priced, e.g. "
                             f"{delta[0]}) - each priced by its own model")
    if absolute:
        return MODEL_ABSOLUTE, (f"{len(absolute)} dish/variant(s) with dish_item $0 priced on "
                                f"modifier_rule, e.g. {absolute[0]}")
    if delta:
        return MODEL_DELTA, (f"dish_item carries the base; {len(delta)} dish/variant(s) with "
                             f"upcharge rules, e.g. {delta[0]}")
    if any(p > 0 for p in dish_item_prices.values()):
        return MODEL_DELTA, "no priced dish-level modifier_rule rows; dish_item carries base prices"
    return MODEL_BROKEN, "no priced modifier rules and no dish_item price - no base price anywhere"


def _check_g11_core(mr_data, dish_item_prices, model=None, dish_groups=None):
    """G11: no (dish, variant) may carry its base twice — a priced dish_item
    AND a priced By Itself rule — or customers pay base + base / full total.
    Decided per (dish, variant) (see _cell_models), so a restaurant can mix
    ABSOLUTE and DELTA dishes; `model` is accepted for backward compatibility.
    dish_groups ({dish: {dish groups}}): also a priced GROUP-level By Itself
    rule (no dish_item) on a dish whose dish_item is priced.
    Returns a list of violation strings (empty = OK)."""
    out = []
    for (dish, variant), (m, by_itself) in sorted(_cell_models(mr_data, dish_item_prices).items()):
        if m == MODEL_DOUBLE:
            dp = _dish_price(dish_item_prices, dish, variant)
            out.append(f"(dish='{dish}', variant='{variant}') dish_item=${dp:.2f} AND By Itself "
                       f"rule ${by_itself:.2f} - DOUBLE CHARGE")
    if dish_groups:
        def _c(row, ci):
            return str(row[ci]).strip() if len(row) > ci and row[ci] is not None else ''
        anchors = []
        for row in mr_data:
            p = _parse(row[PRICE_COL]) if len(row) > PRICE_COL else None
            if not _c(row, DISH_COL) and _c(row, DGROUP_COL) and p is not None and p > 0.001 \
                    and _is_by_itself(_c(row, MOD_COL)):
                anchors.append((_c(row, DGROUP_COL), _c(row, VARIANT_COL), _c(row, MOD_COL), p))
        for (dish, variant), dp in sorted(dish_item_prices.items()):
            if dp is None or dp <= 0.001:
                continue
            for g, av, mod, p in anchors:
                if g in dish_groups.get(dish, ()) and av in ('', variant):
                    out.append(f"(dish='{dish}', variant='{variant}') dish_item=${dp:.2f} AND group-level "
                               f"{mod} rule on '{g}' ${p:.2f} - DOUBLE CHARGE")
    return out


def _load_dish_groups(raw_gsheet_path):
    """{dish: {dish groups}} from the raw's dish_group tab (by header text)."""
    wb = load_workbook(raw_gsheet_path, read_only=True)
    out = defaultdict(set)
    if 'dish_group' in wb.sheetnames:
        rows = list(wb['dish_group'].iter_rows(values_only=True))
        h = [str(c).strip() if c else '' for c in rows[2]] if len(rows) > 2 else []
        if 'dish_group' in h and 'dish_item' in h:
            gi, di = h.index('dish_group'), h.index('dish_item')
            g = ''
            for r in rows[3:]:
                if r and str(r[0] or '').startswith('{') and gi < len(r) and r[gi]:
                    g = str(r[gi]).strip()
                if g and di < len(r) and r[di]:
                    out[str(r[di]).strip()].add(g)
    wb.close()
    return out


def _check_g11(A):
    """G11 wrapper for the analysis payload (cmd_compare path)."""
    return _check_g11_core(A['mr_data'], A['dish_item_prices'], A.get('model'), A.get('dish_groups'))


def _normalize(name: str) -> str:
    s = str(name).lower().strip()
    s = re.sub(r'\s*\([^)]*\)\s*', ' ', s)
    s = s.replace('w.', 'with').replace('&', 'and')
    s = re.sub(r'[^a-z0-9\s]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    s = re.sub(r'\broast pork\b', 'pork', s)
    s = re.sub(r'\bbbq\b', 'barbecue', s)
    s = re.sub(r'\bwings\b', 'wing', s)
    return s


# Sauce phrases used to disambiguate compound "X or Y or Z Sauce" dishes.
# Longest phrases first so 'black bean' wins over 'bean'/'black'. 'bbq' folds
# to 'barbecue' to match _normalize's own bbq→barbecue rewrite.
SAUCE_PHRASES = ['black bean', 'black pepper', 'sweet and sour', 'garlic', 'curry',
                 'honey', 'buffalo', 'barbecue', 'bbq', 'teriyaki', 'oyster',
                 'hoisin', 'brown', 'duck']
# Protein head-nouns ordered most-specific → least, so a compound sauce sibling
# binds to the right family (a wings line must not grab a chicken/crispy dish).
PROTEIN_TOKENS = ['wing', 'rib', 'shrimp', 'scallop', 'crab', 'gizzard', 'nugget',
                  'cutlet', 'fish', 'beef', 'pork', 'chicken']


def _sauce_of(norm: str) -> Optional[str]:
    for s in SAUCE_PHRASES:
        if re.search(r'\b' + re.escape(s) + r'\b', norm):
            return 'barbecue' if s == 'bbq' else s
    return None


def _head_protein(norm: str) -> Optional[str]:
    for t in PROTEIN_TOKENS:
        if re.search(r'\b' + t, norm):
            return t
    return None


def _expand_side_items(items: List[dict]) -> List[dict]:
    """Expand a folded compound sauce line — "<Base> (in|w.|with) X or Y [or Z]
    Sauce" — into one sibling item per sauce choice, each inheriting the parent's
    prices, side_prices and code (they share the printed upcharge). Non-compound
    items pass through unchanged. Fixes the class of miss where e.g. "Chicken
    Wings in Garlic or Curry or Honey Sauce" collapsed to a single row and its
    sibling POS dishes (Honey Wings, Wing in Curry Sauce, Wings with Garlic
    Sauce) never reached the Modifier Rule tab."""
    out = []
    for it in items:
        name = str(it.get('name', ''))
        m = re.search(r'^(.*?\b(?:in|w\.?|with))\s+(.+?\bor\b.+?)\s*sauce\s*$',
                      name, re.I)
        if m:
            base = m.group(1).strip()
            choices = [c.strip() for c in re.split(r'\s+or\s+', m.group(2), flags=re.I)
                       if c.strip()]
            if len(choices) >= 2:
                for c in choices:
                    sib = dict(it)
                    sib['name'] = f'{base} {c} Sauce'
                    out.append(sib)
                continue
        out.append(it)
    return out


def _match_dish_name(ocr_name: str, system_names: Dict[str, str]) -> Optional[str]:
    norm_ocr = _normalize(ocr_name)
    for sys_name, norm_sys in system_names.items():
        if norm_ocr == norm_sys:
            return sys_name
    # Sauce-aware disambiguation (before fuzzy) for names that name a sauce.
    # Restrict candidates to those sharing the sauce AND the head protein, so a
    # wings-in-garlic line binds to a wings dish, not a chicken/crispy dish.
    sauce = _sauce_of(norm_ocr)
    if sauce:
        cands = {n: v for n, v in system_names.items() if _sauce_of(v) == sauce}
        prot = _head_protein(norm_ocr)
        if prot:
            pc = {n: v for n, v in cands.items() if re.search(r'\b' + prot, v)}
            if pc:
                cands = pc
        if len(cands) == 1:
            return next(iter(cands))
        if cands:
            ow = set(norm_ocr.split())
            best, bs = None, 0.0
            for n, v in cands.items():
                sw = set(v.split())
                sc = len(ow & sw) / max(len(ow), len(sw)) if sw else 0
                if sc > bs:
                    bs, best = sc, n
            if best:
                return best
    if len(norm_ocr.split()) <= 3:
        for sys_name, norm_sys in system_names.items():
            if norm_ocr in norm_sys or norm_sys in norm_ocr:
                return sys_name
    ocr_words = set(norm_ocr.split())
    best_score = 0
    best_match = None
    for sys_name, norm_sys in system_names.items():
        sys_words = set(norm_sys.split())
        if not sys_words:
            continue
        overlap = len(ocr_words & sys_words)
        score = overlap / max(len(ocr_words), len(sys_words))
        if score > best_score and score >= 0.5:
            best_score = score
            best_match = sys_name
    return best_match


def _match_column_to_modifiers(col_header: str, modifier_names: List[str]) -> List[str]:
    h = col_header.lower().strip()
    # Strip a leading side-prefix: the spelled-out word "with", or the
    # abbreviation "w." / "w ". The bare "w" MUST be followed by "." or
    # whitespace — otherwise "with french fries" was truncated to
    # "ith french fries" (matched nothing) while "w.pork fried rice" and bare
    # "beef fried rice" matched fine, which is exactly the asymmetry we hit.
    h = re.sub(r'^with\s+|^w\.\s*|^w\s+', '', h)
    parts = re.split(r'\s+or\s+|/', h)
    parts = [p.strip() for p in parts if p.strip()]
    SUFFIXES = ['fried rice', 'lo mein', 'chow mein', 'mai fun']
    detected_suffix = None
    for p in parts:
        for suf in SUFFIXES:
            if p.endswith(suf) and p != suf:
                detected_suffix = suf
                break
        if detected_suffix:
            break
    expanded = []
    for p in parts:
        has_suffix = any(p.endswith(suf) for suf in SUFFIXES)
        is_standalone = p in ('french fries', 'white rice', 'fried rice')
        if has_suffix or is_standalone:
            expanded.append(p)
        elif detected_suffix:
            expanded.append(f'{p} {detected_suffix}')
        else:
            expanded.append(p)
    # a bare "Fried Rice" / "Lo Mein" … column is the plain one
    for base in ('fried rice', 'lo mein', 'chow mein', 'mai fun'):
        if base in expanded:
            expanded.append(f'plain {base}')
    matches = []
    for part in expanded:
        norm_part = _normalize(part)
        for mn in modifier_names:
            norm_mn = _normalize(mn)
            if norm_part == norm_mn and mn not in matches:
                matches.append(mn)
    return matches


# ══════════════════════════════════════════════════════════════════════
#  SHARED ANALYSIS (used by compare)
# ══════════════════════════════════════════════════════════════════════

def _side_sections(ocr):
    """Every multi-column section of the OCR JSON, in paper order. v4.x read
    only the FIRST side_pricing section and silently ignored the rest (e.g. a
    separate Wings table). Falls back to the legacy top-level
    'fried_specialty_sides' block. Each section:
    {'name', 'columns', 'items' (rows carrying side prices), 'coverage'}."""
    sections = []
    for sec in ocr.get('sections', []):
        sp = sec.get('side_pricing')
        if not (sp and sp.get('columns')):
            continue
        name = sec.get('menu_section', '')
        cols = list(sp['columns'])
        all_items = sec.get('items', [])
        kept = [it for it in all_items if it.get('side_prices')]
        # Coverage audit: the paper section's full item list is ground truth.
        # A row with NO side_prices is DROPPED (never reaches the tab); a kept
        # row still missing some declared columns is PARTIAL. Both are OCR
        # gaps to surface — silent filtering here is how D7/D8-type rows vanish.
        dropped = [it for it in all_items if not it.get('side_prices')]
        partial = [(it, [c for c in cols if (it.get('side_prices') or {}).get(c) is None])
                   for it in kept]
        partial = [(it, miss) for it, miss in partial if miss]
        sections.append({'name': name, 'columns': cols, 'items': kept, 'coverage': {
            'section': name,
            'total': len(all_items),
            'kept': len(kept),
            'dropped': [(it.get('code', ''), it.get('name', '')) for it in dropped],
            'partial': [(it.get('code', ''), it.get('name', ''), miss) for it, miss in partial],
        }})
    if not any(s['items'] for s in sections):
        legacy = ocr.get('fried_specialty_sides')
        if legacy and legacy.get('items'):
            sections = [{'name': 'Fried Specialties', 'columns': legacy.get('columns', [])[1:],
                         'items': legacy['items'], 'coverage': None}]
    return [s for s in sections if s['items']]


_SIZE_IN_NAME = re.compile(r'(\d+)\s*(?:pcs?\b|pieces?\b|pc\.)|\((\d+)\)', re.I)


def _variant_sort_key(v):
    """'5 Pieces' before '10 Pieces' (numeric), then by name."""
    m = re.match(r'\s*(\d+)', v or '')
    return (int(m.group(1)) if m else float('inf'), v or '')



def _numeric_prices(node, path='', out=None):
    """Make every price in the OCR JSON a number, in place: "$9.59" / "9.59" →
    9.59 (a model sometimes writes prices as text); anything else ("MP",
    "market") → null, listed in the returned messages — never a crash, never
    silently read as a price. Keys: prices / side_prices (dict or list),
    price, default_price."""
    out = [] if out is None else out

    def num(v, where):
        if v is None or isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip().replace('$', '').replace(',', '')
        try:
            return float(s) if s else None
        except ValueError:
            out.append(f"NOT A PRICE in the OCR JSON: {where} = {v!r} — not read")
            return None

    if isinstance(node, dict):
        name = node.get('name') or node.get('menu_section') or path
        for k, v in list(node.items()):
            if k in ('prices', 'side_prices') and isinstance(v, dict):
                node[k] = {kk: num(vv, f"{name} {k}[{kk}]") for kk, vv in v.items()}
            elif k in ('prices', 'side_prices') and isinstance(v, list):
                node[k] = [num(vv, f"{name} {k}[{i}]") for i, vv in enumerate(v)]
            elif k in ('price', 'default_price') and not isinstance(v, (dict, list)):
                node[k] = num(v, f"{name} {k}")
            elif isinstance(v, (dict, list)):
                _numeric_prices(v, f"{path}/{k}", out)
    elif isinstance(node, list):
        for i, v in enumerate(node):
            _numeric_prices(v, f"{path}[{i}]", out)
    return out


def _analyze(ocr_json_path, raw_gsheet_path):
    """Return the full analysis payload used to render compare's views.

    Everything is keyed per (dish, variant): the pricing model (ABSOLUTE /
    DELTA — _cell_models), the paper observations and the resulting prices.
    A dish priced per size (Cut Wings 5 / 7 / 10 Pieces …) is never collapsed
    to one price per modifier. Each paper row is attributed to ONE variant —
    the key of its Plain price, the dish's only variant, or a size in its name
    ("5pcs") — otherwise it is flagged VARIANT UNKNOWN and not applied."""
    with open(ocr_json_path) as f:
        ocr = json.load(f)
    ocr_not_prices = _numeric_prices(ocr)

    sections = _side_sections(ocr)
    if not sections:
        return None
    side_columns = []                       # union, paper order (for the mapping)
    for sec in sections:
        for c in sec['columns']:
            if c not in side_columns:
                side_columns.append(c)

    # Pricing model per (dish, variant), from the raw, before any arithmetic.
    _mr_data_pre, _dish_item_prices = _load_raw(raw_gsheet_path)
    model, model_evidence = _detect_pricing_model(_mr_data_pre, _dish_item_prices)
    cells = _cell_models(_mr_data_pre, _dish_item_prices)

    # A combo box ("Appetizer Combos — all served w. fries or fried rice
    # $10.95") prints only the price WITH the side; the dish's own price is on
    # its row elsewhere (APPETIZER). Such a row takes its Plain from there —
    # only when the paper names match (plural-folded) and that row prints
    # exactly ONE price; otherwise it has no Plain (fine on an ABSOLUTE cell,
    # where the side rule stores the full total).
    def _paper_key(name):
        return ' '.join(w[:-1] if w.endswith('s') and len(w) > 3 else w
                        for w in _normalize(name).split())

    single_plain = defaultdict(set)         # paper key -> {(price key, price, section)}
    for osec in ocr.get('sections', []):
        for it in osec.get('items', []):
            pr = [(k, pp) for k, p in (it.get('prices') or {}).items() for pp in [_parse(p)] if pp is not None]
            if len(pr) == 1:
                single_plain[_paper_key(it.get('name', ''))].add(
                    (pr[0][0], round(pr[0][1], 2), osec.get('menu_section', '')))
    joined = defaultdict(list)              # section -> ["<name> $x (from <section>)"]

    # One entry per printed row per section. Folded compound "X or Y or Z
    # Sauce" lines are expanded into per-sauce siblings first, so each sibling
    # carries the shared upcharge. 'keys' = every printed Plain price by its
    # variant key; which one is the row's Plain is decided once the dish is known.
    not_a_price = []     # text where the OCR JSON should hold a number ("MP") — not compared
    for sec in sections:
        sec['items'] = _expand_side_items(sec['items'])
        sec['entries'] = []
        for item in sec['items']:
            if 'side_prices' in item:
                keys = [(k, round(pp, 2)) for k, p in (item.get('prices') or {}).items()
                        for pp in [_parse(p)] if pp is not None]
                sp = item['side_prices'] or {}
            else:
                prices = item.get('prices', [])
                p0 = _parse(prices[0]) if prices else None
                keys = [('Specialty', round(p0, 2))] if p0 is not None else []
                sp = {sec['columns'][i]: prices[i + 1]
                      for i in range(len(sec['columns'])) if i + 1 < len(prices)}
            for col, t in sp.items():
                if t not in (None, '') and _parse(t) is None:
                    not_a_price.append(f"{item['name']} / {col}: {t!r}")
            if not keys:
                cands = single_plain.get(_paper_key(item['name']), set())
                if len({(k, p) for k, p, _s in cands}) == 1:
                    k, p, src = next(iter(cands))
                    keys = [(k, p)]
                    joined[sec['name']].append(f"{item['name']} {_fmt(p)} (from {src})")
            item['_entry'] = {'code': item.get('code', ''), 'name': item['name'], 'keys': keys,
                              'totals': {col: round(_parse(t), 2) for col, t in sp.items()
                                         if _parse(t) is not None}}
            sec['entries'].append(item['_entry'])

    # load modifier_rule
    wb_raw = load_workbook(raw_gsheet_path, read_only=True)
    ws_mr = wb_raw['modifier_rule']
    all_mr_rows = list(ws_mr.iter_rows(values_only=True))
    wb_raw.close()
    mr_header = [list(all_mr_rows[0]), list(all_mr_rows[1]), list(all_mr_rows[2])]
    mr_data = all_mr_rows[3:]

    all_mr_dish_names = set()
    for row in mr_data:
        d = str(row[DISH_COL]).strip() if len(row) > DISH_COL and row[DISH_COL] else ''
        if d:
            all_mr_dish_names.add(d)
    system_names = {n: _normalize(n) for n in all_mr_dish_names}
    # An exact (normalized) match to ANY system dish wins — including a dish
    # with no modifier rules, which then simply has nothing to update. The
    # fuzzy steps below only see rule-bearing dishes, so without this a row for
    # a rule-less dish ("BBQ Spare Ribs") was pinned on its closest rule-bearing
    # neighbour ("Barbecue Boneless Spare Ribs").
    exact_any = {}
    for d, _v in _dish_item_prices:
        exact_any.setdefault(_normalize(d), d)
    dish_map = {}
    for sec in sections:
        for e in sec['entries']:
            if e['name'] not in dish_map:
                m = exact_any.get(_normalize(e['name'])) or _match_dish_name(e['name'], system_names)
                if m:
                    dish_map[e['name']] = m
            e['sys'] = dish_map.get(e['name'])
    matched_dishes = {e['sys'] for sec in sections for e in sec['entries'] if e['sys']}

    # System-driven selection is a UNION of all criteria (not limited to
    # 'Specialty'), so folded/misread compound lines can no longer silently
    # drop siblings. A modifier_rule row is selected if ANY of:
    #   (a) its dish_item has a 'Specialty' variant anywhere in modifier_rule,
    #   (b) its dish_item was matched from a paper multi-column table,
    #   (c) its dish_group is a Fried Specialty group (group-level rules that
    #       carry no dish_item are kept too — selected per spec).
    specialty_dishes = set()
    for row in mr_data:
        d = str(row[DISH_COL]).strip() if len(row) > DISH_COL and row[DISH_COL] else ''
        v = str(row[VARIANT_COL]).strip() if len(row) > VARIANT_COL and row[VARIANT_COL] else ''
        if d and v == 'Specialty':
            specialty_dishes.add(d)
    selected_dishes = matched_dishes | specialty_dishes

    def _is_fried_specialty_group(g):
        return bool(g) and 'fried special' in g.lower()

    mr_dish_rows = defaultdict(list)
    all_modifier_items = set()
    rule_variants = defaultdict(set)       # dish -> variants its modifier_item rules use
    by_itself_mods = defaultdict(list)     # (dish, variant) -> By Itself anchor modifier_item(s)
    for ri, row in enumerate(mr_data):
        dish = str(row[DISH_COL]).strip() if len(row) > DISH_COL and row[DISH_COL] else ''
        grp = str(row[DGROUP_COL]).strip() if len(row) > DGROUP_COL and row[DGROUP_COL] else ''
        variant = str(row[VARIANT_COL]).strip() if len(row) > VARIANT_COL and row[VARIANT_COL] else ''
        mod = str(row[MOD_COL]).strip() if len(row) > MOD_COL and row[MOD_COL] else ''
        selected = dish in selected_dishes or _is_fried_specialty_group(grp)
        if not selected:
            continue
        # Key by dish_item when present; group-level rules (no dish_item) bucket
        # under a synthetic "[group] <name>" key so they render on the tab too.
        key = dish if dish else f'[group] {grp}'
        mr_dish_rows[key].append((ri + 4, list(row)))
        if mod:
            all_modifier_items.add(mod)
            rule_variants[key].add(variant)
            if dish and _is_by_itself(mod) and mod not in by_itself_mods[(dish, variant)]:
                by_itself_mods[(dish, variant)].append(mod)

    col_to_mods = {}
    for col in side_columns:
        col_to_mods[col] = _match_column_to_modifiers(col, list(all_modifier_items))

    flags = []     # callback items — conflicts, never auto-resolved (G14)
    notes = list(ocr_not_prices) + [f"NOT A PRICE in the OCR JSON side prices: {x} — not compared"
                                    for x in not_a_price]
    hold = set()   # ((dish, variant), modifier) left at the current price
    dish_groups = _load_dish_groups(raw_gsheet_path)
    # dish group → the modifier items its enabled, priced GROUP-level rules price
    # (a rule that only attaches a modifier group prices nothing itself)
    group_side_mods = defaultdict(set)
    for row in mr_data:
        if (len(row) > MOD_COL and row[DGROUP_COL] and not row[DISH_COL] and row[MOD_COL]
                and str(row[ENABLED_COL] or '').strip().upper() != 'FALSE' and _parse(row[PRICE_COL]) is not None):
            group_side_mods[str(row[DGROUP_COL]).strip()].add(str(row[MOD_COL]).strip())

    def _label(dish, variant):
        """'Cut Wings [5 Pieces]' for a dish priced per variant, else the dish."""
        return f'{dish} [{variant}]' if variant and len(rule_variants.get(dish, ())) > 1 else dish

    def _model_of(dish, variant):
        if dish.startswith('[group] '):
            return MODEL_DELTA                      # group-level rules are add-ons
        for k in ((dish, variant), (dish, '')):
            if k in cells:
                return cells[k][0]
        dp = _dish_price(_dish_item_prices, dish, variant)
        return MODEL_DELTA if dp is not None and dp > 0.001 else MODEL_ABSOLUTE

    def _attribute(e):
        """(variant, plain) of a matched paper row; variant None = cannot tell."""
        sizes = {v for v in rule_variants.get(e['sys'], ()) if v}
        for k, p in e['keys']:
            if k in sizes:
                return k, p
        std = (next(((k, p) for k, p in e['keys'] if k in ('Specialty', 'Order', 'Small')),
                    e['keys'][0]) if e['keys'] else (None, None))
        if len(sizes) <= 1:
            return (next(iter(sizes)) if sizes else ''), std[1]
        m = _SIZE_IN_NAME.search(e['name'])
        if m:
            n = m.group(1) or m.group(2)
            hit = [v for v in sizes if re.match(rf'{n}\b', v)]
            if len(hit) == 1:
                return hit[0], std[1]
        return None, std[1]

    # What the paper says per (system dish, variant, modifier), gathered across
    # EVERY section. On an ABSOLUTE cell the Plain price is the By Itself
    # rule's value (the base lives there; dish_item stays $0 — G11) and a side
    # rule takes the full printed total; on a DELTA cell it takes total − Plain.
    obs = defaultdict(lambda: defaultdict(list))   # (dish, variant) -> mod -> [(value, total, where)]
    side_seen = defaultdict(list)                  # (dish, variant) -> [(column, its modifiers, total, where)]
    plain_obs = defaultdict(list)                  # (dish, variant) -> [(plain, where)]
    unknown_size = set()                           # dishes with a paper row of unknown size
    for sec in sections:
        for e in sec['entries']:
            std = (next(((k, p) for k, p in e['keys'] if k in ('Specialty', 'Order', 'Small')),
                        e['keys'][0]) if e['keys'] else (None, None))
            e['plain'], e['variant'], e['model'] = std[1], None, ''
            e['values'] = ({col: round(t - std[1], 2) for col, t in e['totals'].items()}
                           if std[1] is not None else dict(e['totals']))
            sd = e['sys']
            if sd and sd not in mr_dish_rows and e['totals']:
                # printed with side prices, but the dish has no rules of its own:
                # nothing to update — say so, never drop the paper silently
                where = f"{sec['name']} {e['code'] or e['name']}".strip()
                shared = sorted(g for g in dish_groups.get(sd, set()) if group_side_mods.get(g)
                                and any(_match_column_to_modifiers(col, sorted(group_side_mods[g]))
                                        for col in e['totals']))
                if shared:
                    # its sides are priced by its dish group's rules: the paper may
                    # disagree with them, and one paper row never changes a group
                    flags.append(f"NO DISH RULES: {sd} — paper row {where} prints side prices, but the "
                                 f"dish has no modifier rules of its own; its sides are priced by the "
                                 f"group-level rules of {', '.join(shared)} (shared by the whole group, "
                                 f"not changed from one paper row) — not applied. Check its side prices "
                                 f"on the callback.")
                else:
                    notes.append(f"NO SIDE RULES: {sd} — paper row {where} prints side prices, but no "
                                 f"modifier rule prices a side for this dish (its side options, if any, "
                                 f"are priced as modifier items — Modifier Summary) — nothing compared.")
                e['values'], e['variant'], e['model'] = {}, '—', 'no dish rules'
                continue
            if not sd or sd not in mr_dish_rows:
                continue          # unmatched: nothing to update
            where = f"{sec['name']} {e['code'] or e['name']}".strip()
            variant, plain = _attribute(e)
            e['plain'] = plain
            if variant is None:
                sizes = ', '.join(sorted((v for v in rule_variants[sd] if v), key=_variant_sort_key))
                flags.append(f"VARIANT UNKNOWN: {sd} is priced per size ({sizes}) and paper row "
                             f"{where} does not say which size — not applied. Key its Plain price "
                             f"by the size (e.g. \"5 Pieces\") in the OCR JSON and re-run compare.")
                e['values'] = {}
                unknown_size.add(sd)
                continue
            m = _model_of(sd, variant)
            e['variant'], e['model'] = variant, m
            if plain is None and m != MODEL_ABSOLUTE:
                flags.append(f"NO PLAIN: {_label(sd, variant)} — paper row {where} prints only "
                             f"the price with the side, and the upcharge (DELTA) needs the dish's "
                             f"own price — not applied. Add its Plain to the OCR JSON and re-run.")
                e['values'] = {}
                continue
            e['values'] = (dict(e['totals']) if m == MODEL_ABSOLUTE
                           else {col: round(t - plain, 2) for col, t in e['totals'].items()})
            key = (sd, variant)
            if plain is not None:
                plain_obs[key].append((plain, where))
            for col, val in e['values'].items():
                side_seen[key].append((col, col_to_mods.get(col, []), e['totals'][col], where))
                for mod in col_to_mods.get(col, []):
                    obs[key][mod].append((val, e['totals'][col], where))
            if m == MODEL_ABSOLUTE and plain is not None:
                for mod in by_itself_mods.get(key, []):
                    obs[key][mod].append((plain, plain, where))

    sys_plain = {}
    for key, seen in plain_obs.items():
        sys_plain[key] = seen[0][0]
        if len({p for p, _ in seen}) > 1:
            flags.append(f"PAPER CONFLICT: {_label(*key)} Plain printed differently — "
                         + '; '.join(f"{w}: {_fmt(p)}" for p, w in seen)
                         + ". Check the dish price on the callback.")
        if _model_of(*key) == MODEL_ABSOLUTE and not by_itself_mods.get(key):
            flags.append(f"NO BASE RULE: {_label(*key)} — dish_item is $0.00 and its side rules "
                         f"carry full totals, but it has no By Itself rule, so its paper Plain "
                         f"{_fmt(sys_plain[key])} has nowhere to go (dish_item stays $0.00, G11). "
                         f"Confirm how the plain dish is priced on the callback.")

    mod_known = defaultdict(dict)       # (dish, variant) -> mod -> paper value (all prints agree)
    printed_total = defaultdict(dict)   # (dish, variant) -> mod -> total actually printed
    for key, mods in obs.items():
        for mod, seen in mods.items():
            if len({v for v, _, _ in seen}) == 1:
                mod_known[key][mod] = seen[0][0]
                printed_total[key][mod] = seen[0][1]
            else:
                # Same dish printed in two places (two sections, or two paper
                # rows) with different prices for this modifier: not guessed.
                hold.add((key, mod))
                flags.append(f"PAPER CONFLICT: {_label(*key)} / {mod} printed differently in more "
                             f"than one section/row — "
                             + '; '.join(f"{w}: {_fmt(t)}" for _, t, w in seen)
                             + " — left at the current price. Confirm on the callback.")

    final_prices = defaultdict(dict)    # (dish, variant) -> mod -> new price
    price_sources = defaultdict(dict)
    models = {}                         # (dish, variant) -> model its rows are priced by
    mod_old_by_key = {}                 # (dish, variant) -> mod -> old price
    for dish, rows in mr_dish_rows.items():
        by_var = defaultdict(dict)
        for _, row in rows:
            variant = str(row[VARIANT_COL]).strip() if len(row) > VARIANT_COL and row[VARIANT_COL] else ''
            mod = str(row[MOD_COL]).strip() if len(row) > MOD_COL and row[MOD_COL] else ''
            price = _parse(row[PRICE_COL]) if len(row) > PRICE_COL else None
            models.setdefault((dish, variant), _model_of(dish, variant))
            if mod and price is not None:
                by_var[variant][mod] = price
        for variant, mod_old in by_var.items():
            key = (dish, variant)
            mod_old_by_key[key] = dict(mod_old)
            price_groups = defaultdict(list)
            for mod, price in mod_old.items():
                price_groups[price].append(mod)
            known = mod_known.get(key, {})
            for old_price, members in price_groups.items():
                # A printed member ALWAYS keeps its own printed value. Only
                # unprinted members are propagated, and only when every printed
                # member of the group agrees (G14: a split group is not resolved
                # by picking a side).
                printed = {m: known[m] for m in members if m in known}
                agreed = set(printed.values())
                unprinted = []
                for m in members:
                    if m in printed:
                        final_prices[key][m] = printed[m]
                        price_sources[key][m] = 'PAPER MENU'
                    elif (key, m) not in hold and len(agreed) == 1:
                        final_prices[key][m] = next(iter(agreed))
                        price_sources[key][m] = 'PROPAGATED'
                    else:
                        final_prices[key][m] = old_price
                        price_sources[key][m] = 'NO DATA'
                        if (key, m) not in hold:
                            unprinted.append(m)
                if len(agreed) > 1:
                    split = ', '.join(f'{m} {_fmt(v)}' for m, v in sorted(printed.items()))
                    if unprinted:
                        hold.update((key, m) for m in unprinted)
                        flags.append(f"GROUP CONFLICT: {_label(*key)} — rows that shared "
                                     f"${old_price:.2f} were printed at different prices ({split}); "
                                     f"each printed row keeps its own price. Unprinted "
                                     f"{', '.join(unprinted)} left at ${old_price:.2f} (not "
                                     f"propagated). Confirm on the callback.")
                    else:
                        notes.append(f"SPLIT GROUP: {_label(*key)} — rows that shared "
                                     f"${old_price:.2f} were printed at different prices ({split}); "
                                     f"each keeps its own printed price.")

    # A printed side the dish has no rule for (the paper prints "w. Lo Mein",
    # the dish has fries / rice rules only): nothing to update — a callback
    # item, never dropped silently.
    for key, seen in side_seen.items():
        have = set(mod_old_by_key.get(key, {})) | set(mod_old_by_key.get((key[0], ''), {}))
        done = set()
        for col, mods_, total, where in seen:
            if set(mods_) & have or (col, total) in done:
                continue
            done.add((col, total))
            grp = sorted(g for g in dish_groups.get(key[0], set()) if set(mods_) & group_side_mods.get(g, set()))
            why = (f"its group-level rule on {', '.join(grp)} prices it — shared by the whole group, not "
                   f"changed from one paper row" if grp else
                   "no rule prices it" if mods_ else "no modifier matches that column")
            flags.append(f"NO RULE FOR SIDE: {_label(*key)} — paper row {where} prints '{col}' {_fmt(total)}, "
                         f"but the dish has no modifier rule for it ({why}) — not applied. Check on the callback.")

    # Price-twin propagation (Q1): a Specialty/Fried-Specialty dish with NO paper
    # data borrows the delta from a "price-twin" — a dish that DID get paper data
    # for the SAME modifier_item at the SAME old base price, priced by the SAME
    # model. Fried Specialties is a uniform upcharge family, so equal (modifier,
    # old_price) pairs move together. Build the twin table from resolved (PAPER
    # MENU/PROPAGATED) rows, then fill NO DATA rows whose (modifier, old_price)
    # has a unique twin delta. Ambiguous cases are never guessed — they are
    # reported as TWIN CONFLICT. Rows held by a PAPER / GROUP CONFLICT are never
    # twin-filled either. Twins are looked up DISH-GROUP-FIRST, then globally.
    # Group scope is the tighter cohort (a dish group is priced as a family),
    # and it is strictly additive: a globally-unanimous delta stays unanimous
    # inside any subset, so group-first can only resolve cases the global key
    # left ambiguous — it can never contradict a global-unique answer.
    dish_group_of = {}
    for dish, rows in mr_dish_rows.items():
        for _, row in rows:
            g = str(row[DGROUP_COL]).strip() if len(row) > DGROUP_COL and row[DGROUP_COL] else ''
            if g:
                dish_group_of[dish] = g
                break

    twin_delta = defaultdict(set)        # (model, modifier, old_price) -> {delta, ...}
    twin_delta_grp = defaultdict(set)    # (group, model, modifier, old_price) -> {delta, ...}
    for key, mods in final_prices.items():
        for m, new_p in mods.items():
            if price_sources[key].get(m) in ('PAPER MENU', 'PROPAGATED'):
                op = mod_old_by_key.get(key, {}).get(m)
                if op is not None:
                    d = round(new_p - op, 2)
                    twin_delta[(models[key], m, round(op, 2))].add(d)
                    twin_delta_grp[(dish_group_of.get(key[0], ''), models[key], m,
                                    round(op, 2))].add(d)
    twins_applied = 0
    twins_group_scoped = 0
    for key, mods in final_prices.items():
        # an unprinted size of a dish whose paper row gave no size may be the
        # very row the paper printed — not twin-filled (VARIANT UNKNOWN asks)
        if key[0] in unknown_size and key not in obs:
            continue
        for m, cur in list(mods.items()):
            if price_sources[key].get(m) != 'NO DATA' or (key, m) in hold:
                continue
            op = round(cur, 2)
            grp_seen = twin_delta_grp.get((dish_group_of.get(key[0], ''), models[key], m, op)) or set()
            all_seen = twin_delta.get((models[key], m, op)) or set()
            if len(grp_seen) == 1:
                deltas_seen, scoped = grp_seen, True
            elif len(all_seen) == 1:
                deltas_seen, scoped = all_seen, False
            else:
                if all_seen:
                    flags.append(
                        f"TWIN CONFLICT: {_label(*key)} / {m} @ ${cur:.2f} — candidate twins "
                        f"disagree {sorted(all_seen)}"
                        + (f" (same dish group: {sorted(grp_seen)})" if grp_seen else '')
                        + "; left at the current price. Confirm on the callback.")
                continue
            d = next(iter(deltas_seen))
            if abs(d) > 0.001:
                final_prices[key][m] = round(cur + d, 2)
                price_sources[key][m] = 'PRICE TWIN (group)' if scoped else 'PRICE TWIN'
                twins_applied += 1
                twins_group_scoped += 1 if scoped else 0
    for sname, rows_ in joined.items():
        notes.append(f"PLAIN FROM ANOTHER SECTION: {len(rows_)} '{sname}' row(s) print only the "
                     f"price with the side; their Plain was taken from the same dish's own row: "
                     + '; '.join(rows_))
    if twins_applied:
        notes.append(f"PRICE-TWIN: filled {twins_applied} no-data modifier prices "
                     f"by matching (modifier_item, base price) to a paper-data twin "
                     f"({twins_group_scoped} resolved within the same dish group).")

    return dict(sections=sections, side_columns=side_columns, dish_map=dish_map,
                sys_plain=sys_plain, dish_groups=dish_groups, mr_header=mr_header, mr_dish_rows=mr_dish_rows,
                col_to_mods=col_to_mods, final_prices=final_prices,
                price_sources=price_sources, printed_total=printed_total,
                models=models, mod_old_by_key=mod_old_by_key,
                labels={k: _label(*k) for k in models},
                flags=flags, notes=notes, cells=cells,
                model=model, model_evidence=model_evidence,
                mr_data=_mr_data_pre, dish_item_prices=_dish_item_prices)


def _default_crop_path():
    """Resolve the default multi-column pricing crop shipped alongside the
    scripts: <skill>/crops/special_full.png. Returns the path if it exists,
    else None."""
    here = os.path.dirname(os.path.abspath(__file__))
    cand = os.path.join(os.path.dirname(here), 'crops', 'special_full.png')
    return cand if os.path.isfile(cand) else None


def _embed_pricing_image(ws, img_path, anchor_row=1):
    """Anchor the paper-menu multi-column pricing photo top-right of the tab,
    past the validation block. No-op (with a printed note) when the image or the
    Pillow-backed image support is unavailable, so the report still builds."""
    if not img_path:
        return False
    if not _HAVE_XLIMAGE:
        print("  (note: skipped pricing image — openpyxl image support/Pillow "
              "unavailable)")
        return False
    if not os.path.isfile(img_path):
        print(f"  (note: skipped pricing image — not found: {img_path})")
        return False
    try:
        img = XLImage(img_path)
        # Cap width so a large scan doesn't sprawl; keep aspect ratio.
        max_w = 900
        if getattr(img, 'width', 0) and img.width > max_w:
            scale = max_w / img.width
            img.width = int(img.width * scale)
            img.height = int(img.height * scale)
        anchor = f'{get_column_letter(IMG_ANCHOR_COL)}{anchor_row}'
        ws.add_image(img, anchor)
        print(f"  Embedded pricing image at {anchor}: {os.path.basename(img_path)}")
        return True
    except Exception as e:
        print(f"  (note: skipped pricing image — {e})")
        return False


def _dish_item_cells(wb):
    """{(dish, variant): [(dish_group, cell, ref_cell)]} — where each dish's
    price sits on the report's Dish Item tab, as absolute cross-sheet refs
    ("'Dish Item'!$F$12"), with the hidden _ref_ baseline of that cell. A dish
    listed under two dish groups has two rows (one live dish)."""
    out = defaultdict(list)
    if wb is None or 'Dish Item' not in wb.sheetnames:
        return out
    ws = wb['Dish Item']
    hdr = [str(c.value).strip() if c.value is not None else '' for c in ws[1]]
    name_col = hdr.index('Dish Name:Primary') + 1 if 'Dish Name:Primary' in hdr else 2
    group_col = hdr.index('Dish Group') + 1 if 'Dish Group' in hdr else 1
    meta = {'Dish Group', 'Dish Name:Primary', 'Dish Name:Secondary', ''}
    var_cols = {h: i + 1 for i, h in enumerate(hdr) if h not in meta and not h.startswith('_ref_')}
    ref_cols = {h[len('_ref_'):]: i + 1 for i, h in enumerate(hdr) if h.startswith('_ref_')}
    for r in range(2, ws.max_row + 1):
        dish = str(ws.cell(row=r, column=name_col).value or '').strip()
        if not dish:
            continue
        group = str(ws.cell(row=r, column=group_col).value or '').strip()
        for v, c in var_cols.items():
            if ws.cell(row=r, column=c).value in (None, ''):
                continue
            ref = ref_cols.get(v)
            out[(dish, v)].append((group, f"'Dish Item'!${get_column_letter(c)}${r}",
                                   f"'Dish Item'!${get_column_letter(ref)}${r}" if ref else None))
    return out


def _dish_price_expr(cands, prefer_group):
    """Excel expression for a dish's price on the Dish Item tab: the copy in
    the rule's own dish group first, but any other copy the specialist edited
    (value ≠ its hidden _ref_) wins over it — the rule Part 2 applies to a dish
    listed under two groups."""
    cands = sorted(cands, key=lambda c: c[0] != prefer_group)
    expr = cands[0][1]
    for _g, cell, ref in reversed(cands[1:]):
        if ref:
            expr = f'IF({cell}<>{ref},{cell},{expr})'
    return expr


def _render_mrv_tab(ws, A, img_path=None):
    """Compact 5-column edit view (Enabled/Price/Dish/Variant/Modifier) + an
    OCR validation block to the right. Part 2's apply() reads the five data
    columns by their friendly header label (OUT_HEADERS), so the linkage does
    not depend on the raw modifier_rule column positions."""
    counts = {'updated': 0, 'confirmed': 0, 'no_data': 0}
    r = 1
    ws.cell(row=r, column=1, value='MODIFIER RULE — MULTI-COLUMN SPECIALTY UPCHARGES').font = Font(bold=True, size=13)
    r += 1
    ws.cell(row=r, column=OUT_PRICE_COL, value='← Specialist edits Price here (Part 2 apply reads it)').font = Font(italic=True, color='FF0000')
    dv = ws.cell(row=r, column=VAL_PLAIN_COL,
                 value='↓ VALIDATION — NOT IMPORTED BY PART 2 (right of blank spacer col F) ↓')
    dv.font = Font(bold=True, italic=True, color='7F6000')
    dv.fill = DIVIDER_FILL
    r += 1
    ws.cell(row=r, column=1, value=('Blue = increase | Pink = decrease | Green = confirmed no change | Orange = '
                                    'specialist edit | Red text = not a valid price (not applied)')
            ).font = Font(italic=True, color='666666')
    ws.cell(row=r, column=VAL_PLAIN_COL,
            value=("'by itself' is live — DELTA: the dish's price on the Dish Item tab; ABSOLUTE: its "
                   "By Itself row's Price. Edit it there and Expected / Match? follow.")
            ).font = Font(italic=True, color='7F6000')
    r += 2
    header_row = r
    # Single friendly header row for the five data columns
    for col, label in OUT_HEADERS.items():
        c = ws.cell(row=header_row, column=col, value=label)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    # Validation-block headers
    expected_label = {MODEL_ABSOLUTE: 'Expected (raw total)',
                      MODEL_MIXED: 'Expected (charged total)'}.get(A.get('model'),
                                                                   'Expected (by itself+Δ)')
    # hidden columns, found by these labels in Part 2 (never by position)
    for col, label in ((VAL_BASELINE_COL, '_ref_baseline'), (VAL_LIVE_COL, '_ref_live'),
                       (VAL_KEY_COL, '_ref_key')):
        ws.cell(row=header_row, column=col, value=label).font = Font(color='808080', size=8)
    for col, label in ((VAL_PLAIN_COL, 'by itself'), (VAL_OCRTOT_COL, 'Paper menu Total'),
                       (VAL_EXPECT_COL, expected_label),
                       (VAL_MATCH_COL, 'Match?')):
        c = ws.cell(row=header_row, column=col, value=label)
        c.fill = DIVIDER_FILL
        c.font = HEADER_FONT
    r += 1
    data_start = r
    ordered = []
    for dish, rows in sorted(A['mr_dish_rows'].items()):
        # a dish priced per size lists its sizes in order (5, 7, 10 … Pieces)
        for _, row in sorted(rows, key=lambda x: _variant_sort_key(
                str(x[1][VARIANT_COL]).strip() if len(x[1]) > VARIANT_COL and x[1][VARIANT_COL] else '')):
            ordered.append((dish, row))
    A['mrv_rows'] = {}          # (dish_item, variant, modifier_item) -> tab row
    di_cells = _dish_item_cells(getattr(ws, 'parent', None))
    row_info = []               # (tab row, dish_item, variant, dish_group, (dish, variant))
    for dish, row in ordered:
        enabled = str(row[ENABLED_COL]).strip() if len(row) > ENABLED_COL and row[ENABLED_COL] is not None else ''
        dgroup = str(row[DGROUP_COL]).strip() if len(row) > DGROUP_COL and row[DGROUP_COL] else ''
        dish_item = str(row[DISH_COL]).strip() if len(row) > DISH_COL and row[DISH_COL] else ''
        variant = str(row[VARIANT_COL]).strip() if len(row) > VARIANT_COL and row[VARIANT_COL] else ''
        mod = str(row[MOD_COL]).strip() if len(row) > MOD_COL and row[MOD_COL] else ''
        old_price = _parse(row[PRICE_COL]) if len(row) > PRICE_COL else None
        # Compact data columns (dish_item is the row's own value; empty for
        # group-level rules, which carry only a dish_group).
        A['mrv_rows'].setdefault((dish_item, variant, mod), r)
        row_info.append((r, dish_item, variant, dgroup, (dish, variant)))
        ws.cell(row=r, column=OUT_ENABLED_COL, value=enabled)
        ws.cell(row=r, column=OUT_DGROUP_COL, value=dgroup)
        ws.cell(row=r, column=OUT_DISH_COL, value=dish_item)
        ws.cell(row=r, column=OUT_VARIANT_COL, value=variant)
        ws.cell(row=r, column=OUT_MOD_COL, value=mod)
        key = (dish, variant)       # prices are kept per (dish, variant) — never collapsed
        new_price = A['final_prices'].get(key, {}).get(mod)
        source = A['price_sources'].get(key, {}).get(mod, '')
        price_cell = ws.cell(row=r, column=OUT_PRICE_COL)
        # Write the price (G1: preserve blank group-gate cells — never fabricate a price)
        if old_price is not None:
            price_cell.value = _fmt(new_price if new_price is not None else old_price)
        price_cell.number_format = '0.00'
        if new_price is not None and old_price is not None:
            if abs(new_price - old_price) > 0.001:
                price_cell.fill = BLUE_FILL if new_price > old_price else PINK_FILL
                counts['updated'] += 1
            elif source != 'NO DATA':
                price_cell.fill = GREEN_FILL
                counts['confirmed'] += 1
            else:
                counts['no_data'] += 1
        else:
            counts['no_data'] += 1
        # --- Validation columns, live formulas referencing this row's Price cell ---
        plain = A['sys_plain'].get(key)
        # Paper menu Total = the total actually PRINTED for this row (the side
        # column price; the Plain price for an ABSOLUTE By Itself row). It is
        # never recomputed from the generated price — a recomputed total always
        # "matches" and hid printed prices being overwritten. PROPAGATED /
        # PRICE TWIN rows have no printed total, so theirs is blank (Match? N/A).
        printed = (A['printed_total'].get(key, {}).get(mod)
                   if source == 'PAPER MENU' else None)
        pL = get_column_letter(OUT_PRICE_COL)      # 'B' — editable price
        aa = get_column_letter(VAL_PLAIN_COL)
        ab = get_column_letter(VAL_OCRTOT_COL)
        ac = get_column_letter(VAL_EXPECT_COL)
        if plain is not None:
            ws.cell(row=r, column=VAL_PLAIN_COL, value=round(plain, 2)).number_format = '0.00'
        if printed is not None:
            ws.cell(row=r, column=VAL_OCRTOT_COL,
                    value=round(printed, 2)).number_format = '0.00'
        # Expected = what the customer pays with the live-edited Price (strips
        # any $ so it re-reads edits): by-itself + Price on a DELTA row; the
        # Price itself on an ABSOLUTE row, where the rule already stores the
        # full total (dish_item is $0). Decided per row's (dish, variant). Written on EVERY data row unconditionally — the
        # formula is always present regardless of whether a Paper menu Total
        # exists; IFERROR yields "" only when its inputs aren't numeric.
        # blank-safe: Google Sheets treats "" as 0 in arithmetic and VALUE("")
        if A['models'].get(key) == MODEL_ABSOLUTE:
            exp_formula = (f'=IF(TRIM({pL}{r}&"")="","",'
                           f'IFERROR(VALUE(SUBSTITUTE({pL}{r}&"","$","")),""))')
        else:
            exp_formula = (f'=IF(OR(TRIM({aa}{r}&"")="",TRIM({pL}{r}&"")=""),"",'
                           f'IFERROR({aa}{r}+VALUE(SUBSTITUTE({pL}{r}&"","$","")),""))')
        exp = ws.cell(row=r, column=VAL_EXPECT_COL, value=exp_formula)
        exp.number_format = '0.00'
        # Match? — N/A when no OCR reference (or no Expected), else TRUE/FALSE
        # within 0.5c tolerance
        ws.cell(row=r, column=VAL_MATCH_COL,
                value=(f'=IF(OR({ab}{r}="",{ac}{r}=""),"N/A",'
                       f'IF(ABS({ac}{r}-{ab}{r})<=0.005,"TRUE","FALSE"))'))
        # Baseline (hidden) — drives orange-on-specialist-edit.
        baseline = new_price if (new_price is not None and old_price is not None
                                 and abs(new_price - old_price) > 0.001) else old_price
        if baseline is not None:
            ws.cell(row=r, column=VAL_BASELINE_COL, value=round(baseline, 2)).number_format = '0.00'
        if old_price is not None:
            ws.cell(row=r, column=VAL_LIVE_COL, value=round(old_price, 2)).number_format = '0.00'
        ws.cell(row=r, column=VAL_KEY_COL, value='|'.join((dgroup, dish_item, variant, mod)))
        r += 1
    data_end = r - 1
    # 'by itself' (H) as a live link, so an edited price moves Expected / Match?:
    # DELTA — the dish's own price on the Dish Item tab; ABSOLUTE — the Price of
    # its By Itself row on this tab. Group-level rules (no dish) keep the value.
    by_itself_row = {}
    for (d, v, mod), rr in A['mrv_rows'].items():
        if d and _is_by_itself(mod):
            by_itself_row.setdefault((d, v), rr)
    pL = get_column_letter(OUT_PRICE_COL)
    for rr, dish_item, variant, dgroup, key in row_info:
        if not dish_item:
            continue
        if A['models'].get(key) == MODEL_ABSOLUTE:
            bi = by_itself_row.get((dish_item, variant))
            src = f'{pL}{bi}' if bi else None
        else:
            cands = di_cells.get((dish_item, variant))
            if not cands and not variant:
                # a rule for every size: link when the dish has exactly one priced size
                sizes = [k for k in di_cells if k[0] == dish_item]
                cands = di_cells[sizes[0]] if len(sizes) == 1 else None
            src = _dish_price_expr(cands, dgroup) if cands else None
        if src:
            c = ws.cell(row=rr, column=VAL_PLAIN_COL,
                        value=f'=IF(TRIM(({src})&"")="","",IFERROR(VALUE(SUBSTITUTE(({src})&"","$","")),""))')
            c.number_format = '0.00'
    # Conditional formatting: Match? green TRUE / red FALSE / grey N/A
    if data_end >= data_start:
        mcol = get_column_letter(VAL_MATCH_COL)
        rng = f'{mcol}{data_start}:{mcol}{data_end}'
        ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"TRUE"'], fill=GREEN_FILL))
        ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"FALSE"'], fill=RED_FILL))
        ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"N/A"'], fill=GREY_FILL))
        # Orange the edited Price cell whenever it differs from the hidden baseline.
        pL = get_column_letter(OUT_PRICE_COL)
        bL = get_column_letter(VAL_BASELINE_COL)
        prng = f'{pL}{data_start}:{pL}{data_end}'
        # Red text first (stops there): what apply ignores — text that is not a
        # number (x included: not used on this tab), characters a price never
        # has, 1000 or more, 5× / a fifth of the live price (a misplaced decimal
        # point). Negatives are fine on rules. Same rules as _rule_entry.
        # 5× / a fifth against the LIVE rule price at Part 1 (hidden _ref_live)
        c, base = f'${pL}{data_start}', f'${get_column_letter(VAL_LIVE_COL)}{data_start}'
        val = f'VALUE(SUBSTITUTE(TRIM({c}),"$",""))'
        has = lambda ch: f'ISNUMBER(SEARCH("{ch}",{c}))'
        ratio = lambda x: (f'{base}<>0,ROUND({x},2)<>0,'          # rounded first, like apply
                           f'OR(ROUND(ABS({x}),2)>=5*ABS({base}),ROUND(ABS({x}),2)*5<=ABS({base}))')
        for f in (f'AND(ISTEXT({c}),TRIM({c}&"")<>"",ISERROR({val}))',
                  f'AND(ISTEXT({c}),OR({has(",")},{has("%")},{has(":")},{has("(")},{has("e")},{has("/")}))',
                  f'IFERROR(AND(ISNUMBER({c}),ROUND(ABS({c}),2)>=1000),FALSE)',
                  f'AND(ISTEXT({c}),TRIM({c}&"")<>"",IFERROR(ROUND(ABS({val}),2)>=1000,FALSE))',
                  f'IFERROR(AND(ISNUMBER({c}),ISNUMBER({base}),{ratio(c)}),FALSE)',
                  f'IFERROR(AND(ISTEXT({c}),ISNUMBER({base}),{ratio(val)}),FALSE)'):
            assert len(f) < 255, (len(f), f)
            ws.conditional_formatting.add(prng, FormulaRule(formula=[f], fill=INVALID_FILL, font=INVALID_FONT,
                                                            stopIfTrue=True))
        # (a conditional-format formula is stored WITHOUT a leading '=' — Excel
        # drops a rule written with one when it opens the file)
        orange_formula = (f'IFERROR(AND(${bL}{data_start}<>"",'
                          f'VALUE(SUBSTITUTE(${pL}{data_start}&"","$",""))<>${bL}{data_start}),FALSE)')
        ws.conditional_formatting.add(prng, FormulaRule(formula=[orange_formula], fill=ORANGE_FILL))
        ws.column_dimensions[bL].hidden = True
        ws.column_dimensions[get_column_letter(VAL_LIVE_COL)].hidden = True
        ws.column_dimensions[get_column_letter(VAL_KEY_COL)].hidden = True
    # Column widths
    ws.column_dimensions[get_column_letter(OUT_ENABLED_COL)].width = 10
    ws.column_dimensions[get_column_letter(OUT_PRICE_COL)].width = 10
    ws.column_dimensions[get_column_letter(OUT_DGROUP_COL)].width = 20
    ws.column_dimensions[get_column_letter(OUT_DISH_COL)].width = 28
    ws.column_dimensions[get_column_letter(OUT_VARIANT_COL)].width = 12
    ws.column_dimensions[get_column_letter(OUT_MOD_COL)].width = 28
    for c in (VAL_PLAIN_COL, VAL_OCRTOT_COL, VAL_EXPECT_COL, VAL_MATCH_COL):
        ws.column_dimensions[get_column_letter(c)].width = 16
    # Paper-menu multi-column pricing photo, top-right past the Match? column.
    # Embedded here (inside the render) so it persists across the del/create
    # regeneration of this tab. Resolve default crop when none is passed.
    _embed_pricing_image(ws, img_path or _default_crop_path())
    return counts


PABS_RULES = 'P-ABS — price on modifier rules'      # action3's Category for these rows
EDIT_FONT = Font(bold=True, color='C00000')


def _point_pabs_rows(wb, A):
    """Questionnaire P-ABS rows priced on modifier rules: fill in the exact
    Modifier Rule row to edit (the dish's By Itself row), now that the tab
    exists. action3 wrote the rows with the dish as "<dish> [<variant>]" and a
    provisional pointer in the Details column. Idempotent."""
    if 'Questionnaire' not in wb.sheetnames:
        return 0
    ws = wb['Questionnaire']
    hdr = [str(c.value).strip() if c.value else '' for c in ws[1]]
    if 'Details' not in hdr or 'Dish Name' not in hdr:
        return 0
    dcol, ncol, ccol = hdr.index('Details') + 1, hdr.index('Dish Name') + 1, 2
    pointed = 0
    for r in range(2, ws.max_row + 1):
        qid = str(ws.cell(row=r, column=1).value or '')
        if not qid.startswith('P-ABS-') or ws.cell(row=r, column=ccol).value != PABS_RULES:
            continue
        m = re.match(r'^(.*) \[(.*)\]$', str(ws.cell(row=r, column=ncol).value or ''))
        if not m:
            continue
        dish, variant = m.group(1), m.group(2)
        old = str(ws.cell(row=r, column=dcol).value or '')
        tail = old.split(' | ', 1)[1] if ' | ' in old else old
        # its own size first, else a By Itself rule for every size (no size)
        hits = (sorted(row for (d, v, mod), row in A.get('mrv_rows', {}).items()
                       if d == dish and v == variant and _is_by_itself(mod))
                or sorted(row for (d, v, mod), row in A.get('mrv_rows', {}).items()
                          if d == dish and v == '' and _is_by_itself(mod)))
        cell = ws.cell(row=r, column=dcol)
        if hits:
            where = (f"EDIT → 'Modifier Rule' tab · row {hits[0]} · column B (out_unit_price) — "
                     f"the By Itself row of {dish} [{variant}], the price of the plain dish.")
            cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{MRV_TAB}'!B{hits[0]}")
        else:
            where = (f"No By Itself row for {dish} [{variant}] on the Modifier Rule tab — write the "
                     f"plain price you want in Notes → Wonka on this row.")
        cell.value = f'{where} | {tail}'
        cell.font = EDIT_FONT
        pointed += 1
    return pointed


def _append_fried_specialty(ws, A):
    """Append delta + propagation below Action-3's own section (idempotent)."""
    # Strip any prior compare append (rows at/after the sentinel).
    sentinel_row = None
    for row in ws.iter_rows(min_col=1, max_col=1):
        if row[0].value == FS_SENTINEL:
            sentinel_row = row[0].row
            break
    if sentinel_row:
        ws.delete_rows(sentinel_row, ws.max_row - sentinel_row + 1)
    # a re-run starts where the previous block started (never one row lower)
    r = sentinel_row if sentinel_row else ws.max_row + 2
    # Section 2: Price difference — one block per multi-column paper section
    ws.cell(row=r, column=1, value=FS_SENTINEL).font = Font(bold=True, size=12)
    r += 1
    model = A.get('model')
    if model == MODEL_ABSOLUTE:
        note = ('ABSOLUTE model — side rules store the FULL total, so the values below are the '
                'totals written to modifier_rule, not differences. Plain goes on the By Itself rule.')
    elif model == MODEL_MIXED:
        note = ('MIXED model, decided per dish and size — ABSOLUTE rows store the FULL total (Plain '
                'goes on the By Itself rule); DELTA rows store the difference (Plain goes on '
                'dish_item). See the Model column.')
    else:
        note = None
    if note:
        ws.cell(row=r, column=1, value=note).font = Font(italic=True, color='7F6000')
        r += 1
    std_variants = ('', 'Specialty', 'Order', 'Small')
    for sec in A['sections']:
        ws.cell(row=r, column=1, value=f"Section: {sec['name']}").font = Font(bold=True)
        r += 1
        label = {MODEL_ABSOLUTE: '{} (total)', MODEL_MIXED: '{} (→ rule)'}.get(model, 'Δ {}')
        headers = ['Code', 'Name', 'Plain'] + [label.format(c) for c in sec['columns']]
        # Size / model per row, when they are not the same for every row
        per_row = model == MODEL_MIXED or any(
            e.get('sys') and e.get('variant') not in std_variants for e in sec['entries'])
        if per_row:
            headers += ['Size / variant', 'Model']
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font = HEADER_FONT
            c.fill = HEADER_FILL
        r += 1
        for item in sec['items']:
            ocr_name = item['name']
            e = item.get('_entry')
            ws.cell(row=r, column=1, value=item.get('code', ''))
            ws.cell(row=r, column=2, value=A['dish_map'].get(ocr_name, f'⚠ {ocr_name}'))
            ws.cell(row=r, column=3, value=_fmt(e['plain'] if e else None))
            for si, col in enumerate(sec['columns']):
                d = e['values'].get(col) if e else None
                if d is not None:
                    cell = ws.cell(row=r, column=4 + si, value=('$+%.2f' % d) if d >= 0 else _fmt(d))
                    cell.fill = BLUE_FILL if d > 0 else (PINK_FILL if d < 0 else GREEN_FILL)
            if per_row and e and e.get('sys'):
                vcol = 4 + len(sec['columns'])
                ws.cell(row=r, column=vcol, value=e['variant'] if e['variant'] is not None
                        else '⚠ size?')
                ws.cell(row=r, column=vcol + 1, value=e['model'] or '—')
            r += 1
        r += 1
    # Column → modifier + dish-name mappings.
    # Both blocks are anchored at column B (col index 2), not A: column A is too
    # narrow to read the longer column headers / dish names comfortably.
    ws.cell(row=r, column=2, value='COLUMN → MODIFIER MAPPING').font = Font(bold=True, size=11)
    r += 1
    for col, mods in A['col_to_mods'].items():
        ws.cell(row=r, column=2, value=col)
        ws.cell(row=r, column=3, value='→')
        ws.cell(row=r, column=4, value=', '.join(mods) if mods else '⚠ NO MATCH')
        r += 1
    r += 1
    # Dish name mapping (moved here from the former Delta Analysis tab).
    # List EVERY OCR dish from every multi-column section — not just matched
    # ones. Matched dishes show the resolved system name; a dish captured in OCR
    # that has no modifier_rule counterpart (i.e. it was never in the wsheet
    # modifier_rule tab) is flagged '⚠ NO MATCH' so the specialist sees it
    # instead of it silently vanishing from the table.
    ws.cell(row=r, column=2, value='DISH NAME MAPPING (OCR → System)').font = Font(bold=True, size=11)
    r += 1
    _seen_dn = set()
    for item in (it for sec in A['sections'] for it in sec['items']):
        o = item['name']
        if o in _seen_dn:
            continue
        _seen_dn.add(o)
        ws.cell(row=r, column=2, value=o)
        ws.cell(row=r, column=3, value='→')
        ws.cell(row=r, column=4, value=A['dish_map'].get(o, '⚠ NO MATCH'))
        r += 1
    r += 1
    # Section 3: propagation detail
    ws.cell(row=r, column=1, value='PROPAGATION DETAIL (per dish price-group)').font = Font(bold=True, size=12)
    r += 1
    for key in sorted(A['final_prices'].keys(),
                      key=lambda k: (k[0], _variant_sort_key(k[1]))):
        ws.cell(row=r, column=1, value=A['labels'].get(key, key[0])).font = Font(bold=True)
        r += 1
        for ci, h in enumerate(['Old Price', 'Modifier Item', 'New Price', 'Source', 'Changed?'], 1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font = HEADER_FONT
            c.fill = HEADER_FILL
        r += 1
        mod_old = A['mod_old_by_key'].get(key, {})
        groups = defaultdict(list)
        for mod, price in mod_old.items():
            groups[price].append(mod)
        for old_price in sorted(groups.keys()):
            for mod in groups[old_price]:
                new_p = A['final_prices'][key].get(mod, old_price)
                src = A['price_sources'][key].get(mod, '')
                changed = '→ CHANGED' if abs(new_p - old_price) > 0.001 else ('✓ confirmed' if src != 'NO DATA' else '— kept')
                ws.cell(row=r, column=1, value=_fmt(old_price))
                ws.cell(row=r, column=2, value=mod)
                ws.cell(row=r, column=3, value=_fmt(new_p))
                ws.cell(row=r, column=4, value=src)
                ws.cell(row=r, column=5, value=changed)
                if abs(new_p - old_price) > 0.001:
                    ws.cell(row=r, column=3).fill = BLUE_FILL if new_p > old_price else PINK_FILL
                elif src != 'NO DATA':
                    ws.cell(row=r, column=3).fill = GREEN_FILL
                r += 1
        r += 1
    if A['flags']:
        ws.cell(row=r, column=1, value='⚠ CONFLICTS — NOT AUTO-RESOLVED (left at the current price; '
                                       'confirm on the callback)').font = Font(bold=True, color='FF0000')
        r += 1
        for fl in A['flags']:
            ws.cell(row=r, column=1, value=fl)
            r += 1
        r += 1
    if A.get('notes'):
        ws.cell(row=r, column=1, value='NOTES').font = Font(bold=True)
        r += 1
        for nt in A['notes']:
            ws.cell(row=r, column=1, value=nt)
            r += 1


def _reorder_tabs(wb):
    order = [t for t in CANON_ORDER if t in wb.sheetnames]
    order += [t for t in wb.sheetnames if t not in order]
    wb._sheets.sort(key=lambda s: order.index(s.title))


def _apply_tab_styling(wb):
    """Colour the specialist-edited/Part-2-consumed tabs red and hide the
    read-forward 'modifier' baseline tab. Idempotent."""
    for ws in wb.worksheets:
        if ws.title in RED_TABS:
            ws.sheet_properties.tabColor = RED_TAB
        if ws.title in HIDDEN_TABS:
            ws.sheet_state = 'hidden'


def _strip_dollar_display(wb):
    """DISPLAY-ONLY: convert '$X.XX' string cells to bare numbers with a 0.00
    number format, across every tab. This does NOT change what Part 2 writes to
    the raw gsheet — apply() re-formats prices as '$X.XX' on write to match the
    production out_unit_price format. Part 2's reader uses _parse(), which
    tolerates both bare numbers and '$'-prefixed strings."""
    money_re = re.compile(r'^\$[+-]?\d+(?:\.\d+)?$')
    for ws in wb.worksheets:
        if ws.title == 'Dish Code':
            continue                    # codes are text, never prices
        questionnaire = ws.title == 'Questionnaire'
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and money_re.match(v.strip()):
                    signed = v.strip()[1] == '+'  # '$+1.20' → keep + sign on display
                    num = _parse(v)
                    if num is not None:
                        cell.value = num
                        cell.number_format = '+0.00;-0.00' if signed else '0.00'
                elif (questionnaire and isinstance(v, str) and cell.column in (5, 6, 7)
                      and '$' in v):
                    cell.value = v.replace('$', '')   # 'Small $8.95 / Large $12.95'


# ══════════════════════════════════════════════════════════════════════
#  compare
# ══════════════════════════════════════════════════════════════════════

def _print_coverage_warning(cov):
    """Loudly surface multi-column rows the OCR JSON left without side_prices.
    DROPPED rows never reach the Modifier Rule tab; PARTIAL rows are on the tab
    but missing some declared columns. Both are OCR gaps for the reviewer to
    verify against the paper menu (compare cannot invent the missing prices)."""
    if not cov or (not cov['dropped'] and not cov['partial']):
        return
    print(f"\n  {'⚠'} MULTI-COLUMN COVERAGE — '{cov['section']}'")
    print(f"    Section has {cov['total']} items; {cov['kept']} carry side_prices.")
    if cov['dropped']:
        print(f"    DROPPED ({len(cov['dropped'])}) — no side_prices, NOT on "
              f"Modifier Rule tab:")
        for code, name in cov['dropped']:
            print(f"      - {code or '?'}  {name}")
    if cov['partial']:
        print(f"    PARTIAL ({len(cov['partial'])}) — on tab but missing columns:")
        for code, name, miss in cov['partial']:
            print(f"      - {code or '?'}  {name}  → missing {miss}")
    print(f"    → Verify against the paper menu. If these rows DO have combo-column")
    print(f"      prices, the OCR JSON missed them (see ocr-prompt.md rule 5,")
    print(f"      partial rows). Add their side_prices and re-run compare.")


def cmd_compare(args):
    print(f"OCR JSON: {args.ocr_json}")
    print(f"Raw Gsheet: {args.raw_gsheet}")
    if args.conflict is not None:
        print(f"NOTE: --conflict {args.conflict} is ignored (v4.4): printed prices are never "
              f"overwritten and conflicts are never auto-resolved (G14).")
    A = _analyze(args.ocr_json, args.raw_gsheet)
    if A is not None:
        print(f"\nPricing model: {A['model']}  ({A['model_evidence']})")
        if A['model'] == MODEL_UNKNOWN:
            print("  WARNING: could not detect pricing model - review manually before proceeding.")
        if A['model'] in (MODEL_DOUBLE, MODEL_BROKEN):
            # Same verdict preflight blocks on — compare no longer relies on
            # preflight having been run first.
            print(f"!! BLOCKING VERDICT: {A['model']} - refusing to build the report. "
                  f"Resolve the live data (see preflight) and re-run.")
            sys.exit(1)
        _g11 = _check_g11(A)
        if _g11:
            print(f"\n!! DOUBLE-CHARGE ERROR (G11): {len(_g11)} dish/variant(s) priced on BOTH "
                  f"dish_item and a By Itself rule:")
            for _v in _g11:
                print(f"   - {_v}")
            print("   Refusing to build the report. Fix dish_item to $0 (or reclassify).")
            sys.exit(1)
        _print_cell_models(A['cells'])
    if A is None:
        print("No multi-column pricing data found. Nothing to do.")
        sys.exit(0)

    n_paper = len({e['name'] for sec in A['sections'] for e in sec['entries']})
    print(f"\nMulti-column sections: {len(A['sections'])}")
    for sec in A['sections']:
        print(f"  '{sec['name']}': {len(sec['items'])} items with side pricing; "
              f"columns {sec['columns']}")
    print(f"\nDish matches ({len(A['dish_map'])}/{n_paper}):")
    for o, s in A['dish_map'].items():
        print(f"  '{o}' → '{s}'")
    total_mr = sum(len(v) for v in A['mr_dish_rows'].values())
    print(f"\nModifier_rule rows found: {total_mr} across {len(A['mr_dish_rows'])} dishes")
    print(f"\nColumn → modifier mapping:")
    for col, mods in A['col_to_mods'].items():
        print(f"  '{col}' → {mods}")

    # Resolve the pricing image: explicit path, or default crop, or disabled.
    _img_arg = getattr(args, 'pricing_image', None)
    if _img_arg in ('', 'none', 'None'):
        img_path = None                     # explicitly disabled
    else:
        img_path = _img_arg or _default_crop_path()

    if args.report:
        wb = load_workbook(args.report)
        # Modifier Rule (own red tab — see _apply_tab_styling)
        if MRV_TAB in wb.sheetnames:
            del wb[MRV_TAB]
        ws_mrv = wb.create_sheet(MRV_TAB)
        counts = _render_mrv_tab(ws_mrv, A, img_path=img_path)
        _point_pabs_rows(wb, A)     # P-ABS rows: the exact By Itself row to edit
        # Remove any legacy separate Delta Analysis / Propagation Detail tabs;
        # their content now lives inside the Fried Specialty tab.
        for t in (DELTA_TAB, PROP_TAB):
            if t in wb.sheetnames:
                del wb[t]
        # Fold delta + mappings + propagation into Fried Specialty (below Action-3 section)
        if 'Fried Specialty' in wb.sheetnames:
            _append_fried_specialty(wb['Fried Specialty'], A)
            if wb['Fried Specialty']['A1'].value == FS_DRIFT_NOTE:
                print(f"\n!! DRIFT (still): {FS_DRIFT_NOTE}")
        else:
            # The OCR JSON has a multi-column table but the report has no Fried
            # Specialty tab: action2 / action3 ran on a different OCR JSON.
            # Build the tab from this analysis and say so — never silently.
            ws_fs = wb.create_sheet('Fried Specialty')
            ws_fs.cell(row=1, column=1, value=FS_DRIFT_NOTE).font = Font(bold=True, color='C00000')
            _append_fried_specialty(ws_fs, A)
            print(f"\n!! DRIFT: {FS_DRIFT_NOTE}")
        _reorder_tabs(wb)
        _strip_dollar_display(wb)  # display-only: '$X.XX' → bare 0.00 across all tabs
        _apply_tab_styling(wb)     # red edit-tabs + hide 'modifier'
        if 'Summary' in wb.sheetnames:
            wb.active = wb.sheetnames.index('Summary')  # active sheet must not be hidden
        out = args.output or args.report
        wb.save(out)
        wb.close()
        print(f"\n  Folded into report: {out}")
    else:
        out = args.output or 'modifier_rule_review.xlsx'
        wb = Workbook()
        counts = _render_mrv_tab(wb.active, A, img_path=img_path)
        wb.active.title = MRV_TAB
        wb.active.sheet_properties.tabColor = RED_TAB
        _strip_dollar_display(wb)  # display-only: '$X.XX' → bare 0.00
        wb.save(out)
        wb.close()
        print(f"\n  Output (standalone — --report recommended): {out}")

    print(f"\n{'='*60}")
    print(f"  Dishes matched: {len(A['dish_map'])} / {n_paper}")
    print(f"  Modifier_rule rows found: {total_mr}")
    print(f"  Cells updated (would change): {counts['updated']}")
    print(f"  Cells confirmed (no change): {counts['confirmed']}")
    print(f"  Cells no data (kept as-is): {counts['no_data']}")
    if A['flags']:
        print(f"\n  ⚠ Conflicts NOT auto-resolved (left at the current price — confirm on "
              f"the callback): {len(A['flags'])}")
        for fl in A['flags']:
            print(f"    {fl}")
    if A['notes']:
        print(f"\n  Notes:")
        for nt in A['notes']:
            print(f"    {nt}")
    for sec in A['sections']:
        _print_coverage_warning(sec.get('coverage'))
    print(f"  NOTE: Raw gsheet NOT modified — review only.")
    print(f"{'='*60}")
    print("Done.")


# ══════════════════════════════════════════════════════════════════════
#  apply
# ══════════════════════════════════════════════════════════════════════

def _read_mrv(report_path):
    """Read the approved rows of the compact Modifier Rule tab.

    Returns (entries, has_group): entries = [(key, out_unit_price)], one per
    tab row that carries a price, key = (dish_group, dish_item, variant,
    modifier_item) exactly as the row shows them. A group-level rule has an
    empty dish_item and is identified by its dish_group (v4.x dropped these
    rows: the key required a dish_item). A report made before the dish_group
    column existed yields key = (dish_item, variant, modifier_item) and
    has_group=False. Returns None when the report has no Modifier Rule tab.

    Columns are located by their HEADER LABEL — NOT by fixed position — so the
    tab may be reordered without breaking Part 2. Backward-compatible: if the
    legacy 3-row 'id' header is present instead, fall back to the raw fixed
    offsets."""
    wb = load_workbook(report_path)          # styles: the Price fill says what compare changed
    if MRV_TAB not in wb.sheetnames:
        wb.close()
        return None
    ws = wb[MRV_TAB]
    rows = list(ws.iter_rows(values_only=True))
    fills = [[(str(c.fill.fgColor.rgb or '')[-6:].upper()
               if c.fill is not None and c.fill.fill_type == 'solid' else '') for c in row]
             for row in ws.iter_rows()]
    pct = [['%' in str(c.number_format or '') for c in row] for row in ws.iter_rows()]
    wb.close()
    # A formula typed in Price counts by its saved result (none saved → reported)
    wbc = load_workbook(report_path, read_only=True, data_only=True)
    cached = list(wbc[MRV_TAB].iter_rows(values_only=True))
    wbc.close()
    ignored = []                  # [(key, typed value, why)] — invalid entries, never applied

    def _cellstr(v):
        return str(v).strip() if v is not None else ''

    # Preferred: find the wsheet-name header row. Columns are matched by SUBSTRING
    # (case-insensitive) so labels like 'dish_item', 'modifier_item',
    # 'out_unit_price', 'variant' all resolve. 'dish_item' is checked before the
    # bare 'dish' token so 'dish_group' can't mis-bind. The validation-block
    # labels ('by itself', 'Paper menu Total', 'Match?', ...) contain none of
    # these tokens, so they are never mistaken for a data column.
    def _find(labels, *tokens):
        for tok in tokens:
            for lbl, ci in labels.items():
                if tok in lbl:
                    return ci
        return None

    hdr_idx = None
    for i, r in enumerate(rows):
        if not r:
            continue
        labels = {_cellstr(c).lower(): ci for ci, c in enumerate(r) if _cellstr(c)}
        di = _find(labels, 'dish_item', 'dish')
        mi = _find(labels, 'modifier_item', 'modifier')
        pi = _find(labels, 'out_unit_price', 'price')
        if di is not None and mi is not None and pi is not None:
            vi = _find(labels, 'variant')
            gi = _find(labels, 'dish_group')
            hdr_idx = i
            break

    def _price_of(ri, ci):
        """(kind, number, typed) for a Price cell — formulas by their saved result."""
        v = rows[ri][ci] if ci < len(rows[ri]) else None
        if isinstance(v, str) and v.startswith('='):
            cv = cached[ri][ci] if ri < len(cached) and ci < len(cached[ri]) else None
            if cv is None:
                return 'invalid', None, v
            v = cv
        kind, num = _rule_entry(v)
        return kind, num, v

    def _get(r, ci):
        return _cellstr(r[ci]) if ci is not None and len(r) > ci else ''

    entries = []
    if hdr_idx is not None:
        # hidden columns by their header label; a tab made before the labels:
        # the fixed positions compare used
        hl = {_cellstr(c): ci for ci, c in enumerate(rows[hdr_idx]) if _cellstr(c)}
        bi = hl.get('_ref_baseline', VAL_BASELINE_COL - 1)   # compare's own price for the row
        li = hl.get('_ref_live', VAL_LIVE_COL - 1)           # the live price when Part 1 ran
        ki = hl.get('_ref_key', VAL_KEY_COL - 1)             # the row's own key
        has_baseline = any(len(r) > bi and r[bi] not in (None, '') for r in rows[hdr_idx + 1:] if r)
        has_live = any(len(r) > li and r[li] not in (None, '') for r in rows[hdr_idx + 1:] if r)
        # each row's hidden L / AD / fill, by its key (None: an older tab — by position)
        by_key = None
        if any(len(r) > ki and r[ki] not in (None, '') for r in rows[hdr_idx + 1:] if r):
            by_key = {}
            for ri2, r2 in enumerate(rows[hdr_idx + 1:], hdr_idx + 1):
                if r2 and len(r2) > ki and r2[ki] not in (None, ''):
                    by_key[str(r2[ki])] = r2
        for ri, r in enumerate(rows[hdr_idx + 1:], hdr_idx + 1):
            if not r or all(c in (None, '') for c in r):
                continue
            grp, dish, variant, mod = _get(r, gi), _get(r, di), _get(r, vi), _get(r, mi)
            if not (dish or grp or mod):
                continue
            kind, price, typed = _price_of(ri, pi)
            key = (grp, dish, variant, mod) if gi is not None else (dish, variant, mod)
            if kind == 'blank':
                continue
            if kind == 'price' and ri < len(pct) and pi < len(pct[ri]) and pct[ri][pi]:
                kind = 'invalid'             # typed as a percentage (Excel stores 10% as 0.1)
            if kind != 'price':
                ignored.append((key, typed, 'x is not used on the Modifier Rule tab — type the price, or '
                                'ask Wonka to disable the rule' if kind == 'x' else
                                'a formula with no saved result — open and save the report in Excel / '
                                'Google Sheets, or type the number' if str(typed).startswith('=') else
                                'not a valid price (letters, comma, %, 1000 or more)'))
                continue
            # An instruction = compare changed the price (blue / pink) or the
            # specialist edited it (≠ the hidden baseline). An untouched row
            # holds Part 1's live price — never written back over a newer one.
            hid = r
            if by_key is not None:
                hid = by_key.get('|'.join((grp, dish, variant, mod)), r)
            base = _parse(hid[bi]) if len(hid) > bi else None
            was = _parse(hid[li]) if len(hid) > li else None
            fill = fills[ri][pi] if ri < len(fills) and pi < len(fills[ri]) else ''
            # 5× / a fifth is judged against the LIVE rule price at Part 1 (compare's
            # own price may be the misread being corrected); a tab without it: the
            # baseline. Compare's untouched change is checked too.
            ref_ = was if was is not None else base
            if (ref_ is not None and (base is None or abs(price - base) > 0.001 or abs(base - ref_) > 0.001)
                    and _misplaced(price, ref_)):
                ignored.append((key, typed, f'looks like a misplaced decimal point ({_fmt(price)} vs the live '
                                f'{_fmt(ref_)} — {MISPLACED_RATIO}× or a fifth) — not applied; type the right price, or '
                                f'confirm in Notes → Wonka'))
                continue
            if not has_baseline or base is None or abs(price - base) > 0.001:
                instr = True                 # edited (or a tab too old to tell)
            elif has_live:
                instr = was is not None and abs(base - was) > 0.001    # compare changed it
            else:                            # older tab: compare's blue / pink fill
                instr = fill in (BLUE_FILL.fgColor.rgb[-6:].upper(), PINK_FILL.fgColor.rgb[-6:].upper())
            entries.append((key, price, instr))
        return entries, gi is not None, ignored

    # Fallback: legacy 3-row modifier_rule header (row whose col A == 'id').
    hdr = None
    for i, r in enumerate(rows):
        if r and str(r[0]).strip() == 'id':
            hdr = i
            break
    if hdr is None:
        return [], True, ignored
    for r in rows[hdr + 3:]:
        if not r or all(c in (None, '') for c in r):
            continue
        grp, dish = _get(r, DGROUP_COL), _get(r, DISH_COL)
        variant, mod = _get(r, VARIANT_COL), _get(r, MOD_COL)
        if not (dish or grp or mod):
            continue
        kind, price = _rule_entry(r[PRICE_COL] if len(r) > PRICE_COL else None)
        if kind == 'blank':
            continue
        if kind != 'price':
            ignored.append(((grp, dish, variant, mod), r[PRICE_COL], 'not a valid price'))
            continue
        entries.append(((grp, dish, variant, mod), price, True))
    return entries, True, ignored


def _key_label(key):
    if len(key) == 4:
        g, d, v, m = key
        who = f"dish='{d}'" if d else f"dish_group='{g}' (group-level)"
    else:
        d, v, m = key
        who = f"dish='{d}'"
    return f"({who}, variant='{v}', modifier='{m}')"


def cmd_apply(args):
    print(f"Price Comparison Report: {args.report}")
    print(f"Raw Gsheet: {args.raw_gsheet}")
    read = _read_mrv(args.report)
    # Re-detect model + G11 on the raw being patched (defense in depth before
    # writing). Run this on action5's output: the fresh fetch still holds the
    # old dish_item prices, so a price just written onto an ABSOLUTE $0 dish
    # would pass unnoticed there. Same blocking verdicts as preflight.
    _mr_data, _dip = _load_raw(args.raw_gsheet)
    _model, _evidence = _detect_pricing_model(_mr_data, _dip)
    print(f"Pricing model: {_model} ({_evidence})")
    if _model in (MODEL_DOUBLE, MODEL_BROKEN):
        print(f"!! BLOCKING VERDICT: {_model} - refusing to apply (nothing written). "
              f"Under ABSOLUTE pricing the dish_item must stay $0.00 — the base lives on the "
              f"By Itself rule. Fix the dish_item / report and re-run.")
        sys.exit(1)
    _dgroups = _load_dish_groups(args.raw_gsheet)
    _dishes_of = defaultdict(set)            # dish group → its dishes (a group-level rule prices them all)
    for _d, _gs in _dgroups.items():
        for _g in _gs:
            _dishes_of[_g].add(_d)

    def _free_at_zero(g, d, v):
        """The dishes this rule prices IN FULL — their dish_item is $0.00 or
        has no price for the size (ABSOLUTE): a dish-level rule its own dish, a
        group-level rule every dish of the group, a size-less rule every size."""
        out = []
        for dish in ([d] if d else sorted(_dishes_of.get(g, ()))):
            sizes = [p for (dd, vv), p in _dip.items() if dd == dish and (not v or vv == v)]
            if not sizes or min(abs(p) for p in sizes) <= 0.001:
                out.append(dish)
        return out
    _g11 = _check_g11_core(_mr_data, _dip, _model, _dgroups)
    if _g11:
        for _v in _g11:
            print(f"   - {_v}")
        print(f"!! DOUBLE-CHARGE ERROR (G11): {len(_g11)} violation(s) - refusing to apply.")
        sys.exit(1)
    if read is None:
        print("No 'Modifier Rule' tab in the report. Nothing to apply.")
        sys.exit(0)
    entries, has_group, ignored = read
    print(f"Approved out_unit_price rows: {len(entries)}")
    for k, typed, why in ignored:
        print(f"!! NOT A PRICE — ignored: {_key_label(k)} {str(typed)!r} — {why}")

    def key_of(g, d, v, m):
        return (g, d, v, m) if has_group else (d, v, m)

    tab_prices = defaultdict(list)
    tab_instr = defaultdict(bool)
    for key, price, instr in entries:
        tab_prices[key].append(price)
        tab_instr[key] = tab_instr[key] or instr
    # A report made before the dish_group column cannot identify a group-level
    # rule (no dish_item) — never guess which group it meant.
    unkeyed = [k for k in tab_prices if not has_group and not k[0]]
    for k in unkeyed:
        print(f"!! GROUP-LEVEL RULE NOT APPLIED: {_key_label(k)} — this report has no dish_group "
              f"column (made before v4.4), so the rule's group is unknown. Re-run Part 1 compare, "
              f"or edit it in wsheet directly.")

    wb = load_workbook(args.raw_gsheet)  # writable, preserves all other tabs
    ws = wb['modifier_rule']

    def _cv(row, ci):
        return str(row[ci].value).strip() if len(row) > ci and row[ci].value else ''

    # rows 1-3 are the 3-row header; data starts at row 4 (1-indexed)
    live = defaultdict(list)
    for row in ws.iter_rows(min_row=4):
        live[key_of(_cv(row, DGROUP_COL), _cv(row, DISH_COL), _cv(row, VARIANT_COL),
                    _cv(row, MOD_COL))].append(row)

    writes = []       # (price_cell, new_price)
    not_reverted = [] # untouched rows whose live price changed since Part 1
    held = []         # (key, old, new, why) — not applied by a guard
    no_live_price = []  # edited rows whose live rule has no price (never written, G1)
    matched = set()
    confirmed = 0
    ambiguous = []
    for key, prices in tab_prices.items():
        if key in unkeyed:
            continue
        rows = live.get(key, [])
        if not rows:
            continue                                   # → drift, below
        matched.add(key)
        cells = [r[PRICE_COL] for r in rows]
        olds = [_parse(c.value) for c in cells]
        if len(rows) == 1 and len(prices) == 1:
            old_price, new_price = olds[0], prices[0]
            # G1: never write a blank cell (protects disabled group-gate rows)
            if old_price is None:
                if tab_instr[key]:
                    no_live_price.append(key)
                continue
            # G4: only write when the value actually changed
            if abs(new_price - old_price) <= 0.001:
                confirmed += 1
                continue
            if not tab_instr[key]:
                # untouched row, the live price moved since Part 1 — keep it
                not_reverted.append(key)
                continue
            d_, v_, m_ = (key[1], key[2], key[3]) if has_group else key
            g_ = key[0] if has_group else ''
            free = _free_at_zero(g_, d_, v_) if (old_price > 0.001 and new_price <= 0.001) else []
            if free:
                # G11: under ABSOLUTE this rule IS the price (By Itself = the
                # plain dish, a side rule = the full combo) — at $0.00 it is free
                more = f' (+{len(free) - 1} more)' if len(free) > 1 else ''
                held.append((key, old_price, new_price,
                             f"the full price of {free[0]}{more} [{v_ or 'every size'}] — a dish with no "
                             f"dish price ($0.00 / blank: ABSOLUTE) priced on this rule; at $0.00 it would "
                             f"be free"))
                continue
            writes.append((cells[0], new_price))
            continue
        # The tab key matches several live rows (they differ only in columns
        # the tab does not show, e.g. modifier_group / premod_item). Unchanged
        # is fine; an edit cannot be attributed to one row, so it is not applied.
        if sorted(p for p in olds if p is not None) == sorted(prices):
            confirmed += len(prices)
            continue
        if not tab_instr[key]:
            not_reverted.append(key)       # untouched — the live rows moved since Part 1
            continue
        ambiguous.append(key)
        print(f"!! AMBIGUOUS: Modifier Rule row(s) {_key_label(key)} match {len(rows)} live "
              f"modifier_rule rows (live {[_fmt(p) for p in olds]}, tab {[_fmt(p) for p in prices]}) "
              f"— NOT applied. Edit these in wsheet directly.")

    # G8: drift — approved rows never matched to a live row
    drift = [k for k in tab_prices if k not in matched and k not in unkeyed]
    for k in drift:
        print(f"!! DRIFT WARNING: approved {_key_label(k)} not found in live modifier_rule — "
              f"NOT applied (renamed/removed since Part 1).")

    # G11 on the RESULTING state: this raw's dish_item (post-action5) plus the
    # out_unit_price values about to be written. Nothing is saved if it fails.
    post_mr = [list(r) for r in _mr_data]
    for cell, new_price in writes:
        i = cell.row - 4
        if 0 <= i < len(post_mr) and len(post_mr[i]) > PRICE_COL:
            post_mr[i][PRICE_COL] = _fmt(new_price)
    post_model, post_evidence = _detect_pricing_model(post_mr, _dip)
    post_g11 = _check_g11_core(post_mr, _dip, post_model, _dgroups)
    if post_model in (MODEL_DOUBLE, MODEL_BROKEN) or post_g11:
        for _v in post_g11:
            print(f"   - {_v}")
        print(f"!! BLOCKING: the result would be {post_model} ({post_evidence}) — refusing to "
              f"apply (nothing written).")
        sys.exit(1)

    for cell, new_price in writes:
        cell.value = _fmt(new_price)  # G6: $X.XX
    written = len(writes)

    out = args.output or 'final_raw.xlsx'
    wb.save(out)
    wb.close()

    # ── Deterministic machine-readable summary (anti-hallucination, G12) ──
    # Merge modifier_rule counts into the SAME part2_summary.json action5 wrote,
    # so part2_report.py renders one report from ground truth. Never hand-count
    # or recall these values.
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

    def _as_dict(k):
        if len(k) == 4:
            return {'dish_group': k[0], 'dish': k[1], 'variant': k[2], 'modifier': k[3]}
        return {'dish': k[0], 'variant': k[1], 'modifier': k[2]}

    _existing['modifier_rule'] = {
        'pricing_model': _model,
        'approved_triples': len(entries),
        'matched_live_rows': len(matched),
        'out_unit_price_cells_written': written,
        'confirmed_no_change': confirmed,
        'not_reverted': {
            'total': len(not_reverted),
            'triples': [_as_dict(k) for k in not_reverted[:200]],
        },
        'ignored_input': [dict(_as_dict(k), value=str(t), why=w) for k, t, w in ignored],
        'held': [dict(_as_dict(k), old=_fmt(o), new=_fmt(n), why=w) for k, o, n, w in held],
        'no_live_price': [_as_dict(k) for k in no_live_price],
        'drift': {
            'total': len(drift),
            'triples': [_as_dict(k) for k in drift],
        },
        'not_applied': {
            'total': len(ambiguous) + len(unkeyed),
            'ambiguous': [_as_dict(k) for k in ambiguous],
            'group_level_unkeyed': [_as_dict(k) for k in unkeyed],
        },
        'output_xlsx': out,
    }
    with open(_summary_path, 'w', encoding='utf-8') as _f:
        _json.dump(_existing, _f, indent=2)

    print(f"\n{'='*60}")
    print(f"  Approved rows: {len(entries)}")
    print(f"  Matched live rows: {len(matched)}")
    print(f"  out_unit_price cells written (changed): {written}")
    print(f"  Confirmed (no change, skipped): {confirmed}")
    if not_reverted:
        print(f"  Untouched rows whose live price changed since Part 1 — kept, NOT reverted: "
              f"{len(not_reverted)}")
    for k, o, n, w in held:
        print(f"  !! HELD (G11): {_key_label(k)} {_fmt(o)} → {_fmt(n)} NOT applied — {w}")
    for k in no_live_price:
        print(f"  !! NOT APPLIED: {_key_label(k)} — this rule has no price in the live menu (a disabled "
              f"/ gate rule); a price is never written into it (G1)")
    if ignored:
        print(f"  Not a price — ignored: {len(ignored)}")
    print(f"  Drift warnings (NOT applied): {len(drift)}")
    if ambiguous or unkeyed:
        print(f"  Ambiguous / unkeyed rows (NOT applied): {len(ambiguous) + len(unkeyed)}")
    print(f"  Post-apply check: {post_model}, G11 OK")
    print(f"  Output: {out}")
    print(f"  Summary JSON: {_summary_path}  (source of truth for part2_report.py)")
    print(f"  Only the modifier_rule tab was touched (G3); all other tabs byte-identical.")
    print(f"{'='*60}")
    if drift or ambiguous or unkeyed:
        print("Resolve every DRIFT / AMBIGUOUS / GROUP-LEVEL warning before staging the draft.")
    # NEXT STEP: Part 2 output is a STAGED DRAFT. Commit is a separate step done
    # only on explicit user request, through ACE's gated `wsheet draft commit`
    # (inline approval). See guardrails G9/G10 and the "ACE-native gating" section.
    print()
    print(f"{'#'*60}")
    print("# NEXT STEP — STAGE A DRAFT (do NOT commit proactively)")
    print("#   1. Reverse-patch the TSV from this xlsx (bridge to-tsv).")
    print("#   2. wsheet validate <dir> --level full   (must exit 0)")
    print("#   3. .venv/bin/python scripts/part2_report.py <dir>   (report from JSON)")
    print("#   4. wsheet draft create <dir> --name \"...\"   ← stage the draft")
    print("#   STOP and report the staged draft. Commit ONLY if the user asks to")
    print("#   push to prod, via the gated `wsheet draft commit` (inline approval).")
    print(f"{'#'*60}")
    print("Done.")


def _print_cell_models(cells, limit=8):
    """Which dish/variants are priced ABSOLUTE vs DELTA (a restaurant may mix them)."""
    by_model = defaultdict(list)
    for (dish, variant), (m, _b) in sorted(cells.items()):
        by_model[m].append(f"{dish} [{variant}]" if variant else dish)
    for m in (MODEL_ABSOLUTE, MODEL_DELTA, MODEL_DOUBLE):
        if by_model.get(m):
            names = by_model[m]
            more = f" (+{len(names) - limit} more)" if len(names) > limit else ''
            print(f"  {m:13} {len(names):3} dish/variant(s): {', '.join(names[:limit])}{more}")


def cmd_preflight(args):
    """Mandatory gate: classify the pricing model and refuse to continue on a
    DOUBLE_CHARGE / BROKEN verdict. Exit 0 = safe to proceed."""
    mr_data, dish_item_prices = _load_raw(args.raw_gsheet)
    dish_groups = _load_dish_groups(args.raw_gsheet)
    model, evidence = _detect_pricing_model(mr_data, dish_item_prices)
    print(f"Pricing model: {model}  ({evidence})")
    _print_cell_models(_cell_models(mr_data, dish_item_prices))
    if model in (MODEL_DOUBLE, MODEL_BROKEN):
        for v in _check_g11_core(mr_data, dish_item_prices, model, dish_groups):
            print(f"   - {v}")
        print(f"!! BLOCKING VERDICT: {model} - do NOT run compare/apply or stage a draft.")
        sys.exit(1)
    if model == MODEL_UNKNOWN:
        print("WARNING: UNKNOWN model - no By-Itself anchor found. Review manually; not guessing.")
        sys.exit(2)
    g11 = _check_g11_core(mr_data, dish_item_prices, model, dish_groups)
    if g11:
        for v in g11:
            print(f"   - {v}")
        print(f"!! DOUBLE-CHARGE ERROR (G11): {len(g11)} violation(s). Blocking.")
        sys.exit(1)
    print(f"Preflight OK - {model}. Safe to proceed.")
    sys.exit(0)


def main():
    parser = argparse.ArgumentParser(description='Multi-column pricing → modifier_rule (compare + apply)')
    sub = parser.add_subparsers(dest='cmd', required=True)

    pc = sub.add_parser('compare', help='Part 1: build review views (optionally fold into the report)')
    pc.add_argument('ocr_json', help='OCR JSON file (with side_pricing sections)')
    pc.add_argument('raw_gsheet', help='Raw Gsheet Excel (with modifier_rule tab)')
    pc.add_argument('--report', help='Price Comparison Report to fold the analysis into')
    pc.add_argument('--output', '-o', help='Write to a copy instead of in place / default standalone name')
    pc.add_argument('--conflict', choices=['mode', 'lowest', 'highest'], default=None,
                    help='Ignored since v4.4 (accepted so older runbooks still parse): printed '
                         'prices keep their own value and conflicts are never auto-resolved (G14)')
    pc.add_argument('--pricing-image', default=None,
                    help='PNG of the paper-menu multi-column pricing block to embed on '
                         'the Modifier Rule tab (default: crops/special_full.png if present; '
                         'produce one with crop_section.py). Use "" or "none" to disable.')
    pc.set_defaults(func=cmd_compare)

    pp = sub.add_parser('preflight',
                        help='Detect pricing model + run G11 double-charge check on a raw gsheet')
    pp.add_argument('raw_gsheet', help='Raw Gsheet Excel (P1: the fresh wsheet fetch; '
                                       "P2: action5's final_raw.xlsx)")
    pp.set_defaults(func=cmd_preflight)

    pa = sub.add_parser('apply', help='Part 2: write approved out_unit_price back into modifier_rule')
    pa.add_argument('report', help='Price Comparison Report with the edited Modifier Rule tab')
    pa.add_argument('raw_gsheet', help="action5's output (final_raw.xlsx) to patch — the "
                                       "post-action5 state, so G11 sees the dish_item prices "
                                       "production will get (a copy is written)")
    pa.add_argument('--output', '-o', help='Output path (default: final_raw.xlsx)')
    pa.set_defaults(func=cmd_apply)

    args = parser.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()