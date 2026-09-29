#!/usr/bin/env python3
"""
modifier_rule_pricing.py  (compare + apply)
===========================================

Integrated multi-column-pricing → modifier_rule tool. Dependencies: openpyxl only.

Subcommands
-----------
  compare <ocr_json> <raw_gsheet.xlsx> [--report <price_comparison_report.xlsx>]
          [-o <output.xlsx>] [--conflict mode|lowest|highest]

      Part-1 (review) half. Finds side_pricing in the OCR JSON, computes deltas
      (combo_total - plain) per dish x column, matches OCR dish names + paper
      columns to the system modifier_rule rows, and propagates by price-group.
      Produces the analysis in exactly two places inside the report:
        - "Modifier Rule"  — a COMPACT 5-column edit view for matched
          dishes (Enabled | Price | Dish | Variant | Modifier) under a single
          friendly header row, colour coded (Blue=increase, Pink=decrease,
          Green=confirmed, Orange=specialist edit). To its right (after a blank
          spacer col) sits the OCR validation block (by-itself, Paper menu
          Total, live Expected=by-itself+edited-delta, Match? TRUE/FALSE/N/A) —
          NOT imported by Part 2. The Paper menu Total is filled for every row
          carrying a resolved delta (PAPER MENU printed, or PROPAGATED / PRICE
          TWIN derived); the Expected formula is present on every data row. Its
          own RED tab — the tab the specialist edits and `apply` reads (by
          header label, not fixed position).
        - "Fried Specialty"     — Action-3's section, appended below with the
          price-difference deltas, column→modifier + dish-name mappings, and
          per-dish price-group propagation detail.
      There are NO separate "Delta Analysis" / "Propagation Detail" tabs; if a
      prior run created them they are removed. With --report the content is
      folded INTO the existing Price Comparison Report, keeping one single
      workbook. Re-running replaces only compare's own content. compare NEVER
      modifies the Raw Gsheet.

  apply <price_comparison_report.xlsx> <raw_gsheet.xlsx> [-o <final_raw.xlsx>]

      Part-2 (write) half. Reads the specialist-approved out_unit_price values
      from the report's "Modifier Rule" tab, matches each
      (dish_item, variant, modifier_item) triple to the live modifier_rule rows,
      and overwrites ONLY changed out_unit_price cells in a copy of the Raw
      Gsheet. Guardrails: never writes a blank price (G1), never adds/removes
      rows (G2), touches only the modifier_rule tab (G3), only writes cells that
      actually changed (G4), and prints "!! DRIFT WARNING" for any approved
      triple not found live (G8) without applying it.

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

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import PatternFill, Font
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
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
# The anchor row is not always literally "By Itself"; match a normalized alias.
BY_ITSELF_ALIASES = ['by itself', 'plain', 'no rice', 'no side']
MODEL_ABSOLUTE = 'ABSOLUTE'
MODEL_DELTA = 'DELTA'
MODEL_DOUBLE = 'DOUBLE_CHARGE'
MODEL_BROKEN = 'BROKEN'
MODEL_UNKNOWN = 'UNKNOWN'

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
DELTA_TAB = 'Delta Analysis'
PROP_TAB = 'Propagation Detail'
# canonical tab order used by the report. Delta Analysis + Propagation Detail
# are NOT separate tabs — their content lives inside the Fried Specialty tab.
# 'Modifier' is hidden (read-forward baseline only) so it sorts to the end.
CANON_ORDER = ['Summary', 'Dish Item', 'Questionnaire', 'Price Inference',
               'Modifier Summary', 'Fried Specialty', MRV_TAB,
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
    if val is None:
        return ''
    try:
        return f'${float(val):.2f}'
    except Exception:
        return ''


def _is_by_itself(mod_name: str) -> bool:
    """True if a modifier_item name normalizes to a base-price anchor alias."""
    n = _normalize(mod_name)
    return any(a in n for a in BY_ITSELF_ALIASES)


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


def _detect_pricing_model(mr_data, dish_item_prices):
    """Classify the pricing model from the By-Itself anchor row cross-checked
    against dish_item. Returns (model, evidence_str).

        By-Itself priced + dish_item $0    -> ABSOLUTE
        By-Itself absent  + dish_item >$0  -> DELTA
        By-Itself priced  + dish_item >$0  -> DOUBLE_CHARGE  (block)
        By-Itself absent  + dish_item $0   -> BROKEN         (block)
        no anchor found anywhere           -> UNKNOWN        (warn, don't guess)
    """
    anchor = None   # (dish, variant, price)
    for row in mr_data:
        mod = str(row[MOD_COL]).strip() if len(row) > MOD_COL and row[MOD_COL] else ''
        if mod and _is_by_itself(mod):
            price = _parse(row[PRICE_COL]) if len(row) > PRICE_COL else None
            dish = str(row[DISH_COL]).strip() if len(row) > DISH_COL and row[DISH_COL] else ''
            variant = str(row[VARIANT_COL]).strip() if len(row) > VARIANT_COL and row[VARIANT_COL] else ''
            if price is not None:
                anchor = (dish, variant, price)
                break
    by_itself_priced = anchor is not None and anchor[2] > 0
    if anchor:
        dish, variant, _p = anchor
        dp = dish_item_prices.get((dish, variant))
        if dp is None:
            dp = next((p for (d, v), p in dish_item_prices.items() if d == dish), None)
        dish_priced = dp is not None and dp > 0
    else:
        dish_priced = any(p > 0 for p in dish_item_prices.values())

    if by_itself_priced and not dish_priced:
        return MODEL_ABSOLUTE, f"By-Itself anchor {anchor} priced; dish_item $0"
    if not anchor and dish_priced:
        return MODEL_DELTA, "no By-Itself anchor; dish_item carries base price"
    if by_itself_priced and dish_priced:
        return MODEL_DOUBLE, f"By-Itself anchor {anchor} priced AND dish_item >$0 - DOUBLE CHARGE"
    if not anchor and not dish_priced:
        return MODEL_BROKEN, "no By-Itself anchor and dish_item $0 - no base price anywhere"
    return MODEL_UNKNOWN, "could not resolve a By-Itself anchor; not guessing"


def _check_g11_core(mr_data, dish_item_prices, model):
    """G11 (ABSOLUTE-only): every (dish, variant) with a priced modifier_rule row
    must have dish_item == $0, else customers pay base + full total (~double).
    Returns a list of violation strings (empty = OK). No-op for non-ABSOLUTE."""
    if model != MODEL_ABSOLUTE:
        return []
    violations = []
    for row in mr_data:
        price = _parse(row[PRICE_COL]) if len(row) > PRICE_COL else None
        if price is None or price <= 0:
            continue
        dish = str(row[DISH_COL]).strip() if len(row) > DISH_COL and row[DISH_COL] else ''
        variant = str(row[VARIANT_COL]).strip() if len(row) > VARIANT_COL and row[VARIANT_COL] else ''
        if not dish:
            continue
        dp = dish_item_prices.get((dish, variant))
        if dp is not None and dp > 0:
            violations.append(f"(dish='{dish}', variant='{variant}') dish_item=${dp:.2f} "
                              f"but modifier_rule priced ${price:.2f} - DOUBLE CHARGE")
    return violations


def _check_g11(A):
    """G11 wrapper for the analysis payload (cmd_compare path)."""
    return _check_g11_core(A['mr_data'], A['dish_item_prices'], A.get('model'))


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
    if 'fried rice' in expanded:
        expanded.append('plain fried rice')
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

def _analyze(ocr_json_path, raw_gsheet_path, conflict_mode):
    """Return the full analysis payload used to render compare's views."""
    with open(ocr_json_path) as f:
        ocr = json.load(f)

    side_section = None
    side_columns = []
    side_items = []
    side_coverage = None
    for sec in ocr.get('sections', []):
        sp = sec.get('side_pricing')
        if sp and sp.get('columns'):
            side_section = sec['menu_section']
            side_columns = sp['columns']
            all_items = sec.get('items', [])
            side_items = [it for it in all_items if it.get('side_prices')]
            # Coverage audit: the paper section's full item list is ground truth.
            # A row with NO side_prices is DROPPED (never reaches the tab); a kept
            # row still missing some declared columns is PARTIAL. Both are OCR
            # gaps to surface — silent filtering here is how D7/D8-type rows vanish.
            dropped = [it for it in all_items if not it.get('side_prices')]
            partial = [
                (it, [c for c in side_columns if (it.get('side_prices') or {}).get(c) is None])
                for it in side_items
            ]
            partial = [(it, miss) for it, miss in partial if miss]
            side_coverage = {
                'section': side_section,
                'total': len(all_items),
                'kept': len(side_items),
                'dropped': [(it.get('code', ''), it.get('name', '')) for it in dropped],
                'partial': [(it.get('code', ''), it.get('name', ''), miss)
                            for it, miss in partial],
            }
            break
    if not side_items:
        legacy = ocr.get('fried_specialty_sides')
        if legacy and legacy.get('items'):
            side_section = 'Fried Specialties'
            side_columns = legacy.get('columns', [])[1:]
            side_items = legacy['items']
    if not side_items:
        return None

    # Expand folded compound "X or Y or Z Sauce" lines into per-sauce siblings
    # BEFORE delta computation, so each sibling carries the shared upcharge.
    side_items = _expand_side_items(side_items)

    # Detect pricing model (ABSOLUTE vs DELTA) before computing deltas.
    _mr_data_pre, _dish_item_prices = _load_raw(raw_gsheet_path)
    model, model_evidence = _detect_pricing_model(_mr_data_pre, _dish_item_prices)

    # deltas
    deltas = {}
    plain_prices = {}
    for item in side_items:
        name = item['name']
        if 'side_prices' in item:
            plain = None
            for v in ['Specialty', 'Order', 'Small']:
                if v in item.get('prices', {}):
                    plain = item['prices'][v]
                    break
            sp = item['side_prices'] or {}
        else:
            prices = item.get('prices', [])
            plain = prices[0] if prices else None
            sp = {side_columns[i]: prices[i + 1] for i in range(len(side_columns)) if i + 1 < len(prices)}
        if plain is None:
            continue
        plain_prices[name] = plain
        item_deltas = {}
        for col, total in sp.items():
            if total is not None:
                if model == MODEL_ABSOLUTE:
                    item_deltas[col] = round(total, 2)         # full total, no subtraction
                else:
                    item_deltas[col] = round(total - plain, 2)  # DELTA (existing behavior)
        deltas[name] = item_deltas

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
    dish_map = {}
    for ocr_name in deltas:
        m = _match_dish_name(ocr_name, system_names)
        if m:
            dish_map[ocr_name] = m

    sys_deltas = {}
    sys_plain = {}
    for ocr_name, sys_name in dish_map.items():
        sys_deltas[sys_name] = deltas[ocr_name]
        sys_plain[sys_name] = plain_prices[ocr_name]

    # System-driven selection is a UNION of all criteria (not limited to
    # 'Specialty'), so folded/misread compound lines can no longer silently
    # drop siblings. A modifier_rule row is selected if ANY of:
    #   (a) its dish_item has a 'Specialty' variant anywhere in modifier_rule,
    #   (b) its dish_item was matched from the paper multi-column table (sys_deltas),
    #   (c) its dish_group is a Fried Specialty group (group-level rules that
    #       carry no dish_item are kept too — selected per spec).
    specialty_dishes = set()
    for row in mr_data:
        d = str(row[DISH_COL]).strip() if len(row) > DISH_COL and row[DISH_COL] else ''
        v = str(row[VARIANT_COL]).strip() if len(row) > VARIANT_COL and row[VARIANT_COL] else ''
        if d and v == 'Specialty':
            specialty_dishes.add(d)
    selected_dishes = set(sys_deltas) | specialty_dishes

    def _is_fried_specialty_group(g):
        return bool(g) and 'fried special' in g.lower()

    mr_dish_rows = defaultdict(list)
    all_modifier_items = set()
    for ri, row in enumerate(mr_data):
        dish = str(row[DISH_COL]).strip() if len(row) > DISH_COL and row[DISH_COL] else ''
        grp = str(row[DGROUP_COL]).strip() if len(row) > DGROUP_COL and row[DGROUP_COL] else ''
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

    col_to_mods = {}
    for col in side_columns:
        col_to_mods[col] = _match_column_to_modifiers(col, list(all_modifier_items))

    mod_known_delta = defaultdict(dict)
    for dish, dd in sys_deltas.items():
        for col, delta in dd.items():
            for mod in col_to_mods.get(col, []):
                mod_known_delta[dish][mod] = delta

    final_prices = defaultdict(dict)
    price_sources = defaultdict(dict)
    flags = []
    for dish, rows in mr_dish_rows.items():
        mod_old = {}
        for _, row in rows:
            mod = str(row[MOD_COL]).strip() if len(row) > MOD_COL and row[MOD_COL] else ''
            price = _parse(row[PRICE_COL]) if len(row) > PRICE_COL else None
            if mod and price is not None:
                mod_old[mod] = price
        price_groups = defaultdict(list)
        for mod, price in mod_old.items():
            price_groups[price].append(mod)
        for old_price, members in price_groups.items():
            known = {m: mod_known_delta[dish][m] for m in members if m in mod_known_delta.get(dish, {})}
            if known:
                uniq = list(known.values())
                if len(set(uniq)) > 1:
                    # auto-resolve conflict
                    if conflict_mode == 'lowest':
                        winner = min(uniq)
                    elif conflict_mode == 'highest':
                        winner = max(uniq)
                    else:  # mode, tie-break lowest
                        counts = defaultdict(int)
                        for v in uniq:
                            counts[v] += 1
                        top = max(counts.values())
                        winner = min(v for v, c in counts.items() if c == top)
                    flags.append(
                        f"CONFLICT: {dish} ${old_price:.2f} group had disagreeing new prices "
                        f"{known}; auto-picked ${winner:.2f} (--conflict {conflict_mode}); "
                        f"discarded {sorted(set(uniq) - {winner})}")
                    new_price = winner
                else:
                    new_price = uniq[0]
                for m in members:
                    final_prices[dish][m] = new_price
                    price_sources[dish][m] = 'PAPER MENU' if m in known else 'PROPAGATED'
            else:
                for m in members:
                    final_prices[dish][m] = old_price
                    price_sources[dish][m] = 'NO DATA'

    # Price-twin propagation (Q1): a Specialty/Fried-Specialty dish with NO paper
    # data borrows the delta from a "price-twin" — a dish that DID get paper data
    # for the SAME modifier_item at the SAME old base price. Fried Specialties is
    # a uniform upcharge family, so equal (modifier, old_price) pairs move
    # together. Build the twin table from resolved (PAPER MENU/PROPAGATED) rows,
    # then fill NO DATA rows whose (modifier, old_price) has a unique twin delta.
    # Ambiguous cases are never guessed — they are reported as TWIN CONFLICT.
    # Twins are looked up DISH-GROUP-FIRST, then globally. Group scope is the
    # tighter cohort (a dish group is priced as a family), and it is strictly
    # additive: a globally-unanimous delta stays unanimous inside any subset, so
    # group-first can only resolve cases the global key left ambiguous — it can
    # never contradict a global-unique answer.
    dish_group_of = {}
    for dish, rows in mr_dish_rows.items():
        for _, row in rows:
            g = str(row[DGROUP_COL]).strip() if len(row) > DGROUP_COL and row[DGROUP_COL] else ''
            if g:
                dish_group_of[dish] = g
                break
    mod_old_by_dish = defaultdict(dict)
    for dish, rows in mr_dish_rows.items():
        for _, row in rows:
            m = str(row[MOD_COL]).strip() if len(row) > MOD_COL and row[MOD_COL] else ''
            op = _parse(row[PRICE_COL]) if len(row) > PRICE_COL else None
            if m and op is not None and m not in mod_old_by_dish[dish]:
                mod_old_by_dish[dish][m] = op

    twin_delta = defaultdict(set)        # (modifier, old_price) -> {delta, ...}
    twin_delta_grp = defaultdict(set)    # (group, modifier, old_price) -> {delta, ...}
    for dish, mods in final_prices.items():
        for m, new_p in mods.items():
            if price_sources[dish].get(m) in ('PAPER MENU', 'PROPAGATED'):
                op = mod_old_by_dish[dish].get(m)
                if op is not None:
                    d = round(new_p - op, 2)
                    twin_delta[(m, round(op, 2))].add(d)
                    twin_delta_grp[(dish_group_of.get(dish, ''), m, round(op, 2))].add(d)
    twins_applied = 0
    twins_group_scoped = 0
    twin_conflicts = []
    for dish, mods in final_prices.items():
        for m, cur in list(mods.items()):
            if price_sources[dish].get(m) != 'NO DATA':
                continue
            key = round(cur, 2)
            grp_seen = twin_delta_grp.get((dish_group_of.get(dish, ''), m, key)) or set()
            all_seen = twin_delta.get((m, key)) or set()
            if len(grp_seen) == 1:
                deltas_seen, scoped = grp_seen, True
            elif len(all_seen) == 1:
                deltas_seen, scoped = all_seen, False
            else:
                if all_seen:
                    twin_conflicts.append(
                        f"TWIN CONFLICT: {dish} / {m} @ ${cur:.2f} — candidate twins "
                        f"disagree {sorted(all_seen)}"
                        + (f" (same dish group: {sorted(grp_seen)})" if grp_seen else '')
                        + "; left at the current price. Confirm on the callback.")
                continue
            d = next(iter(deltas_seen))
            if abs(d) > 0.001:
                final_prices[dish][m] = round(cur + d, 2)
                price_sources[dish][m] = 'PRICE TWIN (group)' if scoped else 'PRICE TWIN'
                twins_applied += 1
                twins_group_scoped += 1 if scoped else 0
    if twins_applied:
        flags.append(f"PRICE-TWIN: filled {twins_applied} no-data modifier prices "
                     f"by matching (modifier_item, base price) to a paper-data twin "
                     f"({twins_group_scoped} resolved within the same dish group).")
    flags.extend(twin_conflicts)

    return dict(side_section=side_section, side_columns=side_columns, side_items=side_items,
                deltas=deltas, plain_prices=plain_prices, dish_map=dish_map, sys_plain=sys_plain,
                mr_header=mr_header, mr_dish_rows=mr_dish_rows, col_to_mods=col_to_mods,
                final_prices=final_prices, price_sources=price_sources, flags=flags,
                side_coverage=side_coverage,
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
    ws.cell(row=r, column=1, value='Blue = increase | Pink = decrease | Green = confirmed no change | Orange = specialist edit').font = Font(italic=True, color='666666')
    r += 2
    header_row = r
    # Single friendly header row for the five data columns
    for col, label in OUT_HEADERS.items():
        c = ws.cell(row=header_row, column=col, value=label)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    # Validation-block headers
    for col, label in ((VAL_PLAIN_COL, 'by itself'), (VAL_OCRTOT_COL, 'Paper menu Total'),
                       (VAL_EXPECT_COL,
                        'Expected (raw total)' if A.get('model') == MODEL_ABSOLUTE
                        else 'Expected (by itself+Δ)'),
                       (VAL_MATCH_COL, 'Match?')):
        c = ws.cell(row=header_row, column=col, value=label)
        c.fill = DIVIDER_FILL
        c.font = HEADER_FONT
    r += 1
    data_start = r
    ordered = []
    for dish, rows in sorted(A['mr_dish_rows'].items()):
        for _, row in rows:
            ordered.append((dish, row))
    for dish, row in ordered:
        enabled = str(row[ENABLED_COL]).strip() if len(row) > ENABLED_COL and row[ENABLED_COL] is not None else ''
        dgroup = str(row[DGROUP_COL]).strip() if len(row) > DGROUP_COL and row[DGROUP_COL] else ''
        dish_item = str(row[DISH_COL]).strip() if len(row) > DISH_COL and row[DISH_COL] else ''
        variant = str(row[VARIANT_COL]).strip() if len(row) > VARIANT_COL and row[VARIANT_COL] else ''
        mod = str(row[MOD_COL]).strip() if len(row) > MOD_COL and row[MOD_COL] else ''
        old_price = _parse(row[PRICE_COL]) if len(row) > PRICE_COL else None
        # Compact data columns (dish_item is the row's own value; empty for
        # group-level rules, which carry only a dish_group).
        ws.cell(row=r, column=OUT_ENABLED_COL, value=enabled)
        ws.cell(row=r, column=OUT_DGROUP_COL, value=dgroup)
        ws.cell(row=r, column=OUT_DISH_COL, value=dish_item)
        ws.cell(row=r, column=OUT_VARIANT_COL, value=variant)
        ws.cell(row=r, column=OUT_MOD_COL, value=mod)
        new_price = A['final_prices'].get(dish, {}).get(mod)
        source = A['price_sources'].get(dish, {}).get(mod, '')
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
        plain = A['sys_plain'].get(dish)
        # Paper menu Total = plain + generated delta. Populated for every row that
        # carries a resolved delta, whether the combo price was actually PRINTED
        # (PAPER MENU) or DERIVED — PROPAGATED (delta borrowed from a sibling
        # modifier in the same dish) or PRICE TWIN (delta borrowed from a same-
        # (modifier, base-price) row on another dish).
        has_ref = (source.startswith(('PAPER MENU', 'PROPAGATED', 'PRICE TWIN'))
                   and plain is not None and new_price is not None)
        pL = get_column_letter(OUT_PRICE_COL)      # 'B' — editable price
        aa = get_column_letter(VAL_PLAIN_COL)
        ab = get_column_letter(VAL_OCRTOT_COL)
        ac = get_column_letter(VAL_EXPECT_COL)
        if plain is not None:
            ws.cell(row=r, column=VAL_PLAIN_COL, value=round(plain, 2)).number_format = '0.00'
        if has_ref:
            ws.cell(row=r, column=VAL_OCRTOT_COL,
                    value=round(plain + new_price, 2)).number_format = '0.00'
        # Expected = by-itself + live-edited delta in Price (strips any $ so it
        # re-reads edits). Written on EVERY data row unconditionally — the formula
        # is always present regardless of whether a Paper menu Total exists;
        # IFERROR yields "" only when by-itself/Price aren't both numeric.
        exp = ws.cell(row=r, column=VAL_EXPECT_COL,
                      value=(f'=IFERROR({aa}{r}+VALUE(SUBSTITUTE({pL}{r}&"","$","")),"")'))
        exp.number_format = '0.00'
        # Match? — N/A when no OCR reference, else TRUE/FALSE within 0.5c tolerance
        ws.cell(row=r, column=VAL_MATCH_COL,
                value=(f'=IF({ab}{r}="","N/A",IF(ABS({ac}{r}-{ab}{r})<=0.005,"TRUE","FALSE"))'))
        # Baseline (hidden) — drives orange-on-specialist-edit.
        baseline = new_price if (new_price is not None and old_price is not None
                                 and abs(new_price - old_price) > 0.001) else old_price
        if baseline is not None:
            ws.cell(row=r, column=VAL_BASELINE_COL, value=round(baseline, 2)).number_format = '0.00'
        r += 1
    data_end = r - 1
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
        orange_formula = (f'=AND(${bL}{data_start}<>"",'
                          f'VALUE(SUBSTITUTE(${pL}{data_start}&"","$",""))<>${bL}{data_start})')
        ws.conditional_formatting.add(prng, FormulaRule(formula=[orange_formula], fill=ORANGE_FILL))
        ws.column_dimensions[bL].hidden = True
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
    r = ws.max_row + 2
    # Section 2: Price difference
    ws.cell(row=r, column=1, value=FS_SENTINEL).font = Font(bold=True, size=12)
    r += 1
    headers = ['Code', 'Name', 'Plain'] + [f'Δ {c}' for c in A['side_columns']]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
    r += 1
    for item in A['side_items']:
        ocr_name = item['name']
        ws.cell(row=r, column=1, value=item.get('code', ''))
        ws.cell(row=r, column=2, value=A['dish_map'].get(ocr_name, f'⚠ {ocr_name}'))
        ws.cell(row=r, column=3, value=_fmt(A['plain_prices'].get(ocr_name)))
        for si, col in enumerate(A['side_columns']):
            d = A['deltas'].get(ocr_name, {}).get(col)
            if d is not None:
                cell = ws.cell(row=r, column=4 + si, value=('$+%.2f' % d) if d >= 0 else _fmt(d))
                cell.fill = BLUE_FILL if d > 0 else (PINK_FILL if d < 0 else GREEN_FILL)
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
    # List EVERY OCR dish from the multi-column section — not just matched ones.
    # Matched dishes show the resolved system name; a dish captured in OCR that
    # has no modifier_rule counterpart (i.e. it was never in the wsheet
    # modifier_rule tab) is flagged '⚠ NO MATCH' so the specialist sees it
    # instead of it silently vanishing from the table.
    ws.cell(row=r, column=2, value='DISH NAME MAPPING (OCR → System)').font = Font(bold=True, size=11)
    r += 1
    _seen_dn = set()
    for item in A['side_items']:
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
    for dish in sorted(A['final_prices'].keys()):
        ws.cell(row=r, column=1, value=dish).font = Font(bold=True)
        r += 1
        for ci, h in enumerate(['Old Price', 'Modifier Item', 'New Price', 'Source', 'Changed?'], 1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font = HEADER_FONT
            c.fill = HEADER_FILL
        r += 1
        mod_old = {}
        for _, row in A['mr_dish_rows'][dish]:
            mod = str(row[MOD_COL]).strip() if len(row) > MOD_COL and row[MOD_COL] else ''
            price = _parse(row[PRICE_COL]) if len(row) > PRICE_COL else None
            if mod and price is not None:
                mod_old[mod] = price
        groups = defaultdict(list)
        for mod, price in mod_old.items():
            groups[price].append(mod)
        for old_price in sorted(groups.keys()):
            for mod in groups[old_price]:
                new_p = A['final_prices'][dish].get(mod, old_price)
                src = A['price_sources'][dish].get(mod, '')
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
        ws.cell(row=r, column=1, value='⚠ CONFLICTS (auto-resolved)').font = Font(bold=True, color='FF0000')
        r += 1
        for fl in A['flags']:
            ws.cell(row=r, column=1, value=fl)
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
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and money_re.match(v.strip()):
                    signed = v.strip()[1] == '+'  # '$+1.20' → keep + sign on display
                    num = _parse(v)
                    if num is not None:
                        cell.value = num
                        cell.number_format = '+0.00;-0.00' if signed else '0.00'


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
    A = _analyze(args.ocr_json, args.raw_gsheet, args.conflict)
    if A is not None:
        print(f"\nPricing model: {A['model']}  ({A['model_evidence']})")
        if A['model'] == MODEL_UNKNOWN:
            print("  WARNING: could not detect pricing model - review manually before proceeding.")
        _g11 = _check_g11(A)
        if _g11:
            print(f"\n!! DOUBLE-CHARGE ERROR (G11): {len(_g11)} priced modifier_rule row(s) "
                  f"with a non-zero dish_item under ABSOLUTE pricing:")
            for _v in _g11:
                print(f"   - {_v}")
            print("   Refusing to build the report. Fix dish_item to $0 (or reclassify).")
            sys.exit(1)
    if A is None:
        print("No multi-column pricing data found. Nothing to do.")
        sys.exit(0)

    print(f"\nSection: '{A['side_section']}'")
    print(f"Side columns: {A['side_columns']}")
    print(f"Items with side pricing: {len(A['side_items'])}")
    print(f"\nDish matches ({len(A['dish_map'])}/{len(A['deltas'])}):")
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
        # Modifier Rule (own green tab)
        if MRV_TAB in wb.sheetnames:
            del wb[MRV_TAB]
        ws_mrv = wb.create_sheet(MRV_TAB)
        counts = _render_mrv_tab(ws_mrv, A, img_path=img_path)
        # Remove any legacy separate Delta Analysis / Propagation Detail tabs;
        # their content now lives inside the Fried Specialty tab.
        for t in (DELTA_TAB, PROP_TAB):
            if t in wb.sheetnames:
                del wb[t]
        # Fold delta + mappings + propagation into Fried Specialty (below Action-3 section)
        if 'Fried Specialty' in wb.sheetnames:
            _append_fried_specialty(wb['Fried Specialty'], A)
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
    print(f"  Dishes matched: {len(A['dish_map'])} / {len(A['deltas'])}")
    print(f"  Modifier_rule rows found: {total_mr}")
    print(f"  Cells updated (would change): {counts['updated']}")
    print(f"  Cells confirmed (no change): {counts['confirmed']}")
    print(f"  Cells no data (kept as-is): {counts['no_data']}")
    if A['flags']:
        print(f"\n  ⚠ Conflicts auto-resolved: {len(A['flags'])}")
        for fl in A['flags']:
            print(f"    {fl}")
    _print_coverage_warning(A.get('side_coverage'))
    print(f"  NOTE: Raw gsheet NOT modified — review only.")
    print(f"{'='*60}")
    print("Done.")


# ══════════════════════════════════════════════════════════════════════
#  apply
# ══════════════════════════════════════════════════════════════════════

def _read_mrv(report_path):
    """Read approved (dish, variant, modifier) → out_unit_price from the compact
    Modifier Rule tab. Columns are located by their friendly HEADER LABEL
    (Dish/Variant/Modifier/Price) — NOT by fixed position — so the tab may be
    reordered without breaking Part 2. Backward-compatible: if the legacy 3-row
    'id' header is present instead, fall back to the old fixed offsets."""
    wb = load_workbook(report_path, read_only=True)
    if MRV_TAB not in wb.sheetnames:
        wb.close()
        return None
    ws = wb[MRV_TAB]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    def _cellstr(v):
        return str(v).strip() if v is not None else ''

    # Preferred: find the wsheet-name header row. Columns are matched by SUBSTRING
    # (case-insensitive) so labels like 'dish_item', 'modifier_item',
    # 'out_unit_price', 'variant' all resolve. 'modifier_item' is checked before
    # 'variant'/'price' so a stray substring can't mis-bind. The validation-block
    # labels ('by itself', 'Paper menu Total', 'Match?', ...) contain none of
    # these tokens, so they are never mistaken for a data column.
    def _find(labels, *tokens):
        for tok in tokens:
            for lbl, ci in labels.items():
                if tok in lbl:
                    return ci
        return None

    want = {'dish': None, 'variant': None, 'modifier': None, 'price': None}
    hdr_idx = None
    for i, r in enumerate(rows):
        if not r:
            continue
        labels = {_cellstr(c).lower(): ci for ci, c in enumerate(r) if _cellstr(c)}
        di = _find(labels, 'dish_item', 'dish')
        mi = _find(labels, 'modifier_item', 'modifier')
        pi = _find(labels, 'out_unit_price', 'price')
        if di is not None and mi is not None and pi is not None:
            want['dish'] = di
            want['modifier'] = mi
            want['price'] = pi
            want['variant'] = _find(labels, 'variant')
            hdr_idx = i
            break

    approved = {}
    if hdr_idx is not None:
        di, vi, mi, pi = want['dish'], want['variant'], want['modifier'], want['price']
        for r in rows[hdr_idx + 1:]:
            if not r or all(c in (None, '') for c in r):
                continue
            dish = _cellstr(r[di]) if di is not None and len(r) > di else ''
            variant = _cellstr(r[vi]) if vi is not None and len(r) > vi else ''
            mod = _cellstr(r[mi]) if mi is not None and len(r) > mi else ''
            price = _parse(r[pi]) if pi is not None and len(r) > pi else None
            if dish and mod and price is not None:
                approved[(dish, variant, mod)] = price
        return approved

    # Fallback: legacy 3-row modifier_rule header (row whose col A == 'id').
    hdr = None
    for i, r in enumerate(rows):
        if r and str(r[0]).strip() == 'id':
            hdr = i
            break
    if hdr is None:
        return {}
    for r in rows[hdr + 3:]:
        if not r or all(c in (None, '') for c in r):
            continue
        dish = _cellstr(r[DISH_COL]) if len(r) > DISH_COL else ''
        variant = _cellstr(r[VARIANT_COL]) if len(r) > VARIANT_COL else ''
        mod = _cellstr(r[MOD_COL]) if len(r) > MOD_COL else ''
        price = _parse(r[PRICE_COL]) if len(r) > PRICE_COL else None
        if dish and mod and price is not None:
            approved[(dish, variant, mod)] = price
    return approved


def cmd_apply(args):
    print(f"Price Comparison Report: {args.report}")
    print(f"Raw Gsheet: {args.raw_gsheet}")
    approved = _read_mrv(args.report)
    # Re-detect model + G11 from the fresh raw gsheet (defense in depth before writing).
    _mr_data, _dip = _load_raw(args.raw_gsheet)
    _model, _evidence = _detect_pricing_model(_mr_data, _dip)
    print(f"Pricing model: {_model} ({_evidence})")
    _g11 = _check_g11_core(_mr_data, _dip, _model)
    if _g11:
        for _v in _g11:
            print(f"   - {_v}")
        print(f"!! DOUBLE-CHARGE ERROR (G11): {len(_g11)} violation(s) - refusing to apply.")
        sys.exit(1)
    if approved is None:
        print("No 'Modifier Rule' tab in the report. Nothing to apply.")
        sys.exit(0)
    print(f"Approved out_unit_price triples: {len(approved)}")

    wb = load_workbook(args.raw_gsheet)  # writable, preserves all other tabs
    ws = wb['modifier_rule']
    # rows 1-3 are the 3-row header; data starts at row 4 (1-indexed)
    matched = set()
    written = 0
    confirmed = 0
    for row in ws.iter_rows(min_row=4):
        dish = str(row[DISH_COL].value).strip() if len(row) > DISH_COL and row[DISH_COL].value else ''
        variant = str(row[VARIANT_COL].value).strip() if len(row) > VARIANT_COL and row[VARIANT_COL].value else ''
        mod = str(row[MOD_COL].value).strip() if len(row) > MOD_COL and row[MOD_COL].value else ''
        key = (dish, variant, mod)
        if key not in approved:
            continue
        matched.add(key)
        new_price = approved[key]
        price_cell = row[PRICE_COL]
        old_price = _parse(price_cell.value)
        # G1: never write a blank cell (protects disabled group-gate rows)
        if old_price is None:
            continue
        # G4: only write when the value actually changed
        if abs(new_price - old_price) <= 0.001:
            confirmed += 1
            continue
        price_cell.value = _fmt(new_price)  # G6: $X.XX
        written += 1

    # G8: drift — approved triples never matched to a live row
    drift = [k for k in approved if k not in matched]
    for k in drift:
        print(f"!! DRIFT WARNING: approved (dish='{k[0]}', variant='{k[1]}', modifier='{k[2]}') "
              f"not found in live modifier_rule — NOT applied (renamed/removed since Part 1).")

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
    _existing['modifier_rule'] = {
        'pricing_model': _model,
        'approved_triples': len(approved),
        'matched_live_rows': len(matched),
        'out_unit_price_cells_written': written,
        'confirmed_no_change': confirmed,
        'drift': {
            'total': len(drift),
            'triples': [{'dish': k[0], 'variant': k[1], 'modifier': k[2]} for k in drift],
        },
        'output_xlsx': out,
    }
    with open(_summary_path, 'w', encoding='utf-8') as _f:
        _json.dump(_existing, _f, indent=2)

    print(f"\n{'='*60}")
    print(f"  Approved triples: {len(approved)}")
    print(f"  Matched live rows: {len(matched)}")
    print(f"  out_unit_price cells written (changed): {written}")
    print(f"  Confirmed (no change, skipped): {confirmed}")
    print(f"  Drift warnings (NOT applied): {len(drift)}")
    print(f"  Output: {out}")
    print(f"  Summary JSON: {_summary_path}  (source of truth for part2_report.py)")
    print(f"  Only the modifier_rule tab was touched (G3); all other tabs byte-identical.")
    print(f"{'='*60}")
    if drift:
        print("Resolve every DRIFT WARNING before staging the draft.")
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


def cmd_preflight(args):
    """Mandatory gate: classify the pricing model and refuse to continue on a
    DOUBLE_CHARGE / BROKEN verdict. Exit 0 = safe to proceed."""
    mr_data, dish_item_prices = _load_raw(args.raw_gsheet)
    model, evidence = _detect_pricing_model(mr_data, dish_item_prices)
    print(f"Pricing model: {model}  ({evidence})")
    if model in (MODEL_DOUBLE, MODEL_BROKEN):
        print(f"!! BLOCKING VERDICT: {model} - do NOT run compare/apply or stage a draft.")
        sys.exit(1)
    if model == MODEL_UNKNOWN:
        print("WARNING: UNKNOWN model - no By-Itself anchor found. Review manually; not guessing.")
        sys.exit(2)
    g11 = _check_g11_core(mr_data, dish_item_prices, model)
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
    pc.add_argument('--conflict', choices=['mode', 'lowest', 'highest'], default='mode',
                    help='How to auto-resolve a price-group conflict (default: mode, tie-break lowest)')
    pc.add_argument('--pricing-image', default=None,
                    help='PNG of the paper-menu multi-column pricing block to embed on '
                         'the Modifier Rule tab (default: crops/special_full.png if present; '
                         'produce one with crop_section.py). Use "" or "none" to disable.')
    pc.set_defaults(func=cmd_compare)

    pp = sub.add_parser('preflight',
                        help='Detect pricing model + run G11 double-charge check on a raw gsheet')
    pp.add_argument('raw_gsheet', help='Raw Gsheet Excel (from a fresh wsheet fetch)')
    pp.set_defaults(func=cmd_preflight)

    pa = sub.add_parser('apply', help='Part 2: write approved out_unit_price back into modifier_rule')
    pa.add_argument('report', help='Price Comparison Report with the edited Modifier Rule tab')
    pa.add_argument('raw_gsheet', help='Fresh Raw Gsheet Excel to patch (a copy is written)')
    pa.add_argument('--output', '-o', help='Output path (default: final_raw.xlsx)')
    pa.set_defaults(func=cmd_apply)

    args = parser.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()