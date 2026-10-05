#!/usr/bin/env python3
"""
price_inference.py — Cohort-based price inference for unprinted prices
=======================================================================

Standalone script. Dependencies: openpyxl only.
Runs AFTER action3 (and BEFORE modifier_rule_pricing.py compare) and writes
into the SAME Price Comparison Report workbook — no separate file.

Usage:
    python price_inference.py <menu_preview_2a.xlsx> <ocr_output_2b.xlsx> \
        --report <price_comparison_report.xlsx> [-o <copy.xlsx>] \
        [--mode suggest|apply] [--min-donors N] [--tier2-min-donors N] \
        [--max-delta-abs D] [--max-delta-pct P] [--include-low-priority] \
        [--tray-min-agreement F] [--tray-min-rows N] [--tray-max-offset D] \
        [--no-tray] [--json <path>]

THE PROBLEM IT SOLVES
---------------------
Paper menus are incomplete in two ways that occur on almost every run:

  Scenario 1 — VARIANT GAP (V-GAP). The dish IS on the paper menu, but not
    every variant the restaurant sells is printed (system has Small + Large,
    paper prints only one). Today action3 keeps the system price in the
    unprinted variant and does not flag it: a STALE price ships silently.

  Scenario 2 — DISH GAP (D-GAP). The dish has no paper match at all, but is
    still sold (it is simply not on the paper menu) and, again, its old
    price ships unchanged.

In both cases the price is usually recoverable by pattern, because a dish
group is priced as a family: dishes that shared an old price almost always
share the new one.

    Chicken with Broccoli     $8.95 → $9.95   (printed — donor)
    Chicken with Cashew Nut   $8.95 → ?       (not printed — recipient)
                              ^^^^^ same old price, same dish group
    ⇒ infer $9.95

INFERENCE TIERS (a recipient takes the FIRST tier that resolves)
---------------------------------------------------------------
  T0  SAME-PRICE RULES (maintainer, 2026-09-30) — tried first:
      * General Tso / Orange / Sesame <Chicken|Beef|Shrimp> always cost the
        same: an unprinted one takes its printed siblings' price (same
        protein, same variant), whatever its old price was.
      * a Healthy-group dish ("Steam Chicken with Broccoli") that is NOT on the
        paper menu takes its regular equivalent's new price ("Chicken with
        Broccoli") when both OLD prices were the same. A Healthy dish the
        paper prints (even one size) is left to the tiers below.
      Printed siblings that disagree = FAMILY CONFLICT: no inference.
      (Chow Mein and Chop Suey are one family for T1 below.)

  T1  EXACT OLD-PRICE COHORT — within the same (dish group, variant) — Chow
      Mein and Chop Suey count as ONE dish group here — collect
      every donor (a dish that DID get a paper price) whose OLD price equals
      the recipient's old price. If they all agree on the new price, that is
      the inference. Confidence HIGH at >= 2 donors, MEDIUM at exactly 1.
      Donors disagreeing = COHORT CONFLICT: no inference, reported for the
      specialist. Direct conflicting evidence is never overridden by T2.

  T2  UNIFORM GROUP DELTA — only when the recipient's old price has NO donor
      cohort at all. If every donor in the (dish group, variant) moved by the
      SAME delta and there are at least --tier2-min-donors of them, apply that
      delta. Confidence MEDIUM.

Everything else is reported as SKIPPED with the reason, never guessed. A
cell whose system price is $0.00 is never priced by inference at all (a free
item, or a dish whose price lives on its modifiers): SKIPPED, any tier.

  T-TRAY  PARTY TRAY FORMULA (v4.3) — a separate tier for TRAY_VARIANTS
      columns, which never take T1/T2. Restaurants price a tray as a formula
      off a base column on the SAME row (e.g. Tray = Large × 5 − $2.00). The
      formula (base column, multiplier 1.0–10.0 in 0.5 steps, offset within
      ±--tray-max-offset) is fitted on the system's OLD prices across every
      row that has a tray AND a base price, restaurant-wide. It is trusted
      only when ONE formula fits ≥ --tray-min-agreement (default 70%) of
      those rows, ≥ --tray-min-rows rows, with ≥ 2 distinct base prices, and
      no other formula ties it (a tie = AMBIGUOUS, not guessed — G14).
      An unprinted tray is then derived from its row's NEW printed base,
      only if that row itself follows the formula (outliers are asked, not
      derived). A base that was itself inferred (T1/T2) yields a SUGGESTED
      tray only — inference is never chained into an applied write.
      Printed trays are not touched here: action3 already compared them.

DELIBERATE NON-FEATURE — no cross-variant delta transfer
--------------------------------------------------------
A dish's own printed Small delta is NOT transferred to its unprinted Large.
Restaurants routinely raise Pt by $0.50 and Qt by $1.00, so the per-variant
delta is not a shared quantity. Inference only ever moves ACROSS DISHES,
WITHIN one variant. Do not "improve" this into a row-wise fill.

The single, deliberate exception is T-TRAY: a Party Tray is not an
independently-priced portion but a catering multiple of a base portion, and
T-TRAY copies no delta — it fits a multiplicative formula and only applies it
when the restaurant's own data shows ≥70% of trays follow it. Do not extend
this exception to any other variant pair.

SAFETY POSTURE
--------------
  --mode suggest (DEFAULT)  Nothing is written to Dish Item. Every inference
      lands on the Price Inference tab as a suggestion.
      Zero production risk; the specialist types anything they accept.
  --mode apply              Eligible inferences are ALSO written into the
      Dish Item price cells with a PURPLE fill and the hidden _ref_ mirror
      refreshed (so the apricot "specialist edited this" rule does not
      false-fire). Part 2's action5 then carries them to production like any
      other report value — which is exactly why every one of them is also a
      client-callback confirmation item.

An inference is downgraded from APPLIED to SUGGESTED (never silently applied)
when any of these hold:
  - the inferred delta is a DECREASE (mirrors the P2 "decreases are rare" rule)
  - |delta| > --max-delta-abs, or |delta| / old > --max-delta-pct
  - it would break a Small < Large ordering that held in the system
  - the dish group is operational (Drink / Sauces / Side Order) and
    --include-low-priority was not passed
  T-TRAY uses the same rules, except the absolute cap scales with the
  multiplier (--max-delta-abs × k), since a $1 base rise is a $5 tray rise.

Guardrails honoured: G1 (a blank system cell is never written — a recipient
must already have a price), G3 (Party Tray: T-TRAY only, never T1/T2), G5 (columns
resolved by header text), G11 (a $0.00 dish whose price lives on modifiers —
the Menu Preview's 'absolute_pricing' tab — is never a recipient or a donor),
G13/G14 (see part1.md).
"""

import argparse
import json
import os
import re
import sys
from collections import defaultdict
from typing import Optional

try:
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl required. Install: pip install openpyxl")
    sys.exit(1)


# ── Shared vocabulary with action3 ──────────────────────────────────────
NAME_COLUMNS = {'Dish Group', 'Dish Name:Primary', 'Dish Name:Secondary'}
# Catering-tray columns. Excluded from the T1/T2 cohort tiers (a tray is not
# priced as a family with other trays) and handled instead by T-TRAY, which
# derives the tray from a base column on the SAME row via the restaurant's
# own tray formula. Keep in sync with action2 TRAY_VARIANTS and
# ocr_sanity_check.py TRAY_LABELS.
TRAY_VARIANTS = {'Party Tray'}
NEVER_INFER_VARIANTS = set(TRAY_VARIANTS)      # G3: no T1/T2 on trays
TRAY_MULTIPLIERS = [x / 2 for x in range(2, 21)]   # 1.0× … 10.0× in 0.5 steps
TRAY_TOL = 0.011                               # cent-rounding tolerance
P_TRAY_SENTINEL = 'P-TRAY — PARTY TRAY (CONFIRM ON CALLBACK)'
LOW_PRIORITY_GROUPS = {'drink', 'drinks', 'sauce', 'sauces',
                       'side order', 'side orders'}

PURPLE_FILL = PatternFill('solid', fgColor='E4DFEC')   # inferred (new)
GREEN_FILL = PatternFill('solid', fgColor='C6EFCE')
GREY_FILL = PatternFill('solid', fgColor='D9D9D9')
RED_FILL = PatternFill('solid', fgColor='FFC7CE')
HEADER_FILL = PatternFill('solid', fgColor='D9E1F2')
HEADER_FONT = Font(bold=True, size=11)

INF_TAB = 'Price Inference'
S_SENTINEL = 'AUTO-INFERENCE (price_inference.py)'
CANON_ORDER = ['Summary', 'Dish Item', 'Questionnaire', INF_TAB,
               'Modifier Summary', 'Modifier Rule', 'Fried Specialty', 'Dish Code']


def _parse(val) -> Optional[float]:
    if val is None:
        return None
    s = str(val).strip()
    if not s or s in ('None', 'NULL'):
        return None
    s = s.replace('$', '').replace(',', '')
    try:
        return float(s)
    except ValueError:
        return None


def _fmt(val) -> str:
    if val is None:
        return ''
    try:
        return f'${float(val):.2f}'
    except (TypeError, ValueError):
        return ''


def _fmt_signed(val) -> str:
    """'$+1.00' / '$-0.50' — the '$' leads so modifier_rule_pricing.py's
    _strip_dollar_display() recognizes it and applies a signed number format."""
    if val is None:
        return ''
    return f'${"+" if val >= 0 else "-"}{abs(val):.2f}'


def _variant_columns(headers: list) -> dict:
    """Auto-discover variant columns by header text (G5). Mirrors action3's
    _detect_extra_variants: every non-name, non-_ref_ column is a variant."""
    out = {}
    for i, h in enumerate(headers):
        h = str(h).strip() if h else ''
        if not h or h in NAME_COLUMNS or h.startswith('_ref_') or h in out:
            continue
        out[h] = i
    return out


def _is_low_priority(group: str) -> bool:
    return (group or '').strip().lower() in LOW_PRIORITY_GROUPS


# ── Maintainer same-price rules (2026-09-30) ────────────────────────────
# T0 tiers are tried BEFORE T1. They are the maintainer's explicit business
# rules — the only sanctioned widening besides T-TRAY (G14):
#   * General Tso / Orange / Sesame <protein> (Chicken, Beef, Shrimp) always
#     cost the same: an unprinted one takes its printed siblings' price.
#   * A Healthy dish ("Steam Chicken with Broccoli") NOT on the paper menu
#     takes its regular equivalent's new price ("Chicken with Broccoli") when
#     both OLD prices were the same.
#   * Chow Mein and Chop Suey are one family for T1: same old price → same
#     new price across the two dish groups.
T0_TRIO_TIER = 'T0 same-price family (General Tso / Orange / Sesame)'
T0_HEALTHY_TIER = 'T0 healthy = regular dish (same old price)'
MERGED_COHORT = 'Chow Mein + Chop Suey'
_TRIO_RE = re.compile(r'^(general tso|orange|sesame) (chicken|beef|shrimp)$')


def _dish_norm(name: str) -> str:
    """Lower-case, punctuation-free dish name for the same-price rules:
    "General Tso's Chicken" / "Gen. Tsos Chicken" -> "general tso chicken",
    "Chicken w. Broccoli" -> "chicken with broccoli"."""
    s = str(name or '').lower()
    s = re.sub(r"(?<=\w)['’]s\b", '', s)
    s = re.sub(r'\bgen\.?\s+tso', 'general tso', s)
    s = re.sub(r'\btsos\b', 'tso', s)
    s = re.sub(r'\([^)]*\)', ' ', s)
    s = s.replace('&', ' and ').replace('w/', ' with ').replace('w.', ' with ')
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    s = re.sub(r'\bflavou?red\b', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def _healthy_base(name: str) -> str:
    """A Healthy dish name without its diet words — "Steam Shrimp with
    Broccoli" and the paper's "Diet Shrimp w/ Broccoli (Steamed)" both give
    "shrimp with broccoli"."""
    words = [w for w in _dish_norm(name).split() if w not in ('steam', 'steamed', 'diet', 'healthy')]
    return ' '.join(words)


def _qualifier(name: str) -> str:
    """The bracketed qualifier of a dish name — 'white meat' for "General Tso
    Chicken (White Meat)". Same-price families are only formed between dishes
    with the SAME qualifier: a White Meat dish is a different, dearer product."""
    q = ' '.join(re.findall(r'\(([^)]*)\)', str(name or '').lower()))
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9 ]', ' ', q)).strip()


def _trio_protein(dish: str):
    """'chicken' for General Tso's / Orange / Sesame Chicken, else None."""
    m = _TRIO_RE.match(_dish_norm(dish))
    return m.group(2) if m else None


def _is_healthy_group(group: str) -> bool:
    return 'healthy' in (group or '').lower()


def _cohort_group(group: str) -> str:
    """T1 cohort key: Chow Mein and Chop Suey share one cohort."""
    g = (group or '').lower()
    return MERGED_COHORT if ('chow mein' in g or 'chop suey' in g) else group


# ══════════════════════════════════════════════════════════════════════
#  Analysis — pure, no workbook writes
# ══════════════════════════════════════════════════════════════════════

def analyze(rows_2a, rows_2b, opts):
    """Return (records, stats). One record per recipient (row, variant) cell
    that had a system price but no paper price."""
    headers = [str(c).strip() if c else '' for c in rows_2a[0]] if rows_2a else []
    var_cols = _variant_columns(headers)
    name_col = headers.index('Dish Name:Primary') if 'Dish Name:Primary' in headers else 1
    group_col = headers.index('Dish Group') if 'Dish Group' in headers else 0
    absolute_cells = opts.get('absolute_cells') or set()

    # ── Pass 1: classify every cell as donor / recipient ────────────────
    donors = defaultdict(lambda: defaultdict(list))   # (g,v) → old → [(dish,new)]
    group_deltas = defaultdict(dict)                  # (g,v) → {dish: delta}
    paper_any = defaultdict(bool)                     # ri → any paper price, trays included
    recipients = []                                   # candidate cells
    row_meta = {}                                     # ri → (group, dish, matched)
    old_by_row = defaultdict(dict)                    # ri → {variant: old}
    new_by_row = defaultdict(dict)                    # ri → {variant: printed new}

    for ri in range(1, len(rows_2a)):
        r2a = list(rows_2a[ri]) if ri < len(rows_2a) else []
        r2b = list(rows_2b[ri]) if ri < len(rows_2b) else []
        dish = str(r2a[name_col]).strip() if len(r2a) > name_col and r2a[name_col] else ''
        group = str(r2a[group_col]).strip() if len(r2a) > group_col and r2a[group_col] else ''
        if not dish:
            continue                                  # separator row

        matched = False
        pending = []
        for v, ci in var_cols.items():
            if ci < len(r2b) and _parse(r2b[ci]) is not None:
                paper_any[ri] = True                  # printed in some size (a tray too)
        if dish.lower() in (opts.get('unplaced_dishes') or ()):
            paper_any[ri] = True                      # printed in a size the dish does not have (P-SIZE)
        conflicted = {v for v in var_cols if (dish, v) in (opts.get('paper_conflicts') or ())}
        if conflicted:
            paper_any[ri] = True                      # printed — twice, differently (P-PRICE)
            matched = True
        for v, ci in var_cols.items():
            if v in NEVER_INFER_VARIANTS:             # G3
                continue
            if (dish, v) in absolute_cells:           # G11: $0, priced on modifiers
                continue
            if v in conflicted:                       # asked (P-PRICE): no guess, no donor
                old = _parse(r2a[ci]) if ci < len(r2a) else None
                if old is not None:
                    old_by_row[ri][v] = old
                continue
            old = _parse(r2a[ci]) if ci < len(r2a) else None
            new = _parse(r2b[ci]) if ci < len(r2b) else None
            if old is not None:
                old_by_row[ri][v] = old
            if new is not None:
                matched = True
                new_by_row[ri][v] = round(new, 2)
                if old is not None:
                    # one live dish listed under two groups is ONE donor
                    lst = donors[(_cohort_group(group), v)][round(old, 2)]
                    if all(d != dish for d, _n, _g in lst):
                        lst.append((dish, round(new, 2), group))
                    group_deltas[(group, v)][dish] = round(new - old, 2)
            elif old is not None:
                # G1: only a cell that already carries a price is inferable.
                pending.append((v, ci, old))

        row_meta[ri] = (group, dish, matched)
        for v, ci, old in pending:
            recipients.append({'ri': ri, 'variant': v, 'ci': ci, 'old': old,
                               'group': group, 'dish': dish})

    # ── Pass 1b: families for the T0 same-price rules ───────────────────
    trio_rows = defaultdict(list)                     # (protein, qualifier) → [ri]
    regular_by_name = defaultdict(list)               # (normalized name, qualifier) → [ri] (non-healthy)
    for ri, (group, dish, _m) in row_meta.items():
        protein = _trio_protein(dish)
        if protein:
            trio_rows[(protein, _qualifier(dish))].append(ri)
        if not _is_healthy_group(group):
            regular_by_name[(_dish_norm(dish), _qualifier(dish))].append(ri)

    def _t0_peers(ri, group, dish):
        """(tier, peer rows, needs equal old price) for a recipient, or None.
        Peers never include another row of the same dish (a dish listed under
        two groups is one dish, not its own sibling)."""
        protein = _trio_protein(dish)
        if protein:
            return T0_TRIO_TIER, [rj for rj in trio_rows[(protein, _qualifier(dish))]
                                  if row_meta[rj][1] != dish], False
        if _is_healthy_group(group):
            base = re.sub(r'^steam(ed)?\s+', '', _dish_norm(dish))   # (same as _healthy_base)
            peers = [rj for rj in regular_by_name.get((base, _qualifier(dish)), [])
                     if row_meta[rj][1] != dish]
            if peers:
                return T0_HEALTHY_TIER, peers, True
        return None

    def _t0_miss(t0, v, old):
        """Why a same-price rule matched the dish but gave it no price."""
        _tier, peers, equal_old = t0
        printed = [(row_meta[rj][1], old_by_row[rj].get(v)) for rj in peers if v in new_by_row[rj]]
        if not printed:
            return ('same-price rule: its ' + ('regular dish' if equal_old else
                    'General Tso / Orange / Sesame siblings') + f' not printed in {v}')
        d, o = printed[0]
        return (f'healthy rule not used: regular {d} was '
                f'{_fmt(o) if o is not None else "not sold"} in {v}, not {_fmt(old)} '
                f'(needs the same old price)')

    # ── Pass 2: resolve each recipient ─────────────────────────────────
    records = []
    for rec in recipients:
        ri, group, v, old = rec['ri'], rec['group'], rec['variant'], rec['old']
        matched = row_meta[ri][2]
        rec['scenario'] = 'V-GAP' if matched else 'D-GAP'
        rec['tier'] = rec['confidence'] = rec['reason'] = ''
        rec['inferred'] = None
        rec['donor_n'] = 0
        rec['evidence'] = ''
        rec['status'] = 'SKIPPED'
        if abs(old) < 0.005:
            # A $0.00 system price is never priced by inference: it is a free
            # item (straws, "Free Egg Roll") or a dish priced on its modifiers,
            # and a group moving +$0.50 says nothing about either.
            rec['tier'] = 'not inferred ($0.00)'
            rec['reason'] = ('system price is $0.00 (free item, or priced on its modifiers) — '
                             'never priced by inference; ask on the callback if it should '
                             'now cost something')
            rec['delta'] = None
            rec['blockers'] = []
            records.append(rec)
            continue

        t0 = _t0_peers(ri, group, rec['dish'])
        t0_note = ''
        p1_name = ((opts.get('p1_suggested') or {}).get(rec['dish'].lower())
                   or ((opts.get('p1_healthy') or {}).get(_healthy_base(rec['dish']))
                       if _is_healthy_group(group) else None))
        if t0 and t0[0] == T0_HEALTHY_TIER and (matched or paper_any[ri] or p1_name):
            # Maintainer (2026-09-30): only a Healthy dish that is NOT on the
            # paper menu takes its regular dish's price.
            t0_note = (f"healthy rule not used: {rec['dish']} is on the paper menu"
                       + (f" (P1 '{p1_name}' looks like it)" if p1_name and not (matched or paper_any[ri])
                          else '')
                       + " — only an unprinted Healthy dish takes its regular dish's price")
            t0 = None
        t0_donors = []
        if t0:
            tier, peers, equal_old = t0
            for rj in peers:
                n = new_by_row[rj].get(v)
                o = old_by_row[rj].get(v)
                if n is None or (equal_old and (o is None or abs(o - old) > 0.005)):
                    continue
                if all(d != row_meta[rj][1] for d, _o, _n in t0_donors):
                    t0_donors.append((row_meta[rj][1], o, n))
        cohort = donors.get((_cohort_group(group), v), {})
        bucket = cohort.get(round(old, 2), [])
        all_donors = sum(len(x) for x in cohort.values())

        if t0_donors:
            distinct = sorted({n for _, _, n in t0_donors})
            rec['donor_n'] = len(t0_donors)
            rec['evidence'] = '; '.join(
                f'{d} {_fmt(o)}→{_fmt(n)}' for d, o, n in t0_donors[:4]
            ) + (f' (+{len(t0_donors) - 4} more)' if len(t0_donors) > 4 else '')
            if len(distinct) == 1:
                rec['inferred'] = distinct[0]
                rec['tier'] = tier
                rec['confidence'] = 'HIGH'
            else:
                rec['tier'] = 'T0 conflict'
                rec['reason'] = ('SAME-PRICE FAMILY CONFLICT — printed prices disagree: '
                                 + ', '.join(_fmt(p) for p in distinct)
                                 + '. Not guessed; ask on callback.')
        elif bucket:
            distinct = sorted({new for _d, new, _g in bucket})
            rec['donor_n'] = len(bucket)
            rec['evidence'] = '; '.join(
                f'{d} {_fmt(old)}→{_fmt(n)}' for d, n, _g in bucket[:4]
            ) + (f' (+{len(bucket) - 4} more)' if len(bucket) > 4 else '')
            if len(distinct) == 1:
                rec['inferred'] = distinct[0]
                rec['tier'] = 'T1 exact old-price cohort' + (
                    f' ({MERGED_COHORT})' if _cohort_group(group) == MERGED_COHORT
                    and any(('chop suey' in g.lower()) != ('chop suey' in group.lower())
                            for _d, _n, g in bucket) else '')
                rec['confidence'] = 'HIGH' if len(bucket) >= 2 else 'MEDIUM'
            else:
                rec['tier'] = 'T1 conflict'
                rec['reason'] = ('COHORT CONFLICT — donors at ' + _fmt(old)
                                 + ' disagree: ' + ', '.join(_fmt(p) for p in distinct)
                                 + '. Not guessed; ask on callback.')
        elif all_donors:
            deltas = sorted(set(group_deltas[(group, v)].values()))
            n = len(group_deltas[(group, v)])
            if len(deltas) == 1 and n >= opts['tier2_min_donors']:
                rec['inferred'] = round(old + deltas[0], 2)
                rec['tier'] = 'T2 uniform group delta'
                rec['confidence'] = 'MEDIUM'
                rec['donor_n'] = n
                rec['evidence'] = (f'all {n} priced dishes in {group}/{v} moved '
                                   f'{_fmt_signed(deltas[0])}')
            else:
                rec['tier'] = 'T2 not available'
                rec['reason'] = (f'no donor at {_fmt(old)} in {_cohort_group(group)}/{v}; '
                                 + (f'group deltas not uniform ({n} donors, '
                                    f'{len(deltas)} distinct deltas)' if n else
                                    f'no printed dish in {group}/{v} for a group delta'))
        else:
            rec['reason'] = f'no priced donor at all in {group}/{v}'
        if t0 and not t0_donors and rec['inferred'] is None:
            rec['reason'] = f"{_t0_miss(t0, v, old)}; {rec['reason']}"
        elif t0_note and rec['inferred'] is None:
            rec['reason'] = f"{t0_note}; {rec['reason']}"

        # ── Eligibility screening ──────────────────────────────────────
        if rec['inferred'] is not None:
            old_p, new_p = rec['old'], rec['inferred']
            delta = round(new_p - old_p, 2)
            rec['delta'] = delta
            blockers = []
            if abs(delta) < 0.005:
                rec['status'] = 'NO CHANGE EXPECTED'
                rec['reason'] = ('cohort shows no change — current price looks '
                                 'correct, nothing to write')
                records.append(rec)
                continue
            if rec['donor_n'] < opts['min_donors']:
                blockers.append(f'only {rec["donor_n"]} donor(s), '
                                f'--min-donors {opts["min_donors"]}')
            if delta < 0:
                blockers.append('inferred DECREASE — verify like a P2')
            if abs(delta) > opts['max_delta_abs']:
                blockers.append(f'|delta| {_fmt(abs(delta))} > --max-delta-abs '
                                f'{_fmt(opts["max_delta_abs"])}')
            if old_p > 0 and abs(delta) / old_p * 100 > opts['max_delta_pct']:
                blockers.append(f'|delta| is {abs(delta) / old_p * 100:.0f}% of old '
                                f'price > --max-delta-pct {opts["max_delta_pct"]:.0f}%')
            if _is_low_priority(rec['group']) and not opts['include_low_priority']:
                blockers.append('operational group (Drink/Sauces/Side Order)')
            rec['blockers'] = blockers
            rec['status'] = 'ELIGIBLE' if not blockers else 'SUGGESTED'
            if blockers:
                rec['reason'] = 'suggest-only: ' + '; '.join(blockers)
        else:
            rec['delta'] = None
            rec['blockers'] = []
        records.append(rec)

    # ── Pass 2b: a dish listed under two dish groups is ONE live dish ──
    # Its rows must agree: one inference + rows with no evidence → every row
    # takes it; rows that infer different prices, or any row in conflict →
    # no row is inferred (a guess on one row and "conflict" on another must
    # never ship one of them).
    by_cell = defaultdict(list)
    for rec in records:
        by_cell[(rec['dish'], rec['variant'])].append(rec)
    for (dish, v), recs in by_cell.items():
        if len(recs) < 2 or any(r['tier'] == 'not inferred ($0.00)' for r in recs):
            continue
        vals = {r['inferred'] for r in recs if r['inferred'] is not None}
        clash = any(r['tier'] in ('T1 conflict', 'T0 conflict') for r in recs)
        if clash or len(vals) > 1:
            detail = '; '.join(f"{r['group']}: " + (_fmt(r['inferred']) if r['inferred'] is not None
                                                    else (r['tier'] or 'no inference')) for r in recs)
            for r in recs:
                r.update(inferred=None, status='SKIPPED', tier='conflict (same dish, several groups)',
                         confidence='', delta=None, blockers=[],
                         reason=f'SAME DISH IN SEVERAL GROUPS — its rows disagree ({detail}). '
                                f'Not guessed; ask on callback.')
        elif vals:
            src = next(r for r in recs if r['inferred'] is not None)
            for r in recs:
                if r['inferred'] is None:
                    for k in ('inferred', 'tier', 'confidence', 'evidence', 'donor_n', 'status', 'delta'):
                        r[k] = src[k]
                    r['blockers'] = list(src.get('blockers') or [])
                    r['reason'] = ((src['reason'] + '; ') if src['reason'] else '') + \
                        f"same dish as its {src['group']} row"

    # ── Pass 3: Small < Large ordering sanity, per row ─────────────────
    # Only downgrades; it never invents or changes a price. Uses what the
    # row will hold: its NEW printed prices and the eligible inferences.
    by_row = defaultdict(list)
    for rec in records:
        by_row[rec['ri']].append(rec)
    for ri, recs in by_row.items():
        final = {**old_by_row[ri], **new_by_row[ri]}
        for rec in recs:
            if rec['status'] == 'ELIGIBLE':
                final[rec['variant']] = rec['inferred']
        for lo, hi in (('Small', 'Large'),):
            o_lo, o_hi = old_by_row[ri].get(lo), old_by_row[ri].get(hi)
            n_lo, n_hi = final.get(lo), final.get(hi)
            if None in (o_lo, o_hi, n_lo, n_hi):
                continue
            if o_lo < o_hi and n_lo >= n_hi:
                for rec in recs:
                    if rec['status'] == 'ELIGIBLE' and rec['variant'] in (lo, hi):
                        rec['status'] = 'SUGGESTED'
                        rec['blockers'].append(f'would break {lo} < {hi} ordering')
                        rec['reason'] = 'suggest-only: ' + '; '.join(rec['blockers'])

    # ── Pass 4: T-TRAY (Party Tray formula) ────────────────────────────
    trays = []
    if not opts.get('no_tray'):
        base_recs = {(r['ri'], r['variant']): r for r in records}
        for tv in [v for v in var_cols if v in TRAY_VARIANTS]:
            # action1 emits the full template header list, so a 'Party Tray'
            # column exists in EVERY Menu Preview. Only a restaurant that
            # actually prices trays gets T-TRAY / a P-TRAY section.
            tci = var_cols[tv]
            if not any(_parse(r[tci]) is not None
                       for r in rows_2a[1:] if tci < len(r)):
                continue
            formulas = detect_tray_formulas(rows_2a, var_cols, tv, opts)
            trays.extend(f for g, f in formulas.items() if g and f['population'])
            tray_recs = infer_trays(rows_2a, rows_2b, var_cols, formulas, base_recs, opts)
            # a dish listed under two dish groups is ONE dish with ONE tray price:
            # a tray is derived only when EVERY group's formula derives the same
            # price — a group whose formula is missing, does not fit the dish,
            # or says "no change" disagrees, so no row is derived (asked)
            by_dish = defaultdict(list)
            for rec in tray_recs:
                by_dish[rec['dish']].append(rec)
            for dish, recs in by_dish.items():
                if len(recs) < 2 or all(r['inferred'] is None for r in recs):
                    continue
                if len({r['inferred'] for r in recs}) > 1:
                    detail = '; '.join(f"{r['group']}: " + _tray_outcome(r) for r in recs)
                    for r in recs:
                        r.update(inferred=None, status='SKIPPED', tier='conflict (same dish, several groups)',
                                 delta=None, blockers=[], ask=True,
                                 reason=f'SAME DISH IN SEVERAL GROUPS — their tray formulas disagree '
                                        f'({detail}). Not guessed; ask on callback.')
                elif len({r['status'] for r in recs}) > 1:
                    # the same price, but one group's row is held back (suggest-only):
                    # every row is
                    blockers = sorted({b for r in recs for b in (r.get('blockers') or [])})
                    for r in recs:
                        if r['status'] == 'ELIGIBLE':
                            r['status'] = 'SUGGESTED'
                            r['blockers'] = list(blockers)
                            r['reason'] = 'suggest-only: ' + '; '.join(blockers) + \
                                ' (same dish, held back in another of its groups)'
            records.extend(tray_recs)

    records.sort(key=lambda r: (r['scenario'], r['group'], r['dish'], r['variant']))
    core = [r for r in records if r['scenario'] != 'TRAY']
    tray_recs = [r for r in records if r['scenario'] == 'TRAY']
    stats = {
        'recipient_cells': len(recipients),
        'eligible': sum(1 for r in core if r['status'] == 'ELIGIBLE'),
        'suggested': sum(1 for r in core if r['status'] == 'SUGGESTED'),
        'no_change_expected': sum(1 for r in core if r['status'] == 'NO CHANGE EXPECTED'),
        'skipped': sum(1 for r in core if r['status'] == 'SKIPPED'),
        'cohort_conflicts': sum(1 for r in core if r['tier'] == 'T1 conflict'),
        'family_conflicts': sum(1 for r in core if r['tier'] == 'T0 conflict'),
        'dish_row_conflicts': sum(1 for r in core if r['tier'] == 'conflict (same dish, several groups)'),
        'tier0': sum(1 for r in core
                     if r['tier'].startswith('T0 ') and r['tier'] != 'T0 conflict'
                     and r['status'] in ('ELIGIBLE', 'SUGGESTED')),
        'v_gap': sum(1 for r in core if r['scenario'] == 'V-GAP'),
        'd_gap': sum(1 for r in core if r['scenario'] == 'D-GAP'),
        'tier1': sum(1 for r in core
                     if r['tier'].startswith('T1 exact')
                     and r['status'] in ('ELIGIBLE', 'SUGGESTED')),
        'tier2': sum(1 for r in core
                     if r['tier'].startswith('T2 uniform')
                     and r['status'] in ('ELIGIBLE', 'SUGGESTED')),
        'tray_cells': len(tray_recs),
        'tray_eligible': sum(1 for r in tray_recs if r['status'] == 'ELIGIBLE'),
        'tray_suggested': sum(1 for r in tray_recs if r['status'] == 'SUGGESTED'),
        'tray_no_change': sum(1 for r in tray_recs if r['status'] == 'NO CHANGE EXPECTED'),
        'tray_skipped': sum(1 for r in tray_recs if r['status'] == 'SKIPPED'),
    }
    return records, stats, var_cols, trays


# ══════════════════════════════════════════════════════════════════════
#  T-TRAY — Party Tray derived from the restaurant's own tray formula
# ══════════════════════════════════════════════════════════════════════
#
#  Restaurants price a Party Tray as a formula off a base column on the same
#  dish row, e.g.  Tray = Large × 5 − $2.00  (the specialist's =G23*5 habit).
#  The formula is LEARNED from the system's OLD prices, restaurant-wide, and
#  only trusted when at least --tray-min-agreement (default 70%) of tray rows
#  follow the same (base column, multiplier, offset). It is then applied to
#  the NEW printed base price of rows that themselves follow it.
#
#  This is NOT the forbidden cross-variant DELTA transfer (see docstring):
#  no delta is copied. A multiplicative formula is fitted and gated on
#  restaurant-wide evidence, and every outcome is a callback item (G14).

def _fmt_formula(base, k, d) -> str:
    kk = f'{k:g}'
    if abs(d) < 0.005:
        return f'{base} × {kk}'
    return f'{base} × {kk} {"+" if d > 0 else "−"} {_fmt(abs(d))}'


def _tray_outcome(rec) -> str:
    """What one group's tray formula said about a row, in a few words (for the
    same-dish-in-several-groups conflict)."""
    if rec['inferred'] is not None:
        return _fmt(rec['inferred']) + ('' if rec['status'] in ('ELIGIBLE', 'NO CHANGE EXPECTED')
                                        else ' (suggest-only)')
    if rec['tier'] == 'T-TRAY outlier':
        return "its tray does not follow this group's formula"
    if rec['tier'] == 'T-TRAY unavailable':
        return 'no reliable formula for this group'
    if rec['status'] == 'NO CHANGE EXPECTED':
        return 'tray stays'
    return 'not derived'


def detect_tray_formula(rows_2a, var_cols, tray_v, opts, only_rows=None, group='') -> dict:
    """Fit Tray = Base × k + d over every system row that carries a tray
    price AND at least one other price — the whole menu, or only the rows in
    `only_rows` (one dish group). Pure; no writes."""
    headers = [str(c).strip() if c else '' for c in rows_2a[0]]
    name_col = headers.index('Dish Name:Primary') if 'Dish Name:Primary' in headers else 1
    tray_ci = var_cols[tray_v]
    base_cols = [v for v in var_cols if v not in TRAY_VARIANTS]

    pop = []                                    # (ri, old_tray, {base_v: old})
    rows_of = defaultdict(list)                 # dish → its rows (one dish in two groups)
    for ri in range(1, len(rows_2a)):
        if only_rows is not None and ri not in only_rows:
            continue
        r = list(rows_2a[ri])
        dish = str(r[name_col]).strip() if len(r) > name_col and r[name_col] else ''
        if not dish:
            continue
        rows_of[dish].append(ri)
        if len(rows_of[dish]) > 1:
            continue                            # counted once: agreement is per dish, not per row
        tray = _parse(r[tray_ci]) if tray_ci < len(r) else None
        if tray is None or (dish, tray_v) in (opts.get('absolute_cells') or ()):
            continue
        others = {}
        for v in base_cols:
            ci = var_cols[v]
            p = _parse(r[ci]) if ci < len(r) else None
            if p is not None and p > 0:
                others[v] = p
        if others:
            pop.append((ri, tray, others))

    info = {'variant': tray_v, 'group': group, 'status': 'NOT DETECTED', 'base': None,
            'multiplier': None, 'offset': None, 'fits': 0,
            'population': len(pop), 'agreement': 0.0, 'fit_rows': set(),
            'pop_rows': {ri for ri, _t, _o in pop}, 'source': 'this group' if group else 'the whole menu',
            'describe': '', 'reason': ''}
    if len(pop) < opts['tray_min_rows']:
        info['status'] = 'INSUFFICIENT DATA'
        info['reason'] = (f'only {len(pop)} tray row(s) with a base price '
                          f'(--tray-min-rows {opts["tray_min_rows"]})')
        return info

    # Enumerate every (base, k, d) candidate. For each (base, k) the offsets
    # are rounded to cents; each distinct offset is a candidate, and
    # candidates that capture the same row set collapse to the most common
    # exact offset (so -2.00 and -1.99 are not a spurious tie).
    cands = []
    for base in base_cols:
        pairs = [(ri, t, o[base]) for ri, t, o in pop if base in o]
        if not pairs:
            continue
        for k in TRAY_MULTIPLIERS:
            offs = [round(t - k * b, 2) for _, t, b in pairs]
            exact = defaultdict(int)
            for d in offs:
                exact[d] += 1
            by_rows = {}
            for d in exact:
                if abs(d) > opts['tray_max_offset']:
                    continue
                fit = frozenset(ri for (ri, _, _), od in zip(pairs, offs)
                                if abs(od - d) <= TRAY_TOL)
                prev = by_rows.get(fit)
                if prev is None or exact[d] > exact[prev]:
                    by_rows[fit] = d
            for fit, d in by_rows.items():
                bases = {round(b, 2) for ri, _, b in pairs if ri in fit}
                cands.append({'base': base, 'k': k, 'd': d, 'fit': fit,
                              'distinct_bases': len(bases)})

    # A candidate whose fitting rows all share ONE base price cannot pin
    # down the multiplier (×4 + $8.20 and ×5 − $4.75 fit equally) — drop it.
    live = [c for c in cands if c['distinct_bases'] >= 2]
    if not live:
        info['status'] = 'INSUFFICIENT DATA'
        info['reason'] = ('tray rows do not have at least 2 different base prices '
                          'under any formula — multiplier cannot be determined')
        return info

    best_n = max(len(c['fit']) for c in live)
    top = [c for c in live if len(c['fit']) == best_n]
    best = top[0]
    agreement = best_n / len(pop)
    info.update({'base': best['base'], 'multiplier': best['k'], 'offset': best['d'],
                 'fits': best_n, 'agreement': agreement, 'fit_rows': set(best['fit']),
                 'describe': f'{tray_v} = {_fmt_formula(best["base"], best["k"], best["d"])}'})

    if len(top) > 1:
        info['status'] = 'AMBIGUOUS'
        alts = '; '.join(_fmt_formula(c['base'], c['k'], c['d']) for c in top[:3])
        info['reason'] = (f'{len(top)} different formulas each fit {best_n}/{len(pop)} '
                          f'tray rows ({alts}) — not picking one (G14)')
        info['describe'] = ''
    elif best_n < opts['tray_min_rows']:
        info['status'] = 'NOT DETECTED'
        info['reason'] = (f'best formula fits only {best_n} row(s) '
                          f'(--tray-min-rows {opts["tray_min_rows"]})')
    elif agreement + 1e-9 < opts['tray_min_agreement']:
        info['status'] = 'NOT DETECTED'
        info['reason'] = (f'best formula {info["describe"]} fits {best_n}/{len(pop)} '
                          f'tray rows ({agreement:.0%}) < --tray-min-agreement '
                          f'{opts["tray_min_agreement"]:.0%}')
    else:
        info['status'] = 'DETECTED'
        info['reason'] = f'fits {best_n}/{len(pop)} tray rows ({agreement:.0%})'
    if info['status'] != 'DETECTED':
        info['describe'] = ''        # best candidate is already quoted in 'reason'
    # every row of a fitting dish follows the formula (its copy in another group too)
    first = {rs[0]: rs for rs in rows_of.values()}
    info['fit_rows'] = {rj for ri in info['fit_rows'] for rj in first.get(ri, [ri])}
    return info


def detect_tray_formulas(rows_2a, var_cols, tray_v, opts) -> dict:
    """{dish group: formula} — Party Tray formulas PER DISH GROUP (maintainer,
    2026-10-05: trays usually follow one formula across the menu, but some
    restaurants price them per group — be safe). A group's own formula when
    its own tray rows agree (detect_tray_formula on that group). A group too
    small or too mixed to confirm one uses the whole-menu formula only when
    that fits ≥ --tray-min-agreement of the group's own tray rows; otherwise
    the group has no formula and its changed trays are asked on the callback.
    Key '' = the whole-menu formula (for reporting)."""
    headers = [str(c).strip() if c else '' for c in rows_2a[0]]
    name_col = headers.index('Dish Name:Primary') if 'Dish Name:Primary' in headers else 1
    group_col = headers.index('Dish Group') if 'Dish Group' in headers else 0
    menu = detect_tray_formula(rows_2a, var_cols, tray_v, opts)
    rows_in = defaultdict(set)
    for ri in range(1, len(rows_2a)):
        r = list(rows_2a[ri])
        if len(r) > name_col and r[name_col]:
            rows_in[str(r[group_col]).strip() if len(r) > group_col and r[group_col] else ''].add(ri)
    out = {'': menu}
    for g, ris in rows_in.items():
        own = detect_tray_formula(rows_2a, var_cols, tray_v, opts, only_rows=ris, group=g)
        if own['population'] == 0 or own['status'] == 'DETECTED':
            out[g] = own                      # nothing to derive from / its own formula
            continue
        n = own['population']
        fit = own['pop_rows'] & menu['fit_rows']
        if menu['status'] == 'DETECTED' and len(fit) / n + 1e-9 >= opts['tray_min_agreement']:
            eff = dict(menu)
            eff.update({'group': g, 'source': 'the whole menu', 'fits': len(fit), 'population': n,
                        'agreement': len(fit) / n, 'pop_rows': own['pop_rows'],
                        'fit_rows': ris & menu['fit_rows'],
                        'reason': (f'the whole-menu formula fits {len(fit)}/{n} of this group\'s tray rows '
                                   f'({len(fit) / n:.0%}); on its own the group {own["reason"]}')})
            out[g] = eff
        else:
            own['reason'] += (f'; the whole-menu formula {menu["describe"]} fits only {len(fit)}/{n} of '
                              f'this group\'s tray rows' if menu['status'] == 'DETECTED' else '')
            out[g] = own
    return out


def infer_trays(rows_2a, rows_2b, var_cols, formulas, base_recs, opts) -> list:
    """One record per UNPRINTED tray cell (system has a tray price, paper
    did not print one). Printed trays were already compared by action3.
    base_recs: {(ri, variant): rec} from the T1/T2 pass, used only to SUGGEST
    (never apply) a tray off a base price that was itself inferred."""
    headers = [str(c).strip() if c else '' for c in rows_2a[0]]
    name_col = headers.index('Dish Name:Primary') if 'Dish Name:Primary' in headers else 1
    group_col = headers.index('Dish Group') if 'Dish Group' in headers else 0
    tray_v = formulas['']['variant']
    tray_ci = var_cols[tray_v]
    base_cols = [v for v in var_cols if v not in TRAY_VARIANTS]

    out = []
    for ri in range(1, len(rows_2a)):
        r2a = list(rows_2a[ri])
        r2b = list(rows_2b[ri]) if ri < len(rows_2b) else []
        dish = str(r2a[name_col]).strip() if len(r2a) > name_col and r2a[name_col] else ''
        if not dish:
            continue
        old = _parse(r2a[tray_ci]) if tray_ci < len(r2a) else None
        if old is None:
            continue                                   # G1: no tray slot on this row
        if (dish, tray_v) in (opts.get('absolute_cells') or ()):
            continue                                   # G11: $0, priced on modifiers
        if (_parse(r2b[tray_ci]) if tray_ci < len(r2b) else None) is not None:
            continue                                   # printed → action3 compared it
        if (dish, tray_v) in (opts.get('paper_conflicts') or ()):
            continue                                   # printed twice, differently → asked (P-PRICE)
        group = str(r2a[group_col]).strip() if len(r2a) > group_col and r2a[group_col] else ''
        # the formula of THIS row's dish group (per group, maintainer 2026-10-05)
        formula = formulas.get(group) or formulas['']
        ok = formula['status'] == 'DETECTED'
        base_v, k, d = formula['base'], formula['multiplier'], formula['offset']
        conf = 'HIGH' if formula['agreement'] >= 0.9 else 'MEDIUM'
        fit_line = (f'{formula["describe"]} — fits {formula["fits"]}/{formula["population"]} '
                    f'tray rows of {group or "the menu"} ({formula["agreement"]:.0%}'
                    + (', the whole-menu formula' if formula['source'] == 'the whole menu' and group else '')
                    + ')') if ok else ''
        rec = {'ri': ri, 'variant': tray_v, 'ci': tray_ci, 'old': old,
               'group': group, 'dish': dish, 'scenario': 'TRAY', 'tier': '',
               'confidence': '', 'inferred': None, 'donor_n': 0, 'evidence': '',
               'status': 'SKIPPED', 'reason': '', 'delta': None, 'blockers': [],
               'base_variant': '', 'base_old': None, 'base_new': None,
               'base_source': '', 'ask': False}

        if abs(old) < 0.005:
            rec['reason'] = 'system tray price is $0.00 — never priced by inference'
            out.append(rec)
            continue

        # What moved on this row? (printed, non-tray)
        moved = {}
        printed_any = False
        for v in base_cols:
            ci = var_cols[v]
            o = _parse(r2a[ci]) if ci < len(r2a) else None
            n = _parse(r2b[ci]) if ci < len(r2b) else None
            if n is not None:
                printed_any = True
                if o is not None and abs(n - o) > 0.005:
                    moved[v] = (o, n)
        has_base = any((_parse(r2a[var_cols[v]]) if var_cols[v] < len(r2a) else None)
                       is not None for v in base_cols)

        if not has_base:
            rec['reason'] = 'tray-only dish — no base price on the row to derive from'
            out.append(rec)
            continue

        if not ok:
            rec['tier'] = 'T-TRAY unavailable'
            if moved:
                rec['ask'] = True
                rec['evidence'] = '; '.join(f'{v} {_fmt(o)}→{_fmt(n)}' for v, (o, n) in moved.items())
                rec['reason'] = (f'{group or "menu"} tray formula {formula["status"].lower()}: '
                                 f'{formula["reason"]}. Base price changed; tray left at current price '
                                 f'— ask client.')
            elif printed_any:
                rec['status'] = 'NO CHANGE EXPECTED'
                rec['reason'] = 'no printed base price moved — current tray looks correct'
            else:
                rec['reason'] = 'dish not on paper menu — base unknown (see the Price Inference tab for the base)'
            out.append(rec)
            continue

        rec['tier'] = 'T-TRAY formula'
        rec['base_variant'] = base_v
        old_b = _parse(r2a[var_cols[base_v]]) if var_cols[base_v] < len(r2a) else None
        new_b = _parse(r2b[var_cols[base_v]]) if var_cols[base_v] < len(r2b) else None
        rec['base_old'], rec['base_new'] = old_b, new_b
        fits_row = ri in formula['fit_rows']

        if old_b is None:
            rec['ask'] = bool(moved)
            rec['reason'] = (f'row has no {base_v} price — formula cannot apply'
                             + ('; other prices moved, ask client' if moved else ''))
            out.append(rec)
            continue

        if new_b is not None:                          # base printed on paper
            rec['base_source'] = 'PAPER'
            if abs(new_b - old_b) < 0.005:
                rec['status'] = 'NO CHANGE EXPECTED'
                rec['reason'] = f'{base_v} unchanged on paper — tray stays'
                out.append(rec)
                continue
            if not fits_row:
                would = round(k * new_b + d, 2)
                rec['ask'] = True
                rec['tier'] = 'T-TRAY outlier'
                rec['evidence'] = (f'formula predicts {_fmt(round(k * old_b + d, 2))} for current '
                                   f'{base_v} {_fmt(old_b)}, system tray is {_fmt(old)}; '
                                   f'on new {base_v} {_fmt(new_b)} formula would give {_fmt(would)}')
                rec['reason'] = ('this tray does not follow the restaurant formula — '
                                 'not derived; ask client for the new tray price')
                out.append(rec)
                continue
            rec['inferred'] = round(k * new_b + d, 2)
            rec['confidence'] = conf
            rec['evidence'] = f'{base_v} {_fmt(old_b)}→{_fmt(new_b)} (printed); {fit_line}'
        else:                                          # base not printed
            brec = base_recs.get((ri, base_v))
            if (fits_row and brec and brec.get('inferred') is not None
                    and brec['status'] in ('ELIGIBLE', 'SUGGESTED')):
                rec['base_source'] = 'INFERRED'
                rec['base_new'] = brec['inferred']
                rec['inferred'] = round(k * brec['inferred'] + d, 2)
                rec['confidence'] = 'LOW'
                rec['ask'] = True
                rec['evidence'] = (f'{base_v} {_fmt(old_b)}→{_fmt(brec["inferred"])} '
                                   f'(INFERRED by {brec["tier"]}, not printed); {fit_line}')
                rec['status'] = 'SUGGESTED'
                rec['blockers'] = [f'base {base_v} is itself inferred, not printed']
                rec['reason'] = ('suggest-only: base price was inferred, not printed — '
                                 f'confirm {base_v} first, then tray = {_fmt_formula(base_v, k, d)}')
                rec['delta'] = round(rec['inferred'] - old, 2)
            elif (dish, base_v) in (opts.get('paper_conflicts') or ()):
                rec['ask'] = True
                rec['reason'] = (f'the paper prints two different {base_v} prices (P-PRICE) — nothing to '
                                 f'derive from until the client confirms one; ask for the tray too')
            else:
                rec['ask'] = bool(moved)
                rec['reason'] = (f'{base_v} not printed — nothing to derive from'
                                 + ('; other prices moved, ask client' if moved else ''))
            out.append(rec)
            continue

        # ── Eligibility screening (base printed, row fits formula) ──────
        delta = round(rec['inferred'] - old, 2)
        rec['delta'] = delta
        if abs(delta) < 0.005:
            rec['status'] = 'NO CHANGE EXPECTED'
            rec['reason'] = 'formula gives the current tray price — nothing to write'
            out.append(rec)
            continue
        blockers = []
        if delta < 0:
            blockers.append('inferred DECREASE — verify like a P2')
        cap_abs = opts['max_delta_abs'] * k
        if abs(delta) > cap_abs:
            blockers.append(f'|delta| {_fmt(abs(delta))} > --max-delta-abs × {k:g} = {_fmt(cap_abs)}')
        if old > 0 and abs(delta) / old * 100 > opts['max_delta_pct']:
            blockers.append(f'|delta| is {abs(delta) / old * 100:.0f}% of old tray '
                            f'> --max-delta-pct {opts["max_delta_pct"]:.0f}%')
        if _is_low_priority(group) and not opts['include_low_priority']:
            blockers.append('operational group (Drink/Sauces/Side Order)')
        rec['blockers'] = blockers
        rec['ask'] = True
        rec['status'] = 'ELIGIBLE' if not blockers else 'SUGGESTED'
        if blockers:
            rec['reason'] = 'suggest-only: ' + '; '.join(blockers)
        out.append(rec)
    return out


# ══════════════════════════════════════════════════════════════════════
#  Report writing
# ══════════════════════════════════════════════════════════════════════

def _ref_columns(ws) -> dict:
    """Map a variant header → its hidden _ref_ mirror column index (1-based)."""
    out = {}
    for cell in next(ws.iter_rows(min_row=1, max_row=1)):
        h = str(cell.value).strip() if cell.value else ''
        if h.startswith('_ref_') and not h.startswith(('_ref_sys_', '_ref_key:')):
            out[h[len('_ref_'):]] = cell.column
    return out


def _same_price(a, b) -> bool:
    """Two report cells hold the same price (both empty counts as same)."""
    pa, pb = _parse(a), _parse(b)
    if pa is None or pb is None:
        return pa is None and pb is None and str(a or '').strip() == str(b or '').strip()
    return abs(pa - pb) < 0.005


def _last_content_row(ws) -> int:
    """Last row that actually holds a value. openpyxl's max_row stays stale
    after delete_rows (row dimensions persist), which would make the append
    blocks below drift down one row on every re-run."""
    last = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                last = max(last, cell.row)
                break
    return last


def _name_column(ws) -> int:
    """1-based column of 'Dish Name:Primary' in the report (G5)."""
    for cell in next(ws.iter_rows(min_row=1, max_row=1)):
        if str(cell.value).strip() == 'Dish Name:Primary':
            return cell.column
    return 2


def _write_price(cell, value):
    """Write a price matching the format already in the cell/column, so the
    report stays internally consistent whether or not _strip_dollar_display
    has converted '$X.XX' strings to bare numbers."""
    if isinstance(cell.value, str) and cell.value.strip().startswith('$'):
        cell.value = _fmt(value)
    else:
        cell.value = round(float(value), 2)
        cell.number_format = '0.00'


def _undo_previous(ws):
    """A re-run starts from the report as action3 left it: purple cells from an
    earlier run go back to the system price (the hidden _ref_sys_ mirror) and
    their _ref_ follows, so a re-run with other options (or --mode suggest)
    never keeps a guess it no longer makes. A purple cell the specialist typed
    over (value ≠ its _ref_) is their edit — never undone. → (undone, kept):
    kept = purple cells that stay (typed over, or an older report without the
    mirror)."""
    hdr = [str(c.value).strip() if c.value is not None else '' for c in ws[1]]
    sys_cols = {h[len('_ref_sys_'):]: i + 1 for i, h in enumerate(hdr) if h.startswith('_ref_sys_')}
    ref_cols = {h[len('_ref_'):]: i + 1 for i, h in enumerate(hdr)
                if h.startswith('_ref_') and not h.startswith(('_ref_sys_', '_ref_key:'))}
    price_cols = [i + 1 for i, h in enumerate(hdr) if h and not h.startswith('_ref_') and h not in NAME_COLUMNS]
    name_col = _name_column(ws)

    def fill_of(c):
        f = c.fill
        return str(f.fgColor.rgb or '')[-6:].upper() if f is not None and f.fill_type == 'solid' else ''

    undone = kept = 0
    rows = set()
    for ci in price_cols:
        h = hdr[ci - 1]
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(row=r, column=ci)
            if fill_of(cell) != 'E4DFEC':
                continue
            sc, rc = sys_cols.get(h), ref_cols.get(h)
            if sc is None or (rc and not _same_price(cell.value, ws.cell(row=r, column=rc).value)):
                kept += 1
                continue
            sv = ws.cell(row=r, column=sc).value
            cell.value = sv
            cell.fill = PatternFill(fill_type=None)
            if ref_cols.get(h):
                ws.cell(row=r, column=ref_cols[h]).value = sv
            undone += 1
            rows.add(r)
    for r in rows:                     # name cells: as action3 left them
        fills = {fill_of(ws.cell(row=r, column=c)) for c in price_cols}
        if 'E4DFEC' in fills:
            continue
        nf = (PatternFill(fill_type=None) if fills & {'ADD8E6', 'FFB6C1', 'C6EFCE'}
              else PatternFill('solid', fgColor='FCF7DE'))
        for c in (name_col, name_col + 1):
            ws.cell(row=r, column=c).fill = nf
    return undone, kept


def apply_to_dish_item(ws, records, ref_cols, name_col):
    """Write ELIGIBLE inferences into the Dish Item price cells (purple) and
    refresh the hidden _ref_ mirror so apricot 'specialist edit' CF does not
    false-fire on the script's own value."""
    applied = 0
    for rec in records:
        if rec['status'] != 'ELIGIBLE':
            continue
        # action3 writes rows_2a[ri] to report row ri+1 (row 1 = header).
        row = rec['ri'] + 1
        col = rec['ci'] + 1
        cell = ws.cell(row=row, column=col)
        here = str(ws.cell(row=row, column=name_col).value or '').strip()
        if here != rec['dish']:
            # the report does not line up with the Menu Preview (a stale report)
            rec['status'] = 'SUGGESTED'
            rec['blockers'].append(f"report row {row} is '{here}', not this dish — not written "
                                   f"(re-run action3 on the current Menu Preview)")
            rec['reason'] = 'suggest-only: ' + '; '.join(rec['blockers'])
            continue
        if _parse(cell.value) is None:
            # Defensive G1: the report cell lost its price (edited/blank).
            rec['status'] = 'SUGGESTED'
            rec['blockers'].append('report cell is blank — not written (G1)')
            rec['reason'] = 'suggest-only: ' + '; '.join(rec['blockers'])
            continue
        ref_ci = ref_cols.get(rec['variant'])
        if ref_ci and not _same_price(cell.value, ws.cell(row=row, column=ref_ci).value):
            # the specialist typed over this cell (≠ its _ref_): their price stays
            rec['status'] = 'SUGGESTED'
            rec['blockers'].append('the report cell was edited by the specialist — not overwritten')
            rec['reason'] = 'suggest-only: ' + '; '.join(rec['blockers'])
            continue
        _write_price(cell, rec['inferred'])
        cell.fill = PURPLE_FILL
        if ref_ci:
            _write_price(ws.cell(row=row, column=ref_ci), rec['inferred'])
        # action3 flags a row with no blue/pink/green activity YELLOW on the two
        # name cells ("still needs action"). A row that just received an
        # inference does carry a provisional new price, so leaving it yellow is
        # a false signal — repaint it purple to match the price cell.
        for col in (name_col, name_col + 1):
            ws.cell(row=row, column=col).fill = PURPLE_FILL
        rec['status'] = 'APPLIED'
        applied += 1
    return applied


def render_inference_tab(wb, records, stats, opts, trays=()):
    if INF_TAB in wb.sheetnames:
        del wb[INF_TAB]
    ws = wb.create_sheet(INF_TAB)

    ws.cell(row=1, column=1,
            value='PRICE INFERENCE — prices the paper menu did not print').font = Font(bold=True, size=13)
    ws.cell(row=2, column=1, value=(
        f'mode={opts["mode"]} | min-donors={opts["min_donors"]} | '
        f'tier2-min-donors={opts["tier2_min_donors"]} | '
        f'max-delta-abs={_fmt(opts["max_delta_abs"])} | '
        f'max-delta-pct={opts["max_delta_pct"]:.0f}%'
    )).font = Font(italic=True, color='666666')
    ws.cell(row=3, column=1, value=(
        'APPLIED = written into Dish Item (purple) — CONFIRM EACH ON THE CALLBACK. '
        'SUGGESTED = not written, specialist decides. '
        'SKIPPED = no safe pattern. '
        'V-GAP = dish matched, this variant not printed. '
        'D-GAP = dish not on paper menu at all. '
        'TRAY = Party Tray not printed, derived from the restaurant tray formula (T-TRAY).'
    )).font = Font(italic=True, color='666666')
    lr = 4
    for t in trays:                                  # one formula per dish group
        where = t.get('group') or 'menu'
        line = (f'T-TRAY {where}: {t["status"]} — {t["describe"]} — {t["reason"]}' if t['describe']
                else f'T-TRAY {where}: {t["status"]} ({t["variant"]}) — {t["reason"]}')
        ws.cell(row=lr, column=1, value=line).font = Font(bold=True, color='7030A0')
        lr += 1

    header_row = max(5, lr)
    headers = ['ID', 'Scenario', 'Status', 'Dish Group', 'Dish Name', 'Variant',
               'Current Price', 'Inferred Price', 'Delta', 'Tier', 'Confidence',
               'Donors', 'Evidence', 'Reason / Action']
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=header_row, column=ci, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
    ws.freeze_panes = f'A{header_row + 1}'

    r = header_row + 1
    for i, rec in enumerate(records, 1):
        ws.cell(row=r, column=1, value=f'INF-{i}')
        ws.cell(row=r, column=2, value=rec['scenario'])
        ws.cell(row=r, column=3, value=rec['status'])
        ws.cell(row=r, column=4, value=rec['group'])
        ws.cell(row=r, column=5, value=rec['dish'])
        ws.cell(row=r, column=6, value=rec['variant'])
        ws.cell(row=r, column=7, value=_fmt(rec['old']))
        ws.cell(row=r, column=8, value=_fmt(rec['inferred']))
        ws.cell(row=r, column=9, value=_fmt_signed(rec['delta']) if rec.get('delta') else '')
        ws.cell(row=r, column=10, value=rec['tier'])
        ws.cell(row=r, column=11, value=rec['confidence'])
        ws.cell(row=r, column=12, value=rec['donor_n'] or '')
        ws.cell(row=r, column=13, value=rec['evidence'])
        ws.cell(row=r, column=14, value=rec['reason'])
        fill = {'APPLIED': PURPLE_FILL, 'SUGGESTED': GREY_FILL,
                'NO CHANGE EXPECTED': GREEN_FILL}.get(rec['status'])
        if rec['tier'] in ('T1 conflict', 'T0 conflict', 'conflict (same dish, several groups)'):
            fill = RED_FILL
        if fill:
            for ci in range(1, 4):
                ws.cell(row=r, column=ci).fill = fill
        r += 1
    if not records:
        ws.cell(row=r, column=1, value='(no unprinted prices to infer)')

    for col, w in (('A', 10), ('B', 10), ('C', 20), ('D', 24), ('E', 38), ('F', 12),
                   ('G', 14), ('H', 14), ('I', 10), ('J', 26), ('K', 12), ('L', 9),
                   ('M', 62), ('N', 62)):
        ws.column_dimensions[col].width = w
    ws.sheet_properties.tabColor = 'E4DFEC'
    return ws


def append_questionnaire_tray(wb, records, trays):
    """Data-driven Party Tray callback section (replaces action3's static
    'A-1 Party Tray pricing — confirm or update?'). Idempotent via sentinel.
    Written only when the restaurant has a tray column."""
    if 'Questionnaire' not in wb.sheetnames:
        return 0
    ws = wb['Questionnaire']
    for row in ws.iter_rows(min_col=1, max_col=1):
        if row[0].value == P_TRAY_SENTINEL:
            # an earlier run's block goes first — also when this run has no trays
            # (--no-tray), so no stale tray answers are left behind
            ws.delete_rows(row[0].row, ws.max_row - row[0].row + 1)
            break
    if not trays:
        return 0
    r = _last_content_row(ws) + 2
    ws.cell(row=r, column=1, value=P_TRAY_SENTINEL).font = Font(bold=True, size=11, color='7030A0')
    r += 1
    for t in trays:                                  # one row per dish group's formula
        g = t.get('group') or '(all)'
        if t['status'] == 'DETECTED':
            src = (' — the whole-menu formula, which fits this group' if t.get('source') == 'the whole menu'
                   and t.get('group') else '')
            q = (f'Detected {t["describe"]} for {g} ({t["fits"]}/{t["population"]} tray rows, '
                 f'{t["agreement"]:.0%}{src}). Confirm with the client that these trays still follow '
                 f'this. If any {t["base"]} price below changes on the callback, recompute '
                 f'its tray with the same formula.')
        else:
            q = (f'No reliable tray formula for {g} ({t["status"].lower()}: {t["reason"]}). '
                 f'Ask the client for new tray prices on the rows below.')
        ws.cell(row=r, column=1, value='P-TRAY-0')
        ws.cell(row=r, column=2, value=f'P-TRAY formula ({t["status"].lower()})')
        ws.cell(row=r, column=3, value=g)
        ws.cell(row=r, column=4, value=f'[{t["variant"]}]')
        ws.cell(row=r, column=8, value=q)
        r += 1
    asks = [x for x in records if x['scenario'] == 'TRAY' and x.get('ask')]
    if not asks:
        ws.cell(row=r, column=1, value='(no tray needs a callback answer)')
        return 0
    for i, rec in enumerate(asks, 1):
        ws.cell(row=r, column=1, value=f'P-TRAY-{i}')
        ws.cell(row=r, column=2, value=f'P-TRAY ({rec["status"].lower()} — {rec["tier"] or "n/a"})')
        ws.cell(row=r, column=3, value=rec['group'])
        ws.cell(row=r, column=4, value=f'{rec["dish"]} [{rec["variant"]}]')
        ws.cell(row=r, column=5, value=_fmt(rec['old']))
        ws.cell(row=r, column=6, value=_fmt(rec['inferred']))
        if rec.get('base_variant') and rec.get('base_new') is not None:
            src = 'printed' if rec['base_source'] == 'PAPER' else 'INFERRED'
            ws.cell(row=r, column=7, value=f'{rec["base_variant"]} {_fmt(rec["base_new"])} ({src})')
        ws.cell(row=r, column=8, value='. '.join(x for x in (rec['evidence'], rec['reason']) if x))
        if rec['status'] == 'APPLIED':
            for ci in range(1, 9):
                ws.cell(row=r, column=ci).fill = PURPLE_FILL
        r += 1
    return len(asks)


def append_summary_block(wb, stats, opts, applied, trays=()):
    """Append the inference counts + purple legend to Summary. Idempotent."""
    if 'Summary' not in wb.sheetnames:
        return
    ws = wb['Summary']
    for row in ws.iter_rows(min_col=1, max_col=1):
        if row[0].value == S_SENTINEL:
            ws.delete_rows(row[0].row, ws.max_row - row[0].row + 1)
            break
    r = _last_content_row(ws) + 2
    ws.cell(row=r, column=1, value=S_SENTINEL).font = Font(bold=True, size=12)
    r += 1
    ws.cell(row=r, column=1, value=('Purple = price inferred from the dish group, NOT printed on '
                                    'the paper menu. Confirm on the client callback.')).font = Font(
        italic=True, color='666666')
    r += 1
    for label, val in (
        ('Mode', opts['mode']),
        ('Unprinted prices examined', stats['recipient_cells']),
        ('  of which variant gaps (V-GAP, dish matched)', stats['v_gap']),
        ('  of which dish gaps (D-GAP, no paper match)', stats['d_gap']),
        ('Applied to Dish Item (purple, all tiers incl. trays)', applied),
        ('Suggested only (specialist decides)', stats['suggested']),
        ('No change expected (cohort flat)', stats['no_change_expected']),
        ('Skipped — no safe pattern', stats['skipped']),
        ('  of which cohort conflicts (donors disagree)', stats['cohort_conflicts']),
        ('  of which same-price family conflicts', stats['family_conflicts']),
        ('Resolved by T0 same-price rules (Gen Tso/Orange/Sesame, Healthy)', stats['tier0']),
        ('Resolved by T1 exact old-price cohort', stats['tier1']),
        ('Resolved by T2 uniform group delta', stats['tier2']),
    ):
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=val)
        r += 1
    for t in trays:
        ws.cell(row=r, column=1, value=f'{t["variant"]} formula (T-TRAY), {t.get("group") or "menu"}')
        ws.cell(row=r, column=2, value=(f'{t["describe"]} [{t["reason"]}]' if t['describe']
                                        else f'{t["status"]} — {t["reason"]}'))
        r += 1
    if trays:
        v = trays[0]['variant']
        for label, val in (
            (f'  {v} unprinted, examined', stats['tray_cells']),
            (f'  {v} applied (purple)', stats.get('tray_applied', 0)),
            (f'  {v} suggested only', stats['tray_suggested']),
            (f'  {v} no change expected', stats['tray_no_change']),
            (f'  {v} skipped', stats['tray_skipped']),
        ):
            ws.cell(row=r, column=1, value=label)
            ws.cell(row=r, column=2, value=val)
            r += 1


MONEY_RE = re.compile(r'^\$([+-]?)(\d+(?:\.\d+)?)$')


def _numeric_display(wb):
    """Every price in the report is a plain number shown as 0.00 — never a
    '$X.XX' string — so a specialist just types 12.95 (same rule as action3).
    A signed delta keeps its sign; Questionnaire price columns that list
    several prices lose the '$'."""
    for ws in wb.worksheets:
        if ws.title == 'Dish Code':
            continue                    # codes are text, never prices
        questionnaire = ws.title == 'Questionnaire'
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str):
                    continue
                m = MONEY_RE.match(v.strip())
                if m:
                    cell.value = round(float(m.group(1).replace('+', '') + m.group(2)), 2)
                    cell.number_format = '+0.00;-0.00' if m.group(1) == '+' else '0.00'
                elif questionnaire and cell.column in (5, 6, 7) and '$' in v:
                    cell.value = v.replace('$', '')


def reorder_tabs(wb):
    order = [t for t in CANON_ORDER if t in wb.sheetnames]
    order += [t for t in wb.sheetnames if t not in order]
    wb._sheets.sort(key=lambda s: order.index(s.title))


# ══════════════════════════════════════════════════════════════════════
#  CLI
# ══════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(
        description='Infer unprinted prices from dish-group pricing patterns '
                    'and write them into the Price Comparison Report.')
    ap.add_argument('menu_preview', help='Menu Preview / Excel 2a (system prices)')
    ap.add_argument('ocr_output', help='OCR Output / Excel 2b (paper prices, mirrors 2a)')
    ap.add_argument('--report', required=True,
                    help='Price Comparison Report from action3 (written in place)')
    ap.add_argument('-o', '--output', help='Write a copy instead of updating --report in place')
    ap.add_argument('--mode', choices=['suggest', 'apply'], default='suggest',
                    help='suggest (default): report only. apply: also write purple '
                         'prices into Dish Item.')
    ap.add_argument('--min-donors', type=int, default=1,
                    help='Minimum donors for an APPLIED T1 inference (default 1; '
                         'a single donor is MEDIUM confidence)')
    ap.add_argument('--tier2-min-donors', type=int, default=3,
                    help='Minimum donors before a uniform group delta may be used (default 3)')
    ap.add_argument('--max-delta-abs', type=float, default=5.00,
                    help='Downgrade to suggest-only above this absolute delta (default 5.00)')
    ap.add_argument('--max-delta-pct', type=float, default=40.0,
                    help='Downgrade to suggest-only above this %% of the old price (default 40)')
    ap.add_argument('--include-low-priority', action='store_true',
                    help='Allow applying in Drink / Sauces / Side Order groups')
    ap.add_argument('--tray-min-agreement', type=float, default=0.70,
                    help='T-TRAY: share of tray rows that must follow ONE formula '
                         'before it is trusted (default 0.70)')
    ap.add_argument('--tray-min-rows', type=int, default=3,
                    help='T-TRAY: minimum tray rows fitting the formula (default 3)')
    ap.add_argument('--tray-max-offset', type=float, default=10.00,
                    help='T-TRAY: largest +/- dollar offset a formula may carry, '
                         'e.g. the "- $2.00" in Large x 5 - $2.00 (default 10.00)')
    ap.add_argument('--no-tray', action='store_true',
                    help='Disable T-TRAY (Party Tray derivation) for this run')
    ap.add_argument('--json', dest='json_path',
                    help='Write the deterministic summary JSON (default: '
                         'price_inference_summary.json beside the report)')
    args = ap.parse_args()

    opts = {'mode': args.mode, 'min_donors': args.min_donors,
            'tier2_min_donors': args.tier2_min_donors,
            'max_delta_abs': args.max_delta_abs, 'max_delta_pct': args.max_delta_pct,
            'include_low_priority': args.include_low_priority,
            'tray_min_agreement': args.tray_min_agreement,
            'tray_min_rows': args.tray_min_rows,
            'tray_max_offset': args.tray_max_offset,
            'no_tray': args.no_tray}

    wb_2a = load_workbook(args.menu_preview, read_only=True, data_only=True)
    wb_2b = load_workbook(args.ocr_output, read_only=True, data_only=True)
    rows_2a = list(wb_2a['price'].iter_rows(values_only=True))
    rows_2b = list(wb_2b['price'].iter_rows(values_only=True))
    # $0.00 dish cells whose price lives on modifiers (action1 'absolute_pricing',
    # v4.4): never a recipient, never a donor — an inferred price there would be
    # charged on top of the modifier price (G11).
    opts['absolute_cells'] = set()
    if 'absolute_pricing' in wb_2a.sheetnames:
        abs_rows = list(wb_2a['absolute_pricing'].iter_rows(values_only=True))
        ahdr = [str(c).strip() if c else '' for c in abs_rows[0]] if abs_rows else []
        if 'Dish Name' in ahdr and 'Variant' in ahdr:
            dn, vn = ahdr.index('Dish Name'), ahdr.index('Variant')
            opts['absolute_cells'] = {(str(r[dn]).strip(), str(r[vn]).strip())
                                      for r in abs_rows[1:]
                                      if r and len(r) > max(dn, vn) and r[dn] and r[vn]}
    # P1 rows whose suggested system match is a dish: it may be on the paper
    # under another name — the Healthy rule leaves such a dish alone.
    # dishes the paper prints in a size they do not have (action2 'Unplaced
    # Prices', Questionnaire P-SIZE): on the paper, so the Healthy rule leaves them
    opts['unplaced_dishes'] = ({str(r[1]).strip().lower() for r in
                                list(wb_2b['Unplaced Prices'].iter_rows(values_only=True))[1:] if r and len(r) > 1 and r[1]}
                               if 'Unplaced Prices' in wb_2b.sheetnames else set())
    # sizes the paper printed two different prices for (action2 'Price
    # Conflicts', Questionnaire P-PRICE): left blank and asked — never inferred,
    # never a donor, never a tray derived over; the dish counts as printed
    opts['paper_conflicts'] = ({(str(r[1]).strip(), str(r[2]).strip()) for r in
                                list(wb_2b['Price Conflicts'].iter_rows(values_only=True))[1:]
                                if r and len(r) > 2 and r[1] and r[2]}
                               if 'Price Conflicts' in wb_2b.sheetnames else set())
    opts['p1_suggested'], opts['p1_healthy'] = {}, {}
    if 'Unmatched' in wb_2b.sheetnames:
        for r in list(wb_2b['Unmatched'].iter_rows(values_only=True))[1:]:
            if not (r and str(r[0] or '').startswith('P1-')):
                continue
            if len(r) > 5 and r[5]:
                # action2 writes "Maybe: <system dish> (substring)" — the dish name
                # is between the prefix and the trailing reason
                sug = re.sub(r'^\s*maybe:\s*', '', str(r[5]).strip(), flags=re.I)
                sug = re.sub(r'\s*\([^()]*\)\s*$', '', sug).strip()
                if sug:
                    opts['p1_suggested'].setdefault(sug.lower(), str(r[3] or '').strip())
            if len(r) > 3 and r[3]:
                # "Diet Shrimp w/ Broccoli (Steamed)" → "shrimp with broccoli"
                opts['p1_healthy'].setdefault(_healthy_base(str(r[3])), str(r[3]).strip())
    wb_2a.close()
    wb_2b.close()
    if not rows_2a:
        print('ERROR: Menu Preview price tab is empty.')
        return 2

    records, stats, var_cols, trays = analyze(rows_2a, rows_2b, opts)

    # data_only=False: the report carries live validation formulas (Modifier
    # Rule) that must survive the round-trip.
    wb = load_workbook(args.report)
    if 'Dish Item' not in wb.sheetnames:
        print('ERROR: --report has no "Dish Item" tab — is it an action3 report?')
        return 2
    ws_di = wb['Dish Item']
    undone, kept = _undo_previous(ws_di)
    if undone:
        print(f'  Re-run: {undone} purple cell(s) from an earlier run put back to the system price '
              f'first (the hidden _ref_sys_ mirror), then inferred afresh.')
    if kept:
        print(f'  !! Re-run: {kept} earlier purple cell(s) kept as they are — typed over by the '
              f'specialist, or an older report without the _ref_sys_ mirror (re-run action3 for a '
              f'clean report).')

    applied = 0
    if args.mode == 'apply':
        applied = apply_to_dish_item(ws_di, records, _ref_columns(ws_di),
                                    _name_column(ws_di))
    else:
        for rec in records:
            if rec['status'] == 'ELIGIBLE':
                rec['status'] = 'SUGGESTED'
                rec['reason'] = (rec['reason']
                                 or 'eligible — run --mode apply to write it into Dish Item')
    stats['applied'] = applied
    stats['eligible'] = sum(1 for r in records if r['scenario'] != 'TRAY' and r['status'] == 'ELIGIBLE')
    stats['tray_eligible'] = sum(1 for r in records if r['scenario'] == 'TRAY' and r['status'] == 'ELIGIBLE')
    stats['tray_applied'] = sum(1 for r in records
                                if r['scenario'] == 'TRAY' and r['status'] == 'APPLIED')
    stats['suggested'] = sum(1 for r in records
                             if r['scenario'] != 'TRAY' and r['status'] == 'SUGGESTED')
    stats['tray_suggested'] = sum(1 for r in records
                                  if r['scenario'] == 'TRAY' and r['status'] == 'SUGGESTED')

    render_inference_tab(wb, records, stats, opts, trays)
    p_tray = append_questionnaire_tray(wb, records, trays)
    append_summary_block(wb, stats, opts, applied, trays)
    reorder_tabs(wb)
    _numeric_display(wb)   # every price a plain number — no '$X.XX' strings

    out = args.output or args.report
    wb.save(out)
    wb.close()

    summary = {'mode': args.mode, 'report': out,
               'options': {k: (sorted(v) if isinstance(v, set) else v) for k, v in opts.items()},
               'stats': stats,
               'questionnaire_p_tray_items': p_tray,
               'tray_formulas': [{k: v for k, v in t.items() if k not in ('fit_rows', 'pop_rows')}
                                 for t in trays],
               'records': [{k: v for k, v in r.items() if k not in ('ci',)}
                           for r in records]}
    json_path = args.json_path or os.path.join(
        os.path.dirname(os.path.abspath(out)) or '.', 'price_inference_summary.json')
    with open(json_path, 'w') as f:
        json.dump(summary, f, indent=2)

    print('=' * 60)
    print(f'  PRICE INFERENCE ({args.mode})')
    print(f'  Unprinted prices examined:      {stats["recipient_cells"]}'
          f'  (V-GAP {stats["v_gap"]}, D-GAP {stats["d_gap"]})')
    print(f'  Applied to Dish Item (purple):  {applied}')
    print(f'  Suggested only:                 {stats["suggested"]}')
    print(f'  No change expected:             {stats["no_change_expected"]}')
    print(f'  Skipped (no safe pattern):      {stats["skipped"]}'
          f'  (cohort conflicts {stats["cohort_conflicts"]})')
    print(f'  Tier split:                     T0 {stats["tier0"]}, T1 {stats["tier1"]}, '
          f'T2 {stats["tier2"]}  (T0 = same-price rules; family conflicts '
          f'{stats["family_conflicts"]})')
    for t in trays:
        print(f'  T-TRAY [{t["variant"]}] {t.get("group") or "menu"}:  {t["status"]}'
              + (f' — {t["describe"]}' if t['describe'] else '')
              + f' ({t["reason"]})')
    if trays:
        print(f'    unprinted trays: {stats["tray_cells"]} | applied (purple) '
              f'{stats["tray_applied"]} | suggested {stats["tray_suggested"]} | '
              f'no change {stats["tray_no_change"]} | skipped {stats["tray_skipped"]}')
        print(f'    Questionnaire P-TRAY callback items: {p_tray}')
    print(f'  Report:  {out}')
    print(f'  Summary: {json_path}')
    print('=' * 60)
    if applied:
        print('  ⚠ Every purple cell is a GUESS from the dish group. Read all of')
        print('    them back to the client on the callback before Part 2.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
