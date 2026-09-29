#!/usr/bin/env python3
"""
ocr_sanity_check.py - deterministic plausibility check on the OCR JSON produced
in Step 1 of Menu Change Part 1, BEFORE it is fed to Action 2.

Purpose: surface likely OCR misreads as review items rather than letting them
flow silently into the price comparison. It NEVER changes prices and never runs
the matcher - it only reads the JSON and prints flags. Dependencies: stdlib only.

Checks (all advisory):
  * SIZE ORDER   - Small price > Large price (a size is priced above a bigger one)
  * DROPPED DEC  - whole-number per-dish price >= 100 (likely 895 -> 8.95).
                  Not raised for catering-tray labels (TRAY_LABELS): a $120
                  Party Tray is normal. OUT OF RANGE still applies to them.
  * OUT OF RANGE - price <= 0 or > 500
  * UNCERTAIN    - item carries "uncertain": true (set by the OCR reader)
  * NO PRICE     - item has no price and its container has no default_price
                   (possible skipped row)

Exit: 0 always (advisory) unless --strict, which exits 1 when any flag is raised
or 2 on a malformed/unreadable JSON.
"""
import argparse, json, sys

DROPPED_DECIMAL_MIN = 100.0   # per-dish prices at/above this are suspicious
OUT_OF_RANGE_MAX    = 500.0
# Variant keys whose prices are routinely whole-dollar and >= $100 (v4.3).
# Keep in sync with action2 / price_inference.py TRAY_VARIANTS.
TRAY_LABELS = {'Party Tray'}

def _prices_of(item):
    """Yield (label, value) numeric prices from an item's prices/price/side_prices."""
    out = []
    pr = item.get('prices')
    if isinstance(pr, dict):
        for k, v in pr.items():
            if isinstance(v, (int, float)):
                out.append((k, float(v)))
    if isinstance(item.get('price'), (int, float)):
        out.append(('price', float(item['price'])))
    sp = item.get('side_prices')
    if isinstance(sp, dict):
        for k, v in sp.items():
            if isinstance(v, (int, float)):
                out.append((f'side:{k}', float(v)))
    return out

def _iter_items(data):
    """Yield (section_label, container_has_default, item) across all containers."""
    for sec in data.get('sections', []) or []:
        if not isinstance(sec, dict):
            continue
        label = sec.get('menu_section', '(unnamed section)')
        has_def = sec.get('default_price') is not None
        for it in sec.get('items', []) or []:
            if isinstance(it, dict):
                yield label, has_def, it
    for key in ('lunch_specials', 'combinations'):
        cont = data.get(key)
        if isinstance(cont, dict):
            has_def = cont.get('default_price') is not None
            for it in cont.get('items', []) or []:
                if isinstance(it, dict):
                    yield key, has_def, it
    so = data.get('side_orders')
    if isinstance(so, list):
        for it in so:
            if isinstance(it, dict):
                yield 'side_orders', False, it

def check(data):
    flags = []
    def add(kind, section, name, detail):
        flags.append((kind, section, name, detail))
    for section, has_def, it in _iter_items(data):
        name = it.get('name', '(unnamed)')
        prices = _prices_of(it)
        pdict = it.get('prices') if isinstance(it.get('prices'), dict) else {}

        # SIZE ORDER: Small must not exceed Large
        s, l = pdict.get('Small'), pdict.get('Large')
        if isinstance(s, (int, float)) and isinstance(l, (int, float)) and s > l:
            add('SIZE ORDER', section, name, f'Small {s} > Large {l}')

        for label, v in prices:
            if v <= 0:
                add('OUT OF RANGE', section, name, f'{label}={v} (<= 0)')
            elif (v >= DROPPED_DECIMAL_MIN and float(v).is_integer()
                  and label not in TRAY_LABELS):
                extra = ' and out of range' if v > OUT_OF_RANGE_MAX else ' (or a genuine party-tray/catering price)'
                add('DROPPED DEC', section, name, f'{label}={v:.0f} - likely {v/100:.2f}?{extra}')
            elif v > OUT_OF_RANGE_MAX:
                add('OUT OF RANGE', section, name, f'{label}={v} (> {OUT_OF_RANGE_MAX:.0f})')

        if it.get('uncertain') is True:
            note = it.get('uncertain_note') or 'flagged uncertain by reader'
            add('UNCERTAIN', section, name, note)

        # NO PRICE: nothing captured and no default to inherit
        if not prices and not has_def:
            add('NO PRICE', section, name, 'no price captured and no section default_price - skipped row?')
    return flags

def main(argv=None):
    ap = argparse.ArgumentParser(description='Deterministic plausibility check on the OCR JSON (advisory).')
    ap.add_argument('ocr_json', help='Path to the OCR JSON produced in Step 1')
    ap.add_argument('--strict', action='store_true', help='exit 1 if any flag is raised (default: advisory exit 0)')
    args = ap.parse_args(argv)
    try:
        with open(args.ocr_json, encoding='utf-8') as f:
            data = json.load(f)
    except (OSError, ValueError) as e:
        print(f"ocr_sanity_check: cannot read JSON: {e}", file=sys.stderr)
        return 2

    flags = check(data)
    if not flags:
        print("ocr_sanity_check: OK - no plausibility flags.")
        return 0

    order = ['OUT OF RANGE', 'SIZE ORDER', 'DROPPED DEC', 'NO PRICE', 'UNCERTAIN']
    flags.sort(key=lambda f: (order.index(f[0]) if f[0] in order else 99, f[1], f[2]))
    print(f"ocr_sanity_check: {len(flags)} item(s) to VERIFY before Action 2\n")
    cur = None
    for kind, section, name, detail in flags:
        if kind != cur:
            print(f"[{kind}]"); cur = kind
        print(f"    {section} :: {name}  ->  {detail}")
    print("\nReview each flag against the paper menu. These are advisory - fix the OCR JSON and re-run,")
    print("or confirm the price is correct (e.g. a real party-tray price) and proceed.")
    return 1 if args.strict else 0

if __name__ == '__main__':
    sys.exit(main())
