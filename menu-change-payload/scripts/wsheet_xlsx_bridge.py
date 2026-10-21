#!/usr/bin/env python3
"""
wsheet_xlsx_bridge.py — bridge wsheet TSV content dir <-> Raw-Gsheet .xlsx
=========================================================================

Standalone. Dependency: openpyxl only.

Why this exists (ACE / Option B):
    The menu-change scripts (action1 .. action5) speak Raw-Gsheet `.xlsx`.
    ACE speaks the live `wsheet` TSV content dir (one `<tab>.tsv` per tab,
    each with the 3-row header + JSON root rows). The cell layouts are
    IDENTICAL — only the container differs (one .xlsx workbook with a sheet
    per tab, vs. a directory of .tsv files). This bridge converts between
    them so the existing scripts run unchanged, and edits land back in the
    TSVs for `wsheet validate -> draft create` (Part 2 is DRAFT-ONLY: a human
    reviews and commits the staged draft separately — never `draft commit` here).

Two directions:

  to-xlsx <tsv_dir> -o <raw.xlsx>
      Every `<tab>.tsv` in the dir becomes a worksheet named `<tab>`.
      Cells are written verbatim as strings (no type coercion), preserving
      the 3-row header and JSON root rows. Feed the result to action1 (to
      build a Menu Preview) or use it as the action5 base.

  to-tsv <final_raw.xlsx> <tsv_dir> [--tabs dish_item,modifier_item]
      Patches the TSV dir IN PLACE from the .xlsx (e.g. action5 output).
      For each named tab, compares cell-by-cell and rewrites ONLY the cells
      that changed. Guardrails mirror action5:
        * never write to a cell that is blank in the original TSV (G1/G2) —
          EXCEPT one action5 filled because the specialist typed a new price
          into that empty cell (part2_summary.json action5.added), and only
          with exactly that price
        * never blank a cell — EXCEPT one action5 cleared because the
          specialist typed x over that price: those cells are listed in
          part2_summary.json (action5.removed; next to the xlsx, or
          --summary) and are cleared only if the TSV still holds that price
        * a requested x or new price the TSV no longer matches (the TSV changed
          since to-xlsx — e.g. re-fetched into the same dir) is NOT applied and
          the run exits 1: re-fetch into a fresh dir and re-run Part 2
        * never add or remove rows
        * shape must match (row/col counts) or the tab is skipped with a warning
      All other TSVs are left byte-identical, which is what wsheet wants.

Round-trip safety: `to-xlsx` then `to-tsv` with no edits is a no-op (0 cells
written). Verified in the smoke test at the bottom of the ACE skill notes.
"""

import argparse
import json
import sys
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
except ImportError:
    print("ERROR: openpyxl required. In ACE bootstrap a venv (no pip/root):")
    print("  python3 -m venv .venv && .venv/bin/python get-pip.py && \\")
    print("  .venv/bin/python -m pip install openpyxl")
    sys.exit(1)


def _read_tsv(path: Path) -> list:
    """Read a TSV into a list of row-lists of strings, verbatim.

    wsheet TSVs are raw tab-delimited: JSON root cells like {"id":1} are
    stored UNQUOTED and must not be run through csv quote-handling, or the
    quotes get doubled. Split strictly on the tab character and keep the
    trailing newline off each line.
    """
    text = path.read_text(encoding='utf-8')
    lines = text.split('\n')
    if lines and lines[-1] == '':
        lines.pop()  # drop the trailing empty element from a final newline
    return [line.split('\t') for line in lines]


def _write_tsv(path: Path, rows: list) -> None:
    """Write rows back as raw tab-joined lines (no quoting), trailing newline."""
    text = '\n'.join('\t'.join(cell for cell in row) for row in rows) + '\n'
    path.write_text(text, encoding='utf-8')


def _xlsx_safe(s: str) -> str:
    return ILLEGAL_CHARACTERS_RE.sub('', s)


def _rows_core(rows: list) -> int:
    """Row count without trailing all-empty rows (an extra blank line at the
    end of a TSV, which an xlsx does not keep)."""
    n = len(rows)
    while n and all(str(c).strip() == '' for c in rows[n - 1]):
        n -= 1
    return n


def _cell_str(v) -> str:
    """Normalize an openpyxl cell value to the string a TSV would hold."""
    if v is None:
        return ''
    return str(v)


# ── to-xlsx ────────────────────────────────────────────────────────

def to_xlsx(tsv_dir: str, out: str) -> None:
    d = Path(tsv_dir)
    tsvs = sorted(p for p in d.glob('*.tsv'))
    if not tsvs:
        print(f"ERROR: no .tsv files found in {tsv_dir}")
        sys.exit(1)

    wb = Workbook()
    wb.remove(wb.active)  # drop default sheet
    for p in tsvs:
        tab = p.stem
        # Excel sheet-name limit is 31 chars; truncate defensively
        ws = wb.create_sheet(title=tab[:31])
        for row in _read_tsv(p):
            # write each cell as an explicit string so JSON ids / prices
            # are never coerced to numbers/dates; a control character an xlsx
            # cannot hold is dropped here (to-tsv then treats that cell as
            # unchanged, so the TSV keeps its original bytes)
            ws.append([(_xlsx_safe(c) if c != '' else None) for c in row])

    wb.save(out)
    wb.close()
    print(f"to-xlsx: {len(tsvs)} tabs -> {out}")


# ── to-tsv ─────────────────────────────────────────────────────────

def _price_of(text: str):
    try:
        return round(float(str(text).strip().replace('$', '').replace(',', '')), 2)
    except ValueError:
        return None


def _code_norm(val) -> str:
    """A dish code as action5 reads it (its _norm_code): every run of spaces /
    no-break spaces → one space, none at the ends."""
    return ' '.join(str(val if val is not None else '').split())


def _action5_cells(final_xlsx: str, summary_path: str, key: str) -> list:
    p = Path(summary_path) if summary_path else Path(final_xlsx).resolve().parent / 'part2_summary.json'
    if not p.is_file():
        return []
    try:
        data = json.loads(p.read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return []
    return [c for c in (((data.get('action5') or {}).get(key) or {}).get('cells') or [])
            if c.get('tab') and c.get('row') and c.get('col')]


def _requested_blanks(final_xlsx: str, summary_path: str = None) -> dict:
    """{(tab, row, col): old price} — the cells action5 cleared because the
    specialist typed x over that price (part2_summary.json action5.removed).
    These are the ONLY cells to-tsv may blank. Rows/cols are 1-based."""
    return {(c['tab'], int(c['row']), int(c['col'])):
            (('code', _code_norm(c.get('old'))) if c.get('kind') == 'code' else _price_of(c.get('old')))
            for c in _action5_cells(final_xlsx, summary_path, 'removed')}


def _requested_fills(final_xlsx: str, summary_path: str = None) -> dict:
    """{(tab, row, col): new price text} — the empty cells action5 filled
    because the specialist typed a new price there (action5.added). These are
    the ONLY empty TSV cells to-tsv may write. Rows/cols are 1-based."""
    return {(c['tab'], int(c['row']), int(c['col'])): str(c.get('new', ''))
            for c in _action5_cells(final_xlsx, summary_path, 'added')}


def to_tsv(final_xlsx: str, tsv_dir: str, tabs: list, summary_path: str = None) -> None:
    wb = load_workbook(final_xlsx, read_only=True, data_only=True)
    d = Path(tsv_dir)
    total_cells = 0
    skipped_tabs = []
    stale = []          # requested x / new prices the TSV no longer matches
    blanks = _requested_blanks(final_xlsx, summary_path)
    fills = _requested_fills(final_xlsx, summary_path)

    for tab in tabs:
        if tab not in wb.sheetnames:
            print(f"  WARN: '{tab}' not in xlsx, skipping")
            continue
        tsv_path = d / f"{tab}.tsv"
        if not tsv_path.exists():
            print(f"  WARN: {tsv_path} missing, skipping")
            continue

        orig = _read_tsv(tsv_path)
        xlsx_rows = [[_cell_str(c) for c in row]
                     for row in wb[tab].iter_rows(values_only=True)]

        if _rows_core(xlsx_rows) != _rows_core(orig):
            print(f"  ERROR: '{tab}' row count differs "
                  f"(tsv={_rows_core(orig)} xlsx={_rows_core(xlsx_rows)}) — tab NOT patched (G2)")
            skipped_tabs.append(tab)
            continue
        xlsx_rows = xlsx_rows[:_rows_core(orig)]

        changed = 0
        cleared = 0
        filled = 0
        for ri, (orow, xrow) in enumerate(zip(orig, xlsx_rows)):
            # a short (ragged) TSV row: room for a requested new price past its end
            for (t_, r_, c_), want in fills.items():
                if t_ == tab and r_ == ri + 1 and c_ > len(orow):
                    if c_ <= len(xrow) and xrow[c_ - 1].strip() == want:
                        orow.extend([''] * (c_ - len(orow)))
                    else:
                        stale.append(f"{tab} row {ri + 1} col {c_}: new price {want} not written — "
                                     f"the xlsx now holds {(xrow[c_ - 1] if c_ <= len(xrow) else '').strip() or 'nothing'!r}")
            for ci in range(len(orow)):
                old = orow[ci]
                new = xrow[ci] if ci < len(xrow) else ''
                if old.strip() == '':
                    # G1/G2: never write a blank cell — unless action5 put a new
                    # price there on the specialist's request, and it is exactly
                    # that price
                    want = fills.get((tab, ri + 1, ci + 1))
                    if want and new.strip() == want:
                        orow[ci] = new
                        changed += 1
                        filled += 1
                    elif want:
                        stale.append(f"{tab} row {ri + 1} col {ci + 1}: new price {want} not written — "
                                     f"the xlsx now holds {new.strip() or 'nothing'!r}")
                    continue
                if new.strip() == '':
                    # don't blank an existing cell — unless action5 cleared it on
                    # an explicit x AND the TSV still holds that same price
                    want = blanks.get((tab, ri + 1, ci + 1), 'no')
                    is_code = isinstance(want, tuple)        # a dish code: compared as text
                    if want != 'no' and (_code_norm(old) == want[1] if is_code else
                                         (_price_of(old) is not None and _price_of(old) == want)):
                        orow[ci] = ''
                        changed += 1
                        cleared += 1
                    elif want != 'no':
                        stale.append(f"{tab} row {ri + 1} col {ci + 1}: x not applied — the TSV now "
                                     f"holds {old!r}, not the {'code' if is_code else 'price'} that was removed")
                    continue
                if old == new or _xlsx_safe(old) == new:
                    continue          # unchanged (or only a dropped control character)
                if (tab, ri + 1, ci + 1) in fills:
                    # action5 filled a cell that was EMPTY; the TSV has a price there now
                    stale.append(f"{tab} row {ri + 1} col {ci + 1}: new price {new!r} not written — the "
                                 f"TSV now holds {old!r} (it was empty when Part 2 read it)")
                    continue
                orow[ci] = new
                changed += 1

        if changed:
            _write_tsv(tsv_path, orig)
        print(f"  {tab}: {changed} cells patched"
              + (f" ({cleared} cleared on request — x in the report)" if cleared else '')
              + (f" ({filled} new price(s) in empty cells, on request)" if filled else ''))
        total_cells += changed

    wb.close()
    print(f"to-tsv: {total_cells} cells written across {len(tabs)} tab(s)")
    if stale:
        print(f"ERROR: {len(stale)} requested change(s) NOT applied — the TSV dir changed since "
              f"to-xlsx (re-fetched into the same dir?):")
        for line in stale[:20]:
            print(f"    {line}")
        print("  Re-fetch into a FRESH dir and re-run Part 2 (do not stage this draft).")
    if skipped_tabs:
        print(f"ERROR: {', '.join(skipped_tabs)} NOT patched — the draft would miss those changes. "
              f"Re-fetch and re-run Part 2 (do not stage this draft).")
    if stale or skipped_tabs:
        sys.exit(1)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)

    a = sub.add_parser('to-xlsx', help='wsheet TSV dir -> Raw-Gsheet .xlsx')
    a.add_argument('tsv_dir')
    a.add_argument('-o', '--output', required=True)

    b = sub.add_parser('to-tsv', help='patch wsheet TSV dir from a Raw-Gsheet .xlsx')
    b.add_argument('final_xlsx')
    b.add_argument('tsv_dir')
    b.add_argument('--tabs', default='dish_item,modifier_item,modifier_rule',
                   help='comma-separated tabs to patch (default: dish_item,modifier_item,modifier_rule '
                        '— an untouched tab writes 0 cells)')
    b.add_argument('--summary', default=None,
                   help='part2_summary.json listing the x removals / new prices (default: next to final_xlsx)')

    args = ap.parse_args()
    if args.cmd == 'to-xlsx':
        to_xlsx(args.tsv_dir, args.output)
    else:
        to_tsv(args.final_xlsx, args.tsv_dir, [t.strip() for t in args.tabs.split(',') if t.strip()],
               args.summary)


if __name__ == '__main__':
    main()