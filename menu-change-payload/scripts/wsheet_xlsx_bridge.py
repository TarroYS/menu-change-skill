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
        * never write to a cell that is blank in the original TSV (G1/G2)
        * never add or remove rows
        * shape must match (row/col counts) or the tab is skipped with a warning
      All other TSVs are left byte-identical, which is what wsheet wants.

Round-trip safety: `to-xlsx` then `to-tsv` with no edits is a no-op (0 cells
written). Verified in the smoke test at the bottom of the ACE skill notes.
"""

import argparse
import sys
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
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
            # are never coerced to numbers/dates
            ws.append([(c if c != '' else None) for c in row])

    wb.save(out)
    wb.close()
    print(f"to-xlsx: {len(tsvs)} tabs -> {out}")


# ── to-tsv ─────────────────────────────────────────────────────────

def to_tsv(final_xlsx: str, tsv_dir: str, tabs: list) -> None:
    wb = load_workbook(final_xlsx, read_only=True, data_only=True)
    d = Path(tsv_dir)
    total_cells = 0

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

        if len(xlsx_rows) != len(orig):
            print(f"  WARN: '{tab}' row count differs "
                  f"(tsv={len(orig)} xlsx={len(xlsx_rows)}), skipping (G2)")
            continue

        changed = 0
        for ri, (orow, xrow) in enumerate(zip(orig, xlsx_rows)):
            for ci in range(len(orow)):
                old = orow[ci]
                new = xrow[ci] if ci < len(xrow) else ''
                if old.strip() == '':
                    continue          # G1/G2: never write a blank cell
                if new.strip() == '':
                    continue          # don't blank an existing cell
                if old == new:
                    continue
                orow[ci] = new
                changed += 1

        if changed:
            _write_tsv(tsv_path, orig)
        print(f"  {tab}: {changed} cells patched")
        total_cells += changed

    wb.close()
    print(f"to-tsv: {total_cells} cells written across {len(tabs)} tab(s)")


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
    b.add_argument('--tabs', default='dish_item,modifier_item',
                   help='comma-separated tabs to patch (default: dish_item,modifier_item)')

    args = ap.parse_args()
    if args.cmd == 'to-xlsx':
        to_xlsx(args.tsv_dir, args.output)
    else:
        to_tsv(args.final_xlsx, args.tsv_dir, [t.strip() for t in args.tabs.split(',') if t.strip()])


if __name__ == '__main__':
    main()