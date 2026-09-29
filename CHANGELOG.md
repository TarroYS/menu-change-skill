# Changelog

## v4.3 — Party Tray now updates (P1), carried to production by P2 unchanged

**Before:** Party Tray prices never changed. action2 discarded the paper tray price and backfilled
the system price; action3 overwrote the tray column with the system price and skipped it in the
compare, coverage, and P3; price_inference never inferred it. The only route to production was a
specialist hand-typing a value or `=G23*5` formula after the callback.

**Now:**
- **Printed trays** are compared like any variant (blue / pink / green) and written by Part 2.
- **Unprinted trays** — new `price_inference.py` tier **T-TRAY**. Learns the restaurant's tray formula
  (`Tray = Base × k + d`, e.g. `Large × 5 − $2.00`) from the system's old prices. Applied (purple)
  only when ONE formula fits ≥ 70% of tray rows, and only to rows that follow it, off a base price
  that was printed. Outliers, ambiguous fits, inferred bases, decreases → callback, not written.
- **Questionnaire:** static `A-1 Party Tray pricing — confirm or update?` replaced by a data-driven
  **P-TRAY** section (formula + per-tray rows). Remaining standing questions renumbered A-1..A-3.
- **ocr_sanity_check:** whole-dollar Party Tray prices ≥ $100 no longer flag `DROPPED DEC`.
- **Guard:** a paper tray price is written only to the tray column, never rerouted to another slot.
- **Part 2:** no code change — action4/action5 were already header-driven.

New flags (price_inference.py): `--tray-min-agreement 0.70`, `--tray-min-rows 3`,
`--tray-max-offset 10.00`, `--no-tray`.

**Regression:** restaurant with no tray prices → OCR output, Dish Item, and final raw byte-identical
to v4.2 (the Menu Preview always carries an empty `Party Tray` header; T-TRAY skips it).

**Design note:** T-TRAY is the one sanctioned exception to "no cross-variant inference" (G14). It
copies no delta; it fits a multiplicative formula gated on restaurant-wide evidence. Do not extend
the exception to other variant pairs.
