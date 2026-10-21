# Changelog

## v4.5 — Fried Specialty / Modifier Rule tabs (2026-10-16)

The maintainer's request, with two decisions:
- the specialist types into the **difference** table (and an absolute table for absolute dishes), not the full-price
  one;
- only restaurants that price the sides **per dish** get the new tabs.

**Both tabs whenever Fried Specialty dishes are priced per dish on their sides.**
- compare finds the pattern from the system alone: dishes of a "Fried Specialt…" dish group whose own rules price
  the items of a "Select…" modifier group (Select Rice: By Itself, French Fries, Pork Fried Rice …).
- The Modifier Rule and Fried Specialty tabs are built when the paper prints:
  - no Fried Specialty table (every price stays today's unless the specialist types a new one);
  - only some of the dishes (every dish is listed, printed or not — also one on "Order", not "Specialty").
- 44 restaurants in the real reports price this way; China Village (9 dishes, absolute) is one.
- Restaurants that price Select Rice once for every dish (on the item itself — Modifier Summary: R85325, R83807,
  ER87669) are unchanged.
- Part 1 now runs compare every time; it says "Nothing to do" when there is nothing.

**The Fried Specialty tab, rebuilt.** One row per dish and size, one column per side (Select Rice order):
- **FULL PRICE** — the paper menu, for reference; grey = not printed (today's price).
- **PRICE DIFFERENCE** — the dishes priced by difference.
  - Type the side price minus Plain: plain numbers, no + sign (a + is accepted).
  - Plain = the dish's Dish Item price, linked (edit it there).
- **ABSOLUTE PRICE** — the absolute dishes. Plain = the By Itself rule; each side its full price.
- **NOTES → WONKA** — a column on every row, and a box under the tables, for instructions or formulas.
- Then how the paper was read (mappings, propagation, callback flags), as before.

Grey cells = that size has no rule for that side. Same colours as the Modifier Rule tab: blue / pink / green from
Part 1, orange = an edit, red = not applied. Questionnaire **P-FS** lists the dishes the paper does not print, to ask
the client.

**Part 2 reads the tables.**
- Each cell is one rule. An edit (≠ its hidden value) is written, and beats an untouched Part 1 value.
- The same rule edited differently on the Modifier Rule tab and the Fried Specialty tab is not applied; it is
  reported (`tab_conflicts`, part2_report).
- Refused and reported: not a price; x; a misplaced decimal point; a value in a grey cell; a typed-over Plain of a
  difference-priced dish.
- Rows are found by a hidden key, so a sorted table is safe.
- The tab's notes go to Wonka with the Questionnaire notes (specialist_requests → part2_summary.json →
  part2_report). part2.md: a price note or formula is carried out by editing the tables and re-running Part 2.

**Independent review (same day): 10 problems found, all fixed.**
- **High — an unprinted dish's side prices could change while P-FS said "today's".** The price-twin rule (as in
  v4.3) copies the printed dishes' same side change to an unprinted dish. Such cells are now purple (not printed —
  confirm), and P-FS names those dishes as filled. Part 2 still writes them unless the specialist changes them.
- A value refused on the Fried Specialty tab (not a price, x, misplaced decimal) holds its rule, like the same input
  on the Modifier Rule tab. Before, Part 1's untouched change was still written.
- The Plain of an absolute size binds to its priced By Itself rule, not the first alias in the system ($0 No Rice).
- A blank row inside a table no longer ends it; edits and notes below it were lost.
- With the Modifier Rule tab deleted, the Fried Specialty tables are still applied (they were ignored).
- The colours stay right after a quick sort: each rule looks its row's hidden values up by dish and size.
- A value typed over a difference row's Plain (a link to Dish Item) shows red, as Part 2 refuses it.
- compare refuses to rebuild tabs that already hold the specialist's edits or notes (`--rebuild` to force). A re-run
  used to wipe them silently.
- No false DRIFT warning when the paper has no side table.
- An optional add-on group called "Select …" is not a side choice. A side choice holds a plain option (By Itself /
  No Rice) or is a required pick.
- Smaller: P-FS names the dishes' own groups; a size with no priced side is left out; a rule for every size of a
  multi-size dish shows no single full price.

**Checked.**
- New suite `tests/fried_specialty` (30 checks):
  - scenario A, plus a blank row and a deleted Modifier Rule tab;
  - scenario B with edits, refusals, a sort (values and colours), a conflict, notes and the re-run guard;
  - the twin fill;
  - a By-Itself-only dish;
  - the shared-price restaurant unchanged.
- Every check fails when its piece is reverted (34 single reverts).
- The whole suite passes.
- China Village with its paper and without its side table: both tabs, P-FS, unedited Part 2 writes 0 cells.

## v4.4 (cont. 4) — dish codes, Party Tray per dish group, the maintainer's answers (2026-10-05)

**Dish codes are now part of the menu change** (maintainer's decisions DC1–DC5, HANDOFF §5c).
- **Part 1.** The report has a **Dish Code** tab, row-aligned with Dish Item:
  - the left block holds each size's current code (`unit_code`, editable);
  - the right block (`<size> · paper`) holds the code printed on the paper: blue = differs from the
    system, yellow = the system has none, green = the same;
  - sizes without a price are grey (a code needs a price).

  Questionnaire **P-CODE** points at the tab, and lists any size the paper printed two different codes
  for.
- **Where a printed code goes.** A code goes to the system size its price landed in. One number on a
  two-size row → both sizes (DC2); an `L26` in the Lunch section → Lunch only, a `D26` → Dinner
  (DC5). Two different codes for one size are asked, not proposed.
- **Part 2 writes a code only where the specialist put one in a code cell** (DC1: fills and
  corrections alike are confirmed by a person, per Wonka's brief). The paper block is never read.
  Same conventions as prices: copy a paper code over to accept it (a whole paper column pastes
  safely), `x` removes a code, a code typed into an empty cell adds one.
- **Part 2 guards.** Each refusal is reported:
  - not a code: a line break, more than 128 characters, or only `# - .` and spaces (red in the tab);
  - a code on a size with no price;
  - a code changed in wsheet since Part 1 (code drift);
  - two copies of a dish edited to different codes (conflict).

  A dish under two groups is one dish (the edited copy wins). Sorting the tab is safe (hidden row
  keys). to-tsv fills / clears a code cell only while the TSV still holds what Part 2 saw.
- **DC4:** a code left on a moved / removed price is always asked (`code_followups`), never moved
  automatically. The old automatic overwrite path (off since it made mistakes) is gone. Dish-level
  codes (`name_complex.code`) are not touched.
- action1 writes the system codes as a `System Codes` sheet. Deliberately not `dish_code`: that
  name would switch on action2's dormant code tie-break in name matching.

**Party Tray formula per dish group (B: "usually one formula, sometimes per group — be safe").**
- T-TRAY finds a formula for each dish group from that group's own trays.
- A group too small or too mixed to confirm its own uses the whole-menu formula only when that fits
  ≥ 70% of its own trays; otherwise its changed trays are asked.
- A dish under two groups whose formulas disagree is not derived.
- P-TRAY-0, the Price Inference tab and the Summary list one formula per group (the tab used to
  overwrite all but one line).
- On the R99999 fixture the Beef group (3 trays, one an outlier) no longer derives Pepper Steak /
  Mongolian Beef. They are asked instead — the only change to the Party Tray final raw vs v4.3.

**G: a report that lost its hidden data stops, with a clear message.** The 5 real reports that
action4 stopped on have lost the hidden Modifier tab AND every hidden `_ref_` column (re-saved by
another app). action4 now says so and tells the specialist to re-make the report with Part 1.

**Recorded answers:** H hold, I yes, J keep, F (= DC4) always ask, G keep stopping, DC1–DC5, K two printed
prices for one size → callback (HANDOFF §6).

**Two different prices printed for one size are asked, never picked** (maintainer, same day: "be safe —
require callback"). Until now the last one read was written. Cases:
- a dish printed in two sections at different prices;
- an Order price routed onto a Large that was also printed;
- a Lunch-based menu's Small routed onto a lunch special's Lunch price.

Neither price is written, so Dish Item keeps the system price. Questionnaire **P-PRICE** names both prices
and where each was printed. price_inference never guesses or uses that cell, and never derives a Party
Tray over it. A tray whose base price is the conflicted one is asked too. The same price printed twice is
no conflict. A $0 dish priced on modifiers stays under P-ABS, which then names both prices. China Village
has no such size, so its report is unchanged.

**Independent review of the above (same day): 7 problems found, all fixed.**
- **Pasting a whole paper column lost its codes (high).** Copying `Order · paper` by its column letter
  also pastes the header, so Part 2 could not find the `Order` column. Its codes were silently dropped.
  - A code column is now found by its header, or by its place when its paper header was pasted over it.
  - A code column found neither way stops Part 2 with a message. Nothing is written.
- **Trays: a dish in two groups is derived only when every group derives the same tray.** Before, a
  group whose formula did not fit the dish, had no formula, or said "no change" did not block the
  other group's tray, so it was still written. Now such a dish is asked. A tray that one group holds
  back (suggest-only) is held back for the whole dish.
- **A code starting with `=`.** A formula in a code cell now counts by its saved value (action4, as on
  Dish Item). Text starting with `=` is not a code (red), so it can never reach wsheet as a formula.
- **Code follow-ups (DC4)** list only codes still sitting on the removed / moved size. A code the
  specialist already removed or moved on the Dish Code tab is no longer listed.
- **What the paper proposes.**
  - A number at the start of a printed name counts as a code only when it has a letter (`A5`) or a
    `.` / `)` after it (`12.`). `10 Fried Dumplings` is a count, not code 10.
  - A lunch-special code never lands on a dish's other sizes.
  - Two codes reaching one size through different paper sizes are asked, not proposed. Examples:
    `15` printed with Order (→ Large) next to `S7` printed with Large; row `26` routed Small→Lunch next
    to lunch special `L26`. If the paper's only codes are such conflicts, the tab is still made, so the
    P-CODE question has somewhere to be answered.
- **The sheet, action5 and to-tsv read a code the same way.**
  - Spaces: no-break spaces count as spaces, runs of spaces become one, none at the ends (Excel TRIM).
  - Text `1.0` stays `1.0`; only a number cell `38.0` reads as `38`.
  - Paper codes are compared in any script (`春1` ≠ `夏1`).
  - An untouched system code `X` is no longer styled as "remove".
  - A code typed on a grey (no price) cell is red. Part 2 refuses it unless a price for that size is
    added on Dish Item too.
  - An `x` on an empty code cell is reported as "no code to remove".

**Checked.** New suite `tests/codes` (34 checks, each mutation-checked: 22 single-fix reverts for the
review fixes, all caught); per-group tray and two-price checks in `tests/audit` (the two-price rule: 7 single-fix reverts, all caught);
332 checks in all (`tests/run_all.sh`). Real data:
- China Village fresh Part 1 + unedited Part 2 still writes 0 cells (269 confirmed); its Dish Code
  tab shows all 340 system codes.
- The 4 real report pairs are unchanged.
- The Party Tray final raw differs from v4.3 only in the two Beef trays above.

## v4.4 (cont. 3) — self-audit: bugs found and fixed (2026-09-30, night)

Three independent reviews of the v4.4 build, a LibreOffice evaluation of the report's sheet rules,
mutation testing of the test suites, then two more reviews of the fixes themselves (their 16
findings are under "Second review round" below). Every fix has a regression check in the new
`tests/audit` suite (or a strengthened check in an existing suite) that fails when the fix is
reverted.

**Part 2 — prices written that nobody asked for, or asked-for changes lost**
- **Part 1 changes were read from the cell colour.** A specialist who recoloured cells (or sorted the
  sheet) silently dropped Part 1's changes: in one repro 0 of 14 were written. Dish Item now has a
  hidden `_ref_sys_<size>` mirror of the system price; "Part 1 changed it" = `_ref_` ≠ `_ref_sys_`.
  Modifier Rule has a hidden `_ref_live` (the live rule price at Part 1). The colour is only the
  fallback for reports made before this.
- **Sorting the report mixed up the baselines.** The hidden `_ref_` block started three columns after
  the data; Excel's quick sort stops at the gap, so the visible rows moved and the baselines did not.
  The hidden block now touches the data, and every row carries its own key (`_ref_key:group` /
  `_ref_key:name`; Modifier Rule `_ref_key`). Part 2 finds a row's baselines by its key.
- **Modifier Rule prices were applied unchecked.** `1,50` became $150, `1395` $1395, `-1`, `1.5%` and
  `x` were all written. They are now ignored and reported, like on the other tabs. A formula counts by
  its saved result; one with no saved result is reported.
- **A misplaced decimal point under $10 shipped.** `195` typed for 1.95 is below the 1000 cap, so
  $195 was written. Now held on every tab, and red in the report:
  - an edit (or Part 1's own price) that is 5× or a fifth of the SYSTEM price;
  - a new price in an empty cell at 10× or a tenth of the item's other prices (Dish Item trays
    exempt).

  See "Second review round" for why the system price is the baseline.
- **A cell formatted as a percentage** (Excel stores a typed `10%` as 0.1) shipped as $0.10. Now
  invalid.
- **`0` or `x` typed into an EMPTY report cell** removed or zeroed a price wsheet had added since
  Part 1, and a new price overwrote it. These cells are now judged by what the specialist saw: a
  since-added live price is never touched (`new_cell_drift`, says "live now has $X").
- **Negative modifier prices were rejected** ("Less Whole Wing" −$1.70; 7 real restaurants). They
  are valid on Modifier Summary and Modifier Rule; a negative dish price stays invalid.
- **Modifier Summary formulas were ignored**; they now count by their saved value.
- **Same-name live rows were written** (an availability row repeating a dish's root, a same-name
  modifier twin). Only the row the report shows is written; the others keep their prices and are
  listed (`twin_rows_skipped`).
- **Silent drops now reported:** an edit in a size the live menu does not have; `x` on a modifier
  whose live cell is already empty; an edited rule whose live rule has no price (`no_live_price`).
- **`x` on a dish's last price** is allowed but flagged (`no_price_left`): taking a dish off the menu
  is a strikethrough.
- **G11 in apply:** By Itself of a $0.00 (ABSOLUTE) dish set to 0 is held — the plain dish would be
  free. This includes a By Itself rule with no size; the second round widens it to every rule that
  is a dish's full price.
- **Wrong order and re-runs are refused or flagged:** action5 on a report that skipped action4 →
  exit 1 (Modifier Summary edits would be lost). action4 on its own output → exit 1. A second action5
  run in the same directory → warning. part2_report with no action5 result → exit 2, never zeros. A
  previous run's modifier_rule counts are dropped so part2_report never shows them. A report with no
  hidden baselines runs with a loud warning.
- **Bridge:** a control character in a TSV crashed to-xlsx. A trailing blank TSV line made to-tsv
  skip the whole tab, and a skipped tab still exited 0. Now a tab that cannot be patched exits 1, and
  the default `--tabs` include modifier_rule. New: an `x` or a new price the TSV no longer matches
  (the dir was re-fetched since to-xlsx) is not applied, and to-tsv exits 1 ("re-fetch into a FRESH
  dir").
- **Modifier Summary sizes:** the modifier tab's own name columns ('Modifier Group', 'Modifier
  Name:Primary') were treated as price sizes; a column the specialist added was too. Sizes are now the
  columns Part 1 made (with a `_ref_`). The group / item columns are found by header.
- **Drift wording:** "the live cell has no price any more" was printed for every drift, including a
  new price typed where wsheet has since added one. Each line now says which case it is.
- **Rounding like the sheet.** Part 2 rounded typed prices with Python's `round()`, which works on
  the binary value: 3.675 (shown 3.68 in the report) was written $3.67. All three parsers now round
  half away from zero on the decimal value, like Excel / Google Sheets.
- **Modifier Rule tab: red text for what apply ignores** (it had none, so a typo showed orange as an
  edit and was then held). Same rules as apply: not a number (incl. `x`), characters a price never
  has, 1000+, 5× / a fifth of the live rule price; negatives are fine. LibreOffice: 58 cases, 0
  mismatches. Its orange "edited" rule was stored with a leading `=` since v4.3, which Excel drops
  when it opens the file; fixed.

**Part 1 — the $0 guard, G11, compare**
- **A group-level add-on made DELTA $0 dishes "ABSOLUTE".** "Sauce / With +$1.00" on the Fried
  Specialties group (real R85325, R83807, China Village) held their $0.00 tray placeholders, so the
  paper's tray price was never compared. Group-level rules now count only when a priced By Itself is
  among them, or they price a required choice (added in the second round, below). The Fried
  Specialties name rule needs "some live $0.00 dish has priced rules holding its price". action1 and
  action5 make the same test, including a dish in no dish group (Party Platter).
- **G11 missed a priced group-level By Itself rule on a priced dish** (a double charge). preflight
  and apply now block it.
- **compare:** printed side prices on a dish without rules of its own were dropped silently or
  labelled "⚠ size?". Now:
  - a callback flag `NO DISH RULES` when an enabled, priced group-level rule of its group prices one
    of those sides;
  - otherwise a `NO SIDE RULES` note. China Village keeps 0 callback conflicts.
- **compare re-runs** pushed the Fried Specialty block down one row each time; the 'by itself' link
  was missing for a rule with no size (it now links when the dish has one priced size); Expected /
  Match? were not blank-safe in Google Sheets.
- **side_rules:** a single add-on ("Spicy +$1.50") was listed as a side table, so action2 warned
  about the wrong dishes. Only a By Itself or a real side counts, dish-level or group-level (see
  the sushi item in the second round).
- **The report's red-text rules now match Part 2 exactly.** There are 10 rules per price column:
  text that is not a number, characters a price never has, out of range after rounding, 0 typed into
  an empty cell, 5× / a fifth of the system price, and a new price at 10× / a tenth of the item's
  other prices. Each comes in number and text forms; Dish Item trays are exempt from the new-price
  rule, as in action5. Every formula is IFERROR-safe and under 255 characters.

  LibreOffice evaluated 140 cases per tab, 72 new-price cases and 58 Modifier Rule cases in en-US:
  0 mismatches against Part 2's parsers. The apricot "edited" rule compares numbers as numbers (text
  "9.45" over 9.45 is no edit).

**Second review round (two reviews of the fixes above; each finding reproduced, fixed, pinned)**
- **$0 guard regression (high).** After the add-on fix, a $0.00 dish whose full price sits on
  group-level rules of a REQUIRED choice ("Select Side: French Fries $9.95 / Fried Rice $10.45", no
  By Itself) was no longer held, so the paper price went onto the dish: a double charge. Priced rules
  now hold a $0.00 dish's price when a By Itself is among them OR they price a required choice; an
  optional add-on is still an upcharge. Same test in action1, action5 and "the restaurant prices on
  modifier rules".
- **G11 in apply (high).** apply could still set an ABSOLUTE rule to $0.00 three ways: a group-level
  By Itself, a size whose dish_item cell is blank, a full-price side rule (wings + fries free). Any
  rule going to $0 is now held when it is the full price of a dish with no dish price. A DELTA
  upcharge set to 0 (a free side) is still applied.
- **Misplaced decimal vs an OCR misread.** The check compared with the report's price. So when the
  paper read 19.50 for 1.95, the specialist's correction to 1.95 was rejected, and the untouched
  misread was written. Edits and Part 1 changes are now judged against the SYSTEM price (hidden
  `_ref_sys_`; Modifier Rule `_ref_live`):
  - the band is 5× / a fifth, which catches a slipped decimal whether the price rose or fell (0.95
    for 9.50 when the system had 9.25);
  - an untouched Part 1 price that far from the system's is held as a likely misread;
  - none of 9,943 real edits, and none of 708 real Part 1 vs live differences, come within 3×.

  New prices are judged against the item's other prices as the report shows them, now in both
  directions.
- **Double charge through a new size.** action5 added a size on top of a group-level By Itself
  rule; without a Modifier Rule tab no preflight runs. It is now held.
- **A new Modifier Summary price overwrote a price wsheet added since Part 1.** It is now drift, not
  written (as on Dish Item).
- **Negative prices** were written `-2.00` / `$-2.50`; now `-$2.00` / `-$2.50`, as wsheet holds them.
- **Silent drops, now asked:**
  - action2 dropped a printed price for a size the dish does not have (the dish has two or more
    other sizes). It is now Questionnaire **P-SIZE**: the dish's sizes are listed, and the specialist
    edits the matching size or adds the new one. It is skipped when it equals the By Itself price of
    the dish's $0.00 size (China Village's rib tips), and price_inference counts the dish as on the
    paper.
  - compare dropped a printed side the dish has no rule for. It is now `NO RULE FOR SIDE`, naming
    the group rule if one prices it.
- **Healthy rule:** the "a P1 row looks like it" guard never matched (action2 writes `Maybe: X
  (substring)`).
- **False warning on sushi menus (real ER87669).** Roll toppings counted as a side table. A side
  table now needs a By Itself or a real side (fries / rice / lo mein / chow mein / mei fun /
  noodles).
- **Low:**
  - a `--no-tray` re-run left the old P-TRAY block;
  - the P-ABS pointer missed a By Itself rule with no size;
  - a price written as text in the OCR JSON ("$9.59", "MP") crashed action2 / action3 / compare.
    Prices are now normalised at load, and anything that is not a number is reported;
  - to-tsv lost a requested new price past the end of a short TSV row;
  - part2.md §4b said to re-run in the same directory (now: re-fetch into a fresh one).

**price_inference**
- "(White Meat)" dishes joined the plain General Tso / Orange / Sesame family and took the plain
  price. Families are now keyed by the bracketed qualifier.
- A dish under two dish groups could get different inferences on its two rows, and either could
  ship. The rows must now agree, or no row is inferred (`conflict (same dish, several groups)`, red).
- The Small < Large check used the OLD printed prices; it now uses the new ones.
- A re-run kept the previous run's purple cells (even in `--mode suggest`). They are undone first
  (back to `_ref_sys_`), but never a cell the specialist typed over; a re-run never overwrites a
  specialist's value either.
- T-TRAY agreement counted rows, so a dish in two groups counted twice; it now counts dishes.
  Donors are deduplicated the same way. The "(Chow Mein + Chop Suey)" label appears only when the
  cohort really mixes the two.
- Healthy rule: a Healthy dish printed only as a tray, or under another name (a P1 row like "Diet
  Shrimp w/ Broccoli (Steamed)"), counts as on the paper and is not inferred.
- A stale report (rows shifted since action3) received inferences on the wrong dish; each write now
  checks the row's dish name. Summary counts are recomputed after writing.

**Tests.** New `tests/audit` (regression checks for every finding above) and `tests/lib/xlformula.py`
(a small Excel formula evaluator, so tests check what a spreadsheet shows, not formula text; it agrees
with LibreOffice on all 4,982 rule evaluations of the parity runs, and `cf_kind` tells which sheet
rule a cell gets).
Strengthened checks that could not fail before:
- GROUP CONFLICT must be in the conflicts section;
- P-TRAY-0 must give its reason;
- the Healthy no-twin case now has a near-miss regular dish;
- older-report notes are tested on a synthetic older layout (no `REAL_REPORT` needed);
- the edit rules expect 6 not-a-price rules.

Mutation check: each audit fix reverted one at a time → its check fails.

**Real data (unchanged results).** Party Tray final raw cell-identical to v4.3. China Village fresh
Part 1 + unedited Part 2: 0 cells changed in dish_item / modifier_item / modifier_rule, 269 confirmed,
0 callback conflicts. Real report pairs: R87669 / R85325 identical to v4.3; R81988 / R83807 differ
only where v4.3 reverted newer live prices. 217-report sweep: action4 identical to v4.3 on all 212
reports it can read (5 lack the hidden Modifier tab, as before).

## v4.4 (cont. 2) — new prices in empty cells, duplicates, live links, guards (2026-09-30)

**Part 2 writes only instructions (found on real data, critical).** action5 used to write EVERY report
price that differed from live, and `apply` every Modifier Rule row. So a price someone changed in
wsheet between Part 1 and Part 2 was silently put back to the report's snapshot. The earlier
"0 is a price" fix made it worse for modifiers: a $0.00 in the hidden Modifier tab overwrote a live
$11.95 (China Village, Family Dinner). Now only these are written:
- a specialist edit (cell ≠ its hidden `_ref_`, or a Modifier Summary edit resolved by action4);
- a price Part 1 changed (Dish Item blue / pink / purple; Modifier Rule blue / pink).

Untouched prices are never written; a live change made since Part 1 is kept and counted
(`not_reverted`). On two real restaurants' older reports plus current raws, v4.3 would have reverted
339 live prices (R83807: 117 Party Trays; China Village: 222 dish and modifier prices). Unedited Part 2
on a fresh China Village report still writes 0 cells.

**(c) A dish or modifier listed under two groups is ONE item.** action5 kept the LAST report row of a
dish, so an untouched lower copy put the old price back over an edit made on the upper copy. That
happened in 10 of the real edited reports in Downloads, 145 cells in all (R89386: 100). action4
replaced a modifier's whole set of edits with the last row's (edits on two rows were lost). Now each
size is resolved across all rows of the item:
- an edited copy wins over untouched copies;
- copies edited to different prices → CONFLICT, nothing applied, reported;
- untouched copies carrying different Part 1 changes → CONFLICT;
- Modifier Summary edits on two rows are merged per size.

**(a) New prices in empty cells, moves, stray keystrokes.**
- A price typed into an EMPTY Dish Item / Modifier Summary cell is written (vivid yellow in the
  report; G1's second exception). To move a price, paste it into the empty cell and type `x` over
  the old one; `x` now works in Modifier Summary too.
- A modifier's new price goes only to the live row the report shows, never to an unpriced same-name
  twin.
- Anything that is not a valid price is ignored and reported, and shows red text in the report:
  letters, comma, %, negative, 1000 or more, or 0 in an empty cell. `1395` (a missing decimal point)
  can never ship; no real report in Downloads has a price of 1000+.
- The bridge fills an empty TSV cell only with a new price action5 recorded (`action5.added`).
- A removed or moved size keeps its dish code (the scripts never write codes) → `code_followups`,
  which Wonka settles with the specialist in the same draft.

**(e) Modifier Rule 'by itself' is a live link.** DELTA rows point at the dish's Dish Item cell (the
edited copy when the dish is in two groups); ABSOLUTE rows point at its By Itself row's Price. An edit
there moves Expected / Match? (Match? is N/A without Expected). On China Village: 474 links, every one
points at the right cell; Match? TRUE on all 67 printed rows.

**(f) $0.00 Fried Specialties never take a dish price in ABSOLUTE restaurants.** The guard looked
only at modifier_rule rows naming the dish. Now it also covers:
- group-level rules (only a dish_group);
- any $0.00 Fried Specialties dish in a restaurant that prices dishes on modifier rules (the
  maintainer's rule), even one with no rules of its own.

This applies in action1 (absolute_pricing → P-ABS), action2 (held) and action5. action5 also never:
- adds a new size to such a dish;
- removes an ABSOLUTE size;
- removes or zeroes an option of a required group that prices a $0.00 dish (it would become free).

**(d) Fried Specialty tab triggers made visible.** The tab exists only when the OCR JSON has a
`side_pricing` section with side_prices items (action2 → action3). 89 real reports with a Fried
Specialties group have no tab, and 11 have a Modifier Rule tab without one.
- action1 writes a `side_rules` tab; action2 warns `⚠ NO MULTI-COLUMN TABLE IN THE OCR JSON` when
  side-rule dishes are read without a table.
- compare builds a missing Fried Specialty tab and prints `!! DRIFT` (action2/action3 ran on a
  different JSON).
- part1.md step 3b now keys off action2's `Multi-column sections passed: N`.

**Healthy rule:** only a Healthy dish NOT on the paper menu takes its regular dish's price (the
maintainer confirmed the two R81587 cells were missed by the specialist).

**Wonka:** struck dishes are always made unavailable in the same draft.

**Checked:** 5 test suites, 171 checks (`tests/run_all.sh`; new `tests/edits`, 47 checks; v4.3 5/34
there). The Party Tray final raw is cell-identical to v4.3. action4 is cell-identical to v4.3 on 212
real reports. Real Part 2 runs on R81988 / R87669 / R83807 / R85325: every difference from v4.3
explained (format, ABSOLUTE guard, not reverted).

## v4.4 (cont.) — the specialist's edit sheet + same-price inference rules (2026-09-30)

**Prices are plain numbers everywhere on the edit sheet.** Dish Item, Modifier Summary, Modifier
Rule, Fried Specialty, Price Inference and the Questionnaire's price columns hold numbers (`9.45`,
`0`) with a `0.00` format (`+0.00` for signed deltas) — never `$9.45` text, so the specialist types
only the number (no `'$`). Down the line: action4 / action5 / apply read numbers and `$` strings
alike; **action5 writes each price in the live cell's own format** (`$9.45`) and apply writes
`$X.XX`, so the final raw / TSV are unchanged (Party Tray final raw still all `$X.XX`); `0` is a real
price (action4 / action5 used to read a `0` as blank — fixed).

**`x` / `X` over a Dish Item price = remove it.** Dark grey (595959) + white bold text, ranked ahead
of the apricot edit colour. action5 clears that live cell — the only blank write (G1 exception) —
and lists it in `part2_summary.json` `action5.removed`; `bridge to-tsv` blanks exactly those cells
(and only while the TSV still holds the recorded old price), so no other blank can reach the draft.
`x` on an empty cell → ignored + reported; `x` in Modifier Summary → ignored with a warning.

**Strikethrough on a Dish Item dish name (column B) = make the dish permanently unavailable.** New
`specialist_requests.py` (Part 2 step 0b) lists every struck dish name (whole cell or any struck
part — a partial strike is flagged) into `part2_summary.json` and prints it. Wonka makes each one
permanently unavailable with its own tools (part2.md §4b; the scripts never add / remove rows);
part2_report shows the count. Specialists already do this: 7 of the real edited reports in the
maintainer's Downloads carry struck dish names (R86659: 9).

**Questionnaire: column I "Notes → Wonka" = instructions for Wonka.** Part 1 never writes it (light
yellow, header comment). specialist_requests.py lists every note with its row's ID / dish / prices /
Details; Wonka carries each out within the guardrails (a price → edit that report cell and re-run
action4 → action5, never by hand; never past G11 / a commit / a row delete) and reports done /
not done. A pre-v4.4 report (no Details column, its "Notes" is script text) → notes not read.

**Questionnaire restructured.** A–H are the script's (Question ID / Category / Dish Group / Dish Name /
Current / Expected / Menu Price / Details), I is the specialist's. **P1-TRAY** is the first P1 row
when the restaurant has Party Tray prices ("confirm or update?", moved up from Additional
Questions). **Removed:** P3 (in system, not on menu), Additional Questions, P5, P6 (inferences stay on
the Price Inference tab + purple cells). Unmatched P1 rows show the paper prices.

**P-ABS says exactly where to edit.** Dish Name `Dish [Size]`, Current `0`, Menu = paper price, and a
red Details line, hyperlinked: `EDIT → 'Modifier Rule' tab · row N · column B (out_unit_price) — the
By Itself row of Dish [Size]` (rows found per size by compare), or `EDIT → 'Modifier Summary' tab ·
rows a–b (group 'G') · column 'Size' (letter)` for a dish priced by a required choice. No such row
/ tab → "write the price in Notes → Wonka on this row". (action1's `absolute_pricing` and action2's
`Absolute Held` tabs now carry the Modifier Group, so the pointer can find the choice's rows.)

**price_inference — maintainer same-price rules (T0, tried before T1; G14-sanctioned):**
- General Tso / Orange / Sesame `<Chicken | Beef | Shrimp>` always the same price — an unprinted one
  takes its printed siblings' price whatever its old price (names normalized: "General Tsos",
  "Gen. Tso's").
- A Healthy "Steam X" takes the regular "X"'s new price when both OLD prices were the same (per size).
- Chow Mein and Chop Suey are one T1 cohort (same old price → same new price).
- Printed family members that disagree → `T0 conflict` (red), not guessed, not retried by T1.
- On R81587 New Choi Hee (the maintainer's edited example): the rules reproduce all 22 Chop Suey
  hand edits and Orange Chicken Lunch / Combo. Two Healthy Small cells differ from the specialist's
  hand edits — Steam Mixed Vegetable 7.00 → 7.25 and Steam Pork with Broccoli 8.00 → 8.50 are
  inferred where the specialist kept the old price (maintainer to confirm).

**Checked:** 4 test suites (`tests/run_all.sh`: party_tray 19, modifier_rule 50 + 3 real-raw,
price_inference rules 14, report 34 + 2 older-report); a sweep of 217 real reports in Downloads —
specialist_requests exits 0 on all, action4's Final Review is cell-identical to v4.3 on the 212 it
can read (5 lack the hidden 'Modifier' tab, v4.3 fails the same).

## v4.4 — Fried Specialty / modifier_rule: six fixes (issues #1–#6)

**#1 ABSOLUTE restaurants were double-charged (critical).** A $0.00 dish whose price lives on
modifiers took the paper Plain price like any other increase, so the customer paid it on top of the
side rule's full total (wings $7.99 + wings-with-fries rule $9.39 = **$17.38** for a $9.39 combo).
Validation hid it ("Paper menu Total" $17.38, Match? TRUE).
- **action1** writes an `absolute_pricing` tab: every $0.00 dish cell whose price lives on
  modifiers — priced dish-level modifier_rule rows (ABSOLUTE), or a required modifier group whose
  every item is priced (the R85325 Seafood Boil pattern).
- **action2** never writes a paper price onto those cells; it lists them in `Absolute Held`.
  **action3** asks about them in a new Questionnaire section **P-ABS** (they count as matched, never P3).
- **price_inference** never infers onto, or from, those cells (it used to suggest $7.99 / could apply
  a group delta onto the $0).
- **compare** (ABSOLUTE): the paper Plain goes on the dish's **By Itself** rule; validation shows the
  printed total, and Expected = the rule price itself. Blocks DOUBLE_CHARGE / BROKEN like preflight.
- **Part 2:** `preflight` and `apply` now run on **action5's `final_raw.xlsx`** (post-action5 state).
  `apply` refuses DOUBLE_CHARGE / BROKEN and any G11 violation in the resulting state (exit 1,
  nothing written). **action5** refuses to write any price onto a live $0.00 dish whose price lives
  on modifiers (`!! ABSOLUTE GUARD`, recorded in part2_summary.json, shown by part2_report).

**#2 Printed prices overwritten by price-group conflicts (high).** A printed row now always keeps
its own printed price. Unprinted members of an old-price group are propagated only when every
printed member agrees; otherwise they keep the current price + `GROUP CONFLICT` for the callback
(G14). "Paper menu Total" is the total actually printed (blank for derived rows → Match? N/A), so an
overwrite can no longer validate as TRUE. `--conflict` is accepted but ignored.

**#3 "Plain Fried Rice" falsely blocked DELTA restaurants.** By Itself anchors are matched on the
whole normalized name, not a substring (compare / preflight / apply; action1 uses the same rule).

**#4 Group-level rule edits were silently dropped in Part 2.** Modifier Rule rows are keyed by
(dish_group, dish_item, variant, modifier_item), so a group-level rule (no dish_item) is applied by
its dish_group. A tab row matching more than one live row → `!! AMBIGUOUS`, not applied; a
pre-v4.4 report without the dish_group column → `!! GROUP-LEVEL RULE NOT APPLIED`. Never silent.

**#5 Only the first multi-column section was processed.** action2, action3 and compare handle
every `side_pricing` section (one block each on Fried Specialty). A dish printed in two sections is
merged; if the sections disagree for a modifier it keeps its current price + `PAPER CONFLICT`.

**#6 Tab order.** compare's canonical order now matches part1.md / action3 / price_inference:
… Modifier Summary, **Modifier Rule, Fried Specialty**.

**Pricing model per (dish, variant), not per restaurant** (found on the real ABSOLUTE restaurant
ER81988 China Village). v4.3 decided the model from the FIRST By Itself-type row in modifier_rule;
China Village's first one is a restaurant-wide `Plain $0.00` rule, so preflight said UNKNOWN and
compare did DELTA arithmetic on ABSOLUTE rules (nuggets + fries $10.95 → $4.00). China Village also
mixes models inside ONE dish: Cut Wings 5–30 pcs are ABSOLUTE, 40/50/100 pcs are priced on
dish_item with upcharge rules. And compare keyed prices by (dish, modifier) without the size, so
every Cut Wings size collapsed to one price (5-pc … 50-pc + fried rice all → $3.70).
- `_cell_models`: each (dish, variant) with priced dish-level rules is DELTA (dish_item priced),
  ABSOLUTE (dish_item $0) or DOUBLE_CHARGE (dish_item AND By Itself priced — blocks). Only
  dish-level By Itself rows are anchors. Restaurant verdict: DELTA / ABSOLUTE / **MIXED** (exit 0)
  or DOUBLE_CHARGE / BROKEN (block). preflight lists which dishes/sizes are which.
- compare keeps every price per (dish, variant). A paper row is attributed to one variant (its Plain
  price key, the dish's only variant, or a size in its name like "5pcs"); otherwise
  `VARIANT UNKNOWN`, not applied. Sizes the paper did not print keep their own prices; they are not
  twin-filled when the dish has a size-less paper row. ocr-prompt.md rule 5: key a sized row's Plain
  by its size column.
- On the real China Village raw (test OCR, invented prices): MIXED (19 ABSOLUTE + 4 DELTA), every
  printed combo charged exactly its paper total, no dish_item written, no other size touched.

**Combo boxes and the real China Village menu.** Run against the real paper menu + raw
(ER81988): its "Appetizer Combos — all served w. french fries or fried rice $10.95" box prints only
the with-side price; each dish's own price is in APPETIZER. Read as a one-price box (floating
`default_price`), the $10.95 became the dishes' own prices: Barbecue Spare Ribs Small $8.95 →
$10.95, Rib Tip / Boneless Large $11.95 → $10.95, and P-ABS quoted $10.95 as the Plain.
- ocr-prompt.md rule 5: a combo box is side pricing — its own section, `side_pricing`, items with
  `"prices": {}` + `side_prices`; never `prices` / `default_price`.
- action2: a row carrying `side_prices` never takes a section's floating price.
- compare: a combo row with no Plain is kept — on an ABSOLUTE cell the side rule takes the full
  total; its By Itself takes the Plain from the same dish's own row elsewhere on the menu, only when
  the paper names match (plural-folded) and that row prints exactly ONE price (noted as PLAIN FROM
  ANOTHER SECTION). A DELTA row without a Plain → NO PLAIN, not applied.
- compare: an exact (normalized) name match to ANY system dish wins, so a row for a dish without
  modifier rules ("BBQ Spare Ribs") no longer lands on a similar rule-bearing dish ("Barbecue
  Boneless Spare Ribs"). A bare "Lo Mein" (Chow Mein, Mai Fun) column maps to the plain one, like
  "Fried Rice" → Plain Fried Rice.
- Result on the real menu (paper = system): 269 prices confirmed, 0 changed, 0 cells written in
  Part 2, no flags.
- Maintainer decision (2026-09-30): price twins are safe to follow for China Village.

**price_inference never prices a $0.00 cell** (free items: "Free Egg Roll", utensils). A group whose
printed dishes all moved +$0.50 used to give a freebie an APPLIED purple $0.50 via T2 (v4.3 too).
Now SKIPPED, any tier incl. T-TRAY.

Also: Fried Specialty separates **conflicts (not auto-resolved)** from informational **notes**
(the old "Conflicts auto-resolved" count included the PRICE-TWIN summary line).

**Regression (as of these six fixes):** Party Tray harness (normal + mixed fixture) — OCR output,
report and final raw cell-identical to v4.3 (only the Summary timestamp differs). A restaurant with
no $0.00 modifier-priced dishes gets the same Dish Item / Questionnaire / Price Inference / Summary as
v4.3. (The edit-sheet batch above then changes the report on purpose; the final raw stays the same.)
Tests: `tests/run_all.sh` (see tests/README.md).

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
