# Tests

```bash
tests/run_all.sh                         # must end "ALL CHECKS OK"
REAL_RAW=~/Downloads/R85325_final_raw_gsheet.xlsx tests/run_all.sh   # + real-data checks
REAL_REPORT="$HOME/Downloads/R81587_NEW_CHOI_HEE_price_comparison (edited).xlsx" tests/run_all.sh  # + an older edited report
PYTHON=.venv/bin/python tests/run_all.sh # any python with openpyxl >= 3.1
tests/run_all.sh /path/to/other/bundle   # run the same checks against another bundle
```

Synthetic fixtures only; nothing here talks to wsheet. Each suite builds its fixture in a temp dir,
runs Part 1 and Part 2 end-to-end with the bundle's scripts, and checks the report and the final raw.

| Suite | Fixture | What it pins |
|---|---|---|
| `party_tray/check_party_tray.py` | R99999 — Party Tray = Large × 5 − $2.00, one outlier, tray-only dishes; plus a `mixed` fixture with no dominant ratio | v4.3 T-TRAY: printed trays compared, unprinted trays derived purple only off a printed base on rows that follow the formula; outlier / decrease / unprinted base left for the callback; P-TRAY section |
| `modifier_rule/check_modifier_rule.py` | R99901 "Delta Wok" (DELTA), R99902 "Absolute Wok" (ABSOLUTE, + a $0 freebie in a uniformly-moving group), R99904 "Village Wok" (the shape of real ABSOLUTE restaurant ER81988 China Village: model differs per dish AND per size, restaurant-wide `Plain $0.00` rule, a combo box, bare "Lo Mein") | Fried Specialty / modifier_rule issues #1–#6, per-(dish, variant) models, $0 never inferred (v4.4). Headline invariant: for every price the paper printed, the final raw charges exactly the paper total (dish_item + modifier_rule) |
| `price_inference/check_rules.py` | R99905 "Rules Wok" | the maintainer's same-price rules (2026-09-30): General Tso / Orange / Sesame <protein> always equal (beats a disagreeing T1 cohort; disagreeing siblings → conflict); Healthy "Steam X" = "X" only at the same old price and only when "Steam X" is not on the paper (the skip reason says why); Chow Mein + Chop Suey one cohort; $0 freebie never priced; no P6 |
| `edits/check_edits.py` | R99906 "Edit Wok" (dish codes, dishes under two dish groups, a modifier under two modifier groups, a same-name unpriced modifier twin, live prices changed between Part 1 and Part 2), R99907 "Abs Wok" (ABSOLUTE, group-level rules, a $0 Fried Specialties dish without rules, a required priced group), R99901 | the second 2026-09-30 batch: new prices in empty cells + moves (Dish Item and Modifier Summary, through action5 and the TSV bridge), invalid input ignored + reported, the report's entry rules (order, < 255 chars), a dish / modifier in two groups resolved as one item (edited copy wins, conflicts held, per-size merge), untouched prices never overwrite a live change (action5 and apply), dish codes left on moved / removed sizes listed, the widened $0 guard (group-level rules, Fried Specialties rule, new size, required options), Modifier Rule 'by itself' live links (ABSOLUTE + DELTA, evaluated), action2's missing-table warning, compare's FS-tab DRIFT |
| `audit/check_audit.py` | small made-up menus per finding (R99910–R99920) + R99901 / R99906 / R99907 | the 2026-09-30 self-audit: one check per confirmed bug (Part 2 instructions and typed input, the sheet's colours = Part 2's verdict, twins, refusals, the TSV bridge incl. a re-fetched TSV dir, the $0 guard, G11, compare flags / links / re-runs, side_rules, every price_inference fix), plus Party Tray per dish group (2026-10-05: own / whole-menu formula, a dish in two groups derived only when every group agrees), and two different prices printed for one size (asked in P-PRICE, never written or inferred). Each check fails when its fix is reverted |
| `codes/check_codes.py` | R99930 "Code Wok" (codes in the raw and on the paper: a match, one number for a two-size dish, a code the system lacks on a dish in two groups, Lunch / Dinner codes, two codes printed for one size), a no-code restaurant, a system-codes-only one; R99933–R99935 (the independent review: counts vs codes, routed codes, a pasted paper column, formulas, odd spaces, grey cells) | dish codes (maintainer 2026-10-05): the Dish Code tab's layout and colours, DC2 / DC5 placement, P-CODE; Part 2 writes only what the specialist confirmed (fills, corrections, x), refuses not-a-code / no price / code drift / conflicts, is sort-safe, the sheet colours = Part 2's verdicts, to-tsv fills / clears codes only while the TSV still holds them; the review fixes (a paper column pasted header and all, `=` / formulas, no-break and double spaces, `1.0`, a code on a grey cell, follow-ups after the specialist moved a code, a printed count is not a code, lunch codes only on Lunch, two codes reaching one size through routing, `春1` ≠ `夏1`) |
| `report/check_report.py` | R99999 (trays), R99902 (ABSOLUTE), R99901 (DELTA) — then a specialist-edited copy of R99901 through Part 2 **and the wsheet TSV bridge** (like ACE) | the edit sheet: every price a plain number, P1-TRAY, no P3 / Additional / P5 / P6, empty Notes → Wonka column, P-ABS "EDIT →" pointers + links, x colour. Part 2: `$9.25` typed the old way and plain `8.95` both apply, raw keeps `$X.XX`, `0` is a price, `x` clears the raw AND the TSV cell (x on an empty cell ignored), strikethrough (whole / part) and notes reach part2_summary.json + part2_report |

`--real-raw` (via `REAL_RAW`) adds three checks on a real wsheet export with a $0.00 dish whose price
lives on a required modifier group (R85325 "Seafood Boil"). `--real-report` (via `REAL_REPORT`) adds
two on an edited report from before 2026-09-30 (specialist_requests must not read its script-text
Notes column; action4 must still read it) — the same two also run on a synthetic older layout
without it. Neither file is stored in the repo.

`lib/xlformula.py` is a small Excel formula evaluator (IF / AND / OR / IFERROR / VALUE / ROUND / MAX /
ranges …; it agrees with LibreOffice on every sheet rule the report writes): tests check what a
spreadsheet would SHOW, and `cf_kind` tells which conditional-format rule a cell gets.

Current: party_tray 20/20, modifier_rule 51/51 (54 with `REAL_RAW`), price_inference rules 15/15,
report 36/36 (38 with `REAL_REPORT`), edits 67/67, audit 104/104, codes 34/34 — 327 checks (332 with both
real files). Against the v4.3 bundle: 16/20, 11/51, 4/15, report crashes, 13/51, 24/96, codes crashes — a crash counts as a failure. The Party Tray harness scripts (`build_fixture.py`, `run_p1p2.sh`, `show.py`) still work
standalone; see `party_tray/README.md`.
