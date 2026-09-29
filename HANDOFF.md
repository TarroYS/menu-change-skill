# Handoff — menu-change skill maintenance

Carried over from a claude.ai session (2026-09-28). Maintainer: **Yee Sern**, Senior Menu Change Specialist, Wonders (KL). Not the original author of these scripts; design intent often has to be reconstructed from the code.

**Current state:** v4.3 (Party Tray) done and tested on synthetic data. Fried Specialty / modifier_rule fully mapped, six issues confirmed by test, none fixed yet. **Next task: issue #1**, after the open question below.

---

## 1. Repo map

```
menu-change-workflow-part1.md   thin skill: compare (OCR → Price Comparison Report)
menu-change-workflow-part2.md   thin skill: apply (edited report → staged draft)
menu-change-payload/
  scripts/        11 Python scripts (the pipeline)
  references/     ocr-prompt, cuisine modules, price-inference, wsheet guide
  hooks/          stop-gate hook
  MANIFEST.txt    file list + "fidelity lock" note + v4.3 amendment
CHANGELOG.md, SPLIT.md          (SPLIT.md is a 2026-09-15 draft note; its
                                 "script bytes unchanged" line is stale after v4.3)
tests/                          synthetic fixtures + pass/fail checks
```

**Part 1:** wsheet fetch → `bridge to-xlsx` → `action1` (Menu Preview = "2a", system prices) → vision OCR → `ocr_sanity_check` → `action2` (OCR Output = "2b", paper prices) → `action3` (2a vs 2b → Price Comparison Report) → `price_inference --mode apply` (purple fills) → if multi-column pricing: `modifier_rule_pricing preflight` + `compare --report`.

**Part 2:** edited report → `action4` (Final Review) → `action5` (Final Raw: dish_item, modifier_item) → `bridge to-tsv` → §2c `modifier_rule_pricing preflight` + `apply` → `validate` → `price_queue_gate` → `wsheet draft create` → `part2_report.py`.

Report colours: blue = increase, pink = decrease, green = confirmed no change, **purple = inferred** (callback item), apricot/orange = specialist edit.

---

## 2. Party Tray — v4.3 (DONE)

**Before:** Party Tray never updated. action2 discarded the paper tray price and backfilled the system price; action3 overwrote it and excluded trays from compare, coverage and P3; price_inference never inferred it (`NEVER_INFER_VARIANTS`). The only route to production was a specialist hand-typing a value or `=G23*5` formula after the callback. The freeze dated back to the original skill ("G3. Party Tray — do not update from paper menu").

**Decisions (maintainer's):**
- Printed trays update like any other price.
- Unprinted trays: infer from the **old ratio**, write **purple**, but **only if ≥70% of tray values share the same ratio**. Ratios vary (e.g. 500% or 400%, sometimes minus a few dollars).
- Tray inference must live in **Part 1**.
- Replace the static Questionnaire row `A-1 Party Tray pricing — confirm or update?` with **data-driven rows**.
- Scope: `Party Tray` only, via one list (`TRAY_VARIANTS`). Half Tray / Full Tray / Family Pack were never frozen; don't extend the ratio rule to them without evidence.

**Implementation:**
- `action2` `map_prices_to_2a_row`: tray price goes only to the tray column, never rerouted (a $60 tray must not land in Order). Removed the system-price backfill.
- `action3`: removed all six suppression sites; trays compared normally. Standing questions renumbered A-1..A-3.
- `price_inference.py` — new tier **T-TRAY** (`detect_tray_formula`, `infer_trays`, `append_questionnaire_tray`):
  - Fits `Tray = Base × k + d` on the system's OLD prices, restaurant-wide. k ∈ 1.0–10.0 in 0.5 steps; |d| ≤ `--tray-max-offset` ($10).
  - Trusted only when ONE formula fits ≥ `--tray-min-agreement` (0.70) of tray rows, ≥ `--tray-min-rows` (3) rows, with ≥2 distinct base prices. Ties → AMBIGUOUS, never guessed.
  - Applied only to rows that follow the formula, off a **printed** base. Base itself inferred → SUGGESTED only (no chaining). Decrease / caps → SUGGESTED. Outliers → asked, with what the formula would give.
  - Questionnaire **P-TRAY** section: `P-TRAY-0` states the formula (or why none), then one row per tray needing an answer.
  - `--no-tray` disables it. Runs only when the restaurant actually has tray prices — action1 emits an empty `Party Tray` header in *every* Menu Preview (a bug in the first draft, caught by the no-tray regression check).
- `ocr_sanity_check`: whole-dollar trays ≥ $100 no longer flag DROPPED DEC (`TRAY_LABELS`).
- **Part 2: no code change** — action4/action5 were already header-driven.
- Documented as the single sanctioned exception to G14 "no cross-variant inference" (the original rule exists because Pt/Qt rise by different amounts; a tray is a multiple, not a portion).

**Assumptions not yet confirmed by the maintainer:** measured restaurant-wide (not per dish group); base column auto-detected; tray-only dishes excluded from the 70% denominator.

**Not yet done:** validation on a real restaurant with Party Tray prices.

**Possible follow-ups (not requested):** flag printed trays that disagree with the detected formula (catches OCR misreads); per-dish-group formulas; write purple trays as live formulas so a callback edit to the base recomputes the tray (risky: action4 relies on the `data_only` cache).

---

## 3. Fried Specialty / modifier_rule — how it works

**Concept.** A multi-column paper table (Plain | w.French Fries | w.Pork Fried Rice …). The Plain price is on the dish; each side price is a row in the raw **modifier_rule** tab — one per (dish, variant, side), price in `out_unit_price`. Only `modifier_rule_pricing.py` touches this tab.

**Two storage models** (must be detected — same paper, different numbers):

| Model | dish_item | By Itself rule | Side rule stores |
|---|---|---|---|
| DELTA | base ($7.99) | none | upcharge (`total − plain`) |
| ABSOLUTE | $0.00 | base ($7.99) | full total |

`_detect_pricing_model` → DELTA / ABSOLUTE / DOUBLE_CHARGE / BROKEN / UNKNOWN. `preflight` exits 0 OK, 1 block (DOUBLE/BROKEN/G11), 2 UNKNOWN.

**Data path.** OCR JSON `section.side_pricing{columns}` + `item.side_prices{column: total}` (ocr-prompt.md rule 5) → action2 passes it through as a "Fried Specialty Sides" sheet → action3 builds the **Fried Specialty** tab → `compare` builds the **Modifier Rule** tab and appends its analysis to Fried Specialty. The Plain price also flows as an ordinary `Specialty` price into Dish Item.

**`_analyze` (compare), in order:**
1. **Row selection** — a modifier_rule row is shown if its dish has a `Specialty` variant, OR it was matched from the paper, OR its dish_group contains "fried special" (pulls in group-level rules with no dish_item).
2. **Dish matching** (`_match_dish_name`) — exact normalized → sauce + protein aware → substring (≤3 words) → ≥50% word overlap. `_expand_side_items` first splits "X in Garlic or Curry or Honey Sauce" into siblings.
3. **Column → modifier** (`_match_column_to_modifiers`) — strip `with`/`w.`, split on `or` and `/`, share suffix ("Pork or Chicken Fried Rice"), exact normalized match. "Fried Rice" also matches "Plain Fried Rice".
4. **Number** — DELTA: `total − plain`; ABSOLUTE: `total`.
5. **Fill**, first match wins: PAPER MENU → PROPAGATED (same dish, same old price) → PRICE TWIN (group) → PRICE TWIN (global) → NO DATA. Disagreeing twins → left + TWIN CONFLICT. Disagreeing *printed* members of one price group → auto-resolved by `--conflict mode|lowest|highest` (see #2).

**Modifier Rule tab.** A–F imported by Part 2 (matched by header label): enabled, **price (B — the only editable cell)**, dish_group, dish_item, variant, modifier_item. B turns orange when edited (vs hidden baseline in L). H–K = validation, ignored by Part 2: by itself, "Paper menu Total", Expected (live = H + B), Match?. Paper crop image at N1 if `crops/special_full.png` exists. Note: on untouched rows Match? is always TRUE (the "total" is computed from the same generated price), so it cannot detect OCR errors.

**Part 2 `apply`** reads approved (dish, variant, modifier) → price, re-detects the model and G11, writes only changed cells (G1/G4), prints DRIFT WARNING for unmatched triples, and merges counts into `part2_summary.json`.

---

## 4. Open issues — confirmed by test

Each has checks in `tests/modifier_rule/check_modifier_rule.py` (currently XFAIL). Fix directions are suggestions, not decisions.

**#1 — ABSOLUTE restaurants get double-charged. (Critical: production money error.)**
action2 isn't model-aware: it writes the paper Plain into dish_item over the $0 (looks like an ordinary blue increase). compare never updates the By Itself rule. P2 preflight and apply's G11 both inspect the fresh raw (dish_item still $0), and `cmd_apply` writes even when it detects DOUBLE_CHARGE — it only blocks G11 violations. Fixture result: wings + fries $17.38 instead of $9.39; the validation block labels $17.38 "Paper menu Total" with Match? TRUE.
*Directions:* `cmd_apply` should block DOUBLE_CHARGE/BROKEN like preflight; G11 in Part 2 must check the post-action5 state; under ABSOLUTE, the Plain price belongs on the By Itself rule, not dish_item.
*Depends on open question A.*

**#2 — Printed prices overwritten by price-group conflicts; validation hides it. (High.)**
In `_analyze`'s price-group loop, when members of one old-price group have different paper values, the winner is written to *all* members, including those with their own printed value, still labelled PAPER MENU. "Paper menu Total" is then recomputed from the wrong price, so Match? is TRUE. Fixture: Wing Pork Fried Rice printed +$2.20, shipped +$1.90.
*Direction:* printed members keep their own value; propagate only to unknown members, and only when the knowns agree (else NO DATA + flag — consistent with G14). Show the actual printed total in the validation column.

**#3 — "Plain Fried Rice" falsely blocks a DELTA restaurant. (Medium; fails safe.)**
`_is_by_itself` matches aliases (`plain`, `no rice`, …) as substrings, so a normal Plain Fried Rice upcharge reads as a By Itself anchor → DOUBLE_CHARGE → preflight exit 1.
*Direction:* whole-name alias match after normalization.

**#4 — Edits to group-level rules silently dropped in Part 2. (Medium.)**
Group-level rules (no dish_item) are shown on the tab, but `_read_mrv` requires a dish and `cmd_apply` keys on (dish, variant, modifier). No drift warning.
*Direction:* key group-level rules by dish_group, or at minimum warn loudly.

**#5 — Only the first multi-column section is processed. (Medium; silent.)**
Both action2 (Fried Specialty Sides passthrough) and `_analyze` `break` after the first `side_pricing` section. A second one (e.g. Wings) is ignored with no warning.
*Direction:* process every section; watch for the same dish appearing in two sections.

**#6 — Tab order contradicts the doc. (Low.)**
`modifier_rule_pricing.CANON_ORDER` puts Fried Specialty before Modifier Rule; part1.md, action3 and price_inference say the reverse. compare runs last, so its order wins.

**Suggested order:** #1 → #2 → #3/#4/#5 (all in compare) → #6.

### Found by reading, not yet tested
- Raw modifier_rule columns read by fixed position (`PRICE_COL=2`, `DISH_COL=9`, …), against G4's header-text rule.
- `final_prices` keyed by (dish, modifier) without variant → a dish with rules on two variants gets one price for both.
- The whole restaurant's model is decided by the first By Itself row found; a By Itself row at $0 lands as UNKNOWN rather than DELTA.
- `compare` doesn't block DOUBLE_CHARGE/BROKEN itself — it relies on preflight having been run.
- "Conflicts auto-resolved" count includes informational lines and TWIN CONFLICTs that weren't resolved.
- ABSOLUTE validation block shows plain + total.
- action3's delta-section title equals compare's `FS_SENTINEL`, so compare replaces action3's section. Works, by coincidence.
- part2.md `apply <report> <raw.xlsx> -o final.xlsx` doesn't say which raw; the uploaded `final_raw.xlsx` artifact won't contain modifier_rule changes.
- Stale comments: Modifier Rule tab called "green"; it's red.

### Other doc drift
- part1.md still describes Questionnaire Question/Recommendation/Answer columns; action3 removed them (`_interview` is a stub).

---

## 5. Open questions for the maintainer

- **A.** Do ABSOLUTE-model restaurants exist in practice, and does their dish_item hold `$0.00` or a blank? Decides whether #1 is live or theoretical. (Checking a real restaurant's raw gsheet would settle it.)
- **B.** Do trays follow one formula across the whole menu, or different ones per dish group (e.g. chicken ×5, seafood ×4)? Per-group would need a T-TRAY change.
- **C.** Party Tray v4.3 needs one real-restaurant run before rollout.

## 6. Tests

`tests/run_all.sh` → must end `ALL CHECKS OK`. Current: Party Tray 21/21 PASS; modifier_rule 5 PASS, 10 XFAIL (issues 1–6). Against the old v4.2 bundle the Party Tray suite scores 9/21, so it discriminates. See `tests/README.md`.
