# Handoff — menu-change skill maintenance

Carried over from a claude.ai session (2026-09-28). Maintainer: **Yee Sern**, Senior Menu Change Specialist, Wonders (KL). Not the original author of these scripts; design intent often has to be reconstructed from the code.

**Current state (2026-10-16):** v4.5 — the Modifier Rule and Fried Specialty tabs are built whenever Fried Specialty dishes are priced per dish on their sides, printed on the paper or not; the Fried Specialty tab is now editable price tables (FULL PRICE / PRICE DIFFERENCE / ABSOLUTE PRICE) with Notes → Wonka, and Part 2 reads them (§3, §6 M). **Committed.** `main` is always the latest version (the maintainer, 2026-10-21); the original v4.3 is tagged `v4.3`, and branches `v4.4` / `v4.5` mark those versions. GitHub remote: `origin` = github.com/TarroYS/menu-change-skill (private). The skill package is `menu-change-skill-v4_5-part1and2.zip` (same layout as the v4.3 / v4.4 zips, built from the commit — not in git).

**Earlier (2026-10-05):** the maintainer answered every open question (§6). Built: **dish codes** (§5c — a Dish Code tab in the report; Part 2 writes only the codes the specialist confirmed), **Party Tray formula per dish group** (§6 B), a clear stop for reports that lost their hidden data (§6 G). An independent review of that work found 7 problems (one high: a paper column pasted with its header lost its codes) — all fixed, each pinned by a check (§5c). Two different prices printed for one size are now asked, never picked (§6 K). `tests/run_all.sh` → ALL CHECKS OK (§7). **Committed** on branch `v4.4` (local repo, no remote; `main` = the v4.3 baseline); the skill package is `menu-change-skill-v4_4-part1and2.zip` (same layout as the v4.3 zip, built from the commit — not in git). Still open: §6 C (a live end-to-end run before rollout), §6 L (a count-prefixed paper name).

**Earlier (2026-09-30, night):** self-audit done (§5d) — three code reviews, a LibreOffice evaluation of the report's sheet rules, mutation testing of the suites, and two reviews of the fixes; every confirmed bug fixed and pinned by a regression check (`tests/audit`). Real-data results unchanged (China Village unedited Part 2 = 0 cells). Decisions wanted: §6 F, G, H, I, J. **Next feature: dish codes** (§5c). Nothing committed.

**Earlier (2026-09-30, evening):** second edit-sheet batch done (§5b): new prices in empty cells + moves, stray keystrokes ignored, a dish/modifier under two groups resolved as one item, Part 2 writes only real instructions (never reverts a live change), Modifier Rule 'by itself' live-linked, the $0 Fried Specialties guard widened, Fried Specialty tab triggers surfaced. **Next feature: dish codes** — see §5c (what exists today + the questions to settle first).

**Earlier (2026-09-30):** v4.4 — all six Fried Specialty / modifier_rule issues (#1–#6) fixed; pricing model decided per (dish, variant) (real ABSOLUTE restaurant ER81988 China Village, real paper menu run end to end); price_inference never prices a $0.00 cell. Plus the **edit-sheet batch** (§5): plain-number prices, `x` = remove, strikethrough = make unavailable, Questionnaire Notes → Wonka, P1-TRAY, P3/Additional/P5/P6 removed, P-ABS edit pointers, and the maintainer's same-price inference rules (T0). All covered by tests (`tests/run_all.sh` → ALL CHECKS OK); final raw still cell-identical to v4.3 on the Party Tray fixtures. Changes are uncommitted in the local git repo (baseline commit = v4.3 as handed off; `git diff` shows v4.4). **Next:** maintainer calls D / E in §6; a real-restaurant run with a real paper menu (C).

---

## 1. Repo map

```
menu-change-workflow-part1.md   thin skill: compare (OCR → Price Comparison Report)
menu-change-workflow-part2.md   thin skill: apply (edited report → staged draft)
menu-change-payload/
  scripts/        13 Python scripts (the pipeline)
  references/     ocr-prompt, cuisine modules, price-inference, wsheet guide
  hooks/          stop-gate hook
  MANIFEST.txt    file list + "fidelity lock" note + v4.3 amendment
CHANGELOG.md, SPLIT.md          (SPLIT.md is a 2026-09-15 draft note; its
                                 "script bytes unchanged" line is stale after v4.3)
tests/                          synthetic fixtures + pass/fail checks (run_all.sh)
```

**Part 1:** wsheet fetch → `bridge to-xlsx` → `action1` (Menu Preview = "2a", system prices, + `absolute_pricing` tab) → vision OCR → `ocr_sanity_check` → `action2` (OCR Output = "2b", paper prices) → `action3` (2a vs 2b → Price Comparison Report) → `price_inference --mode apply` (purple fills) → if multi-column pricing: `modifier_rule_pricing preflight` + `compare --report`.

**Part 2:** edited report → `specialist_requests` (struck dish names + Questionnaire Notes → Wonka, into part2_summary.json) → `action4` (Final Review) → `action5` (Final Raw: dish_item, modifier_item) → §2c `modifier_rule_pricing preflight final_raw.xlsx` + `apply <report> final_raw.xlsx -o final_raw.xlsx` (v4.4: on action5's output) → `bridge to-tsv` → `validate` → `price_queue_gate` → `wsheet draft create` → Wonka carries out the specialist requests (§4b) → `part2_report.py`.

Report colours: blue = increase, pink = decrease, green = confirmed no change, **purple = inferred** (callback item), apricot/orange = specialist edit, **dark grey = `x` (remove this price)**. Strikethrough on a Dish Item dish name = make it permanently unavailable. Every price is a plain number.

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

The model is decided **per (dish, variant)** (`_cell_models`, v4.4): only dish-level By Itself rows are anchors; dish_item priced + By Itself priced = DOUBLE_CHARGE. `_detect_pricing_model` → DELTA / ABSOLUTE / **MIXED** (both — real example: China Village, even inside one dish) / DOUBLE_CHARGE / BROKEN. `preflight` exits 0 OK (lists which dishes/sizes are which), 1 block (DOUBLE/BROKEN/G11).

**Data path.** OCR JSON `section.side_pricing{columns}` + `item.side_prices{column: total}` (ocr-prompt.md rule 5) → action2 passes it through as a "Fried Specialty Sides" sheet → action3 builds the **Fried Specialty** tab → `compare` builds the **Modifier Rule** tab and appends its analysis to Fried Specialty. The Plain price also flows as an ordinary `Specialty` price into Dish Item.

**`_analyze` (compare), in order (v4.4):**
1. **Sections** — EVERY `side_pricing` section (`_side_sections`), each with its own columns and coverage audit.
2. **Row selection** — a modifier_rule row is shown if its dish has a `Specialty` variant, OR it was matched from the paper, OR its dish_group contains "fried special" (pulls in group-level rules with no dish_item).
3. **Dish matching** (`_match_dish_name`) — exact normalized → sauce + protein aware → substring (≤3 words) → ≥50% word overlap. `_expand_side_items` first splits "X in Garlic or Curry or Honey Sauce" into siblings.
4. **Column → modifier** (`_match_column_to_modifiers`) — strip `with`/`w.`, split on `or` and `/`, share suffix ("Pork or Chicken Fried Rice"), exact normalized match. "Fried Rice" also matches "Plain Fried Rice".
5. **Variant + model** — each paper row is attributed to ONE (dish, variant): its Plain price key, the dish's only variant, or a size in its name ("5pcs"); else VARIANT UNKNOWN (flagged, not applied). The model is that cell's own (`_cell_models`): DELTA (dish_item priced) → `total − plain`; ABSOLUTE (dish_item $0) → `total`, and the Plain goes on the By Itself rule. Observations are pooled per (dish, variant, modifier) across sections; disagreeing prints → PAPER CONFLICT (kept, flagged).
6. **Fill**: a printed row always keeps its own value (PAPER MENU). Unprinted members of an old-price group → PROPAGATED only if every printed member agrees, else kept + GROUP CONFLICT. Then PRICE TWIN (group) → PRICE TWIN (global) → NO DATA; disagreeing twins → TWIN CONFLICT. Conflicted rows are never twin-filled. Nothing is auto-resolved; `--conflict` is ignored (G14). Callback items go to "⚠ CONFLICTS — NOT AUTO-RESOLVED", informational lines to "NOTES".

**Modifier Rule tab.** A–F imported by Part 2 (matched by header label): enabled, **price (B — the only editable cell)**, dish_group, dish_item, variant, modifier_item. B turns orange when edited (vs hidden baseline in L). H–K = validation, ignored by Part 2: by itself, "Paper menu Total" (the total actually printed; blank for derived rows → Match? N/A), Expected (DELTA: H + B; ABSOLUTE: B), Match?. Paper crop image at N1 if `crops/special_full.png` exists. On untouched printed rows Match? is TRUE by construction — it validates the specialist's edits, it cannot catch OCR misreads.

**Part 2 `apply`** (run on action5's `final_raw.xlsx`) refuses DOUBLE_CHARGE / BROKEN and any G11 violation in the resulting state (exit 1, nothing written); keys rows by (dish_group, dish_item, variant, modifier_item) so group-level rules apply; a key that matches >1 live row → `!! AMBIGUOUS`, not applied; unmatched → DRIFT WARNING; merges counts into `part2_summary.json`.

**v4.5 (2026-10-16): always, and editable.** `_fs_pattern` finds the pattern from the raw alone — dishes of a
"Fried Specialt…" dish group whose own priced rules price items of a "Select…" modifier group (Select Rice) — so
`compare` runs (Part 1 now always runs it) and builds both tabs with no paper table, or with only some dishes
printed. The Fried Specialty tab (`_render_fs_tab`) then holds FULL PRICE (paper; grey = today's), PRICE DIFFERENCE
(DELTA sizes, editable) and ABSOLUTE PRICE (ABSOLUTE sizes, editable; Plain = By Itself), a Notes → Wonka column +
box, then the paper-reading analysis; Questionnaire P-FS lists what the paper does not print. `apply` reads the
tables (`_read_fs`, hidden `_ref_key` / `_ref:` / `_live:`), merges them with the Modifier Rule tab (an edit beats an
untouched Part 1 value; different edits on both = `tab_conflicts`, not applied); specialist_requests collects the
tab's notes. Restaurants pricing Select Rice once for every dish (on modifier_item) are unchanged. An independent review found 10 problems (one high: an unprinted dish's twin-filled side prices were written while P-FS said "today's" — now purple and named in P-FS); all fixed, each pinned by a check (CHANGELOG v4.5).

**$0.00 dishes priced on modifiers (v4.4, G11).** `action1` lists them in the Menu Preview's `absolute_pricing` tab: priced dish-level modifier_rule rows (ABSOLUTE), or a required modifier group whose every item is priced (R85325 Seafood Boil → Select Seafood Boil, $9–$42). action2 never writes a paper price onto them (`Absolute Held` → Questionnaire **P-ABS**); price_inference never infers onto or from them; action5 refuses to write a price onto such a live cell (`!! ABSOLUTE GUARD`, shown by part2_report).

---

## 4. Issues #1–#6 — FIXED in v4.4

All six reproduced on v4.3 first (`tests/modifier_rule` today: 11/47 on v4.3 → 47/47; 50/50 with the R85325 checks). Summary in `CHANGELOG.md`.

| # | Was | Now |
|---|---|---|
| 1 | ABSOLUTE wings + fries charged $17.38 (dish $7.99 + rule $9.39); By Itself never updated; Part 2 checked the fresh raw; apply wrote through DOUBLE_CHARGE | dish_item stays $0.00 (P-ABS); Plain → By Itself; Part 2 checks post-action5 state; apply blocks DOUBLE_CHARGE/BROKEN + post-state G11; action5 guard. Final raw charges $9.39 |
| 2 | printed +$2.20 overwritten by the group winner +$1.90, validated TRUE | printed rows keep their value; split-group unprinted members kept + GROUP CONFLICT; Paper menu Total = printed total |
| 3 | "Plain Fried Rice" read as By Itself → DELTA blocked | whole-name alias match |
| 4 | group-level rule edits dropped silently | keyed by dish_group; ambiguous / pre-v4.4 rows warned, never silent |
| 5 | only the first multi-column section used | every section; cross-section disagreement → PAPER CONFLICT |
| 6 | compare put Fried Specialty before Modifier Rule | matches part1.md / action3 / price_inference |

Also fixed on the way: compare blocks DOUBLE_CHARGE/BROKEN itself; the "Conflicts auto-resolved" count no longer mixes in the PRICE-TWIN note; ABSOLUTE validation no longer shows plain + total; part2.md says which raw §2c uses (so the uploaded `final_raw.xlsx` now carries the modifier_rule changes); the stale "green tab" comment; part1.md's Questionnaire description.

### Found on the real ABSOLUTE restaurant ER81988 China Village — FIXED
- v4.3 / early v4.4 read China Village as UNKNOWN (first By Itself-type row = a restaurant-wide `Plain $0.00` rule) → compare did DELTA arithmetic on ABSOLUTE rules (nuggets + fries $10.95 → $4.00). Now per-(dish, variant) models: MIXED, 19 ABSOLUTE + 4 DELTA cells.
- It mixes models inside one dish: Cut Wings 5–30 pcs ABSOLUTE, 40/50/100 pcs priced on dish_item (By Itself $0 / off, $3–$5 upcharges, $0.50 Wing Sauce). compare used to collapse all sizes to one price per modifier (all fried-rice combos → $3.70); prices are now kept per (dish, variant).
- The paper row must say its size: ocr-prompt.md rule 5 now keys a sized row's Plain by the size column; a size-less row → VARIANT UNKNOWN.
- Price twins fill unprinted dishes from flat totals (China Village's $10.95 combos move together) and, under ABSOLUTE, By Itself bases from same-old-base dishes. **Maintainer decision (2026-09-30): safe to follow for China Village.** They are blue on the Modifier Rule tab like printed rows; the source is in Fried Specialty's propagation detail.
- Real paper menu (IMG_3914, 2026-09-21) run end to end on the real raw: 269 prices confirmed, 0 changed, 0 cells written — paper = system. Its "Appetizer Combos" box prints only the with-side price ($10.95; own prices in APPETIZER). Read as a one-price box it put $10.95 on Barbecue Spare Ribs Small / Rib Tip Large / Boneless Large and into P-ABS; fixed (ocr-prompt rule 5 combo box, action2 floating-price guard, compare combo rows + Plain join, exact-name match first, bare "Lo Mein" column).
- Still open from that run (**maintainer confirmed both, 2026-09-30 — left open, not fixed**): action2 leaves "Fried Chicken Nuggets" / "Fried Shrimp" as P1 (system "Chicken Nuggets" / "Fried Jumbo Shrimp" — P1 suggests the stripped-"Fried" name); and action2 silently routes a paper `Order` price onto Small/Large by distance when the dish has no Order column (that is how the combo price landed on Small/Large) — worth a P-COLUMN-style flag.

### price_inference and $0.00 freebies — FIXED
- A $0.00 freebie ("Free Egg Roll") in a group whose printed dishes all moved +$0.50 got an APPLIED purple $0.50 via T2 (v4.3 too). A $0.00 cell is now never priced by inference (SKIPPED, any tier incl. T-TRAY).

### Still open — found by reading, not yet tested
- Raw modifier_rule columns read by fixed position (`PRICE_COL=2`, `DISH_COL=9`, …) in modifier_rule_pricing.py, against G4's header-text rule. (action1 / action5's new G11 code reads by header text.)
- action3's delta-section title equals compare's `FS_SENTINEL`, so compare replaces action3's section. Works by design; keep the two strings identical.
- ~~action5 writes report prices as it reads them (bare "7.99" after compare's `_strip_dollar_display`), so the `$X.XX` normalization in part2.md step 2d is load-bearing.~~ FIXED (2026-09-30): action5 writes the live cell's own format (`_price_str`); apply always wrote `$X.XX`. Step 2d no longer normalizes.
- Found by the 2026-09-30 sweep of real reports: 5 reports in Downloads have lost action3's hidden `Modifier` tab (re-saved by another tool — "no default style"); action4 stops with `No 'Modifier' tab` (v4.3 the same). Not addressed.

## 5. The specialist's edit sheet + same-price rules (2026-09-30) — DONE

Requested by the maintainer after reviewing a real edited report (R81587 New Choi Hee). Summary in `CHANGELOG.md`; conventions in part1.md ("Edit-sheet conventions") and part2.md (§0b / §4b, G1, G5, G6).

| Request | Implementation |
|---|---|
| Prices without `$` everywhere (no `'$` typing) | action3 / compare / price_inference write numbers (`_numeric_display`, `_strip_dollar_display`). action5 writes the live cell's format (`$X.XX`), apply `$X.XX`; `0` is a price (action4/5 truthiness bug fixed) |
| Strikethrough on Dish Item col B = permanently unavailable | **Python + md:** `specialist_requests.py` detects (whole / partial, rich text) → part2_summary.json; part2.md §4b tells Wonka to do it with its own tools (scripts never add/remove rows); part2_report counts it |
| `x` / `X` = remove that price, own colour | dark grey 595959 + white, CF ahead of apricot; action5 clears (G1 exception) + records; bridge to-tsv blanks only recorded cells |
| General Tso / Orange / Sesame <protein> same price | price_inference T0 trio (before T1) |
| Healthy = regular dish when old prices equal | T0 healthy ("Steam"/"Steamed" prefix dropped, names normalized) |
| Chow Mein + Chop Suey same old price → same new | merged T1 cohort |
| Questionnaire Notes = prompts for Wonka | column I `Notes → Wonka`; read by specialist_requests (only reports with a Details column); part2.md §4b + G5 |
| Party Tray question into P1 | P1-TRAY, first P1 row, only when tray prices exist |
| Remove P3 / Additional Questions / P5 / P6 | removed (action3, price_inference); Price Inference tab + purple remain |
| P-ABS: obvious what to edit | Details `EDIT → '<tab>' · row · column`, hyperlinked (compare fills the By Itself row per size) |

**Validation:** R81587 rebuilt from its raw + paper: the new rules reproduce all 22 Chop Suey hand edits and Orange Chicken Lunch / Combo; 2 Healthy Small cells differ (question D). Sweep of 217 real reports: specialist_requests exit 0 on all (7 already carry struck dish names — R86659 has 9), action4 cell-identical to v4.3 on 212 (5 = missing hidden tab, above).

## 5b. Second batch (2026-09-30, evening) — DONE

| Request | Implementation |
|---|---|
| a. Type into an EMPTY Dish Item / Modifier Summary cell (move = paste + x); intentional + obvious; letters skipped | vivid-yellow CF on a value in a cell whose `_ref_` is empty; red-text CF on anything not a price (3 rules, each < 255 chars). action5 writes a Dish Item new price only when the cell was empty in the report too; action4 resolves Modifier Summary new prices (hidden `Part 2 Intent` sheet), action5 writes them only on the live row the report shows (never a same-name twin). Not a valid price (letters, comma, %, negative, ≥ 1000 — `PRICE_CAP`, 0 in an empty cell) → ignored + reported. Bridge fills an empty TSV cell only with a recorded new price. `x` now works in Modifier Summary |
| b. Understand dish codes | §5c |
| c. Same dish / modifier in two groups: an untouched lower copy must not restore the old price | action5 resolves each size across all rows of a dish (edited copy wins; different edits / different Part 1 changes = CONFLICT, not applied); action4 merges Modifier Summary edits per (item, size) (it used to drop the upper row's edits). Real data: 10 edited reports / 145 cells had this pattern |
| d. When does Fried Specialty appear? | only when the OCR JSON has a `side_pricing` section with side_prices items (action2 → action3); Modifier Rule only when compare ran. New: action1 `side_rules` tab + action2 `⚠ NO MULTI-COLUMN TABLE` warning; compare builds a missing Fried Specialty tab + `!! DRIFT`; part1.md 3b keys off action2's "Multi-column sections passed: N" |
| e. Dish Item edit → Modifier Rule totals update | H 'by itself' = live link (DELTA: Dish Item cell, edited copy preferred; ABSOLUTE: By Itself row's B); Match? N/A when Expected empty |
| f. $0 Fried Specialties never priced in ABSOLUTE restaurants | was covered only for dish-level rules / required groups; now also group-level rules and the maintainer's name rule (any $0.00 Fried Specialties dish once the restaurant prices on rules); action5 also guards new sizes, removing an ABSOLUTE size, removing / zeroing an option of a required group that prices a $0.00 dish |

**Found on the way (real data) — FIXED:** action5 / apply wrote every report value that differed from live, so a live price changed since Part 1 was reverted to the report's snapshot (and, after the "0 is a price" fix, a stale $0.00 modifier overwrote a live $11.95 at China Village). Now only instructions are written (G15): edits + Part 1 changes (blue/pink/purple; Modifier Rule blue/pink). Real pairs (old report + current raw): v4.3 would have reverted 117 (R83807) + 222 (China Village) live prices.

**Known limits:** a report without `_ref_` columns (very old) falls back to "any difference is written". Formulas in the Modifier Rule tab need a spreadsheet app to calculate (openpyxl stores no cached values) — Part 2 never reads H–K.

## 5d. Self-audit (2026-09-30, night) — DONE

Requested by the maintainer ("do self-audits and find bugs on the latest build"). Three reviews + a LibreOffice evaluation + mutation testing, then two reviews of the fixes themselves (16 more findings, incl. 2 high — all fixed). Full list in `CHANGELOG.md` "v4.4 (cont. 3)"; each fix has a check in `tests/audit` (or a strengthened check elsewhere) that fails when the fix is reverted.

| Area | What was wrong (worst first) | Now |
|---|---|---|
| Part 2: what is an instruction | Part 1's changes were read from the cell colour, so recolouring or sorting dropped them; a sort misaligned every baseline (the hidden block sat 3 columns away) | hidden `_ref_sys_` system-price mirror + per-row keys (`_ref_key:*`, Modifier Rule `_ref_live` / `_ref_key`); colour only for older reports |
| Part 2: typed input | Modifier Rule apply took anything ('1,50' → $150, 1395); a misplaced decimal under $10 shipped (195 for 1.95); judged against the report's price, an OCR misread (19.50 for 1.95) could not be corrected and shipped when untouched; %-formatted cells; 0 / x into an empty cell hit a price wsheet added since; Python rounding ≠ the sheet's (3.675 → $3.67) | one rule set in all three parsers = the sheet's red text (LibreOffice: 0 mismatches); 5× / a fifth of the SYSTEM price held (edits and Part 1's own price), new prices 10× / a tenth of the item's other prices; rounding like the sheet |
| Part 2: structure | same-name twin / availability rows written; edits in sizes the live menu lacks and x on empty modifier cells dropped silently; action5 without action4, action4 on its own output, re-runs accepted; the bridge skipped a tab (trailing blank line) and still exited 0, crashed on control characters; with a TSV dir re-fetched since to-xlsx, a requested x was silently skipped and a new price written over the newer value | only the row the report shows; reported; refused / exit 1 |
| $0 guard / G11 | a group add-on ("Sauce / With +$1.00", real R85325 / R83807 / China Village) made DELTA $0 dishes ABSOLUTE, so their paper tray prices were never compared — and the first fix then let a $0 dish priced on a REQUIRED choice's group rules take a dish price; a priced group-level By Itself on a priced dish passed preflight; apply could set an ABSOLUTE rule (By Itself, full-price side, group-level) to $0; action5 could add a size on top of a group-level By Itself | rules hold a $0 dish's price with a By Itself or on a required choice; blocked; held; held |
| compare | printed side prices on a dish without rules were dropped or labelled "⚠ size?"; the Fried Specialty block moved down on every re-run; size-less rules had no live link; no red text on Modifier Rule; its orange rule (since v4.3) was stored with a leading `=`, which Excel drops | `NO DISH RULES` flag only when a group rule prices that side, else a `NO SIDE RULES` note (China Village: 0 conflicts); fixed |
| price_inference | "(White Meat)" joined the plain family; a dish in two groups could be inferred two ways; Small < Large checked on old prices; re-runs kept old guesses (and could overwrite a specialist's value); trays / donors counted per row; a Healthy dish printed only as a tray or under a P1 name treated as unprinted; a stale report written on the wrong dish | fixed |

**Evidence:** suites in §7; mutation check (each fix reverted → its check fails: 49 mutations, all caught); LibreOffice parity of every sheet rule vs Part 2's parsers (Dish Item / Modifier Summary 140 cases each, new prices 72, Modifier Rule 58 — 0 mismatches); Party Tray final raw cell-identical to v4.3; China Village fresh Part 1 + unedited Part 2 = 0 cells changed, 269 confirmed, 0 callback conflicts; R87669 / R85325 real pairs identical to v4.3, R81988 / R83807 differ only where v4.3 reverted newer live prices; action4 identical to v4.3 on all 212 readable real reports; of 9,943 real specialist edits in 101 edited reports, none is 10× or a tenth of the report price.

**Known limits (not fixed):**
- A %-formatted cell is caught by Part 2, but the sheet cannot see a number format (it shows "15%", which is visible anyway).
- Modifier Rule 'by itself' links point at fixed Dish Item cells: after a sort of the Dish Item tab they show another row's price. Display only — Part 2 never reads them.
- In a de-DE spreadsheet locale, text with a decimal point is coloured by the locale's rules; Part 2's verdict is unaffected.
- A typed TRUE / FALSE in a price cell shows as an edit (apricot); Part 2 ignores it as not a price.
- to-tsv protects the x / new-price cells of a TSV dir re-fetched since to-xlsx (exit 1); other cells take the xlsx's value. part2.md: always re-fetch into a FRESH dir.

## 5c. Dish codes — BUILT (2026-10-05)

**How it works now** (CHANGELOG "v4.4 (cont. 4)"):
- **Part 1:** action1 writes a `System Codes` sheet (each size's `unit_code`, row-aligned with `price`; not named `dish_code`, which would switch on action2's dormant code tie-break). action2 sends each printed code to the system size its price landed in (`src` from `map_prices_to_2a_row`, the Small→Lunch routing undone): one number on a two-size row → both sizes (DC2), a Lunch-section `L26` → Lunch, a `D26` → Dinner (DC5); two different codes for one size → `Code Conflicts` (asked as P-CODE, never proposed). Written to a `Paper Codes` sheet. action3 builds the **Dish Code** tab (row-aligned with Dish Item): left block = current codes (editable, text format), right block `<size> · paper` (blue differs / yellow new / green same; omitted when the paper printed no codes), grey = no price on that size, hidden `_ref_<size>` (the system code) + `_ref_key:*`; CF: x grey, not-a-code red, new yellow, edited apricot (case-sensitive EXACT). Questionnaire P-CODE points at the tab.
- **Part 2:** action5 reads the left block only — an instruction = a cell ≠ its hidden baseline (DC1: every fill and correction confirmed by a person). Guards: not a code (line break, > 128, only `# - .` / spaces), no price on the size, code drift (live ≠ baseline), conflicting copies; `x` removes. Fills / removals go to `action5.added` / `removed` with `kind: 'code'` so to-tsv fills / clears only while the TSV still holds what Part 2 saw. `code_followups` (a code left on a moved / removed price) is only listed — always asked (DC4). Dish-level codes (`name_complex.code`) are never touched.
- **Independent review (same day): 7 problems found and fixed** (CHANGELOG "v4.4 (cont. 4)").
  - A code column whose paper column was pasted over it, header included, was silently skipped. It is now
    found by its place, and a code column found neither way stops Part 2.
  - A dish in two groups gets a derived tray only when every group derives the same one.
  - A formula in a code cell counts by its saved value. Text starting with `=` is not a code.
  - Follow-ups list only codes still on the removed / moved size.
  - The paper proposes a printed number only when it is clearly a code, a lunch-special code only on Lunch,
    and asks when two codes reach one size through routing.
  - The sheet, action5 and to-tsv read a code the same way (Excel TRIM).
  - A code typed on a grey (no price) cell is red.
  Each fix has a check that fails when the fix is reverted (22 reverts, all caught).
- Tests: `tests/codes` (34 checks) + `tests/audit` tray_two_groups (5). China Village (real, 340 codes):
  unedited Part 2 writes 0 cells. Every system code of R85325 / China Village printed on a made-up paper, 20%
  changed, every paper column pasted: 206 / 162 code cells written, each as the paper said.

**Background (read 2026-09-30):**

Source: Wonka's brief (`~/Downloads/dish-codes.md`) + the scripts + the real raws.

- **Three fields, never collapsed.** `unit_code` = the code of one dish at one size (the price link; max 128, may hold `#` `-` `.`) — searched, picks the size from a call transcript when unique on that dish, printed on the ticket. Dish `name_complex.code` (max 6, no `#-.`) = a code for the dish with no size. Size-entity `name_complex.code` (`Sm`) — never a paper number. Never put a code in the dish name; never set `print_dish_code_styles` as a side effect. Lookup key = letters + digits, lower-case (`L-12` = `l12`); duplicates allowed.
- **In the raw:** dish_item has a `unit_code` column next to every size's `unit_price` (row 2 `wprop.unit_code`, row 3 the size), plus `root_wprop.name_complex.code` and `print_dish_code_styles:en/zh_cn`. Real: R85325 529 unit_codes on 841 priced cells + 22 dish-level codes (AS1…); China Village 340 on 620, 0 dish-level. **A code never sits on a size without a price** in either raw. modifier_item has unit_code columns, 0 used.
- **The pack today:** action1 reads every dish's unit_codes (`dish_codes`) and has an unused dish_code-sheet builder; action2 strips a printed prefix into the OCR `code` field (`L15. General Tso` → `L15`), keeps each match's paper code + price slots in its match log (not written anywhere), and has a dormant code tie-break (P4 flag) that needs a dish_code sheet no Part 1 produces; action3 has no code output (part1.md's `P-CODE` exists only as text); action5 `APPLY_DISH_CODES = False` (the dormant path only overwrote existing codes, never filled a blank); part2_report prints "codes: DISABLED". New in §5b: action5 lists `code_followups` when a removed / moved size keeps its code.
- **Settled by the maintainer (2026-10-05):** (1) both — fill missing codes from the paper AND correct wrong ones, every cell confirmed by the specialist (nothing written on its own); (2) one printed number for a dish with two sizes → the same `unit_code` on both sizes; (3) the specialist confirms on a Dish Code tab with the same conventions as prices; (4) a code on a moved / removed price → always asked (F); (5) Lunch / Dinner codes (L26 vs D26) → each on its own size's `unit_code`.

## 6. Open questions for the maintainer

- **A. ANSWERED (2026-09-29):** ABSOLUTE restaurants exist; their dish_item holds `$0.00`, not a blank. Maintainer's rule: whenever a dish is $0 and its full price lives on modifiers, do not update the dish item price. Real example R85325: Seafood Boil $0.00 with the price on the required `Select Seafood Boil` group (modifier_item prices) — covered by the same guard.
- **B. ANSWERED (2026-10-05):** trays usually follow one formula across the menu, sometimes one per group — "be safe, do it per group". T-TRAY now finds a formula per dish group; a group too small or too mixed to confirm its own uses the whole-menu formula only if that fits ≥ 70% of its own trays, else its changed trays are asked; a dish in two groups whose formulas disagree is not derived.
- **C.** Party Tray v4.3 and the v4.4 fixes need one real-restaurant run before rollout (ideally one ABSOLUTE restaurant with a multi-column Fried Specialty table).
- **D. ANSWERED (2026-09-30):** the two R81587 Healthy cells were missed by the specialist — the rule is right. Added: only a Healthy dish NOT on the paper menu takes its regular dish's price.
- **E. ANSWERED (2026-09-30):** make-unavailable (and every other Wonka follow-up) always goes in the same draft.
- **F. ANSWERED (2026-10-05):** a code on a moved / removed price is always asked (`code_followups`, Wonka asks) — never moved or removed automatically.
- **G. ANSWERED (2026-10-05):** keep stopping. Those 5 reports lost ALL hidden data (the Modifier tab and every `_ref_` column), so Part 2 could not tell what was edited; action4 now says so and tells the specialist to re-make the report with Part 1.
- **H. ANSWERED (2026-10-05): hold.** The misplaced-decimal rule holds a price that is 5× or a fifth of the SYSTEM price — a specialist's edit, or Part 1's own paper price (a likely OCR misread) — and a new size at 10× / a tenth of the item's other prices (Dish Item trays exempt). None of 9,943 real specialist edits, and none of 708 real Part 1-vs-live differences, come within 3×; a genuine change that big goes through Notes → Wonka.
- **I. ANSWERED (2026-10-05): yes.** Negative prices are accepted on Modifier Summary and Modifier Rule (real "Less Whole Wing" −$1.70 in 7 restaurants), never on Dish Item.
- **J. ANSWERED (2026-10-05): keep.** A Healthy dish counts as on the paper when an unmatched P1 row looks like it once the diet words are dropped ("Diet Shrimp w/ Broccoli (Steamed)" → "shrimp with broccoli").
- **K. ANSWERED (2026-10-05): two different prices printed for one size → require a callback.** Neither is written;
  Questionnaire P-PRICE asks which is right (a dish in two sections, an Order price routed onto a printed
  Large, a routed Small vs a lunch special). Before, the last one read was written (v4.3 too).
- **L. OPEN:** a count-prefixed paper name ("2 Egg Rolls") matches the single dish ("Egg Roll"), and its
  price is proposed as an increase (blue, not asked) — v4.3 behaviour, found in the 2026-10-05 review.
- **M. ANSWERED (2026-10-16): Fried Specialty tab.** The specialist types into the PRICE DIFFERENCE table (and ABSOLUTE PRICE for absolute dishes), not the full-price one; only restaurants that price the sides per dish get the tables (shared Select Rice prices stay on Modifier Summary).
- **Decided (2026-09-30):** price twins safe to follow (China Village); the two action2 items in §4 are real and stay open.

## 7. Tests

`tests/run_all.sh` → must end `ALL CHECKS OK` (`PYTHON=` any python with openpyxl ≥ 3.1; `REAL_RAW=` path to R85325 adds real-data checks; `REAL_REPORT=` an older edited report). Current: party_tray 20/20, modifier_rule 51/51 (54 with R85325), price_inference rules 15/15, report 36/36 (38 with an older report), edits 67/67, audit 104/104, codes 34/34, fried_specialty 30/30 — 357 checks, 362 with both real files. Against v4.3: 16/20, 11/51, 4/15, report crashes, 13/51, 24/96, codes crashes, fried_specialty crashes (1/8). `tests/lib/xlformula.py` evaluates the report's formulas and sheet rules (agrees with LibreOffice). Fixture R99904 copies China Village's structure; the real China Village export (`~/Downloads/ER81988 CHINA VILLAGE.xlsx`) is not in the repo. See `tests/README.md`.
