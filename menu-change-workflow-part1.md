---
name: menu-change-workflow-part1
description: Paper-menu → Price Comparison Report. Triggers on: price update, OCR menu, Wonders workflow, price comparison, price inference, P6. Not Part 2 / Action 4/5 / apply / draft / post-callback.
---

# Menu Change Workflow — Part 1 (compare)

SCOPE: Paper-menu → specialist review `.xlsx`. Stop after upload. Sister skill `menu-change-workflow-part2` applies edits; do not run it here.
Cuisine-agnostic (variant cols auto-discovered from headers, no whitelist).
Union runbook, P1-only. Scripts live in shared `menu-change-payload/` — not in this file.

## ⛔ REPORT FIDELITY LOCK (runtime — this run)
The Price Comparison Report `.xlsx` MUST be produced by the bundle scripts — same tabs, tab order, RED edit-tab colors, blue/pink/green/**purple** fills, hidden `_ref_` columns, apricot CF, Modifier Rule 5-col layout + embedded paper crop, Fried Specialty delta/propagation/twin sections, Price Inference tab, formula resolution.
Guaranteed ONLY by executing the shipped scripts. SCRIPT CATALOG is a RUN reference, not a spec to reimplement live. Never rewrite, minify, regenerate, or reimplement these scripts, and never build the report by hand or via a subagent. Ship exactly as action3 → `price_inference.py --mode apply` → (if side_pricing) `modifier_rule_pricing.py compare --report`.

## ENV
- SCRIPTS ARE NOT IN THIS FILE. Shared payload: `menu-change-payload/` (`setup.sh`, `scripts/`, `references/`, `hooks/`). Materialize those paths unchanged. `chmod +x setup.sh`.
- ACE: no pip/root, PEP668. `bash setup.sh` once/cold-start → `.venv`. Invoke `.venv/bin/python scripts/<x>.py` (NOT bare python3). `.venv` does not persist.
- Cold-start every run: never reuse a prior fetch/draft/intermediate. Re-resolve, re-fetch, rebuild.
- ACE inputs: user supplies paper menu (+optional Wonders Dictionary). Raw + Menu Preview from live `wsheet`. Standalone (no `wsheet`): user uploads Raw/Preview; deliverable is still the `.xlsx`.

## RESTAURANT RESOLUTION (mandatory, every run)
1. `find_restaurants` (by name or bare R#). Bare R# on menu/report `restaurant_enterprise` tab is OFTEN NOT fetchable (e.g. R88312 fails; RMCR88312 works).
2. STOP — specialist confirms name+identifier+city.
3. `wsheet fetch <exact id find_restaurants returned>`. Never fetch bare R# silently. Link restaurant to context.

## DEFINITION OF DONE
ONE `.xlsx` from action3 + `price_inference.py --mode apply` (+modifier_rule tabs if multi-col), `upload_artifact`ed. A `.md`/prose diff = FAIL.
Never delegate the price comparison to a subagent. Subagents may ONLY OCR images → OCR JSON.
Stop-hook `hooks/menu_change_done_gate.py` blocks stop if inbound menu img/PDF present but pipeline/xlsx missing (fails open on error; loop-safe via stop_hook_active).
Do not run action4/action5, do not stage a draft, do not run `price_inference` in a later P2 session.

## FILES
Raw Gsheet(.xlsx,40+ tabs,3-row hdr,JSON root rows) · Menu Preview(price+dish_code+modifier tabs, from action1) · OCR Output(preview+paper prices) · Price Comparison Report(deliverable).
Report tab order: **Summary, Dish Item, Questionnaire, Price Inference, Modifier Summary, Modifier Rule, Fried Specialty** (hidden Modifier last). Edit-source tabs (Dish Item, Modifier Summary, Modifier Rule) colored RED. Purple cells on Dish Item = inferred (not printed) — callback items.

## SCRIPT CATALOG (P1)
- `setup.sh` — `bash setup.sh`. Idempotent venv+openpyxl bootstrap.
- `wsheet_xlsx_bridge.py to-xlsx <tsv_dir> -o raw.xlsx` — each tab.tsv→sheet, cells verbatim str (no coercion).
- `action1_raw_to_2a.py <raw.xlsx> [-o]` — Raw→Menu Preview. Price cols = row2 `unit_price` markers. Include dish iff: in dish_group, exactly 1 root row, ≥1 non-empty price. "Select "-prefix modifier groups replicated per dish_group; orphans w/ prices at end; sections sorted reverse-alpha.
- `action2_ocr_to_2b_0625v2.py <ocr.json> <menu_preview.xlsx> [--dict <wonders.xlsx>] [-o]` — writes to ANY header variant col (no whitelist). Native POS-code match (dish_code tab, Bouncer-Logic bypass on split items) → fuzzy. Tier1 cleanse, bubble default_price fallback, 10-step name resolution, compound splitting, per-row section-aware size_based/lunch_based (Small→Lunch only when row has Lunch slot & no Small — Soup/White Rice keep Small), dual-column→size_based, section prefix/suffix, section aliasing.
- `ocr_sanity_check.py <ocr.json> [--strict]` — advisory pre-action2: Small>Large, whole-number ≥100 (dropped decimal; not for Party Tray), out-of-range, `uncertain:true`, no-price rows. Never changes prices. Default exit 0; `--strict` exits 1 if any flag.
- `action3_compare_and_merge.py <menu_preview.xlsx> <ocr_output.xlsx> [-o]` — Price Comparison Report. Tabs: Dish Item(blue=inc/pink=dec/green=confirmed-nochange), Modifier(hidden baseline), Summary, Questionnaire(free-text Question/Recommendation/Answer/Answer Detail — NO dropdowns), Fried Specialty(if side_pricing), Modifier Summary(deduped, excl all-$0 groups & $0 items). Hidden `_ref_` cols + apricot CF on manual edits. Party Tray compared like any variant (v4.3). Does NOT write purple / Price Inference — that is the next script.
- `price_inference.py <menu_preview.xlsx> <ocr_output.xlsx> --report <report.xlsx> --mode apply` — EVERY restaurant, after action3, before §3b compare. Fills unprinted prices (V-GAP / D-GAP) from dish-group donors, + **T-TRAY**: unprinted Party Tray from the restaurant's own tray formula (e.g. Large×5−$2), only when ≥70% of tray rows fit ONE formula. Writes purple Dish Item cells + Price Inference tab + Questionnaire P6 + P-TRAY + Summary AUTO-INFERENCE. Script default is `--mode suggest` (no writes) — this workflow uses `--mode apply`. Details: `references/price-inference.md`.
- `modifier_rule_pricing.py` (multi-col only)
  - `preflight <raw.xlsx>` — MANDATORY G11 gate, must exit 0. DOUBLE_CHARGE/BROKEN/G11-violation=nonzero→STOP+escalate. UNKNOWN=exit2 (no By-Itself anchor→manual, don't guess).
  - `compare <ocr.json> <raw.xlsx> --report <report.xlsx> [-o] [--conflict mode|lowest|highest]` — REQUIRED --report writes INTO that workbook (one file, forbidden to make separate). Adds Modifier Rule tab (5-col + OCR validation + embedded crop @M1) + folds delta+propagation into Fried Specialty. Detection by structure (`side_pricing`), not name. Delta=round(combo_total−plain,2). Propagation: same-dish old-price group, then PRICE TWIN (group) then global PRICE TWIN; TWIN CONFLICT left at current price (not guessed). `with`/`w.`/`w ` prefix stripped safely (bare `w` must be followed by `.` or space). Does NOT modify Raw. Idempotent re-run. Also re-applies canon tab order (incl. Price Inference).
- `crop_section.py <img> -o crops/<n>.png [--title X][--next Y][--rows top|middle|bottom][--box L,T,R,B][--min-price-cols N][--sidecar][--debug]` — derives crop from content. Also usable per priced section for Opus vision (see ocr-prompt IMAGE PREPARATION). Optional embed target for compare.

## PIPELINE
0. Resolve+confirm → fetch → `setup.sh` → bridge to-xlsx → action1 → Menu Preview.
1. Step0 inspect: read price-tab row1 (variant cols), skim dish groups for naming/col convention, infer cuisine (ChAm/sushi/other), load `references/cuisine-*.md` + always `references/ocr-prompt.md`.
2. OCR paper: read live dish list first, then dual-pass (faithful + system-aligned name). Map each price to the RIGHT variant col (wrong col=silently dropped). floating price→`default_price`; multi-col table→Plain in `prices`+others in `side_prices` dict + section `side_pricing`. Flag smudged glyphs with `"uncertain": true`. Save `R<x>_menu_ocr.json`.
   2b. `ocr_sanity_check.py` — resolve/confirm flags (fix JSON + re-check) before action2.
3. action2 → review P1 unmatched (near-name matches → edit JSON, re-run) → action3 → **`price_inference.py --mode apply`** (every restaurant; see `references/price-inference.md`).
   3b IF side_pricing: preflight (exit0) → compare --report. 3c optional crop (also: per-section crops for OCR).
4. Iterate: coverage <~80% w/ 5+ dishes = suspect; ship only when P1 small & each remaining P1 plausibly new. Quote inference counts from `price_inference.py` stdout / `price_inference_summary.json` only. `upload_artifact` the ONE .xlsx.
5. Feedback → `#proj-agentic-content-editor` (no Slack connector in ACE). STOP.

## GUARDRAILS
- G1 NEVER write blank cells (incl code slots, group-gate FALSE rows).
- G2 NEVER add/remove rows — overwrite existing only.
- G3 Highlight: blue=increase, pink=decrease, green=confirmed-no-change, **purple=inferred (not printed — confirm on callback)**. Party Tray (v4.3): printed → compared + written like any variant; unprinted → T-TRAY only (purple). Key tray prices `Party Tray`, never `Order`.
- G4 Use SYSTEM names + paper prices. Parse by header text, not position.
- G6 $X.XX format; codes verbatim if ever present.
- G11 ABSOLUTE double-charge guard (model-aware). DELTA: base on dish_item, no By-Itself row, side=upcharge (out_unit_price=combo_total−plain). ABSOLUTE: dish_item=$0, By-Itself row carries base, side=full total (out_unit_price=paper_total). Auto-detect from By-Itself anchor (aliases: By Itself/Plain/No Rice/No Side, normalized) vs dish_item. ABSOLUTE rule: every priced modifier_rule (dish,variant) MUST have dish_item==$0 else DOUBLE-CHARGE ERROR (blocks compare/preflight, nonzero). No-op for DELTA. Never route around nonzero.
- G12 No hand-authored numbers. Inference counts → `price_inference.py` stdout / `price_inference_summary.json` only. If a number isn't in this turn's tool output, don't write it.
- G13 Never skip multi-col analysis or claim it belongs to Part 2 — that is a critical failure. If OCR JSON has side_pricing, running modifier_rule_pricing compare is mandatory here.
- G14 Inferred prices: every purple / P6 row is a callback question (never present as if printed). Never widen the inference rules (no cross-variant delta — sole exception is the shipped T-TRAY tier, do not extend it — no resolving a cohort/twin conflict by picking a side, no hand-filling SKIPPED rows).

## QUESTIONNAIRE CATS (Questionnaire sheet, dishes)
P1=on paper, unmatched. P-COLUMN=variant col blank in system row→routed to only populated col (verify). P2=decreases (rare, verify). P3=in system, no paper match (maybe discontinued, w/ expected price). **P6=unprinted prices inferred by `price_inference.py` (V-GAP / D-GAP) — every row is a callback item.** **P-TRAY=Party Tray callbacks: P-TRAY-0 states the detected formula (or why none), then each derived / outlier / unresolved tray (replaces old static A-1).** +Additional, P5=drinks/sauces/side orders. P-CODE (dish code) → not auto-written (G1).

## OCR RULES (ocr-prompt.md core)
JSON schema = full schema at bottom of `references/ocr-prompt.md` (emit exactly), including optional `uncertain` / `uncertain_note`.
Capture EVERY item; names exact as printed (script normalizes w./&/hyphen/plural/word-flip). Variant key must match a Menu Preview header. Note codes (L1,C1). Multi-col: capture FULL bordered block out to ruled border incl right-most col; preserve with/w. verbatim; capture every printed price incl partial rows (null only for genuinely blank cols); read ragged rows by dotted leaders not x-position; side_prices:null ONLY when Plain-only & truly no side prices; self-check #side_prices rows == #paper rows w/ >1 price. Strip decorative tokens ((4),(6),reading aids) unless disambiguating. Digit care: .25/.75, .00/.50, 8/9, 0/6, 1/7, 3/8, 5/6, 9/4; whole-number ≥100 is a dropped decimal until proven. Prices ~never decrease (re-verify). Count price cols per section first; emit one item per printed row (`null` not skip). Flag uncertain instead of guessing. Crop per section for Opus (~2576px long edge) — do NOT feed system prices into the read. Compound auto-expansion (parenthetical, or, slash-dual, comma multi-protein, bare-protein-under-header→prepend section [or emit `<Protein> <SectionWord>` for suffix sections], choice-of-sauce→one dish/choice).

## CUISINE MODULES (load per Step0)
- ChAm: cols Order/Small/Large/Lunch/Combo/Specialty/Party Tray (catering — capture every printed tray). Small=Pt/1st, Large=Qt/2nd, Order=single, Specialty=Plain-col; lunch_specials & combinations sections. Expand Moo Shu/EFY(reverse)/Chow Mein-Chop Suey(keep combined). Suffix sections: Pancit/Yat (House Yat not House Special Yat). Aliases: Diet→Healthy Dishes, Combination Plates→Combo, Chef's→Chef Specialties, Specialties→Fried Specialties, Sweet&Sour→Sweet Sour, drinks in Side Orders→Drinks; Garlic/Broccoli/Hunan/Szechuan/Lo Mein word-flips, Roast Pork→Pork, Chow Mei Fun→Mei Fun, phonetic (Shu mai/Shao Mai, Kung Po/Pao/Gong Bao, Tso/Tso's). Wing sauce→Honey/Buffalo/Barbecue/General Tso/Garlic Wing. Protein tiers Veg<Pork≈Chicken<Beef≈Shrimp<House/Seafood — re-verify breaks. Lunch = single bubble price→lunch_specials.price (unless tiered).
- Sushi: cols Roll/Regular Roll/Dinner/Bowl + restaurant combo col (verbatim, e.g. Seapark). "Roll" in name ≠ Roll col (Lobster/Crab Sub Rolls=Order, sandwiches). Regular Roll=Any3/6/9 party bundles. CROSS-SECTION DUPLICATES (critical): Party Tray For 2/4/6 → also Sushi Sashimi Entrees Don rows (Sushi For 2, Sashimi For 2, Sushi and Sashimi For N); named rolls in 2 groups→map price to all. Aliases: Crab→Crab Sub Roll, Hawaii↔Hawaiian, apostrophe/plural, chowder/bisque+Soup. Fish tiers Veg<Salmon/Tuna<Yellowtail/Eel<specialty<Lobster; Lobster/Crab uniform (outlier=misread); Box A<B.
- Other: discovery — read variant col headers, per-group inspect existing rows for convention, substring/word-flip/suffix/plural/accent match, walk P1 after action2. Promote recurring cuisine (2+) to own module.
