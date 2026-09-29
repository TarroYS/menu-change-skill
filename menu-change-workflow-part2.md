---
name: menu-change-workflow-part2
description: Post-callback apply of a specialist-edited Price Comparison Report → staged wsheet draft. Triggers on: Part 2, Action 4/5, apply modifier edits, final raw gsheet, final review, post-callback. Not OCR / paper-menu compare / P6 / price inference.
---

# Menu Change Workflow — Part 2 (apply)

SCOPE: Specialist-edited Price Comparison Report → staged `wsheet` draft. Sister skill `menu-change-workflow-part1` produced the report; do not OCR, do not run `price_inference.py`.
Cuisine-agnostic. Union runbook, P2-only. Scripts live in shared `menu-change-payload/` — not in this file.

## ⛔ APPLY FIDELITY LOCK (runtime — this run)
Never rewrite, minify, regenerate, or reimplement `action4` / `action5` / `modifier_rule_pricing.py apply` / `wsheet_xlsx_bridge.py` / `part2_report.py`. SCRIPT CATALOG is a RUN reference, not a spec to reimplement live. Purple Dish Item cells are already in the report — inherit them as ordinary values. Do not re-derive inference.

## ENV
- SCRIPTS ARE NOT IN THIS FILE. Shared payload: `menu-change-payload/` (`setup.sh`, `scripts/`, `references/`, `hooks/`). Materialize those paths unchanged. `chmod +x setup.sh`.
- ACE: no pip/root, PEP668. `bash setup.sh` once/cold-start → `.venv`. Invoke `.venv/bin/python scripts/<x>.py` (NOT bare python3). `.venv` does not persist.
- Cold-start every run: never reuse a prior fetch/draft/intermediate. Re-resolve, re-fetch, rebuild from the attached edited report.
- ACE inputs: edited Price Comparison Report via `get_conversation_attachments`. Raw from live `wsheet`. Standalone (no `wsheet`): deliverable is `RXXXXX_final_raw_gsheet.xlsx` instead of a draft.

## RESTAURANT RESOLUTION (mandatory, every run)
1. `find_restaurants` (by name or bare R#). Bare R# on the report `restaurant_enterprise` tab is OFTEN NOT fetchable (e.g. R88312 fails; RMCR88312 works).
2. STOP — specialist confirms name+identifier+city.
3. `wsheet fetch <exact id find_restaurants returned>`. Never fetch bare R# silently. Link restaurant to context.

## DEFINITION OF DONE
A STAGED `wsheet draft` (validate exit 0 first). Commit only on explicit user request.
Do not run `price_inference.py`. Do not OCR. Do not rebuild the comparison report.

## FILES
Price Comparison Report (edited) · Final Review (action4) · Raw Gsheet (bridge to-xlsx of fresh fetch) · Final Raw Gsheet (action5; ACE = TSV patch source).
Report already carries purple Dish Item cells and, if multi-col, a Modifier Rule tab. action5 patches `dish_item`/`modifier_item` only; modifier_rule writes are §2c.

## SCRIPT CATALOG (P2)
- `setup.sh` — `bash setup.sh`. Idempotent venv+openpyxl bootstrap.
- `wsheet_xlsx_bridge.py`
  - `to-xlsx <tsv_dir> -o raw.xlsx` — each tab.tsv→sheet, cells verbatim str (no coercion).
  - `to-tsv <final.xlsx> <tsv_dir> --tabs dish_item,modifier_item` — patch ONLY changed cells; never write blank cell, never add/remove rows, shape-mismatch tab skipped w/ warn, other TSVs byte-identical, JSON root cells unquoted. No-op=0 cells. Add `modifier_rule` to `--tabs` when §2c ran.
- `action4_apply_modifier_changes.py <report.xlsx> [-o]` — reads Modifier Summary variant prices (NOT Notes/Questionnaire/Price Inference), diffs vs orig, expands 1 summary→every replicated instance, blue/pink, resolves Excel formulas via data_only=True, carries tabs → Final Review. Purple Dish Item cells pass through as ordinary values.
- `action5_to_raw_gsheet.py <final_review.xlsx> <raw.xlsx> [-o]` — copies whole Raw untouched; patches dish_item + modifier_item price cells ($X.XX) where differ. APPLY_DISH_CODES=False (disabled — no Dish Code tab read/write). Never blank/add/remove. Auto-detects R#. Drift G8: report item not in live tabs → `!! DRIFT WARNING` (resolve before staging); clean = "Drift check…OK." Writes part2_summary.json. Applies purple cells like any other edited price.
- `modifier_rule_pricing.py` (only if the report has a Modifier Rule tab)
  - `preflight <raw.xlsx>` — MANDATORY G11 gate, must exit 0. DOUBLE_CHARGE/BROKEN/G11-violation=nonzero→STOP+escalate. UNKNOWN=exit2 (no By-Itself priced anchor → usually DELTA, don't guess an ABSOLUTE By-Itself).
  - `apply <report.xlsx> <raw.xlsx> -o final.xlsx` — writes approved out_unit_price back into modifier_rule only. Re-runs G11. G1/G2/G8-drift.
- `price_queue_gate.py <dir>` — read-only scope reporter, prints staged price cell counts, ALWAYS exit 0. No block, no marker. Never chain with `&&` expecting a gate.
- `part2_report.py <dir> --draft-id <uuid> [--restaurant-name X]` — renders staged-draft report + Slack block from part2_summary.json. exit2=JSON missing→re-run action5, never type counts.

## PIPELINE
0. `get_conversation_attachments` (edited report) → setup → resolve+confirm → fetch.
1. action4 → Final Review.
2. bridge to-xlsx→raw.xlsx → action5 → final_raw.xlsx → bridge to-tsv (--tabs dish_item,modifier_item; normalize bare prices to $X.XX).
2c IF Modifier Rule tab: preflight(exit0) → modifier_rule_pricing apply → to-tsv --tabs modifier_rule (all edits ride ONE draft).
3. `wsheet validate <dir> --level full` → exit0 mandatory. Pre-existing untouched error (e.g. R81888 TarroDelivery mismatch) still exit1 → tell user, confirm patch touched only price tabs, proceed only on explicit OK.
4. `price_queue_gate.py <dir>` (audit) → `wsheet draft create <dir> --name "…"` (or `draft update <uuid>`). Capture draft_id. Stale base → re-fetch→re-apply→re-validate→update. upload_artifact final_raw.xlsx.
5. `part2_report.py <dir> --draft-id <uuid>` → paste verbatim. STOP: staged draft = deliverable.
6. Commit ONLY on explicit request: `wsheet draft commit <uuid> --message "…"` (inline-approval gated). Reject→stop+ask (revise or archive). Stale→re-fetch/apply/validate/update/re-commit. No other route to prod. NO marker files, NO shell instructions to user.
7. Do NOT run `price_inference.py`. Purple cells are already in the report — inherit them. Specialist overwrite/blank is authoritative (blank → not written, G1).

## GUARDRAILS
- G1 NEVER write blank cells (incl code slots, group-gate FALSE rows).
- G2 NEVER add/remove rows — overwrite existing only.
- G3 Purple cells pass through as ordinary values (already inferred in P1), incl. Party Tray (T-TRAY) cells. Do not re-color, re-derive, or recompute trays.
- G4 Only touch dish_item/modifier_item (+modifier_rule via §2c). Parse by header text, not position.
- G5 Ignore Questionnaire / Notes / Price Inference as a write source. Read only Modifier Summary variant prices. Dish-code write is DISABLED — do not read or write a Dish Code tab. Never read Questionnaire.
- G6 $X.XX format; codes verbatim if ever present. Normalize bare `6.50` to `$6.50` when the live cell used `$`.
- G8 Resolve every DRIFT WARNING before staging.
- G9 Commit only explicit + only via gated `wsheet draft commit`. Default terminal = staged draft.
- G10 No `.price_queue_approved` marker, no `touch`, no shell asks (ACE has no terminal). price_queue_gate = read-only.
- G11 ABSOLUTE double-charge guard (model-aware). DELTA: base on dish_item, no By-Itself row, side=upcharge (out_unit_price=combo_total−plain). ABSOLUTE: dish_item=$0, By-Itself row carries base, side=full total (out_unit_price=paper_total). Auto-detect from By-Itself anchor (aliases: By Itself/Plain/No Rice/No Side, normalized) vs dish_item. ABSOLUTE rule: every priced modifier_rule (dish,variant) MUST have dish_item==$0 else DOUBLE-CHARGE ERROR (blocks apply/preflight, nonzero). No-op for DELTA. Never route around nonzero. UNKNOWN preflight (no By-Itself priced anchor) is usually DELTA — do not skip; do not invent an ABSOLUTE By-Itself.
- G12 No hand-authored numbers/state. Counts/report/Slack → part2_report.py stdout only. Draft state → `wsheet draft list <R#>` or draft_id from create output. No cross-run memory. If a number/status isn't in this turn's tool output, don't write it.
- G14 P2 inherits purple / P6 rows and never re-derives. Never widen inference rules.
