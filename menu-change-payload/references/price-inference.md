# Price inference (P1 Step 3a — every restaurant)

Run **after Action 3 and before** `modifier_rule_pricing.py compare`. Not cuisine-conditional. Do **not** run in P2.

```bash
.venv/bin/python scripts/price_inference.py <menu_preview.xlsx> <ocr_output.xlsx> \
    --report <price_comparison_report.xlsx> --mode apply
```

Script default is `--mode suggest` (no Dish Item writes). This workflow uses **`--mode apply`**.

## Gaps it fills

| | Scenario | Action 3 alone | Result |
|---|---|---|---|
| **V-GAP** | Dish IS on the paper menu, but not every system variant is printed | keeps the system price, no highlight | stale price ships |
| **D-GAP** | Dish has no paper match (not on the paper menu) but is still sold | old price unchanged | stale price ships |

A dish group is priced as a family: dishes that shared an old price almost always share the new one.

```
Chicken with Broccoli     $8.95 → $9.95   (printed — donor)
Chicken with Cashew Nut   $8.95 → ?       (not printed — recipient)
                          ^^^^^ same dish group, same variant, same old price
⇒ infer $9.95
```

## Tiers (first match wins)

- **T0 — same-price rules (maintainer, 2026-09-30), tried first.** Business rules, the only sanctioned widenings besides T-TRAY (G14):
  - **General Tso / Orange / Sesame** `<Chicken | Beef | Shrimp>` always cost the same. An unprinted one takes its printed siblings' price (same protein, same variant), **whatever its old price was**, in any dish group. Names are normalized ("General Tso's", "Gen. Tso", "General Tsos" all match). A bracketed qualifier makes its own family: "Orange Chicken (White Meat)" follows "General Tso Chicken (White Meat)", never the plain one. Tier `T0 same-price family (General Tso / Orange / Sesame)`, HIGH.
  - **Healthy** (dish group containing "healthy"): a "Steam X" / "Steamed X" that is **not on the paper menu at all** takes the new price of its regular "X" (a non-healthy dish with the same normalized name) **only when both OLD prices were the same**, per variant. Tier `T0 healthy = regular dish (same old price)`, HIGH. A Healthy dish the paper prints (even one size, only its Party Tray, or only a size it does not have — Questionnaire P-SIZE), or one that looks like an unmatched P1 row ("Diet Shrimp w/ Broccoli (Steamed)", or the row action2 suggests it for), counts as on the paper and is not given its twin's price; old prices differ / no regular twin → falls through to T1 within the Healthy group. The skip reason names the twin and both old prices. (Maintainer, 2026-09-30: confirmed on R81587 — the two Healthy cells the specialist left unchanged were missed, the rule was right.)
  - Printed family members that disagree = **FAMILY CONFLICT** (`T0 conflict`): no guess, red on the Price Inference tab, callback. T1 is not tried after a T0 conflict.
- **T1 — exact old-price cohort.** Same `(dish group, variant)` — **Chow Mein and Chop Suey count as one dish group** (a Chop Suey takes a Chow Mein's new price at the same old price, and vice versa; tier label ends `(Chow Mein + Chop Suey)`). Donors = dishes that DID get a paper price and whose **old** price equals the recipient's old price. All donors agree → that price. Confidence HIGH at ≥2 donors, MEDIUM at exactly 1. Donors disagree = **COHORT CONFLICT**: no guess, report for callback.
- **T2 — uniform group delta.** Only when the recipient's old price has **no** donor cohort. Every donor in that `(dish group, variant)` moved by the **same** delta and there are ≥ `--tier2-min-donors` (default 3) → apply that delta. Confidence MEDIUM.
- Else **SKIPPED** with reason. Never guessed.
- **A dish listed under two dish groups is ONE dish:** it counts once as a donor, and its rows must agree — one row inferred and the others without evidence → every row takes it; rows that infer different prices, or any row in conflict → no row is inferred (`conflict (same dish, several groups)`, red).
- **Never** a recipient (any tier, incl. T-TRAY): a cell whose system price is **$0.00** — a free item ("Free Egg Roll", utensils) or a dish priced on its modifiers. SKIPPED; a whole group moving +$0.50 says nothing about a freebie.
- **Never** a donor either: a $0.00 cell listed in the Menu Preview's `absolute_pricing` tab — the dish's price lives on modifiers (ABSOLUTE), so an inferred dish price would be charged on top (G11).

## T-TRAY — Party Tray formula (v4.3)

Party Tray is **not** a T1/T2 variant. It has its own tier, because a tray is priced as a formula off a base column on the same row, not as a family with other trays.

```
System (old):  General Tso's  Large $12.95   Tray $62.75
               Pepper Steak   Large $14.95   Tray $72.75   ← both = Large × 5 − $2.00
Paper (new):   General Tso's  Large $13.95   Tray (not printed)
⇒ Tray = 13.95 × 5 − 2.00 = $67.75  (purple)
```

**Detect — per dish group** (maintainer, 2026-10-05: trays usually follow one formula across the menu, sometimes one per group — "be safe, do it per group"). For each dish group, fit `Tray = Base × k + d` on the system's OLD prices over the group's rows with a tray and ≥1 other price. `k` ∈ 1.0–10.0 (0.5 steps), `|d|` ≤ `--tray-max-offset` ($10). A group's formula is trusted only if ONE formula fits **≥ 70%** of its rows (`--tray-min-agreement`), ≥ 3 rows (`--tray-min-rows`), with ≥ 2 distinct base prices. Two formulas tied = **AMBIGUOUS** (not picked, G14).

A group too small or too mixed to confirm its own formula uses the **whole-menu** formula (the same fit over every tray row) only when that fits ≥ 70% of the group's OWN tray rows — otherwise the group has no formula and its changed trays are asked on the callback. A dish listed under two groups gets a derived tray only when EVERY group's formula derives the same one — a group whose formula gives another tray, does not fit the dish (outlier), is missing, or says "no change" disagrees, and then no row is derived (`conflict (same dish, several groups)`, asked); the same tray held back in one group (suggest-only) is held back in all. P-TRAY-0, the Price Inference tab and the Summary show one formula line per group.

**Apply** — per unprinted tray:

| Row situation | Outcome |
|---|---|
| Base printed + changed, row follows formula | **APPLIED** purple (decrease / cap → SUGGESTED) |
| Base printed, unchanged | NO CHANGE EXPECTED |
| Row does not follow formula (outlier) | SKIPPED → P-TRAY callback, with what the formula *would* give |
| Base inferred by T1/T2, not printed | SUGGESTED only (never chained into a write) |
| The row's dish group has no formula (and the whole-menu one does not fit the group), base changed | SKIPPED → P-TRAY callback |
| Tray-only dish (no base on row) | SKIPPED |

Printed trays are compared directly by action3 (blue/pink/green) — T-TRAY never touches them. Absolute cap scales with `k` (`--max-delta-abs × k`). Disable for a run with `--no-tray`.

T-TRAY only runs when the restaurant actually has tray prices (every Menu Preview carries a `Party Tray` header from the action1 template, usually empty).

## Deliberate non-feature

A dish's own printed `Small` delta is **not** copied onto its unprinted `Large`. Restaurants routinely raise Pt by $0.50 and Qt by $1.00. Inference only moves **across dishes, within one variant**. Do not "improve" this into a row-wise fill.

**Sole exception: T-TRAY.** It copies no delta — it fits a multiplicative formula gated on ≥70% restaurant-wide agreement. Do not extend the exception to any other variant pair.

## Safety (auto-downgrade to suggest-only, never silently applied)

- inferred delta is a **decrease**
- `|delta|` > `--max-delta-abs` (default $5.00) or `--max-delta-pct` (default 40%)
- would break a `Small < Large` ordering that held in the system (checked against the row's NEW printed prices)
- dish group is operational (Drink / Sauces / Side Order) unless `--include-low-priority`

`--mode apply` writes eligible cells with a **purple** fill (as plain numbers, like every report price) and refreshes the hidden `_ref_` mirror so apricot CF does not false-fire. Part 2 `action5` then ships purple cells like any other value — that is why every purple cell / Price Inference row is a **mandatory callback item** (G14).

## What it adds to the report

- **Price Inference** tab (purple tab color) — one row per unprinted price: Scenario, Status, old → inferred, Delta, Tier, Confidence, donor count, Evidence, Reason.
- **Questionnaire → P-TRAY** — `P-TRAY-0` states the detected formula (or why none); then one row per tray needing a callback answer. Replaces the old static `A-1 Party Tray pricing` question. Only written when the restaurant prices trays.
- **Summary → AUTO-INFERENCE** block + purple legend.
- `price_inference_summary.json` — quote these counts (G12). Never hand-count. Tray counts: `stats.tray_*`; formula: `tray_formulas[]`.

A size the paper printed two different prices for (action2 `Price Conflicts`, Questionnaire P-PRICE) is never inferred, never a donor, and no Party Tray is derived over it; a tray whose base is such a size is asked. No P6 Questionnaire section any more (removed 2026-09-30): the Price Inference tab + purple cells are the callback list. Re-run is idempotent (tab / P-TRAY / Summary block replaced, not duplicated): it first puts its earlier purple cells back to the system price (the hidden `_ref_sys_` mirror), so a re-run with other options never keeps a guess it no longer makes. A cell the specialist typed over is never undone or overwritten. It writes only into a report row that is the same dish as the Menu Preview row; a stale report (rows shifted) gets suggestions only. T-TRAY agreement counts dishes, not rows (a dish in two groups counts once).

## Multi-column sides (already in `modifier_rule_pricing.py compare`)

No extra script. Source column on Fried Specialty propagation:

| Source | Layer |
|---|---|
| `PAPER MENU` | printed — always keeps its own printed price |
| `PROPAGATED` | same dish, same old `out_unit_price`, only when every printed member of that group agrees |
| `PRICE TWIN (group)` | other dish, **same dish group**, same `modifier_item` + old price |
| `PRICE TWIN` | same rule, global fallback |
| `NO DATA` | no safe pattern; keep current. Never auto-resolved, all flagged for the callback: `GROUP CONFLICT` (printed members of one old-price group disagree → unprinted members kept), `PAPER CONFLICT` (same dish printed differently in two sections), `TWIN CONFLICT` (twins disagree) |

Every multi-column section is processed. Under ABSOLUTE the paper Plain goes on the dish's By Itself rule (PAPER MENU) and each side rule takes the full printed total.

What Step 3a adds on top is the **base** price of a multi-column DELTA dish the paper omitted (the Specialty/Order slot the side deltas are computed against). Under ABSOLUTE the base is the By Itself rule, handled by compare — never dish_item.

## P2

Inherit purple cells. Do not re-run this script. Specialist overwrite/blank is authoritative (blank → G1, not written; `x` → price removed, the G1 exception).
