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
| **D-GAP** | Dish has no paper match (Questionnaire P3) but is still sold | old price unchanged | stale price ships |

A dish group is priced as a family: dishes that shared an old price almost always share the new one.

```
Chicken with Broccoli     $8.95 → $9.95   (printed — donor)
Chicken with Cashew Nut   $8.95 → ?       (not printed — recipient)
                          ^^^^^ same dish group, same variant, same old price
⇒ infer $9.95
```

## Tiers (first match wins)

- **T1 — exact old-price cohort.** Same `(dish group, variant)`. Donors = dishes that DID get a paper price and whose **old** price equals the recipient's old price. All donors agree → that price. Confidence HIGH at ≥2 donors, MEDIUM at exactly 1. Donors disagree = **COHORT CONFLICT**: no guess, report for callback.
- **T2 — uniform group delta.** Only when the recipient's old price has **no** donor cohort. Every donor in that `(dish group, variant)` moved by the **same** delta and there are ≥ `--tier2-min-donors` (default 3) → apply that delta. Confidence MEDIUM.
- Else **SKIPPED** with reason. Never guessed.

## T-TRAY — Party Tray formula (v4.3)

Party Tray is **not** a T1/T2 variant. It has its own tier, because a tray is priced as a formula off a base column on the same row, not as a family with other trays.

```
System (old):  General Tso's  Large $12.95   Tray $62.75
               Pepper Steak   Large $14.95   Tray $72.75   ← both = Large × 5 − $2.00
Paper (new):   General Tso's  Large $13.95   Tray (not printed)
⇒ Tray = 13.95 × 5 − 2.00 = $67.75  (purple)
```

**Detect** — fit `Tray = Base × k + d` on the system's OLD prices over every row with a tray and ≥1 other price, restaurant-wide. `k` ∈ 1.0–10.0 (0.5 steps), `|d|` ≤ `--tray-max-offset` ($10). Trusted only if ONE formula fits **≥ 70%** of those rows (`--tray-min-agreement`), ≥ 3 rows (`--tray-min-rows`), with ≥ 2 distinct base prices. Two formulas tied = **AMBIGUOUS** (not picked, G14).

**Apply** — per unprinted tray:

| Row situation | Outcome |
|---|---|
| Base printed + changed, row follows formula | **APPLIED** purple (decrease / cap → SUGGESTED) |
| Base printed, unchanged | NO CHANGE EXPECTED |
| Row does not follow formula (outlier) | SKIPPED → P-TRAY callback, with what the formula *would* give |
| Base inferred by T1/T2, not printed | SUGGESTED only (never chained into a write) |
| No formula detected / ambiguous, base changed | SKIPPED → P-TRAY callback |
| Tray-only dish (no base on row) | SKIPPED |

Printed trays are compared directly by action3 (blue/pink/green) — T-TRAY never touches them. Absolute cap scales with `k` (`--max-delta-abs × k`). Disable for a run with `--no-tray`.

T-TRAY only runs when the restaurant actually has tray prices (every Menu Preview carries a `Party Tray` header from the action1 template, usually empty).

## Deliberate non-feature

A dish's own printed `Small` delta is **not** copied onto its unprinted `Large`. Restaurants routinely raise Pt by $0.50 and Qt by $1.00. Inference only moves **across dishes, within one variant**. Do not "improve" this into a row-wise fill.

**Sole exception: T-TRAY.** It copies no delta — it fits a multiplicative formula gated on ≥70% restaurant-wide agreement. Do not extend the exception to any other variant pair.

## Safety (auto-downgrade to suggest-only, never silently applied)

- inferred delta is a **decrease**
- `|delta|` > `--max-delta-abs` (default $5.00) or `--max-delta-pct` (default 40%)
- would break a `Small < Large` ordering that held in the system
- dish group is operational (Drink / Sauces / Side Order) unless `--include-low-priority`

`--mode apply` writes eligible cells with a **purple** fill and refreshes the hidden `_ref_` mirror so apricot CF does not false-fire. Part 2 `action5` then ships purple cells like any other value — that is why every purple/P6 row is a **mandatory callback item** (G14).

## What it adds to the report

- **Price Inference** tab (purple tab color) — one row per unprinted price: Scenario, Status, old → inferred, Delta, Tier, Confidence, donor count, Evidence, Reason.
- **Questionnaire → P6** — callback list (applied, suggested, cohort conflicts). Trays excluded (own section).
- **Questionnaire → P-TRAY** — `P-TRAY-0` states the detected formula (or why none); then one row per tray needing a callback answer. Replaces the old static `A-1 Party Tray pricing` question. Only written when the restaurant prices trays.
- **Summary → AUTO-INFERENCE** block + purple legend.
- `price_inference_summary.json` — quote these counts (G12). Never hand-count. Tray counts: `stats.tray_*`; formula: `tray_formulas[]`.

Re-run is idempotent (tab / P6 / Summary block replaced, not duplicated).

## Multi-column sides (already in `modifier_rule_pricing.py compare`)

No extra script. Source column on Fried Specialty propagation:

| Source | Layer |
|---|---|
| `PAPER MENU` | printed |
| `PROPAGATED` | same dish, same old `out_unit_price` |
| `PRICE TWIN (group)` | other dish, **same dish group**, same `modifier_item` + old price |
| `PRICE TWIN` | same rule, global fallback |
| `NO DATA` | no safe pattern; keep current. Twin disagreements = `TWIN CONFLICT` (not auto-resolved) |

What Step 3a adds on top is the **base** price of a multi-column dish the paper omitted (the Specialty/Order slot the side deltas are computed against).

## P2

Inherit purple cells. Do not re-run this script. Specialist overwrite/blank is authoritative (blank → G1, not written).
