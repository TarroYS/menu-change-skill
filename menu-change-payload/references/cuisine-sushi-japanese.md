# Sushi / Japanese Cuisine Module

Load this module when Step 0 cuisine inference classifies the restaurant as sushi/Japanese. It contains the sushi-specific OCR rules, variant column conventions, alias patterns, cross-section duplicate rules, and pricing sanity checks.

Tested restaurants in this category: SeaPark Sushi (ER85777).

## Cuisine Inference Signals

You're in this module if the Menu Preview shows any of:

- Variant columns: `Roll`, `Regular Roll`, `Dinner`, `Bowl`, plus a restaurant-specific combo column (e.g., `Seapark`, `Hokkaido`, `Sakura`)
- Dish groups: `Lobster Roll`, `Crab Sub Roll`, `Basic Roll 1/2`, `Special Roll 1/2`, `Spicy Roll`, `Veggie Roll`, `Sushi Bar Entrees`, `Sushi Sashimi Entrees Don`, `Donburi`, `Poke Bowl`, `Bento Box`, `Hibachi`, `Yaki Udon`, `Party Tray Rolls`, `Seapark Sets`
- Dish names containing Japanese terms: `Tekka`, `Sake`, `Unagi`, `Chirashi`, `Yaki Udon`, `Hibachi`, `Tempura`, `Tamago`, `Tonkatsu`, `Bulgogi`, `Maki`, `Nigiri`, `Sashimi`
- Modifier groups like `Premium Add Ons` (avocado, cream cheese, tempura crunch), `Sauce` group, `Roll Add-Ons`

## Variant Column Mapping (Sushi)

| Menu pattern | OCR JSON variant key |
|---|---|
| Sushi rolls (nigiri, maki, base rolls, spicy rolls, classic rolls, special rolls) | `"Roll"` |
| Party Tray "Any N Regular Rolls" bundles (Any 3/6/9 Rolls) | `"Regular Roll"` |
| Dinner Bento Boxes | `"Dinner"` |
| Restaurant-specific combo sets (e.g., "SEAPARK A/B/C/D", "HOKKAIDO PARTY") | `"Seapark"` or whatever the system column is named — use exact header |
| Poke Bowls (only if the system separates `Bowl` from `Order` — many systems use `Order`) | `"Bowl"` or `"Order"` — check the system row |
| Sushi/sashimi platters, A La Carte sushi/sashimi, Donburi, Combo Plates, Lover Plates | usually `"Order"` |
| Lobster Rolls (sandwiches, despite "Roll" in the name) | usually `"Order"` |
| Crab Sub Rolls (sandwiches) | usually `"Order"` |
| Kitchen Entrees, Hibachi Dinner, Yaki Udon | usually `"Order"` |
| Soups, salads, appetizers | usually `"Small"` or `"Order"` |
| Lunch Specials (Sushi/Sashimi/Bento/Teriyaki Lunch + Any 2/3 Rolls Lunch) | `"Lunch"` |

**Important:** "Roll" in the name doesn't always mean the `Roll` variant column. Lobster Rolls and Crab Rolls are sub sandwiches; they live under `Order`. True sushi rolls (Tuna Roll, Dragon Roll, California Roll) live under `Roll`. Always verify against the system row.

**Restaurant-specific combo columns vary.** SeaPark uses `Seapark`; other chains use their own column name. Look at the actual header in the Menu Preview and use it verbatim in the OCR JSON.

## Sushi Protein / Preparation Tier Sanity Check

Sushi menus follow consistent pricing tiers within a fish family:

1. **Veggie / cucumber / avocado rolls** — lowest tier ($7-8)
2. **Salmon / tuna sushi & sashimi** — middle tier ($8-10)
3. **Yellowtail / eel / white tuna** — slightly higher tier ($9-11)
4. **Specialty rolls (named rolls)** — varies $15-25
5. **Lobster rolls / chef specialty rolls** — highest tier ($23-30)

If a price breaks this pattern (e.g., a cucumber roll priced higher than a salmon roll), re-verify the digit.

## Cuisine-Specific Pricing Patterns

### Lobster Roll Uniformity
Restaurants commonly price ALL lobster roll variants at a single uniform price (e.g., all $26 at SeaPark). If most lobster rolls are at one price and one shows different, re-read it — that one is almost certainly an OCR error.

### Crab Roll Uniformity
Same pattern — all Crab Sub Roll variants typically at one price.

### Bento Box A vs B Pattern
- **Box A** typically has simpler sides (mixed green salad, shumai, California roll). Cheaper.
- **Box B** typically has tempura/extra sides (kani salad, shumai, shrimp tempura roll). More expensive by $2-4.
- If you read Box A > Box B in the same protein, you likely swapped them.

### Bubble Tea / Lemonade Uniformity
All bubble teas typically priced identically (e.g., all $9). Same for all lemonades. Check uniformity within a drink family.

### A La Carte Sushi/Sashimi Pricing
The header may say "À La Carte (2 pcs Sushi/Sashimi)" — the price quoted is the 2pc price. Tuna/Salmon/White Tuna/Yellowtail typically uniform; Eel and Ikura (Salmon Roe) usually $1-2 higher; Avocado, Egg (Tamago), Crab Stick (Kani) typically lower.

## Alias Patterns (Sushi)

The system often uses slightly different naming than the paper menu. Check the Menu Preview BEFORE OCR-ing to spot these:

### Crab Sub Roll family
**Paper:** `Traditional Crab Roll`, `Cajun Crab Roll`, `Lemon Pepper Crab Roll`, `Spicy Crab Roll`, `Truffle Crab Roll`, `Cilantro Crab Roll`
**System:** Adds "Sub" — `Traditional Crab Sub Roll`, `Cajun Crab Sub Roll`, etc.

Use the system's exact "Sub" naming in the OCR JSON.

### Hawaii vs Hawaiian
- Paper: `Hawaiian Roll` ↔ System: `Hawaii Roll` (or vice versa)

### Punctuation variants
- `Rock'n' Roll` / `Rock'n Roll` ↔ `Rock N Roll`
- `Spicy Lover's Roll` ↔ `Spicy Lovers Roll` (the script normalizes apostrophes/plurals automatically)

### Spelling drift
- `Sunyflower Roll` ↔ `Sunny Flower Roll`
- `Smoked Salmon` ↔ `Smoke Salmon` (system may drop the "d")

### Soup suffix
- `Clam Chowder` ↔ `Clam Chowder Soup`
- `Lobster Bisque` ↔ `Lobster Bisque Soup`

The system may append "Soup" to all chowders/bisques.

### Symbol normalization (script handles automatically)
- `Sushi & Sashimi` ↔ `Sushi and Sashimi`
- `Kani & Seaweed Salad` ↔ `Kani and Seaweed Salad`
- `Beefin' Time Bowl` ↔ `Beefin Time Bowl`

## Cross-Section Duplicate Mapping (CRITICAL for sushi)

A single item on the paper menu can map to MULTIPLE system rows across different dish groups. Capture the item in the OCR JSON for ALL relevant system names.

### Most common case — Party Tray vs Sushi/Sashimi Entrees Don

The paper menu's "Party Tray → For Two/Four/Six" sushi packages typically appear in the system as BOTH:

1. **The Party Tray section** (where the paper menu shows them) — under "Party Tray Rolls" or similar dish group
2. **Individual entries under `Sushi Sashimi Entrees Don`** with these exact names:
   - `Sushi For 2`, `Sushi For 4`
   - `Sashimi For 2`, `Sashimi For 4`
   - `Sushi and Sashimi For 2`, `Sushi and Sashimi For 4`, `Sushi and Sashimi For 6`

When you see Party Tray "For 2/4/6" entries on the paper menu, add OCR JSON entries for BOTH locations. Otherwise the Sushi Sashimi Entrees Don rows silently keep their old prices.

### Rolls in multiple groups

Some named rolls appear in two dish groups at the same price:
- `Spicy Lobster Roll` may appear in both "Lobster Roll" and "Basic Roll 1"
- `NJ 08226 Roll`, `Hawaiian Roll` may appear in both "Basic Roll 2" and "Special Roll 1"
- `California Roll`, `Crispy Veggie Roll` may appear in both "Basic Roll 1" and "Special Roll 1"

Map the paper price to all occurrences. The OCR JSON can have the same item listed under both menu sections (e.g., one entry under "Special Rolls", one under "Classic Roll") — the script matches independently.

### General rule for sushi

Before finalizing the OCR JSON, search the Menu Preview for every paper menu item by its core name (e.g., "Spicy Lobster", "California", "Hawaii") and check if it appears in more than one dish group. If yes, include all locations in the OCR JSON.

## Common Sushi Section Mapping

| Paper menu section header | Likely system dish group(s) | Variant column |
|---|---|---|
| Soup | Soup | Small |
| Salads | Salads | Small |
| Salad Add-Ons (modifier) | Salad add-on modifiers | N/A (modifier tab) |
| Kitchen Appetizers / Sushi Bar Appetizers | Appetizers | Order |
| Kitchen Entrees | Kitchen Entrees | Order |
| Hibachi Dinner | Hibachi | Order |
| Yaki Udon | Yaki Udon | Order |
| Dinner Bento Box | Dinner Bento Box | Dinner |
| Lunch Special | Lunch Specials | Lunch |
| Lobster Rolls / Lobster Roll Entrees | Lobster Roll | Order |
| Crab Roll | Crab Sub Roll | Order |
| Base Rolls | Basic Roll 1 | Roll |
| Spicy Rolls | Basic Roll 1 (Spicy) | Roll |
| Veggie / Simple Rolls | Basic Roll 1 (Veggie) | Roll |
| Classic Roll | Basic Roll 2 | Roll |
| Special Rolls | Special Roll 1 / Special Roll 2 | Roll |
| Sushi Bar Entrees — Combo Plates | Sushi Sashimi Entrees Don | Order |
| Sushi Bar Entrees — Lover Plates | Sushi Sashimi Entrees Don | Order |
| Donburi | Sushi Sashimi Entrees Don | Order |
| Poke Bowl | Poke Bowl | Order (or Bowl, check system) |
| À La Carte (2 pcs Sushi/Sashimi) | Sushi A La Carte / Sashimi A La Carte | Order |
| Party Tray Rolls (Any 3/6/9) | Party Tray Rolls | Regular Roll |
| Party Tray For 2/4/6 | (1) Party Tray Sets AND (2) Sushi Sashimi Entrees Don | varies — see cross-section rule above |
| Seapark Sets (or restaurant equivalent) | Seapark Sets | Seapark (or restaurant-specific column) |
| Drinks $9 (Bubble Tea / Lemonade) | Drinks | Order |

This table is illustrative — always verify against the actual Menu Preview before relying on it.

## Notes for the OCR Pass

1. **Variant column lookup is critical.** Putting a sushi roll price under `Order` when the system uses `Roll` means the update is silently dropped (Action 2 writes nothing). Always look at the existing system row for the dish group and mirror its variant column.
2. **"Roll" in the name ≠ Roll variant column.** Lobster Rolls and Crab Sub Rolls are sandwiches; they use `Order`.
3. **Cross-section check.** Before finalizing the OCR JSON, scan the Menu Preview for each paper menu item by its core noun and check for duplicates across dish groups. The Party Tray ↔ Sushi Sashimi Entrees Don case is the most common miss.
4. **Look at uniformity.** If 13 of 14 lobster rolls are $26 and one shows $20, re-read the $20 one.