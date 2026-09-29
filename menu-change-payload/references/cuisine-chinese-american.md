# Chinese-American Cuisine Module

Load this module when Step 0 cuisine inference classifies the restaurant as Chinese-American. It contains the ChAm-specific OCR rules, alias patterns, and pricing sanity checks that live on top of the general `ocr-prompt.md` guidance.

This covers ~80% of restaurants in the system (No.1 Chinese, China One, Foo Yung Express, Bei Jing, Chang Xing, Jade Garden, Best Food In Town, etc.).

## Cuisine Inference Signals

You're in this module if the Menu Preview shows any of:

- Variant columns: `Order`, `Small`, `Large`, `Lunch`, `Combo`, `Specialty`
- Dish groups: `Lo Mein`, `Chow Mein`, `Egg Foo Young`, `Moo Shu`, `Combination Plates`, `Chef Specialties`, `Cantonese Style`, `Hunan & Szechuan`, `Soup`, `Egg Roll`, `Fried Rice`
- Dish names with protein-prefix patterns: `Chicken with X`, `X with Garlic Sauce`, `Hunan Chicken`, `Beef Lo Mein`
- Modifier groups like `Add Egg Roll`, `with Soup`, `Spice Level`

## Variant Column Mapping (ChAm)

| Menu pattern | OCR JSON variant key |
|---|---|
| Pt / first column in two-column section | `"Small"` |
| Qt / second column in two-column section | `"Large"` |
| Single-price appetizers, sides, entrees | `"Order"` |
| Specialties "Plain" column | `"Specialty"` |
| Lunch Special prices | `lunch_specials` section with `"price"` field |
| Combination prices | `combinations` section with `"price"` field |
| Chef / Cantonese specialty single prices | `"Order"` |
| Party Tray catering | `"Party Tray"` |

**Party Tray (v4.3): capture every printed tray price.** They are no longer ignored — action3 compares them and Part 2 writes them. Key them `"Party Tray"` only (never `"Order"`, or a $60 tray lands in a dish slot). Unprinted trays are derived by `price_inference.py` T-TRAY when the restaurant's trays follow one formula (e.g. `Large × 5 − $2.00`). Whole-dollar tray prices ≥ $100 are normal and no longer raise `DROPPED DEC`.

**Variant-style detection.** The Action 2 script auto-detects whether this restaurant uses:
- `size_based` — Pt/Qt → Small/Large columns. Most common.
- `lunch_based` — first column → Lunch (when the menu shows Lunch/Dinner two-column sections instead of Pt/Qt).
- `size_based (dual)` — some items have BOTH Small/Large AND Lunch prices in the system (e.g., Best Food In Town R81001). Pt prices go to Small; Lunch Specials route independently via the `lunch_specials` JSON key.

**Write the variant key the section actually uses** — `"Small"`/`"Large"` for genuine cup/bowl or Pt/Qt portions (Soup, White Rice), `"Order"` for single-price items. The Small→Lunch routing is **section-aware**: under `lunch_based`, the script reroutes a paper `"Small"` to `Lunch` **only on system rows that actually have a `Lunch` slot and no `Small` slot** (i.e. mains). Rows that genuinely use Small/Large — Soup, White Rice — keep their `"Small"` price. (Previously the remap was global, which silently dropped legitimate Soup `Small` prices because the remapped `Lunch` value had nowhere to land — fixed in `action2_ocr_to_2b_0625v2.py`.)

## Compound Item Expansion (ChAm-specific patterns)

### Moo Shu
"Moo Shu (Vegetables, Pork or Chicken)" at one price → THREE independent dishes:
- `Moo Shu Vegetable`
- `Moo Shu Fresh Pork`
- `Moo Shu Chicken`

Each at the same price. Expand in the OCR JSON as separate items.

"Moo Shu (Beef or Shrimp)" → `Moo Shu Beef` + `Moo Shu Shrimp`, same price.

### Egg Foo Young
Reverse pattern — the parenthetical comes BEFORE the dish name:
"(Vegetable, Chicken or Roast Pork) Egg Foo Young" → THREE dishes:
- `Vegetable Egg Foo Young`
- `Chicken Egg Foo Young`
- `Pork Egg Foo Young`

### Chow Mein / Chop Suey
Items like "Beef Chow Mein / Chop Suey" — keep the COMBINED name as one OCR entry. The script auto-matches to BOTH the Chow Mein AND Chop Suey rows in the system.

### Multi-protein (with commas + "or")
"Broccoli with Chicken, Pork, Beef or Shrimp" → 4 separate rows.
"Hunan Style Chicken or Beef or Shrimp" → 3 rows.
Enter the FULL name as printed; the script splits automatically. Do NOT manually split in the OCR JSON.

### Bare protein names under section headers
Under a section header like "Stir Fried Udon" with rows that just say "Vegetable", "Chicken", "Beef", "Shrimp" — enter these exactly as printed. The script will prepend the section name to form the full dish name (e.g., "Stir Fried Udon with Vegetable"). Make sure the `menu_section` field accurately captures the section header.

### Suffix sections — emit `<Protein> <SectionWord>`, not bare protein
Some ChAm sections store the section word as a SUFFIX. A blind prepend produces the wrong order and lands in P1. For these, write the full `<Protein> <SectionWord>` name in the OCR JSON:
- **Pancit** (often in the Mai Fun group): `Singapore` → `Singapore Pancit`; `Chicken` → `Chicken Pancit`; `House Special` → `House Special Pancit`.
- **Yat / Yat Gaw Mein**: `Chicken` → `Chicken Yat`; `Beef` → `Beef Yat`. Note the system uses `House Yat`, NOT `House Special Yat` — drop `Special` here.
- `Chicken or Pork` on one line → emit both `Chicken Yat` + `Pork Yat` (or the matching section word).

(Lo Mein, Chow Mein, Egg Foo Young, Fried Rice already work as `<Protein> <Section>` — emit full names there too for safety.)

### Section-name aliasing for irregulars
Beyond Diet→Healthy Dishes and Combination Plates→Combo (below), these paper section headers don't match the system group name:
- **Stir Fried Udon** (paper) → system **Stir Fry Udon Noodle** suffix: `Chicken` → `Chicken Stir Fry Udon Noodle`. **Irregular exception:** the `Vegetable` item is `Vegetable Stir Fry Noodle` (no "Udon") — emit it that way.
- **Sweet & Sour** (paper, single item) → system group **Sweet Sour**, dish `Sweet and Sour Chicken`.
- **Chef's Specialties** (paper) → **Chef Specialties** (no apostrophe).
- **Specialties** (paper) → **Fried Specialties** (system) — this section usually has multi-column side pricing.

### Sub-variant "Choice of sauce" → dish-name map (ChAm wings)
For the S2 "Chicken Wings w. Sauce — Choice of …" pattern (see `ocr-prompt.md`), map each sauce choice to its system dish:
- `Honey Sauce` → `Honey Wings`
- `Hot Wings` / `Buffalo Sauce` → `Buffalo Wing`
- `BBQ Sauce` → `Barbecue Wing`
- `General Tso's Sauce` → `General Tso Wing`
- `Garlic Sauce` → `Wing in Garlic Sauce` (for wings; elsewhere the appropriate `X with Garlic Sauce`)

### Diet Menu / Special Diet Menu / Healthy Dishes
The paper menu may label a section as "Diet Menu" or "Special Diet Menu". In the system, this is called "Healthy Dishes". Use the section name exactly as printed — the script auto-maps any section starting with "Diet".

### Special Combination Plates / Combination Plates / Combination Dinners
Combos appear under various headers. Place these in the `combinations` JSON key with `"price"` per item. The script auto-detects these section names and routes prices to the `Combo` variant column.

### Drinks in Side Orders
Beverages (tea, soda, juice, water) under "Side Orders" on the paper menu map to "Drinks" in the system. Include them in the normal `sections` array — the script handles the mapping.

## Protein Tier Pricing Sanity Check (ChAm)

ChAm menus follow consistent pricing tiers within a dish family:

1. **Vegetable** dishes — lowest tier
2. **Pork / Chicken** dishes — middle tier (usually identical price)
3. **Beef / Shrimp** dishes — higher tier
4. **House Special / Seafood** — highest tier

If a price breaks this pattern (e.g., Pork costs more than Beef in the same section), re-verify the digit. This is the most common OCR error pattern for ChAm menus.

## Alias Patterns (ChAm)

The script handles these aliases automatically, but avoiding duplicates in the OCR JSON helps the matcher:

- `Garlic Sauce with Chicken` = `Chicken with Garlic Sauce`
- `Broccoli with Chicken` = `Chicken with Broccoli`
- `Hunan Style with Chicken` = `Hunan Chicken`
- `Szechuan Style with Chicken` = `Szechuan Chicken`
- `Lo Mein with Chicken` = `Chicken Lo Mein`
- `Egg Foo Young with Chicken` = `Chicken Egg Foo Young`
- `Sweet and Sour Chicken` (Lunch Special) = same row as `Sweet and Sour Chicken` (regular)
- `Spare Rib Tip` = `Barbecue Spare Rib Tips`
- `Chow Mei Fun` → system often drops "Chow" → `Mei Fun`
- `Plain Fried Rice` = `Fried Rice`
- `Roast Pork w. X` = `Pork with X` (Pork section drops the leading `Roast`: `Roast Pork w. Broccoli` → `Pork with Broccoli`)
- `Bean Curd Szechuan Style` = `Szechuan Bean Curd` (strip trailing `Style`, then word-flip)
- `Bean Curd w. Home Style` = `Bean Curd Home Style`
- `Scallops & Beef` = `Beef with Scallops`; `Fried Bean Curd w. Broccoli` = `Broccoli with Bean Curd` (`A & B`/`A with B` ↔ reordered `B with A`)
- `Beef in Garlic Sauce` = `Beef with Garlic Sauce` (`in` → `with`)
- `Bourbon Chicken` → `Street Bourbon Chicken` (paper name is a complete suffix of one system row)

**Phonetic / spelling drift** (treat as the same dish):
- `Shu mai` ↔ `Shao Mai`
- `Kung Po` ↔ `Kung Pao` ↔ `Gong Bao`
- `Hunan` ↔ `Hu Nan`; `Szechuan` ↔ `Sichuan`
- `Tso` ↔ `Tso's` ↔ `Tsos` (drop apostrophe, optional trailing `s`)

If the dish is the same food item, merge prices into the existing row.

## Lunch Special Format (ChAm)

- Lunch specials are typically NOT in a standard table format.
- The price is usually displayed as a SINGLE prominent callout (e.g., "$8.95" in a cloud/banner) that applies to ALL items in the section.
- When OCR-ing: find the single displayed price, then record each item with that same price in the `lunch_specials` JSON section using the `"price"` field (single value), NOT the `"prices"` dict.
- Some restaurants have tiered lunch pricing (e.g., $7.95 chicken, $8.95 shrimp). If individual prices appear next to items, use those per-item prices instead.

## Notes for the OCR Pass

1. **Count price columns per section BEFORE extracting.** Some sections have Pt/Qt (two columns); some have just one Order price.
2. **Cross-verify against neighboring items** in the same section — most ChAm sections have internally consistent pricing tiers.
3. The Action 2 script's name resolution pipeline handles `w.`→`with`, `&`→`and`, abbreviations (`veg`→`vegetable`, `bbq`→`barbecue`), hyphens→spaces, plural toggles, word-order flips. You don't need to normalize names in the OCR JSON.