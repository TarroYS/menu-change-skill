# Other Cuisines — Deterministic Discovery Playbook

Load this module when Step 0 cuisine inference does NOT classify the restaurant as Chinese-American or sushi/Japanese. Examples: Mexican, Italian, pizza, Korean, Thai, Vietnamese, Indian, Mediterranean, fusion, etc.

This module doesn't contain memorized cuisine rules — instead, it teaches a **discovery technique** for inferring the right OCR conventions from the Menu Preview itself. If you run this skill on enough restaurants of a new cuisine, the patterns you discover should be promoted to a dedicated cuisine module.

## The Discovery Technique (4 steps)

### Step 1: Read the variant column headers

Open the Menu Preview's `price` tab, look at row 1. List every column that isn't `Dish Group`, `Dish Name:Primary`, or `Dish Name:Secondary`. These are the variant columns the system uses for this restaurant.

Examples you might see:
- `Order`, `Small`, `Large` — generic
- `Pizza Small`, `Pizza Medium`, `Pizza Large`, `Pizza XL` — pizza
- `Half Tray`, `Full Tray` — Italian / catering
- `Mild`, `Medium`, `Hot` — wings (heat level as variant)
- `Family Pack`, `Party Pack` — Korean BBQ
- `2 Ounce`, `8 Ounce`, `12 Ounce` — drinks by volume

Don't try to guess what they mean from the name alone. Move to Step 2.

### Step 2: For each dish group, inspect existing rows to learn the convention

Pick a dish group (e.g., "Appetizers"). Look at 3-5 existing rows in that group. Note which variant column(s) hold their prices.

Example: "Appetizers" group — all rows have `Order` populated, others blank. → Convention: appetizers use `Order`.

Example: "Pizzas" group — all rows have `Pizza Small`, `Pizza Medium`, `Pizza Large` populated. → Convention: pizzas use those three columns.

Example: "Pasta" group — all rows have `Half Tray` AND `Full Tray` populated. → Convention: pasta uses both.

**Rule:** Whatever column(s) the existing system rows use, that's the convention. Mirror it in the OCR JSON for paper menu prices in the same dish group.

### Step 3: Match paper menu names to system names via substring search

For each paper menu section/dish:
1. Take the core noun (drop "Roll", "Sub", "Special", proteins, etc.).
2. Search the Menu Preview's Primary column for any row containing that substring.
3. If you find a near-match, use the system's exact name in the OCR JSON.

Try these transformations when a direct match fails:
- Substring matches (e.g., paper "Margherita Pizza" → system has "Pizza Margherita")
- Word-order flips (paper "Chicken Garlic" → system "Garlic Chicken")
- Suffix variants (Sub, Soup, Roll, Bowl, Plate, Sandwich)
- Plural toggles (Wing/Wings, Taco/Tacos)
- Accent stripping (Crème → Creme, Café → Cafe)
- Apostrophe variants (Chef's → Chefs)
- Punctuation drops (& → and, w. → with — script handles automatically)
- Spelling drift (one letter off)

### Step 4: After Action 2, walk the P1 list

P1 items in unknown cuisines often turn out to be alias misses. For each P1 item, repeat the substring/transformation search in Step 3. Update the OCR JSON and re-run Action 2 if you find a system match.

If a P1 item really doesn't exist in the system after thorough search, it's a genuinely new dish. Leave it as P1.

## Things to Look For (cuisine-agnostic patterns)

These patterns recur across cuisines and are worth checking:

### Family / size variants
Many cuisines have "Small/Medium/Large" pricing or "Half Tray/Full Tray". Look at the variant columns to see how this restaurant encodes it.

### Heat / preparation variants
Some restaurants encode heat level or preparation as a variant (e.g., Mild/Medium/Hot wings, Single/Double/Triple meat burritos). This usually appears as separate variant columns rather than separate dish rows.

### Cross-section duplicates
Items may appear in multiple dish groups at the same price — e.g., a "Margherita Pizza" listed both in "Classic Pizzas" and "Lunch Specials". Always search the full Menu Preview for each paper menu item.

### Combo / platter sections
Combo/platter/family sections often use a dedicated variant column (e.g., `Combo`, `Platter`, `Family Pack`). Don't put combo prices in `Order` without checking.

### Drinks consolidated under one variant
Beverages often live in a `Drinks` dish group with a single variant (`Order`, `Can`, `Bottle`). Bubble teas, smoothies, juices may share a uniform price.

## Pricing Sanity Checks (cuisine-agnostic)

These checks apply regardless of cuisine:

1. **Within a dish family, watch for outliers.** If all wings are $14 and one shows $4, re-read it.
2. **A vs B variants.** Restaurants commonly use "A" (simpler) vs "B" (deluxe) suffixes, with B priced $2-5 higher. If you read A > B, you likely swapped them.
3. **Uniform pricing within a "specials" section.** Lunch specials, party platters, and drink families often have uniform pricing. An outlier within a uniform section is suspect.
4. **Price digit checks.** Common OCR digit confusions: .25 vs .75, .00 vs .50, 8 vs 9, 0 vs 6, 1 vs 7. These appear regardless of cuisine.
5. **Never decrease without evidence.** Menu prices almost never decrease year-over-year. If you read a decrease, re-verify the digit.

## When to Promote to a Dedicated Cuisine Module

If this restaurant's cuisine appears 2+ times across runs, consider promoting the patterns you discovered into a dedicated `cuisine-<name>.md` module. Document:

- Cuisine inference signals (variant columns, dish group names, term vocabulary)
- Variant column conventions per section type
- Common alias patterns (system naming drift)
- Cross-section duplicate patterns
- Pricing tier / uniformity rules
- Cuisine-specific compound item expansion (if any)

This keeps the discovery cost paid once rather than per-restaurant.