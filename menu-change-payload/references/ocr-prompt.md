# OCR Prompt for Paper Menu Reading (General)

Use this prompt when Claude reads paper menu images to produce the structured OCR JSON. This file contains the **universal** OCR guidance that applies regardless of cuisine. Cuisine-specific rules live in separate modules — load the appropriate one based on Step 0 cuisine inference (see `SKILL.md`).

| Cuisine | Reference module |
|---|---|
| Chinese-American | `references/cuisine-chinese-american.md` |
| Sushi / Japanese | `references/cuisine-sushi-japanese.md` |
| Other / unknown | `references/cuisine-other.md` |

---

You are reading paper menu images for a restaurant. The same JSON schema and universal rules below apply to any cuisine. Cuisine-specific OCR conventions (variant column mapping, alias patterns, compound item expansion, pricing tier checks) are layered on top via the cuisine module identified in Step 0.

Extract EVERY dish, price, and menu section into the JSON schema at the bottom of this file.

## STEP 0: INSPECT THE MENU PREVIEW BEFORE OCR-ING

Before reading the paper menu, **open the Menu Preview's `price` tab and look at row 1**. The headers there are the variant columns the system uses for THIS restaurant. The Action 2 script writes to ANY column in that header row — there is no hardcoded whitelist. But it can only write where your OCR JSON tells it to write. **Putting a price under the wrong variant column means the price update silently does not happen.**

For each dish group, look at the existing system rows and note which column holds their price. Mirror that same column as the key in the OCR JSON's `prices` dict. Examples:

- System: `[Appetizers, Edamame, _, Order=$8.00]` → OCR JSON: `"prices": {"Order": 10.50}`
- System: `[Special Roll 1, Dragon Roll, _, Roll=$16.00]` → OCR JSON: `"prices": {"Roll": 17.00}`
- System: `[Dinner Bento Box, Chicken Bento Box A, _, Dinner=$25.00]` → OCR JSON: `"prices": {"Dinner": 27.00}`
- System: `[Lunch Specials, Sweet and Sour Chicken, _, Lunch=$8.95]` → OCR JSON: `lunch_specials` section with `"price": 9.50`

Refer to the cuisine module for typical variant column conventions for the cuisine at hand.

## UNIVERSAL CRITICAL RULES

1. **Capture EVERY item** — do not skip any dish or price.
2. **Use exact names as printed** (preserve "w." abbreviations, "&" symbols, apostrophes, etc.). The Action 2 script handles normalization automatically.
3. **Map price columns correctly** — the variant key in the OCR JSON must match a column header in the Menu Preview. The cuisine module documents typical conventions for that cuisine. The system's column choice is always the source of truth — if unsure, look at any existing row in that dish group and use whatever column holds its existing price.
4. **Note the menu code numbers** (L1, C1, 72, etc.) in the `"code"` field.
5. **If a section has MULTIPLE price columns** (e.g., Fried Specialties with Plain/w.French Fries/w.Pork FR/...):
   - Extract the FIRST column (Plain) as `"Specialty"` in the `prices` dict — EXCEPT for a dish the system prices per size (its Menu Preview row has size columns such as `5 Pieces` / `10 Pieces`, e.g. Cut Wings): emit one item per printed size and key that row's Plain by the size column (`{"name": "Cut Wings", "prices": {"5 Pieces": 6.75}, "side_prices": {…}}`). Keep the count — here it is the size, not a decorative token (rule 7). A sized row without its size is flagged `VARIANT UNKNOWN` and not applied.
   - **A combo box prices dishes WITH a side, at ONE price** — e.g. `Appetizer Combos — All served w. french fries or fried rice … $10.95`, while each dish's own price is printed in another section (APPETIZER). That price is side pricing, not the dishes' own price: emit the box as its own section with `"side_pricing": {"columns": ["w. french fries or fried rice"]}` (header as printed) and one item per listed dish with `"prices": {}` and `"side_prices": {"w. french fries or fried rice": 10.95}`. Never put it in `prices` or `default_price` — it would overwrite each dish's own price (Small / Large / Plain). Keep each dish's own price in its own section, as printed; the compare step pairs the two by name.
   - Capture the other columns in a `side_prices` dict on each item, keyed by the column header as printed
   - Add a `side_pricing` object at the section level listing the column headers
   - **Capture the FULL bordered block, border included.** Multi-column pricing tables are usually enclosed in a ruled box/frame. Read (and, when cropping for the embedded reference image, frame) the block all the way OUT TO its ruled border on every side — top, bottom, left, and the right-most price column. Never trim to the inner whitespace: the outer-most price column (e.g. `w.Shrimp Fried Rice`) sits flush against the right border and is the first thing lost when the crop stops at inside whitespace. If the table has a visible border/rule, the whole border should be inside the captured region.
   - **Header spelling — preserve `with` / `w.` exactly as printed.** Column headers may be spelled out (`with French Fries`) or abbreviated (`w.Pork Fried Rice`); capture them verbatim. The downstream matcher strips a leading `with`/`w.`/`w ` before matching, so either form resolves — but do NOT paraphrase one into the other.
   - **Capture EVERY price a row actually prints, even a PARTIAL row.** Rows do NOT all fill the same number of columns — kids' meals, small/large-only items, and half-portions commonly print prices in only the first few columns and leave the right-most columns blank. Capture whatever side-column prices ARE printed; set the value to `null` ONLY for the specific columns that are genuinely blank on that row.
   - **Read ragged rows by the dotted leaders, not by fixed x-position.** A shorter row's prices may not line up under the same columns as the full rows above it. Follow each price along its `.......` leader to the correct column header. Do NOT assume a row that stops early has no side prices — count the printed price values and map them left-to-right onto the columns they align with.
   - `"side_prices": null` is reserved for rows that have a Plain price and **truly no** side-column prices at all (e.g. a bare "White Rice (S)3.00 (L)5.00" line). If a row has even one side-column price, use a `side_prices` dict with per-column values (`null` for the blank ones) — never collapse the whole row to `null`.
   - **Self-check before finishing the section:** the number of items carrying `side_prices` should equal the number of paper rows that visibly show more than one price. If a row shows 2+ prices but you emitted `side_prices: null`, you misread it — re-read that row. (The `modifier_rule_pricing.py compare` step also prints a COVERAGE warning listing any multi-column row left without `side_prices`, as a backstop.)

   *Example — full vs. partial vs. plain-only rows in one section:*
   ```
   D1  Fried Chicken Wing (4) ... 7.59 ... 8.79 ... 9.29 ... 10.29 ... 10.59   (full: all columns)
   D7  Kid's Meal for Chicken Wing (2) ... 6.99 ... 7.29 ... 7.99 ... 8.29     (partial: right columns blank)
   D10 White Rice ...(S)3.00 (L)5.00                                          (plain-only: side_prices null)
   ```
   D1 → every column filled. D7 → first columns filled, remaining columns `null` (still a `side_prices` dict — NOT null). D10 → `side_prices: null`.
6. **Verify the restaurant name, address, and phone** from the menu.
7. **Strip decorative noise tokens from dish names** — these never appear in POS names:
   - Trailing portion-count markers: `(4)`, `(6)`, `(8)`, `(10)`, etc. (`Fried Chicken Nuggets (8)` → `Fried Chicken Nuggets`).
   - Parenthetical reading aids: `(Chopped Wings)`, `(Chicken on the Sticks)`, `(Cheese Only)` — **UNLESS** the parenthetical is what distinguishes two otherwise-identical dishes in the same section, in which case keep it.
   - Section-redundant words the system drops (cuisine-specific; e.g., leading `Roast` in a Pork section — see cuisine module).
   Keep `w.`/`&`/apostrophes as printed otherwise — the script normalizes those.

## UNIVERSAL PRICE READING RULES (prevents OCR errors)

7. **READ EACH PRICE DIGIT CAREFULLY.** Common misreads across all cuisines:
   - .25 vs .75 — look similar at low resolution. Double-check by context.
   - .00 vs .50 — verify the cents digits.
   - 8.xx vs 9.xx — the first digit can blur. Check neighboring items.
   - 10.xx vs 11.xx — same issue at the tens boundary.
   - 0 vs 6, 1 vs 7 — common at low scan quality.
   - 3 vs 8, 5 vs 6, 9 vs 4 — common menu-font confusions; when a first digit is ambiguous, check the section's price band.
   - **Dropped decimal point** — `895` or `1295` almost always means `8.95` / `12.95`, not `895.00`. Any whole-number price ≥ 100 on a per-dish item is a misread until proven otherwise.

8. **PRICE DECREASE SANITY CHECK.** Menu prices almost NEVER decrease year-over-year. If you read a price that seems lower than what a dish "should" cost, re-read it. Decreases do happen occasionally (e.g., promotions, repricing) but are flagged for client verification in P2 — they should be rare.

9. **PRICE PATTERN SANITY** by cuisine. The cuisine module describes the expected pricing tiers for the cuisine in question (e.g., ChAm protein tiers, sushi fish tiers, pizza size tiers). If a price breaks the expected pattern, re-verify the digit.

10. **COUNT THE PRICE COLUMNS** per section BEFORE extracting. Some sections have two columns (Pt/Qt, Small/Large, Lunch/Dinner); some have one. Knowing the count up front prevents misalignment errors.

11. **VERIFY against neighboring items** in the same section — most menu sections have internally consistent pricing patterns.

12. **DECLARE THE COLUMNS, THEN READ EVERY ROW.** Before extracting a section, write its `price_columns` explicitly (e.g. `["Small","Large","Lunch"]`), then read each dish ACROSS those columns. **Emit one item per printed row even when a price is blank** — set the blank variant to `null` rather than skipping the row. A silently dropped row is the most common long-list OCR error; a `null` is auditable, a missing row is not.

13. **FLAG UNCERTAIN READS INSTEAD OF GUESSING.** When a glyph is smudged, cut off, or genuinely ambiguous after re-reading, set `"uncertain": true` on that item (optionally `"uncertain_note": "8.95 vs 8.25"`). Do NOT silently pick one. `scripts/ocr_sanity_check.py` collects every `uncertain` item into a review list for the specialist — a surfaced question is safe, a silent misread is not.

## IMAGE PREPARATION (Opus vision — improves digit accuracy)

Claude Opus reads the menu image directly, and its vision encoder **downscales any image whose long edge exceeds ~2576 px**, preserving aspect ratio. A dense full-page photo therefore arrives with the price digits shrunk — the root cause of `8.95`/`8.25` and dropped-decimal misreads. To keep digits sharp:

- **Read one section at a time, not the whole page.** Crop each priced section to its own image so its text fills more of the pixel budget. `scripts/crop_section.py` already crops between two title anchors (`--title` / `--next`); use it per priced section, not only for the multi-column block.
- **Target ~2576 px on the long edge per crop on Opus** — do not pre-shrink to the old 1568 px limit; Opus uses the extra pixels. Keep every edge above 200 px (vision degrades below that).
- **Optional deterministic cross-check:** `crop_section.py` can run `tesseract` when installed. Reading the price column with tesseract and diffing it against your read flags exactly the cells most likely misread — two independent readers catch what one misses.
- **Do NOT feed the system's current prices into the read to "help" it.** That anchors the model toward "no change" and defeats change detection. System prices belong only in the deterministic compare (Action 3), or at most as a flagged tiebreak on a truly illegible glyph.

## UNIVERSAL COMPOUND PATTERNS

These compound patterns can appear in any cuisine and the Action 2 script handles them automatically. Specific named patterns (Moo Shu, Egg Foo Young, Chow Mein/Chop Suey, etc.) are documented in the cuisine module.

### General parenthetical expansion
Any dish written as `Base (Variant1, Variant2 or Variant3)` or `(Variant1, Variant2 or Variant3) Base` at one price means multiple independent dishes sharing that price. The script auto-expands these into separate rows.

Examples (cuisine-agnostic):
- `Burrito (Chicken, Beef or Pork)` → `Burrito Chicken`, `Burrito Beef`, `Burrito Pork`
- `(Mild, Medium or Hot) Wings` → `Mild Wings`, `Medium Wings`, `Hot Wings`

### Multi-variant lists with "or"
Items like `Chicken or Roast Pork w. Broccoli` or `Hunan Style Chicken or Beef or Shrimp` — enter the FULL name as printed. The script splits on `/` and `or` automatically.

### Slash-separated dual items
Items like `Beef Chow Mein / Chop Suey` — keep the combined name. The script auto-matches both system rows. (This is most common in ChAm but the mechanism is general.)

### Multi-protein with commas
`Broccoli with Chicken, Pork, Beef or Shrimp` → 4 separate rows in the system. Enter the full string as printed; the script handles the split.

### Bare protein names under section headers
Section header "Stir Fried Udon" with rows just listing protein names (`Vegetable`, `Chicken`, `Beef`, `Shrimp`) — enter the protein name as printed and make sure `menu_section` accurately captures the header. The script prepends the section name to form the full dish name (e.g., "Stir Fried Udon with Vegetable").

**When the system stores the section word as a SUFFIX, not a prefix**, emit the fully-formed `<Protein> <SectionWord>` name yourself rather than the bare protein. Prepend-only fails for these (e.g., bare "Singapore" under a Pancit section → system "Singapore Pancit", but a blind prepend yields the wrong "Pancit Singapore"). When the section header clearly implies a suffix, write the complete name in the OCR JSON. Cuisine modules list the specific suffix sections (e.g., ChAm Pancit, Yat).

### Sub-variant "Choice of sauce" expansion
A parent dish followed by a sub-line listing sauce/preparation choices — `Choice of X or Y or Z`, `Served with X, Y or Z`, `Sauce: X, Y, Z` — is **one POS dish per choice**, all sharing the parent's price(s) and code. Emit one OCR item per choice and derive each item's name from the choice, not the parent.

Example:
`Choice of White Sauce or Brown Sauce` under "Pan Fried Noodle" → two items: `Pan Fried Noodle White Sauce` and `Pan Fried Noodle Brown Sauce`, each sharing the parent's price and code.

## OCR JSON SCHEMA (output this exact structure)

This is the structure `scripts/action2_ocr_to_2b_0625v2.py` consumes. Unknown keys are ignored, so the `uncertain` flag is safe to include.

```json
{
  "restaurant": {"name": "...", "address": "...", "phone": "...", "r_number": "..."},
  "sections": [
    {
      "menu_section": "Section Name As Printed",
      "price_columns": ["Small", "Large"],
      "default_price": null,
      "items": [
        {"name": "Dish Name As Printed", "prices": {"Small": 7.00, "Large": 10.25}},
        {"name": "Ambiguous Dish", "prices": {"Order": 8.95}, "uncertain": true, "uncertain_note": "8.95 vs 8.25 - smudged"}
      ]
    }
  ],
  "lunch_specials": {"note": "...", "default_price": 9.45, "items": [{"name": "Dish Name", "price": null}]},
  "combinations": {"note": "...", "default_price": 11.75, "items": [{"name": "Dish Name", "price": null}]},
  "side_orders": [{"name": "Steamed Rice", "prices": {"Order": 4.00}}]
}
```

**Multi-column sections** additionally carry, per item, a `side_prices` dict keyed by the column header as printed (see UNIVERSAL CRITICAL RULES §5), plus a section-level `side_pricing` object listing those headers. Use `"side_prices": null` only for a row that truly has no side-column price.
