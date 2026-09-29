# Tested restaurants (historical, from part1/part2)

Load only if useful for a coverage sanity-check. Not required to run the pipeline. Counts are from prior runs — do not quote them as this-run results (G12).

## Part 1

| Restaurant | R-Number | Style | Matched | Changes | Decreases |
|---|---|---|---|---|---|
| No.1 Chinese | R88312 | size_based | 124/209 | 234 inc | 0 |
| China One | R82811 | size_based | 122/~180 | 223 inc | 0 |
| Foo Yung Express | R83067 | lunch_based | 47/~100 | 0 (match) | 0 |
| Bei Jing | R82763 | size_based | 121/~180 | 202 inc, 2 dec | 2 (legit) |
| Chang Xing | R81625 | size_based | 162/~440 | 190 inc, 2 dec | 2 (verify) |
| Jade Garden | R85991 | size_based | 106/~189 | 172 inc, 1 dec | 1 (verify) |
| Best Food In Town | R81001 | size_based (dual) | 173/~289 | 223+ inc, 2 dec | 2 (verify) |
| SeaPark Sushi | ER85777 | sushi (Roll/Dinner/Seapark/Bowl/Regular Roll) | 134/253 | 132 inc, 2 dec | 2 (verify) |
| Kin's Wok | ER88623 | size_based | 120/161 → 150/161 after OCR tuning | — | — |

## Part 2

| Restaurant | R-Number | dish_item cells | dish_item code cells | modifier_item cells |
|---|---|---|---|---|
| No.1 Chinese | R88312 | 234 | — | 172 (simulated) |
| China One | R82811 | 223 | — | 0 (no edits) |
| Foo Yung Express | R83067 | 0 | — | 0 (prices matched) |
| Bei Jing | R82763 | 203 | — | 0 (no edits) |
| Oriental Palace | ER89365 | 43 | 35 (32 items) | 0 |
