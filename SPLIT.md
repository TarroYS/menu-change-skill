# Thin P1/P2 split (draft — not committed)

Status: Yee Sern Goh asked 2026-09-15 to proceed with a thin split of the union pack. **No git commit. No PR.** Approve or send notes.

## Layout

```
menu-change-workflow-part1.md    # P1 compare only (~union P1 sections)
menu-change-workflow-part2.md    # P2 apply only (~union P2 sections)
menu-change-payload/             # UNCHANGED union bytes (scripts/refs/hooks/setup.sh)
SPLIT.md                         # this file
```

Payload is the same directory for both skills. Do not embed scripts in the markdown (that is the 0716 tax).

## What changed vs union `menu-change-workflow.md`

| | Union (combined) | This draft |
|---|---|---|
| Skills | 1 file, P1+P2 triggers in one description | 2 files; P1 description does not mention Action 4/5 / draft; P2 does not mention OCR / P6 |
| Payload | `menu-change-payload/` | same files, same hashes |
| Embeds | none | none |
| Runtime cost | ~4.4k tokens always | P1 ~same as union-minus-P2; P2 much smaller (no OCR/cuisine) |

## What did not change

- Script bytes (fidelity lock).
- P1 pipeline: action3 → `price_inference.py --mode apply` → (if side_pricing) compare `--report`.
- P2 pipeline: action4 → action5 → to-tsv → (if Modifier Rule) apply → validate → staged draft.
- Kickoff is still restaurant + paper menu (P1) or edited report (P2) — not “run this zip”.

## Approve / next

- Approve as-is → say commit / PR if you want it in the monorepo (not done until you say so).
- Notes → reply with edits; still no commit.
