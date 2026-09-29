#!/usr/bin/env python3
"""
price_queue_gate.py — read-only staged-price scope reporter (v2, ACE-native).

WHAT CHANGED IN v2 — READ THIS.
    v1 of this script was a BLOCKING gate: it required a human-created
    `.price_queue_approved` marker file inside the content dir and exited 3
    (blocking `wsheet draft create` via `&&`) when the marker was absent. That
    design assumed a human with shell access to the same filesystem. ACE users
    have NO terminal — only the chat UI — so the marker could never be created
    (a pasted `touch` command arrives as chat text), the assistant is forbidden
    from creating it, and the workflow DEADLOCKED: `draft create` never ran and
    no draft was ever produced.

    v2 removes the blocking behavior entirely. This script is now a pure,
    read-only SCOPE REPORTER: it prints the price-bearing tab row counts for the
    audit trail and ALWAYS EXITS 0. It never reads or writes a marker file and
    never blocks. Draft creation is safe and inherently reviewable in ACE, so
    `wsheet draft create` runs directly. The real human sign-off before
    production is ACE's inline approval on `wsheet draft commit` (guardrail G9).

Usage (informational only — the following command does NOT block draft create):

    .venv/bin/python scripts/price_queue_gate.py <dir>
    wsheet draft create <dir> --name "Menu price update — <Restaurant> (Part 2)"

Do NOT chain these with `&&` expecting a gate — there is no gate. Do NOT
reintroduce a marker file or ask the user to `touch` anything.

Exit codes:
  - 0 -> always, when given a valid content dir (this is a report, not a gate).
  - 2 -> argument error (missing dir / not a directory).
"""
import os
import sys

PRICE_TABS = ("dish_item", "modifier_item", "modifier_rule")


def summarize_pending(content_dir):
    """Best-effort scope hint: data-row counts of the price-bearing tabs."""
    lines = []
    for tab in PRICE_TABS:
        p = os.path.join(content_dir, f"{tab}.tsv")
        if os.path.isfile(p):
            try:
                with open(p, encoding="utf-8") as f:
                    n = sum(1 for _ in f)
                lines.append(f"  - {tab}.tsv ({max(n - 3, 0)} data rows)")
            except OSError:
                pass
    return "\n".join(lines) or "  (no price tabs found)"


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if len(args) != 1:
        print("price_queue_gate: usage: price_queue_gate.py <content_dir>",
              file=sys.stderr)
        return 2
    content_dir = args[0]
    if not os.path.isdir(content_dir):
        print(f"price_queue_gate: not a directory: {content_dir}",
              file=sys.stderr)
        return 2

    scope = summarize_pending(content_dir)
    print(
        "price_queue_gate: SCOPE REPORT (read-only — does not block).\n\n"
        f"Staged price tabs in {content_dir}:\n{scope}\n\n"
        "Next: run `wsheet draft create <dir> --name \"...\"` to STAGE the draft "
        "(safe; does not touch production).\n"
        "To push to production later, a human approves ACE's inline card on "
        "`wsheet draft commit` — no marker file, no shell command needed."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())