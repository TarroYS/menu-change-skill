#!/usr/bin/env python3
"""Stop-hook gate for the Menu Change Workflow (Part 1).

PURPOSE
-------
The Part-1 skill's only valid deliverable is the colored `.xlsx` Price
Comparison Report produced by `scripts/action3_compare_and_merge.py`. In
practice ACE sometimes skips the script pipeline and ships a free-text
markdown "diff" written by a general subagent instead. A SKILL.md instruction
can be ignored; this hook cannot — the harness runs it on every Stop event.

BEHAVIOR
--------
Reads the Stop-hook JSON on stdin (Claude Code contract), parses the
conversation transcript, and:

  1. Decides whether this conversation is a *menu-change run*. The ONLY anchor
     is a real inbound menu image/PDF (an attached file or pasted image) — see
     _is_inbound_menu_file. Pipeline script names, the skill name in prose, and
     subagent-prompt phrasing are deliberately NOT used: they saturate sessions
     that author/test the skill and caused false BLOCKs.
  2. If it is, checks whether the run was completed correctly:
       - `action3_compare_and_merge.py` was actually executed via Bash, AND
       - `price_inference.py` was actually executed via Bash, AND
       - an `upload_artifact` call uploaded a file whose name ends in `.xlsx`.
  3. If it's a menu run but NOT completed correctly, it BLOCKS the stop and
     returns an actionable reason telling ACE to run the pipeline.

Non-menu conversations are never blocked. Loop-safe: if we're already inside a
stop-hook continuation (`stop_hook_active`), we allow the stop to avoid an
infinite block loop.

OUTPUT CONTRACT
---------------
- Allow:  exit 0, no output.
- Block:  exit 0 with JSON {"decision":"block","reason":"..."} on stdout.
On any internal error we FAIL OPEN (allow) so the hook can never wedge a
session — a missing report is recoverable, a stuck agent is not.
"""
import sys
import json
import re


# ── Detection signals ────────────────────────────────────────────────
# THE menu-run anchor is a real *inbound menu image/PDF* — a binary file the
# user attached. This is the one signal development work cannot fabricate:
# building/testing the skill produces .py/.md/.json/.tsv chatter and runs the
# pipeline scripts by name, but it never makes someone *attach a menu photo*.
# Anchoring here is exactly the user's stated rule: "when you receive the file,
# it will always be a menu-change run." We deliberately do NOT key on pipeline
# script names, the skill name in prose, or subagent-prompt phrasing — those
# saturate any session that authors or tests the skill and caused false BLOCKs.
MENU_FILE_EXTS = (
    ".png", ".jpg", ".jpeg", ".webp", ".heic", ".heif",
    ".gif", ".tif", ".tiff", ".bmp", ".pdf",
)
ACTION3_TOKEN = "action3_compare_and_merge"
INFERENCE_TOKEN = "price_inference.py"

# Path fragments that mean a file lives *inside the skill bundle* (reference
# assets, samples) rather than being a fresh inbound menu the user attached.
SKILL_INTERNAL_HINTS = (
    "menu-change-workflow-part",
    "menu-change-workflow",
    "menu-change-payload",
    "/references/",
    "/scripts/",
    "/hooks/",
)


def _is_inbound_menu_file(path):
    """True iff `path` looks like a menu image/PDF the user attached (not a
    skill-internal asset)."""
    p = str(path or "").lower()
    if not p.endswith(MENU_FILE_EXTS):
        return False
    if any(h in p for h in SKILL_INTERNAL_HINTS):
        return False
    return True


def _iter_tool_uses(transcript_path):
    """Yield (tool_name, tool_input_dict) for every tool_use in the transcript.
    Also yields ('__attachment__', entry) for attachment rows and
    ('__image__', {}) for inline image content blocks (pasted images)."""
    with open(transcript_path, "r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
            except Exception:
                continue
            if obj.get("type") == "attachment":
                yield ("__attachment__", obj.get("attachment") or {})
                continue
            msg = obj.get("message") or {}
            content = msg.get("content")
            if not isinstance(content, list):
                continue
            for block in content:
                if not isinstance(block, dict):
                    continue
                btype = block.get("type")
                if btype == "tool_use":
                    yield (block.get("name", ""), block.get("input") or {})
                elif btype == "image":
                    yield ("__image__", {})


def _analyze(transcript_path):
    """Return (is_menu_run, action3_ran, inference_ran, xlsx_uploaded)."""
    is_menu_run = False
    action3_ran = False
    inference_ran = False
    xlsx_uploaded = False

    for name, payload in _iter_tool_uses(transcript_path):
        # Inbound menu image/PDF attachment — the sole menu-run anchor.
        if name == "__attachment__":
            if payload.get("type") == "file" and _is_inbound_menu_file(
                payload.get("filename") or payload.get("displayPath")
            ):
                is_menu_run = True

        # A pasted inline image is also an inbound menu file.
        elif name == "__image__":
            is_menu_run = True

        # Did Action 3 / price_inference actually run?
        elif name == "Bash":
            cmd = str(payload.get("command", "")).lower()
            if ACTION3_TOKEN in cmd:
                action3_ran = True
            if INFERENCE_TOKEN in cmd:
                inference_ran = True

        # Artifact uploads — did a real .xlsx report go out?
        elif name == "mcp__tarro__upload_artifact":
            fp = str(payload.get("file_path", ""))
            fn = str(payload.get("file_name", ""))
            target = (fn or fp).lower()
            if target.endswith(".xlsx"):
                xlsx_uploaded = True

    return is_menu_run, action3_ran, inference_ran, xlsx_uploaded


BLOCK_REASON = (
    "MENU-CHANGE GATE: This looks like a Menu Change Workflow (Part 1) run, but "
    "the Price Comparison Report (.xlsx) pipeline is incomplete. A "
    "markdown/prose diff is NOT a valid deliverable. Before stopping you MUST: "
    "(1) run the real pipeline — setup.sh -> wsheet_xlsx_bridge.py to-xlsx -> "
    "action1_raw_to_2a.py -> (OCR JSON) -> action2_ocr_to_2b_0625v2.py -> "
    "action3_compare_and_merge.py -> price_inference.py --mode apply; "
    "(2) upload the resulting .xlsx via upload_artifact. Do not delegate the "
    "price comparison itself to a general subagent — subagents may only OCR "
    "the images into the OCR JSON. If this is NOT a menu-change run, say so "
    "explicitly and stop."
)


def main():
    # Fail open on any malformed input.
    try:
        raw = sys.stdin.read()
        data = json.loads(raw) if raw.strip() else {}
    except Exception:
        sys.exit(0)

    # Loop guard: never block twice in a row.
    if data.get("stop_hook_active"):
        sys.exit(0)

    transcript_path = data.get("transcript_path")
    if not transcript_path:
        sys.exit(0)

    try:
        is_menu_run, action3_ran, inference_ran, xlsx_uploaded = _analyze(
            transcript_path
        )
    except Exception:
        sys.exit(0)  # fail open

    if is_menu_run and not (action3_ran and inference_ran and xlsx_uploaded):
        print(json.dumps({"decision": "block", "reason": BLOCK_REASON}))
        sys.exit(0)

    sys.exit(0)


if __name__ == "__main__":
    main()