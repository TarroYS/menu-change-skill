#!/usr/bin/env bash
# setup.sh — one-time environment bootstrap for the menu-change workflow.
#
# ACE's sandbox has no pip, no root, and an externally-managed Python
# (PEP 668), so `pip install openpyxl` fails. This script builds a local
# virtualenv with openpyxl, which every action*/bridge script needs.
#
# Idempotent: safe to run every time. If the venv already has openpyxl it
# exits immediately. Run from the skill bundle root (where scripts/ lives):
#
#     bash setup.sh
#     ./.venv/bin/python scripts/action3_compare_and_merge.py ...
#
set -euo pipefail

VENV="${VENV_DIR:-.venv}"
PY="$VENV/bin/python"

# 1. Already good? Bail fast.
if [ -x "$PY" ] && "$PY" -c "import openpyxl" 2>/dev/null; then
    echo "setup.sh: openpyxl already available in $VENV ($("$PY" -c 'import openpyxl;print(openpyxl.__version__)'))"
    exit 0
fi

# 2. Create the venv if missing. Use --without-pip: ensurepip is stripped in
#    this sandbox and a plain `venv` aborts on it. We bootstrap pip in step 3.
if [ ! -x "$PY" ]; then
    echo "setup.sh: creating venv at $VENV"
    python3 -m venv --without-pip "$VENV"
fi

# 3. Ensure pip exists inside the venv (ensurepip is often stripped, so
#    fall back to get-pip.py). Requires network access to pypi.org.
if ! "$PY" -m pip --version >/dev/null 2>&1; then
    echo "setup.sh: bootstrapping pip via get-pip.py"
    if [ ! -f get-pip.py ]; then
        "$PY" - <<'PYEOF'
import urllib.request
urllib.request.urlretrieve("https://bootstrap.pypa.io/get-pip.py", "get-pip.py")
print("setup.sh: downloaded get-pip.py")
PYEOF
    fi
    "$PY" get-pip.py >/dev/null
fi

# 4. Install openpyxl.
echo "setup.sh: installing openpyxl"
"$PY" -m pip install --quiet openpyxl

# 5. Verify.
"$PY" -c "import openpyxl; print('setup.sh: ready, openpyxl', openpyxl.__version__)"