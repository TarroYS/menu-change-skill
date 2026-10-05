#!/bin/bash
# usage: tests/run_all.sh [bundle_dir]        (default: the repo root)
#   REAL_RAW=/path/R85325_final_raw_gsheet.xlsx      + real-data checks (modifier_rule)
#   REAL_REPORT=/path/<an edited older report>.xlsx  + older-report checks (report)
#   PYTHON=/path/to/python                           python with openpyxl >= 3.1
# Must end "ALL CHECKS OK".
set -u
cd "$(dirname "$0")"
B=${1:-..}
PY=${PYTHON:-python3}
fail=0
"$PY" party_tray/check_party_tray.py "$B" || fail=1
"$PY" modifier_rule/check_modifier_rule.py "$B" ${REAL_RAW:+--real-raw "$REAL_RAW"} || fail=1
"$PY" price_inference/check_rules.py "$B" || fail=1
"$PY" report/check_report.py "$B" ${REAL_REPORT:+--real-report "$REAL_REPORT"} || fail=1
"$PY" edits/check_edits.py "$B" || fail=1
"$PY" audit/check_audit.py "$B" || fail=1
"$PY" codes/check_codes.py "$B" || fail=1
if [ $fail = 0 ]; then echo "ALL CHECKS OK"; else echo "CHECKS FAILED"; exit 1; fi
