#!/bin/bash
# usage: run_p1p2.sh <bundle_dir> <raw.xlsx> <outdir>
set -e
S=$1/menu-change-payload/scripts; RAW=$2; O=$3
rm -rf $O && mkdir -p $O && cp $RAW $O/raw.xlsx && cp R99999_menu_ocr.json $O/
cd $O
python3 $S/action1_raw_to_2a.py raw.xlsx -o preview.xlsx            > log_a1.txt 2>&1 || { cat log_a1.txt; exit 1; }
python3 $S/ocr_sanity_check.py R99999_menu_ocr.json                  > log_sanity.txt 2>&1
python3 $S/action2_ocr_to_2b_0625v2.py R99999_menu_ocr.json preview.xlsx -o ocr_out.xlsx > log_a2.txt 2>&1 || { cat log_a2.txt; exit 1; }
python3 $S/action3_compare_and_merge.py preview.xlsx ocr_out.xlsx -o report.xlsx > log_a3.txt 2>&1 || { cat log_a3.txt; exit 1; }
python3 $S/price_inference.py preview.xlsx ocr_out.xlsx --report report.xlsx --mode apply > log_pi.txt 2>&1 || { cat log_pi.txt; exit 1; }
# ---- Part 2 (report unedited = specialist accepted everything) ----
python3 $S/action4_apply_modifier_changes.py report.xlsx -o final_review.xlsx > log_a4.txt 2>&1 || { cat log_a4.txt; exit 1; }
python3 $S/action5_to_raw_gsheet.py final_review.xlsx raw.xlsx -o final_raw.xlsx > log_a5.txt 2>&1 || { cat log_a5.txt; exit 1; }
echo "pipeline OK -> $O"
