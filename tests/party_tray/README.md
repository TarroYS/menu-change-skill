# Party Tray regression harness (v4.3)
Synthetic ChAm restaurant R99999 — Party Tray = Large × 5 − $2.00, one outlier, one tray-only dish.
    python3 build_fixture.py          # R99999_raw.xlsx + R99999_menu_ocr.json
    python3 build_fixture.py mixed    # R99999_raw_mixed.xlsx (no dominant formula)
    ./run_p1p2.sh <bundle_dir> R99999_raw.xlsx out     # action1 → … → action5
    python3 show.py out               # Party Tray at every stage + final raw
Run against the old and new bundle to diff behaviour. Uses python3 + openpyxl directly (no .venv).
