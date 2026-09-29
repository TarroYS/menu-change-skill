#!/usr/bin/env python3
"""
crop_section.py
===============

Reusable, reproducible cropper for the multi-column-pricing block of a paper
menu (e.g. "SPECIAL DISHES" / "FRIED SPECIALTIES" tables).

Why this exists
---------------
`PIL.Image.crop((left, top, right, bottom))` only *executes* a box you already
know; it cannot *find* the right table. Hand-tuned pixel boxes do not transfer
between menus (different scans, resolutions, layouts). This module derives the
box from image content, then uses Pillow to execute + refine it, then validates
the result before saving. Accuracy comes from detection + validation; Pillow is
the executor.

Pipeline
--------
  1. LOAD + NORMALIZE ...... EXIF auto-orient, optional deskew, grayscale copy.
  2. DETECT (coarse box) ... two strategies, best-effort in this order:
       a. OCR-anchored (requires pytesseract + tesseract binary): locate the
          section title word(s) and the next-section word(s); the box spans from
          just above the title to just above the next section, full content
          width. Content-relative => survives resolution/scan differences.
       b. Projection-profile fallback (Pillow + numpy only): find horizontal
          "ink" bands via row-intensity projection; pick the band matching the
          requested vertical position (top/middle/bottom or explicit --rows).
  3. REFINE (Pillow) ....... pad by a margin, then SNAP each edge. Default is
     border-aware: snap each edge OUTWARD to the table's ruled border (an ink
     ridge) so the border and the full block are captured and never clipped;
     edges with no ruled border fall back to the inward whitespace-gutter snap
     so neighbours are still excluded. --no-border forces pure inward snapping.
  4. VALIDATE .............. sanity-check aspect ratio + ink coverage; if OCR is
     available, assert the expected structure (>= N numeric price columns, the
     title captured, first & last row anchors inside). Emit a confidence score.
  5. SAVE .................. write PNG; optionally write a sidecar .json with the
     final box, strategy used, and confidence for reproducibility/audit.

Dependencies
------------
  Required : Pillow, numpy          (both in the skill .venv)
  Optional : pytesseract + tesseract binary  (upgrades detection + validation)
             If absent, the script still runs via the projection-profile path
             and prints a warning that OCR-grade validation was skipped.

CLI
---
  crop_section.py <source_image> -o <out.png>
      [--title "SPECIAL DISHES"] [--next "APPETIZERS"]
      [--rows top|middle|bottom]         # projection-profile band selector
      [--box L,T,R,B]                    # bypass detection (manual box)
      [--min-price-cols N]               # validation threshold (default 3)
      [--margin PX] [--pad PX]
      [--sidecar]                        # also write <out>.json audit file
      [--debug]                          # print detection diagnostics

Exit codes: 0 = crop saved & validated; 1 = saved but validation warned;
            2 = failed (could not detect / bad args).
"""

import argparse
import json
import os
import sys

try:
    import numpy as np
except Exception as e:  # pragma: no cover
    print(f"FATAL: numpy is required ({e})", file=sys.stderr)
    sys.exit(2)

try:
    from PIL import Image, ImageOps
except Exception as e:  # pragma: no cover
    print(f"FATAL: Pillow is required ({e})", file=sys.stderr)
    sys.exit(2)

# Optional OCR. Absence is fine — we degrade to projection-profile detection.
try:
    import pytesseract
    from pytesseract import Output as _TessOutput

    # Probe the actual binary; the python pkg alone is not enough.
    try:
        pytesseract.get_tesseract_version()
        HAVE_OCR = True
    except Exception:
        HAVE_OCR = False
except Exception:
    pytesseract = None
    HAVE_OCR = False


# --------------------------------------------------------------------------- #
# 1. LOAD + NORMALIZE
# --------------------------------------------------------------------------- #
def load_normalized(path, deskew=True, debug=False):
    """Return (rgb_image, gray_ndarray). Auto-orient via EXIF; light deskew."""
    im = Image.open(path)
    im = ImageOps.exif_transpose(im)          # honour camera/scanner rotation
    im = im.convert("RGB")
    if deskew:
        im = _deskew(im, debug=debug)
    gray = np.asarray(im.convert("L"), dtype=np.float32)
    return im, gray


def _deskew(im, max_angle=4.0, step=0.5, debug=False):
    """
    Estimate small scan skew by maximizing the variance of the horizontal
    ink-projection across candidate rotations (well-aligned text => sharp,
    high-variance row profile). Pure Pillow+numpy; no OpenCV needed.
    """
    base = np.asarray(im.convert("L"), dtype=np.float32)
    ink = 255.0 - base                        # ink = dark
    best_angle, best_score = 0.0, -1.0
    a = -max_angle
    while a <= max_angle + 1e-9:
        rot = im.convert("L").rotate(a, resample=Image.BILINEAR,
                                     fillcolor=255, expand=False)
        prof = (255.0 - np.asarray(rot, dtype=np.float32)).sum(axis=1)
        score = float(prof.var())
        if score > best_score:
            best_score, best_angle = score, a
        a += step
    if abs(best_angle) < step:                # nothing meaningful; leave as-is
        return im
    if debug:
        print(f"[deskew] rotating {best_angle:+.2f} deg", file=sys.stderr)
    return im.rotate(best_angle, resample=Image.BICUBIC, fillcolor=(255, 255, 255),
                     expand=False)


# --------------------------------------------------------------------------- #
# 2. DETECT
# --------------------------------------------------------------------------- #
def detect_ocr_anchored(im, title, next_section, debug=False):
    """
    OCR-anchored coarse box. Returns (box, meta) or (None, reason).
    box spans: just above title top -> just above next-section top,
    across the horizontal extent of the words on the title/data rows.
    """
    if not HAVE_OCR:
        return None, "ocr-unavailable"
    data = pytesseract.image_to_data(im, output_type=_TessOutput.DICT)
    words = [
        {"text": (data["text"][i] or "").strip(),
         "l": data["left"][i], "t": data["top"][i],
         "w": data["width"][i], "h": data["height"][i],
         "conf": _safe_float(data["conf"][i])}
        for i in range(len(data["text"]))
        if (data["text"][i] or "").strip()
    ]
    if not words:
        return None, "ocr-no-words"

    def _find_phrase(phrase):
        toks = phrase.lower().split()
        hits = []
        for i in range(len(words)):
            if words[i]["text"].lower() == toks[0]:
                seq = words[i:i + len(toks)]
                if len(seq) == len(toks) and all(
                        seq[k]["text"].lower() == toks[k] for k in range(len(toks))):
                    l = min(s["l"] for s in seq)
                    t = min(s["t"] for s in seq)
                    r = max(s["l"] + s["w"] for s in seq)
                    b = max(s["t"] + s["h"] for s in seq)
                    hits.append((l, t, r, b))
        return hits

    title_hits = _find_phrase(title)
    if not title_hits:
        return None, f"title-not-found:{title!r}"
    tl, tt, tr, tb = title_hits[0]

    top = max(0, tt - int(0.6 * (tb - tt)))     # a little headroom above title
    # bottom: first next-section anchor strictly below the title
    bottom = None
    if next_section:
        for (nl, nt, nr, nb) in sorted(_find_phrase(next_section), key=lambda x: x[1]):
            if nt > tb:
                bottom = max(top + 1, nt - int(0.3 * (nb - nt)))
                break
    if bottom is None:
        # fall back to the lowest data word in the title's column band
        band = [w for w in words if w["t"] > tb and w["l"] < im.width * 0.55]
        bottom = (max(w["t"] + w["h"] for w in band) + 10) if band else im.height

    # width: use ink extent within [top, bottom]
    left, right = _ink_horizontal_extent(im, top, bottom)
    box = (left, top, right, bottom)
    if debug:
        print(f"[detect/ocr] title={title_hits[0]} box={box}", file=sys.stderr)
    return box, {"strategy": "ocr-anchored"}


def detect_projection(gray, rows="top", explicit=None, debug=False):
    """
    Projection-profile fallback (no OCR). Finds horizontal ink bands and picks
    one by position. Returns (box, meta). `explicit`=(t,b) overrides selection.
    """
    h, w = gray.shape
    ink = 255.0 - gray
    row_ink = ink.sum(axis=1)
    thr = row_ink.max() * 0.06                 # rows with >6% of peak ink = text
    mask = row_ink > thr
    bands = _contiguous_true_bands(mask, min_gap=int(0.015 * h), min_len=int(0.03 * h))
    if not bands:
        return None, "no-ink-bands"

    if explicit is not None:
        t, b = explicit
    else:
        bands_sorted = sorted(bands, key=lambda ab: (ab[1] - ab[0]), reverse=True)
        # the pricing table is a tall band; bias by requested vertical region
        if rows == "top":
            cand = sorted(bands, key=lambda ab: ab[0])
        elif rows == "bottom":
            cand = sorted(bands, key=lambda ab: ab[0], reverse=True)
        else:  # middle
            cand = sorted(bands, key=lambda ab: abs((ab[0] + ab[1]) / 2 - h / 2))
        # prefer a reasonably tall band near the requested region
        tall = [ab for ab in cand if (ab[1] - ab[0]) >= 0.10 * h]
        t, b = (tall[0] if tall else cand[0])

    left, right = _ink_horizontal_extent_arr(ink, t, b)
    box = (left, t, right, b)
    if debug:
        print(f"[detect/proj] bands={len(bands)} picked=({t},{b}) box={box}",
              file=sys.stderr)
    return box, {"strategy": "projection-profile"}


# --------------------------------------------------------------------------- #
# 3. REFINE (Pillow-side geometry)
# --------------------------------------------------------------------------- #
def refine_box(gray, box, margin=6, pad=4, border=True, debug=False):
    """
    Refine the detected box's edges. Default (`border=True`) is border-aware: for
    each edge we first look OUTWARD for a ruled table border (an ink ridge) and
    snap to it so the border line — and everything inside it — is captured; this
    is the fix for pricing blocks whose outer rule was being clipped. If no ruled
    border is found on an edge, we fall back to the original inward whitespace-
    gutter snap (so borderless blocks still exclude neighbours). `border=False`
    restores the pure whitespace-gutter behaviour. Returns the refined integer
    box, clamped to image bounds.
    """
    h, w = gray.shape
    l, t, r, b = box
    l = max(0, l - margin); t = max(0, t - margin)
    r = min(w, r + margin); b = min(h, b + margin)

    ink = 255.0 - gray

    def _edge(profile, origin, out_toward, ws_toward, default):
        if border:
            bd = _snap_border_edge(profile, origin, toward=out_toward, default=None)
            if bd is not None:
                return bd
        return _snap_edge(profile, origin, toward=ws_toward, default=default)

    # top/bottom: look for a horizontal ruled border, else row-whitespace
    t = _edge(ink[max(0, t - 20):t + 20, l:r].sum(axis=1), origin=max(0, t - 20),
              out_toward="out_start", ws_toward="down", default=t)
    b = _edge(ink[b - 20:min(h, b + 20), l:r].sum(axis=1), origin=b - 20,
              out_toward="out_end", ws_toward="up", default=b)
    # left/right: look for a vertical ruled border, else column-whitespace
    l = _edge(ink[t:b, max(0, l - 30):l + 30].sum(axis=0), origin=max(0, l - 30),
              out_toward="out_start", ws_toward="down", default=l)
    r = _edge(ink[t:b, r - 30:min(w, r + 30)].sum(axis=0), origin=r - 30,
              out_toward="out_end", ws_toward="up", default=r)

    l = max(0, l - pad); t = max(0, t - pad)
    r = min(w, r + pad); b = min(h, b + pad)
    box = (int(l), int(t), int(r), int(b))
    if debug:
        print(f"[refine] border={border} -> {box}", file=sys.stderr)
    return box


# --------------------------------------------------------------------------- #
# 4. VALIDATE
# --------------------------------------------------------------------------- #
def validate_crop(im, box, min_price_cols=3, row_anchors=("D 1", "D 10"),
                  debug=False):
    """
    Returns (ok: bool, confidence: float 0..1, notes: list[str]).
    Geometry checks always run. If OCR is available, structural checks run too.
    """
    notes = []
    l, t, r, b = box
    W, H = r - l, b - t
    conf = 1.0

    if W < 40 or H < 20:
        return False, 0.0, [f"crop too small ({W}x{H})"]
    ar = W / max(1, H)
    if ar < 1.2:
        conf -= 0.2; notes.append(f"aspect ratio {ar:.2f} low for a wide price table")

    crop = im.crop(box)
    ink_frac = float((np.asarray(crop.convert("L")) < 200).mean())
    if ink_frac < 0.01:
        return False, 0.0, ["crop is almost blank"]
    if ink_frac > 0.6:
        conf -= 0.2; notes.append(f"ink fraction {ink_frac:.2f} high (may include a graphic)")

    if HAVE_OCR:
        txt = pytesseract.image_to_string(crop)
        # count numeric price tokens per line; a multi-col table has lines with >=N
        max_prices = 0
        import re
        for line in txt.splitlines():
            n = len(re.findall(r"\d+\.\d{2}", line))
            max_prices = max(max_prices, n)
        if max_prices < min_price_cols:
            conf -= 0.4
            notes.append(f"only {max_prices} price columns detected (< {min_price_cols})")
        for anchor in row_anchors:
            if anchor and anchor.replace(" ", "").lower() not in txt.replace(" ", "").lower():
                conf -= 0.1
                notes.append(f"row anchor {anchor!r} not found in crop")
    else:
        conf -= 0.15
        notes.append("OCR unavailable: structural validation skipped")

    conf = max(0.0, min(1.0, conf))
    ok = conf >= 0.6 and not any("too small" in n or "blank" in n for n in notes)
    if debug:
        print(f"[validate] ok={ok} conf={conf:.2f} notes={notes}", file=sys.stderr)
    return ok, conf, notes


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def _safe_float(x):
    try:
        return float(x)
    except Exception:
        return -1.0


def _contiguous_true_bands(mask, min_gap, min_len):
    """Merge True runs in a 1-D bool mask, closing gaps < min_gap, keeping runs
    >= min_len. Returns list of (start, end) exclusive-end."""
    idx = np.where(mask)[0]
    if idx.size == 0:
        return []
    bands = []
    start = prev = idx[0]
    for i in idx[1:]:
        if i - prev <= min_gap:
            prev = i
        else:
            bands.append((int(start), int(prev) + 1))
            start = prev = i
    bands.append((int(start), int(prev) + 1))
    return [ab for ab in bands if (ab[1] - ab[0]) >= min_len]


def _ink_horizontal_extent(im, top, bottom, frac=0.02):
    return _ink_horizontal_extent_arr(255.0 - np.asarray(im.convert("L"), np.float32),
                                      top, bottom, frac)


def _ink_horizontal_extent_arr(ink, top, bottom, frac=0.02):
    """Left/right pixels where column ink first/last exceeds a small threshold."""
    band = ink[top:bottom, :]
    col = band.sum(axis=0)
    if col.max() <= 0:
        return 0, ink.shape[1]
    thr = col.max() * frac
    cols = np.where(col > thr)[0]
    if cols.size == 0:
        return 0, ink.shape[1]
    return int(cols[0]), int(cols[-1]) + 1


def _snap_edge(profile, origin, toward, default):
    """
    Given a 1-D ink profile over a small window starting at `origin`, return the
    absolute coordinate of the whitespace minimum. `toward`='down' scans from the
    window start; 'up' from the window end. Falls back to `default`.
    """
    if profile is None or len(profile) == 0:
        return default
    p = np.asarray(profile, dtype=np.float32)
    lo = p.min()
    # candidate gutter positions = near-minimum ink
    gutters = np.where(p <= lo + 0.05 * (p.max() - lo + 1e-6))[0]
    if gutters.size == 0:
        return default
    pick = gutters[0] if toward == "down" else gutters[-1]
    return int(origin + pick)


def _snap_border_edge(profile, origin, toward, default):
    """
    Border-aware edge snap. A ruled table border shows up as a sharp ink RIDGE
    (a sustained peak) rather than a whitespace minimum, so for bordered pricing
    blocks we want the OUTERMOST such ridge — capturing the border line itself,
    not the whitespace just inside it. `toward`='out_start' means the outer edge
    is at the window START (used for top/left, which we extend outward before the
    original edge); 'out_end' means the outer edge is at the window END (used for
    bottom/right). Returns the ridge coordinate, or None if no clear ridge is
    present (caller then falls back to `_snap_edge`).
    """
    if profile is None or len(profile) == 0:
        return None
    p = np.asarray(profile, dtype=np.float32)
    med = float(np.median(p))
    hi = float(p.max())
    # A ruled line is a strong outlier above the local median. Require a real
    # separation so plain text rows/columns don't register as a "border".
    if hi < med + 0.5 * (hi - med + 1e-6) or (hi - med) < 1e-3:
        return None
    ridge = np.where(p >= med + 0.6 * (hi - med))[0]
    if ridge.size == 0:
        return None
    # Pick the ridge nearest the OUTER side of the window so the border line is
    # inside the crop, not cropped away.
    pick = ridge[0] if toward == "out_start" else ridge[-1]
    return int(origin + pick)


# --------------------------------------------------------------------------- #
# orchestration
# --------------------------------------------------------------------------- #
def crop_section(source, out_path, title="SPECIAL DISHES", next_section="APPETIZERS",
                 rows="top", explicit_box=None, min_price_cols=3, margin=6, pad=4,
                 sidecar=False, deskew=True, refine=True, border=True, debug=False):
    im, gray = load_normalized(source, deskew=deskew, debug=debug)

    if explicit_box is not None:
        box, meta = explicit_box, {"strategy": "manual"}
        refine = False                          # honour the user's exact edges
    else:
        box, meta = detect_ocr_anchored(im, title, next_section, debug=debug)
        if box is None:
            if debug:
                print(f"[detect] ocr path unavailable/failed ({meta}); "
                      f"using projection profile", file=sys.stderr)
            box, meta = detect_projection(gray, rows=rows, debug=debug)
            if box is None:
                print(f"FATAL: detection failed ({meta})", file=sys.stderr)
                return 2

    if refine:
        box = refine_box(gray, box, margin=margin, pad=pad, border=border, debug=debug)
    ok, conf, notes = validate_crop(im, box, min_price_cols=min_price_cols, debug=debug)

    crop = im.crop(box)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    crop.save(out_path)

    if sidecar:
        side = os.path.splitext(out_path)[0] + ".json"
        with open(side, "w") as f:
            json.dump({"source": os.path.abspath(source), "box": list(box),
                       "strategy": meta.get("strategy"), "confidence": round(conf, 3),
                       "ocr_available": HAVE_OCR, "notes": notes,
                       "size": list(crop.size)}, f, indent=2)

    status = "OK" if ok else "WARN"
    print(f"[{status}] {out_path}  box={box}  size={crop.size}  "
          f"strategy={meta.get('strategy')}  confidence={conf:.2f}")
    if notes:
        for n in notes:
            print(f"    - {n}")
    if not HAVE_OCR:
        print("    (tip: install tesseract + pytesseract for OCR-grade detection "
              "and validation)")
    return 0 if ok else 1


def _parse_box(s):
    parts = [int(x) for x in s.split(",")]
    if len(parts) != 4:
        raise argparse.ArgumentTypeError("--box must be L,T,R,B")
    return tuple(parts)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Crop the multi-column pricing block "
                                             "from a paper-menu image.")
    ap.add_argument("source", help="source menu image (jpg/png)")
    ap.add_argument("-o", "--out", required=True, help="output PNG path")
    ap.add_argument("--title", default="SPECIAL DISHES",
                    help="section title anchor for OCR detection")
    ap.add_argument("--next", dest="next_section", default="APPETIZERS",
                    help="next-section title anchor (bottom bound)")
    ap.add_argument("--rows", choices=["top", "middle", "bottom"], default="top",
                    help="projection-profile band selector when OCR is unavailable")
    ap.add_argument("--box", type=_parse_box, default=None,
                    help="bypass detection with an explicit L,T,R,B box")
    ap.add_argument("--min-price-cols", type=int, default=3,
                    help="validation: minimum numeric price columns expected")
    ap.add_argument("--margin", type=int, default=6, help="pre-snap padding (px)")
    ap.add_argument("--pad", type=int, default=4, help="final padding (px)")
    ap.add_argument("--no-deskew", action="store_true", help="disable deskew step")
    ap.add_argument("--no-refine", action="store_true",
                    help="disable whitespace-snap refinement (kept exact box)")
    ap.add_argument("--no-border", action="store_true",
                    help="disable border-aware refinement; use pure inward "
                         "whitespace-gutter snap (default snaps OUTWARD to the "
                         "table's ruled border so it is never clipped)")
    ap.add_argument("--sidecar", action="store_true",
                    help="also write <out>.json audit file")
    ap.add_argument("--debug", action="store_true", help="print diagnostics")
    args = ap.parse_args(argv)

    return crop_section(args.source, args.out, title=args.title,
                        next_section=args.next_section, rows=args.rows,
                        explicit_box=args.box, min_price_cols=args.min_price_cols,
                        margin=args.margin, pad=args.pad, sidecar=args.sidecar,
                        deskew=not args.no_deskew, refine=not args.no_refine,
                        border=not args.no_border, debug=args.debug)


if __name__ == "__main__":
    sys.exit(main())