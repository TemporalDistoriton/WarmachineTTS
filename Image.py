"""
Warmachine Image Matcher
=========================
Reads model names from warmachine_units.xlsx, finds the closest matching
image file in the ModelArt folder, copies it and renames it so the path:

    https://raw.githubusercontent.com/TemporalDistoriton/WarmachineTTS/main/ModelArt/<NAME>.png

...resolves correctly for every row in the spreadsheet.

Usage:
    python image_matcher.py

Configuration (edit the variables below):
    EXCEL_FILE   — path to your warmachine_units.xlsx
    IMAGE_DIR    — local path to your ModelArt folder (the repo clone)
    OUTPUT_DIR   — where renamed copies are written  (defaults to IMAGE_DIR)
    REPORT_FILE  — CSV report of every match made / skipped
    DRY_RUN      — set True to preview without copying any files

Matching logic:
    1. Strip common suffixes like numbers, "01"/"02", underscores, spaces
       from both the Excel name and the filename stem.
    2. Score every candidate with fuzzy ratio + token_sort_ratio.
    3. Pick the best score above MIN_SCORE (default 60).
    4. If the target file already exists, skip (no overwrite).
    5. Write a CSV report so you can review every decision.

Requirements:
    pip install pandas openpyxl thefuzz python-Levenshtein
"""

import csv
import re
import shutil
from pathlib import Path

import pandas as pd
from thefuzz import fuzz

# ---------------------------------------------------------------------------
# Configuration — edit these paths
# ---------------------------------------------------------------------------

EXCEL_FILE  = "warmachine_units.xlsx"   # input spreadsheet
IMAGE_DIR   = "ModelArtOld"                # local clone of the ModelArt folder
OUTPUT_DIR  = "ModelArt"                # where renamed copies go (same folder is fine)
REPORT_FILE = "image_match_report.csv"  # audit trail
DRY_RUN     = False                     # True = preview only, no files written
MIN_SCORE   = 60                        # minimum fuzzy match score (0–100)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def normalise(text: str) -> str:
    """
    Reduce a name to its core tokens for comparison.
    - Lowercase
    - Remove trailing numbers / ordinal suffixes  (e.g. "Knight02" → "Knight")
    - Collapse punctuation and extra spaces
    """
    text = text.lower()
    # Remove file extension if present
    text = re.sub(r'\.[a-z]{2,4}$', '', text)
    # Remove trailing digits (and common separators before them)
    text = re.sub(r'[\s_\-]*\d+$', '', text)
    # Replace underscores / hyphens with spaces
    text = re.sub(r'[_\-]+', ' ', text)
    # Collapse whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def score(excel_norm: str, file_norm: str, excel_raw: str, file_raw: str) -> int:
    """
    Return a composite fuzzy score (0–100) between an Excel name and a filename.
    Uses multiple strategies and takes the best.
    """
    scores = [
        fuzz.ratio(excel_norm, file_norm),
        fuzz.partial_ratio(excel_norm, file_norm),
        fuzz.token_sort_ratio(excel_norm, file_norm),
        fuzz.token_set_ratio(excel_norm, file_norm),
        # Also compare raw (unnormalised) strings for exact / near-exact matches
        fuzz.ratio(excel_raw.lower(), file_raw.lower()),
    ]
    return max(scores)


def safe_filename(name: str) -> str:
    """Convert an Excel model name to a valid filename stem."""
    # Replace characters that are illegal in filenames
    name = re.sub(r'[\\/*?:"<>|]', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    excel_path  = Path(EXCEL_FILE)
    image_dir   = Path(IMAGE_DIR)
    output_dir  = Path(OUTPUT_DIR)

    # --- Validate paths ---
    if not excel_path.exists():
        print(f"ERROR: Excel file not found: {excel_path.resolve()}")
        return
    if not image_dir.exists():
        print(f"ERROR: Image directory not found: {image_dir.resolve()}")
        return
    output_dir.mkdir(parents=True, exist_ok=True)

    # --- Load model names ---
    df = pd.read_excel(excel_path)
    if "Name" not in df.columns:
        print("ERROR: No 'Name' column found in the spreadsheet.")
        return

    model_names: list[str] = df["Name"].dropna().unique().tolist()
    print(f"Loaded {len(model_names)} unique model names from {excel_path.name}")

    # --- Index all images in the folder ---
    image_files = [
        f for f in image_dir.iterdir()
        if f.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp") and f.is_file()
    ]
    print(f"Found {len(image_files)} image files in {image_dir.resolve()}")

    if not image_files:
        print("ERROR: No image files found. Check IMAGE_DIR.")
        return

    # Pre-compute normalised stems for all images
    image_index: list[tuple[Path, str]] = [
        (f, normalise(f.stem)) for f in image_files
    ]

    # --- Match and copy ---
    report_rows: list[dict] = []
    copied = skipped = already_exists = no_match = 0

    for model_name in model_names:
        target_stem  = safe_filename(model_name)          # exact target filename
        target_path  = output_dir / f"{target_stem}.png"
        excel_norm   = normalise(model_name)

        # Score every candidate image
        candidates = [
            (img_path, img_norm, score(excel_norm, img_norm, model_name, img_path.stem))
            for img_path, img_norm in image_index
        ]
        candidates.sort(key=lambda x: x[2], reverse=True)

        best_path, best_norm, best_score = candidates[0]

        # --- Decide action ---
        if best_score < MIN_SCORE:
            action  = "NO_MATCH"
            message = f"Best candidate '{best_path.name}' scored {best_score} (below threshold {MIN_SCORE})"
            no_match += 1

        elif target_path.exists() and target_path.resolve() == best_path.resolve():
            action  = "ALREADY_CORRECT"
            message = "Target file already exists and matches"
            already_exists += 1

        elif target_path.exists():
            action  = "SKIPPED_EXISTS"
            message = f"Target '{target_path.name}' already exists (source: '{best_path.name}', score {best_score})"
            skipped += 1

        else:
            action  = "DRY_RUN_COPY" if DRY_RUN else "COPIED"
            message = f"'{best_path.name}' → '{target_path.name}'  (score {best_score})"
            if not DRY_RUN:
                shutil.copy2(best_path, target_path)
            copied += 1

        # Print progress
        status_icon = {"COPIED": "✓", "DRY_RUN_COPY": "~", "NO_MATCH": "✗",
                       "SKIPPED_EXISTS": "→", "ALREADY_CORRECT": "="}.get(action, "?")
        print(f"  [{status_icon}] {model_name!r:45s}  {message}")

        # Top-3 candidates for the report
        top3 = " | ".join(
            f"{p.name} ({s})" for p, _, s in candidates[:3]
        )

        report_rows.append({
            "Model Name":       model_name,
            "Action":           action,
            "Source File":      best_path.name if best_score >= MIN_SCORE else "",
            "Target File":      target_path.name,
            "Score":            best_score,
            "Top 3 Candidates": top3,
            "Message":          message,
        })

    # --- Summary ---
    print()
    print("=" * 60)
    if DRY_RUN:
        print("DRY RUN — no files were written")
    print(f"  Would copy / Copied : {copied}")
    print(f"  Already correct     : {already_exists}")
    print(f"  Skipped (exists)    : {skipped}")
    print(f"  No match found      : {no_match}")
    print(f"  Total models        : {len(model_names)}")
    print("=" * 60)

    # --- Write CSV report ---
    with open(REPORT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=report_rows[0].keys())
        writer.writeheader()
        writer.writerows(report_rows)
    print(f"\nReport saved to: {REPORT_FILE}")

    if no_match > 0:
        print(f"\nTIP: {no_match} model(s) had no match above score {MIN_SCORE}.")
        print("     Lower MIN_SCORE or add images manually for those entries.")
        print("     Check the report CSV for the closest candidates found.")


if __name__ == "__main__":
    main()