#!/usr/bin/env python3
"""
cleanup_stl.py
Scans the directory it lives in (and all subfolders) for *.stl files.
- If a matching *.obj exists  → deletes the *.stl
- If no matching *.obj exists → logs the orphaned *.stl
Prints a summary at the end.
"""

import os
import sys
from pathlib import Path


def main():
    script_dir = Path(__file__).resolve().parent

    stl_files = list(script_dir.rglob("*.stl"))

    if not stl_files:
        print("No .stl files found. Nothing to do.")
        sys.exit(0)

    deleted = []
    orphaned = []

    for stl_path in stl_files:
        # OBJ files may have had spaces removed from the name
        obj_name_nospaces = stl_path.stem.replace(" ", "") + ".obj"
        obj_path_exact = stl_path.with_suffix(".obj")
        obj_path_nospaces = stl_path.parent / obj_name_nospaces

        # Match either "Icy Grip.obj" or "IcyGrip.obj"
        obj_path = obj_path_exact if obj_path_exact.exists() else obj_path_nospaces

        if obj_path.exists():  # will be whichever variant was found
            try:
                stl_path.unlink()
                deleted.append(stl_path)
                print(f"  [DELETED]  {stl_path.relative_to(script_dir)}")
            except OSError as e:
                print(f"  [ERROR]    Could not delete {stl_path.relative_to(script_dir)}: {e}")
        else:
            orphaned.append(stl_path)
            print(f"  [NO OBJ]   {stl_path.relative_to(script_dir)}")

    print()
    print("=" * 50)
    print(f"  STL files deleted (had matching OBJ): {len(deleted)}")
    print(f"  STL files kept   (no matching OBJ):   {len(orphaned)}")
    print("=" * 50)

    if orphaned:
        print("\nOrphaned STL files (no corresponding OBJ):")
        for p in orphaned:
            print(f"  {p.relative_to(script_dir)}")


if __name__ == "__main__":
    main()