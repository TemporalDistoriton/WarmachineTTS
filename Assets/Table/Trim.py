#!/usr/bin/env python3
"""
Centre-crop non-square images to square, and also generate thumbnails.

Default behaviour:
- Reads images from a folder (recursively)
- If width != height, crops a centred square using the smaller dimension
- Optionally downscales the square output (max-size)
- Writes squares to an output folder (or --inplace overwrites originals)
- Writes thumbnails to a THUMBS folder (mirrors subfolders)

Requires:
  pip install pillow

Usage:
  python make_square_crops_and_thumbs.py "Top Images" --out "Top Images (square)"
  python make_square_crops_and_thumbs.py "Top Images" --out "Top Images (square)" --thumb-size 256
  python make_square_crops_and_thumbs.py "Top Images" --inplace --thumb-size 256
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable, Optional

from PIL import Image, ImageOps

SUPPORTED_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif", ".tif", ".tiff"}


def iter_images(root: Path, recursive: bool = True) -> Iterable[Path]:
    if recursive:
        for p in root.rglob("*"):
            if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS:
                yield p
    else:
        for p in root.glob("*"):
            if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS:
                yield p


def centre_crop_square(img: Image.Image) -> Image.Image:
    w, h = img.size
    if w == h:
        return img
    side = min(w, h)
    left = (w - side) // 2
    top = (h - side) // 2
    return img.crop((left, top, left + side, top + side))


def maybe_resize_square(img: Image.Image, max_size: Optional[int]) -> Image.Image:
    if not max_size:
        return img
    w, h = img.size
    if w <= max_size and h <= max_size:
        return img
    # Square by design here, so resize to max_size x max_size
    return img.resize((max_size, max_size), resample=Image.Resampling.LANCZOS)


def make_thumbnail(img: Image.Image, thumb_size: int) -> Image.Image:
    # img is square, create a copy and thumbnail it (keeps aspect ratio)
    t = img.copy()
    t.thumbnail((thumb_size, thumb_size), resample=Image.Resampling.LANCZOS)
    return t


def save_image(out_path: Path, img: Image.Image) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)

    ext = out_path.suffix.lower()
    save_kwargs = {}

    # JPEG cannot store alpha, so flatten if needed
    if ext in {".jpg", ".jpeg"}:
        if img.mode in {"RGBA", "LA"}:
            bg = Image.new("RGB", img.size, (0, 0, 0))
            bg.paste(img, mask=img.split()[-1])
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")
        save_kwargs.update({"quality": 90, "optimize": True, "progressive": True})

    elif ext == ".png":
        save_kwargs.update({"optimize": True})

    img.save(out_path, **save_kwargs)


def main() -> int:
    ap = argparse.ArgumentParser(description="Centre-crop to square and generate thumbnails into THUMBS.")
    ap.add_argument("folder", help="Input folder containing images")
    ap.add_argument("--out", default=None, help="Output folder for square images (default: '<folder> (square)')")
    ap.add_argument("--inplace", action="store_true", help="Overwrite originals in place (careful)")
    ap.add_argument("--no-recursive", action="store_true", help="Do not scan subfolders")
    ap.add_argument("--max-size", type=int, default=None, help="Optional max square size, downscale if larger")
    ap.add_argument("--thumb-size", type=int, default=256, help="Thumbnail size in px (default: 256)")
    ap.add_argument("--thumb-format", choices=["same", "png", "jpg", "webp"], default="same",
                    help="Thumbnail file format (default: same as source/output)")
    ap.add_argument("--dry-run", action="store_true", help="Show what would be done, write nothing")
    args = ap.parse_args()

    in_dir = Path(args.folder)
    if not in_dir.exists() or not in_dir.is_dir():
        print(f"ERROR: Folder not found: {in_dir}")
        return 2

    recursive = not args.no_recursive

    if args.inplace and args.out:
        print("ERROR: Use either --inplace or --out, not both.")
        return 2

    # Destination for square outputs
    if args.inplace:
        square_root = in_dir
    else:
        square_root = Path(args.out) if args.out else Path(str(in_dir) + " (square)")

    # Thumbs folder sits alongside square_root
    thumbs_root = square_root / "THUMBS"

    changed = 0
    skipped = 0
    failed = 0

    for src in iter_images(in_dir, recursive=recursive):
        rel = src.relative_to(in_dir)

        try:
            with Image.open(src) as im:
                im = ImageOps.exif_transpose(im)
                w, h = im.size

                # Create square version (crop if needed)
                sq = centre_crop_square(im)

                # Optional max-size downscale
                sq = maybe_resize_square(sq, args.max_size)

                # Determine whether we need to write square output:
                # - If not square originally, we must write
                # - If max-size forces downscale, we must write
                need_square_write = (w != h) or (args.max_size is not None and sq.size[0] != w)

                # Square output path
                square_path = (square_root / rel) if not args.inplace else src

                # Thumbnail output path (mirrors rel path under THUMBS)
                if args.thumb_format == "same":
                    thumb_suffix = square_path.suffix
                else:
                    thumb_suffix = "." + args.thumb_format

                thumb_path = (thumbs_root / rel).with_suffix(thumb_suffix)

                # Always generate thumbs (from square), even if square write is skipped
                thumb_img = make_thumbnail(sq, args.thumb_size)

                if args.dry_run:
                    action_bits = []
                    if need_square_write:
                        action_bits.append(f"SQUARE {w}x{h}->{sq.size[0]}x{sq.size[1]}")
                    else:
                        action_bits.append("SQUARE skip (already OK)")
                    action_bits.append(f"THUMB -> {thumb_img.size[0]}x{thumb_img.size[1]}")
                    print(f"{rel}: " + ", ".join(action_bits))
                    changed += 1
                    continue

                # Write square output if needed (or always if --inplace and it changed)
                if need_square_write:
                    save_image(square_path, sq)
                    changed += 1
                else:
                    skipped += 1

                # Write thumb
                save_image(thumb_path, thumb_img)

        except Exception as e:
            failed += 1
            print(f"FAILED: {rel} ({e})")

    print(f"Done. Square written: {changed}, Square skipped: {skipped}, Failed: {failed}")
    if not args.dry_run:
        print(f"Square output root: {square_root.resolve()}")
        print(f"Thumbs output root: {thumbs_root.resolve()}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())