#!/usr/bin/env python3
"""
compare_screenshots.py — Pixel-level PNG screenshot comparison for PaperLess.

Compares an original workbook's print-preview screenshot against the edited
workbook's print-preview screenshot and reports every pixel difference.

Usage:
    python compare_screenshots.py original.png edited.png [--tolerance 0] [--output diff.png]

Requirements:
    pip install Pillow
"""

import argparse
import os
import sys
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    print("ERROR: Pillow is required. Install with: pip install Pillow")
    sys.exit(1)


def compare_images(original_path: str, edited_path: str,
                   tolerance: int = 0, output_path: str | None = None) -> dict:
    """
    Compare two PNG images pixel by pixel.

    Returns a dict with:
        - passed: bool
        - total_pixels: int
        - diff_pixels: int
        - max_diff: int (max per-channel difference)
        - avg_diff: float (average per-channel difference across all pixels)
        - diff_percent: float
        - dims_match: bool
        - diff_regions: list of (x, y, w, h) bounding boxes for differing areas
    """
    orig = Image.open(original_path).convert("RGBA")
    edit = Image.open(edited_path).convert("RGBA")

    result = {
        "passed": True,
        "total_pixels": 0,
        "diff_pixels": 0,
        "max_diff": 0,
        "avg_diff": 0.0,
        "diff_percent": 0.0,
        "dims_match": True,
        "diff_regions": [],
        "orig_size": orig.size,
        "edit_size": edit.size,
    }

    # Check dimensions match
    if orig.size != edit.size:
        result["passed"] = False
        result["dims_match"] = False
        result["orig_size"] = orig.size
        result["edit_size"] = edit.size
        return result

    w, h = orig.size
    result["total_pixels"] = w * h

    orig_pixels = orig.load()
    edit_pixels = edit.load()

    total_diff = 0
    diff_coords = []
    max_diff_val = 0

    for y in range(h):
        for x in range(w):
            pr, pg, pb, pa = orig_pixels[x, y]
            er, eg, eb, ea = edit_pixels[x, y]

            diff = abs(pr - er) + abs(pg - eb) + abs(pb - eg) + abs(pa - ea)
            if diff > max_diff_val:
                max_diff_val = diff

            if diff > tolerance:
                result["diff_pixels"] += 1
                total_diff += diff
                diff_coords.append((x, y))

    result["max_diff"] = max_diff_val
    if result["diff_pixels"] > 0:
        result["avg_diff"] = total_diff / result["diff_pixels"]
        result["diff_percent"] = (result["diff_pixels"] / result["total_pixels"]) * 100.0
        result["passed"] = False

    # Compute diff regions (bounding boxes around differing areas)
    if diff_coords and output_path:
        # Create a diff highlight image
        diff_img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        diff_pixels_map = diff_img.load()
        for dx, dy in diff_coords:
            diff_pixels_map[dx, dy] = (255, 0, 0, 255)  # Red highlights
        diff_img.save(output_path)

    return result


def main():
    parser = argparse.ArgumentParser(
        description="Compare two PNG screenshots pixel by pixel"
    )
    parser.add_argument("original", help="Path to original screenshot (PNG)")
    parser.add_argument("edited", help="Path to edited screenshot (PNG)")
    parser.add_argument("--tolerance", type=int, default=0,
                        help="Per-channel tolerance (default: 0)")
    parser.add_argument("--output", "-o", default=None,
                        help="Output path for diff highlight image (PNG)")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Show detailed results")
    args = parser.parse_args()

    for p in [args.original, args.edited]:
        if not os.path.exists(p):
            print(f"ERROR: File not found: {p}")
            sys.exit(1)

    print(f"Comparing: {args.original} vs {args.edited}")
    print(f"Tolerance: {args.tolerance}")
    print()

    result = compare_images(args.original, args.edited,
                            tolerance=args.tolerance,
                            output_path=args.output)

    print(f"Dimensions match: {result['dims_match']}")
    if not result['dims_match']:
        print(f"  Original: {result['orig_size']}")
        print(f"  Edited:   {result['edit_size']}")
        print("\nFAILED: Dimensions differ")
        sys.exit(1)

    print(f"Total pixels:      {result['total_pixels']:,}")
    print(f"Diff pixels:       {result['diff_pixels']:,}")
    print(f"Diff percent:      {result['diff_percent']:.4f}%")
    print(f"Max per-channel:   {result['max_diff']}")
    print(f"Avg diff/pixel:    {result['avg_diff']:.2f}")
    print(f"Result:            {'PASS' if result['passed'] else 'FAIL'}")

    if args.output:
        print(f"Diff image saved:  {args.output}")

    if result["diff_pixels"] > 0 and args.verbose:
        print(f"\n{result['diff_pixels']} differing pixels found.")

    sys.exit(0 if result['passed'] else 1)


if __name__ == "__main__":
    main()
