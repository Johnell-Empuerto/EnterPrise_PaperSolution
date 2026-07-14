import hashlib, json
from typing import List, Optional, Tuple
from CoordinateEngine.models.field_def import FieldGeometry


def validate_deterministic(run1_path: str, run2_path: str) -> Tuple[bool, str, str, str]:
    """
    Compare two engine runs for deterministic output.
    Returns (match, msg, hash1, hash2).
    """
    try:
        with open(run1_path, "rb") as f:
            data1 = f.read()
            h1 = hashlib.sha256(data1).hexdigest()
    except FileNotFoundError:
        return False, f"File not found: {run1_path}", "", ""

    try:
        with open(run2_path, "rb") as f:
            data2 = f.read()
            h2 = hashlib.sha256(data2).hexdigest()
    except FileNotFoundError:
        return False, f"File not found: {run2_path}", h1, ""

    match = h1 == h2
    if match:
        return True, f"Deterministic: SHA-256 matches ({h1[:16]}...)", h1, h2
    else:
        return False, f"NON-DETERMINISTIC: SHA-256 differs ({h1[:16]}... vs {h2[:16]}...)", h1, h2


def validate_fields(geometries: List[FieldGeometry]) -> Tuple[bool, str]:
    """
    Validate field geometries for sanity.
    Checks: no negative coordinates, non-positive dimensions, ratio ranges.
    """
    if not geometries:
        return False, "No fields to validate"

    for g in geometries:
        if g.left_px < 0 or g.top_px < 0:
            return False, f"Negative coordinate for {g.field.addr}: left={g.left_px}, top={g.top_px}"
        if g.width_px <= 0 or g.height_px <= 0:
            return False, f"Non-positive dimension for {g.field.addr}: w={g.width_px}, h={g.height_px}"
        if g.left_ratio < 0 or g.left_ratio > 1.5:
            return False, f"Left ratio out of range for {g.field.addr}: {g.left_ratio}"
        if g.top_ratio < 0 or g.top_ratio > 1.5:
            return False, f"Top ratio out of range for {g.field.addr}: {g.top_ratio}"

    return True, f"All {len(geometries)} fields valid"


def validate_coordinate_consistency(
    geometries: List[FieldGeometry],
) -> Tuple[bool, List[str]]:
    """
    Check internal consistency of coordinates:
    - page_left_pt + page_width_pt should be within page bounds
    - no overlapping fields (unless merged)
    """
    warnings = []

    for g in geometries:
        if g.page_left_pt < 0 or g.page_top_pt < 0:
            warnings.append(f"{g.field.addr}: negative page position")

    # Check for overlaps (non-merged fields)
    for i, a in enumerate(geometries):
        if a.field.is_merge:
            continue
        for b in geometries[i + 1:]:
            if b.field.is_merge:
                continue
            # Simple overlap check
            a_right = a.page_left_pt + a.page_width_pt
            a_bottom = a.page_top_pt + a.page_height_pt
            b_right = b.page_left_pt + b.page_width_pt
            b_bottom = b.page_top_pt + b.page_height_pt

            if (a.page_left_pt < b_right and a_right > b.page_left_pt
                    and a.page_top_pt < b_bottom and a_bottom > b.page_top_pt):
                warnings.append(
                    f"Possible overlap: {a.field.addr} and {b.field.addr}"
                )

    return len(warnings) == 0, warnings


def compute_deterministic_hash(file_path: str) -> Optional[str]:
    """Compute SHA-256 hash of a runtime.json file."""
    try:
        with open(file_path, "rb") as f:
            return hashlib.sha256(f.read()).hexdigest()
    except FileNotFoundError:
        return None
