# Phase 12.0 — Implement the Legacy Coordinate Engine

## Status: Pending (Blocked on Phase 11J.18)

## Objective

Replace the current experimental coordinate calculation with the verified legacy ConMas coordinate generation algorithm. Transition from investigation to production implementation.

## Preconditions

**Proceed only if Phase 11J.18 concluded:**
- `def_file` is the verified workbook used by ConMas
- `background_image_file` is a direct (or functionally equivalent) Excel export
- The source coordinate space has been identified
- The complete coordinate transformation is understood
- No additional hidden processing stages remain

If any prerequisite is not met, stop implementation and document the remaining unknown.

---

## Goals

The new coordinate engine must generate coordinates that match the legacy database without:
- PNG calibration
- Pixel scanning
- Image analysis
- Manual offsets
- Hardcoded template-specific corrections
- Post-processing adjustments

The coordinate system must come from a single deterministic source.

---

## Remove Experimental Logic

Remove all investigation-era temporary logic:
- `AdjustCoordinatesFromPng()`
- Runtime coordinate calibration
- Pixel-based correction
- Manual translation offsets
- Temporary debug calculations
- Investigation-only logging
- Coordinate correction flags
- Experimental scaling patches

Final pipeline must contain only production code.

---

## Production Pipeline

```
Upload Workbook
        │
        ▼
Open Workbook (Excel COM)
        │
        ▼
Read Workbook Geometry
        │
        ▼
Read Page Setup
        │
        ▼
Calculate Legacy Page Coordinates
        │
        ▼
Generate Runtime Coordinates
        │
        ├── runtime.json
        ├── PDF
        └── PNG Preview
```

All outputs must originate from the same coordinate system.

---

## Coordinate Generation

Implement the verified legacy algorithm exactly as determined in Phase 11J.18:

1. Read worksheet geometry from Excel COM
2. Read print layout information (PageSetup, PrintArea, margins, centering)
3. Transform worksheet coordinates into page coordinates
4. Convert page coordinates into runtime coordinates
5. Produce identical coordinates regardless of rendering DPI

**Do not introduce inferred constants** unless Phase 11J.18 proved they are part of the legacy algorithm.

---

## Validation

Run against every available legacy template:

| Template | Fields | Page | Centered | Notes |
|----------|--------|------|----------|-------|
| 546 | 6 | Letter | Yes | Reference |
| 501 | 1 | Letter | Yes | Single field |
| 448 | 5 | Letter | No | FitToPages |
| 228 | 2 | Letter | No | |
| 186 | 1 | A4 Landscape | No | |
| 142 | 1 | A4 Portrait | Yes | |

For every template, compare:
- Generated coordinates vs legacy database coordinates
- PDF overlay alignment

Generate `Validation/` directory with overlays and `validation_report.md`.

---

## Acceptance Criteria

| Criterion | Threshold |
|-----------|-----------|
| Average position error | ≤ 1 pixel |
| Maximum position error | ≤ 2 pixels |
| Overlay visual match | Matches legacy PDF |
| Template-specific logic | None |
| Runtime calibration | None |
| Pixel scanning | None |

---

## Cleanup

Remove investigation-only code:
- Investigation logging
- Temporary comparison methods
- Experimental coordinate utilities
- Reverse-engineering helper methods no longer required

Keep only reusable diagnostic logging useful in production.

---

## Deliverables

| File | Description |
|------|-------------|
| `ExcelCaptureService.cs` | Opens workbook, reads geometry via COM |
| `CoordinateEngine.cs` | Orchestrates coordinate generation pipeline |
| `CoordinateTransformer.cs` | Transforms worksheet → page → runtime coordinates |
| `RuntimeCoordinateGenerator.cs` | Produces runtime.json, PDF, PNG preview |
| `Validation/template_*_overlay.png` | Overlay for each validation template |
| `validation_report.md` | Per-field error table, acceptance criteria results |

---

## Final Verification

| Question | Answer |
|----------|--------|
| Does the new engine reproduce the legacy coordinates? | |
| Does the generated overlay align with the legacy PDF? | |
| Is the coordinate engine deterministic? | |
| Has all experimental correction logic been removed? | |
| Can the engine process new templates without template-specific adjustments? | |

If every answer is **YES**, declare the reverse-engineering project complete and promote the new coordinate engine to production.
