# Phase 3 — Production Validation & Geometry Refinement Report

**Date:** 2026-07-11
**Status:** Complete
**Engine:** CoordinateEngine (Python) v1.0

---

## Overview

Phase 3 transforms the Coordinate Engine from a simple coordinate calculator into a **self-validating, self-debugging geometry system**. Every coordinate is now traceable to its source, cross-validated against all available data, and scored for confidence.

**Architecture unchanged.** No frontend, API, or runtime.json schema changes.

---

## New Modules

### `cross_validator.py`
Multi-source geometry cross-validation engine. For each field:
- Derives expected position from **workbook XML** (column widths + row heights)
- Finds matching content in **PDF vector/text elements** near the expected area
- Scans **PNG pixel content** near the expected area
- Compares all three sources and reports differences
- Computes accuracy statistics: mean error, median error, max error, RMS error, standard deviation
- Produces per-source confidence scores based on agreement level

### `geometry_tracer.py`
Per-field derivation trace generator. For every field, produces a complete **derivation path** showing:
- Column definitions from workbook XML (explicit vs default)
- Range position (Σ column widths, Σ row heights)
- Page setup (margins, paper size, printable area)
- PDF content analysis (effW, effH, centering)
- Scale computation (effW/PAW, effH/PAH)
- Origin computation (margin ± centering)
- Page position (origin + range × scale)
- Pixel conversion (page × DPI/72)
- Ratio computation

Each step shows the formula, intermediate values, and data source. Output is `geometry_trace.md`.

### `debug_visualizer.py`
Enhanced debug image generator producing five visualizations:
| Image | Description |
|-------|-------------|
| `debug_overlay.png` | Blue field rectangles on rendered PDF |
| `debug_grid.png` | Green cell boundary grid with address labels on PNG |
| `debug_pdf_vectors.png` | Colored highlights: green=text, blue=drawings, red=images |
| `debug_field_centers.png` | Red crosshair markers at each field center |
| `debug_errors.png` | Error heatmap: green(<2pt), yellow(<8pt), orange(<20pt), red(≥20pt) |

### `confidence_scorer.py`
Diagnostic confidence scoring based on:
- **Workbook geometry** (35%): checks for explicit col widths, reasonable margins, print area existence, row height validity
- **PDF geometry** (35%): validates effective dimensions, content-to-page ratio, centering detection
- **PNG geometry** (30%): cross-validates PNG vs PDF content position per field
- Overall confidence is weighted average, with penalties for excessive warnings or unusual scale factors

### `error_diagnostics.py`
Root cause analysis for misalignment. Automatically identifies:
- Workbook geometry mismatch (missing explicit column widths)
- Margin mismatch (margins consume >50% of page)
- Scaling mismatch (scale factor deviates from 1.0)
- Centering mismatch (gap asymmetry >10pt)
- Merge range issues (merged cells with no content)
- Print Area issues (no configured print area)

Generates `diagnostic_report.md` with severity classification and recommendations.

### `template_suite.py`
Automated template validation runner:
- Discovers templates from directory (looks for .xlsx + matching .pdf/.png)
- Runs each template through the full pipeline
- Measures execution time per template
- Performs regression comparison against previous runs (SHA-256 hash comparison)
- Generates `validation_report.md` (pass/fail, field counts, page sizes)
- Generates `performance_report.md` (timing breakdown per template)
- Exit code indicates success/failure for CI integration

### `workbook_reader.py` (recreated)
The XLSX reader module was missing from the file system. Recreated with:
- ZIP archive parsing (xml.etree.ElementTree, no openpyxl dependency)
- Column width extraction from `<cols>` elements (with `char_to_pt` conversion)
- Row height extraction from `<row>` elements
- Merge range parsing from `<merges>` elements
- Page dimensions from `<pageSetup>` (paper size → points mapping)
- Margin parsing from `<pageMargins>` (inches → points)
- Default column width / row height from `<sheetFormatPr>`
- Sheet name resolution via `xl/workbook.xml` + `xl/_rels/workbook.xml.rels`

---

## Enhanced Existing Modules

### `main.py`
Now includes:
- **`--validate` flag**: enables cross-validation (slower, requires PDF+PNG)
- **`--trace` flag**: generates geometry trace report
- **`--diagnose` flag**: runs error diagnostics
- Performance timing for every pipeline stage
- SHA-256 deterministic hash output
- Integrated confidence scoring (stored as `_validation` in runtime.json, not in schema)
- Enhanced debug visualizations via `debug_visualizer.py`
- Error diagnostic report even on failure

### `validator.py`
Enhanced with:
- `validate_coordinate_consistency()` — overlap detection for non-merged fields
- `compute_deterministic_hash()` — SHA-256 for any runtime.json
- Better error messages in field validation

---

## Pipeline Stages

| Stage | Module | Description |
|-------|--------|-------------|
| 1 | `workbook_reader` | Parse XLSX → WorkbookInfo |
| 2 | `pdf_reader` | Measure PDF content bounding box → ContentBounds |
| 3 | `geometry_engine` | Compute layout: origin, scale, effective dims → LayoutInfo |
| 4 | `field_detector` | Detect fields from sheet XML → FieldDef list |
| 5 | `geometry_engine` | Per-field geometry → FieldGeometry list |
| 6 | `validator` | Basic field validation (negative coords, ratios) |
| 7 | `cross_validator` (--validate) | Multi-source comparison + accuracy stats |
| 8 | `runtime_generator` | Build runtime.json (with internal validation metadata) |
| 9 | `debug_visualizer` | 5 debug images + coordinate_dump.json + debug_report.md |
| 10 | `geometry_tracer` (--trace) | Per-field derivation trace → geometry_trace.md |
| 11 | `error_diagnostics` (--diagnose) | Root cause analysis → diagnostic_report.md |

---

## CLI Reference

```bash
python CoordinateEngine/main.py \
  --excel workbook.xlsx \
  --pdf page.pdf \
  --png page.png \
  --output runtime.json \
  --dpi 300 \
  --debug-dir ./debug \
  --validate \
  --trace \
  --diagnose
```

### Flags
| Flag | Purpose |
|------|---------|
| `--excel` | Workbook path (required) |
| `--pdf` | PDF path (required) |
| `--png` | PNG path (required) |
| `--output` | Output runtime.json path (required) |
| `--dpi` | Rendering DPI (default: 300) |
| `--debug-dir` | Debug output directory |
| `--validate` | Enable cross-validation (multi-source comparison + accuracy stats) |
| `--trace` | Generate geometry_trace.md per-field derivation report |
| `--diagnose` | Generate diagnostic_report.md root cause analysis |

---

## Debug Output

Each run with `--debug-dir` produces:

| Artifact | Description |
|----------|-------------|
| `debug_overlay.png` | Blue field rectangles on rendered PDF |
| `debug_grid.png` | Green grid lines at field boundaries with cell labels |
| `debug_pdf_vectors.png` | PDF elements highlighted (text=d green, drawings=blue, images=red) |
| `debug_field_centers.png` | Red crosshair markers at each field center |
| `debug_errors.png` | Color-coded error heatmap with PDF content bounds |
| `coordinate_dump.json` | Full per-field geometry (worksheet, page, pixel, ratio) |
| `debug_report.md` | Summary report with page setup, layout, field list |
| `geometry_trace.md` | (--trace) Complete derivation path per field |
| `diagnostic_report.md` | (--diagnose) Root cause analysis with recommendations |

---

## Template Suite

```bash
python -m CoordinateEngine.template_suite \
  --templates-dir ./templates \
  --output-dir ./results \
  --previous-dir ./previous_results \
  --dpi 300
```

Outputs:
- `validation_report.md` — pass/fail, field counts, page sizes, regression status
- `performance_report.md` — timing breakdown per template
- Per-template subdirectories with runtime.json and debug artifacts

---

## Deterministic Verification

Every run outputs the SHA-256 hash of the generated `runtime.json`. This enables automated regression detection:

```text
SHA-256: a1b2c3d4e5f6...
```

The `template_suite.py` automatically compares hashes against previous runs and flags any regression.

---

## Deliverables Checklist

| Deliverable | Status |
|-------------|--------|
| `runtime.json` | ✅ (existing, unchanged schema) |
| `coordinate_dump.json` | ✅ (existing) |
| `debug_overlay.png` | ✅ (existing, enhanced) |
| `debug_grid.png` | ✅ (new) |
| `debug_pdf_vectors.png` | ✅ (new) |
| `debug_field_centers.png` | ✅ (new) |
| `debug_errors.png` | ✅ (new) |
| `geometry_trace.md` | ✅ (new with `--trace`) |
| `validation_report.md` | ✅ (via template_suite.py) |
| `performance_report.md` | ✅ (via template_suite.py) |
| `diagnostic_report.md` | ✅ (new with `--diagnose`) |

---

## Success Criteria

| Criteria | Status |
|----------|--------|
| Average error ≤ 1px | ⬜ Engine output accuracy depends on PDF/PNG quality — cross-validation measures this |
| All templates pass | ⬜ Requires template artifacts + validation suite run |
| Fully traceable coordinates | ✅ geometry_trace.md provides complete derivation per field |
| Deterministic output | ✅ SHA-256 verification on every run |
| No C# coordinate calculations | ✅ Phase 2 completed this |
| Frontend unchanged | ✅ No schema or API changes |
