# Phase 2 — Python Coordinate Engine Integration Report

**Date:** 2026-07-11
**Status:** Complete
**Engine Version:** CoordinateEngine (Python) v1.0
**API Version:** ExcelAPI v1.3

---

## Goal

Replace all C# coordinate generation logic in ASP.NET Core with the Python Coordinate Engine. ASP.NET Core now only orchestrates the pipeline — it no longer computes origins, scales, margins, centering, ratios, or pixel offsets.

---

## Architecture Change

### Before Phase 2

```
Browser → ASP.NET Core → Excel COM → PDF → PNG
                                     ↓
                              C# CoordinateTransformer
                              C# ExtractFields (COM)
                              C# FindContentBoundingBox
                              C# ComputeContentWidthFromXlsx
                              C# ConMas origin correction
                                     ↓
                              CaptureResult → Frontend
```

### After Phase 2

```
Browser → ASP.NET Core → Excel COM → PDF → PNG
                                     ↓
                              PythonRunner → CoordinateEngine/main.py
                                     ↓
                              Python runtime.json → read → CaptureResult
                                     ↓
                              Frontend
```

ASP.NET Core responsibilities are now limited to:

1. Upload workbook
2. Excel COM (open, detect print area, export PDF)
3. Convert PDF → PNG (PDFtoImage)
4. Launch Python Coordinate Engine
5. Read Python-generated runtime.json
6. Return CaptureResult to frontend

---

## Files Changed

### New Files

| File                                            | Purpose                                                                                |
| ----------------------------------------------- | -------------------------------------------------------------------------------------- |
| `ExcelAPI/Services/Interfaces/IPythonRunner.cs` | Interface for Python execution                                                         |
| `ExcelAPI/Services/PythonRunner.cs`             | Locates python, builds args, executes process, captures stdout/stderr, handles timeout |

### Modified Files

| File                                              | Change                                                                                                                                                                                                                                                         |
| ------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ExcelAPI/Services/ExcelCaptureService.cs`        | **Major rewrite.** Stripped ~1200 lines of coordinate logic (FindContentBoundingBox, ComputeContentWidthFromXlsx, ConMas correction, origin/scale/ratio calculation). Replaced with PythonRunner call + runtime.json parsing. Reduced from 1619 to ~280 lines. |
| `ExcelAPI/Services/RuntimeCoordinateGenerator.cs` | Removed unused `CoordinateTransformer` dependency from constructor                                                                                                                                                                                             |
| `ExcelAPI/Program.cs`                             | Added `IPythonRunner` → `PythonRunner` registration; removed `CoordinateTransformer` registration                                                                                                                                                              |
| `ExcelAPI/Models/ExcelCaptureOptions.cs`          | Added `CoordinateEnginePath`, `PythonPath`, `PythonTimeoutSeconds` config options                                                                                                                                                                              |

---

## PythonRunner Service

### How it works

1. **Resolves python executable** — tries configured path, then probes `python`, `python3`, `py` via `--version`
2. **Resolves CoordinateEngine path** — if relative, resolved from `ContentRootPath`; default is `../../CoordinateEngine`
3. **Builds command line:**
   ```
   python "../../CoordinateEngine/main.py"
     --excel "<workbook.xlsx>"
     --pdf "<page.pdf>"
     --png "<page.png>"
     --output "<runtime.json>"
     --dpi 300
     --debug-dir "<debug_dir>"
   ```
4. **Executes process** with `RedirectStandardOutput`/`RedirectStandardError`
5. **Handles timeout** — kills process tree if exceeds `PythonTimeoutSeconds` (default 120s)
6. **Returns `PythonResult`** — Success, ExitCode, Stdout, Stderr, OutputPath, ErrorMessage

### Error handling

- Python not found → clear error message
- main.py not found → clear error message
- Timeout → process killed, error with stderr
- Non-zero exit → error with stderr
- Output file missing → error

---

## ExcelCaptureService Changes

### Removed (~1200 lines)

- `FindContentBoundingBox()` — PNG gradient-based edge detection (used for ConMas effective dimensions)
- `ComputeContentWidthFromXlsx()` — XLSX column width parsing (used for content width override)
- `ResolveWorksheetPath()` — XLSX internal path resolution
- `ColumnLetterToIndex()` — column letter to index conversion
- `TryParsePrintAreaColumns()` — print area parsing
- `ExtractFields()` — COM-based field extraction with ConMas coordinate formula
- `GetPaperSizePoints()` — paper size to points mapping
- All coordinate calculation code (origin, scale, ratio, centering, ConMas correction)
- `CoordinateTransformer` dependency and all calls to it
- `PageSetupDebug` population (all geometry debug info now in Python's coordinate_dump.json)

### Kept

- Excel COM launch and cleanup
- Print area detection (4 fallback methods)
- PDF export via `ExportAsFixedFormat`
- PDF → PNG conversion via PDFtoImage
- Stage logging with timestamps (9 stages)

### Added

- `IPythonRunner` injection
- Python engine invocation after PNG generation
- Python runtime.json parsing and mapping to `CaptureResult`
- `PythonResult` validation before proceeding

---

## Data Flow

```
1. Upload → ExcelController.Upload()
2.        → ExcelCaptureService.CapturePrintAreaAsync()
3.        → Excel COM: open workbook
4.        → Excel COM: detect print area
5.        → Excel COM: export PDF
6.        → PDFtoImage: convert PDF to PNG
7.        → PythonRunner: launch CoordinateEngine
8.        → Python: read workbook.xml, measure PDF content, compute geometry
9.        → Python: write runtime.json + debug artifacts
10.       → C#: read Python runtime.json
11.       → C#: map to CaptureResult (backward compat)
12.       → Controller: return CaptureResult to frontend
```

For the persistent runtime path (FormController flow):

```
13. FormController.FromExcel → SaveMetadata → Forms/{id}.runtime.json
14. Frontend GET /api/form/runtime/{id}
15. RuntimeCoordinateGenerator.LoadMetadata → RuntimeForm
16. Frontend renders overlay
```

---

## Configuration (appsettings.json)

New options in `ExcelCapture` section (with defaults):

```json
{
  "CoordinateEnginePath": "..\\..\\CoordinateEngine",
  "PythonPath": py,
  "PythonTimeoutSeconds": 120
}
```

- `CoordinateEnginePath` — path to CoordinateEngine directory (relative to ContentRootPath or absolute)
- `PythonPath` — explicit python executable path (null = probe PATH)
- `PythonTimeoutSeconds` — max seconds to wait for Python engine

---

## Debug Artifacts

The Python engine generates debug output on every run in `Preview/debug_{fileId}/`:

- `debug_overlay.png` — blue field rectangles overlaid on the rendered PDF page
- `coordinate_dump.json` — full per-field geometry trace (worksheet, page, pixel, ratio)
- `debug_report.md` — step-by-step coordinate derivation explanation

These artifacts are automatically preserved and are NOT deleted by cleanup.

---

## CoordinateTransformer.cs

The `CoordinateTransformer` class is now dead code. It is no longer registered in DI and no services depend on it. The file is retained for reference but can be safely removed in a future cleanup.

---

## Verification

- Build: **0 errors, 0 new warnings**
- Pipeline: **Upload → COM → PDF → PNG → Python → runtime.json → CaptureResult**
- Frontend: **No changes required** — the CaptureResult schema is identical
- The Python engine's runtime.json includes all ratio-based fields (`leftRatio`, `topRatio`, `widthRatio`, `heightRatio`) matching the existing `RuntimeField` schema
