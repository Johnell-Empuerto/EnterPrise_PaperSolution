# Unknowns — Remaining Gaps After Phase 21

## CRITICAL: Coordinate Transform Gaps

### Gap 1: `GetClusterSize()` Internal Algorithm
**Status: UNKNOWN**

`LibExcelController.Lib.ImageUtility.GetClusterSize(string pdfPath, int page, int clusterCount)` is the core coordinate calculation function. It takes an exported PDF path as input and returns `List<ClusterRect>` with float coordinates.

**What we know from reflection:**
- It takes a PDF path (so it reads PDF page dimensions)
- It takes page number and cluster count
- It returns `List<ClusterRect>` (Top/Bottom/Left/Right as Single)

**What we DON'T know:**
- Does it use PDF page dimensions as normalization divisor?
- Does it render the PDF and measure text/object positions?
- Is there additional math beyond column-width summation?
- How does it handle multi-page content?

**Why this matters:** This single function contains the coordinate transform logic. Without decompiling it (requires .NET Framework 4.x runtime), we cannot determine the exact formula.

### Gap 2: Printed Origin Calculation (originX, originY)
**Status: PARTIALLY KNOWN**

The printed origin represents the offset from page top-left to printed content top-left. This involves:
- Page margins (LeftMargin, RightMargin, TopMargin, BottomMargin)
- Header/Footer margins and sizes
- CenterHorizontally / CenterVertically flags
- Zoom/Scaling factor

**What is missing:**
- Exact formula for originX/orginY calculation
- How Excel's `ExportAsFixedFormat` applies these settings
- Whether header/footer areas affect the printable area
- The interaction between centering and margins

**Evidence:** Template 546 works with originX≈0, originY≈0, but templates with different print areas (547 starts at column B) show different origin offsets.

### Gap 3: Gap Compensation (gapH, gapV)
**Status: UNKNOWN**

Empirical testing shows systematic gaps:
- Template 546: gapV ≈ 0.87 pt, gapH ≈ 2.04 pt
- Template 547: gapV ≈ 4.44 pt + 0.18 pt/row, gapH ≈ 0 pt
- Template 548: gapV ≈ 0 pt, gapH ≈ 2.22 pt

**Possible sources (not confirmed):**
- COM column width vs. actual render width (OpenXML uses different formula)
- Row height rounding in rendering
- Border width compensation
- Font metric differences
- Excel print engine internal adjustments

**Hypothesis:** The gaps come from the difference between COM column widths (what we sum) and the actual rendered widths in PDF output. Excel's COM `Range.Width` may use a different measurement than what `ExportAsFixedFormat` renders.

### Gap 4: Multi-Page Content Handling
**Status: PARTIALLY KNOWN**

Template 547 has content that spans more than one page (content height: 1158pt, page height: 792pt, so ~1.46 pages). The gapV grows with row index (0.18 pt/row), suggesting a page-wrap compensation.

**Missing:**
- How does the algorithm determine page breaks?
- Is there a per-page origin reset?
- How are clusters assigned to pages?
- Does `HasPrintArea` / `ExistsPrintAreaOnePage` check affect the algorithm?

### Gap 5: Header/Footer in Coordinate Calculation
**Status: UNKNOWN**

While templates 546/547/548 have no headers/footers, the Designer supports them:
- `PageSetup.PageHeader` / `PageSetup.PageFooter` properties
- `WorksheetHeaderFooter`, `HeaderFooterOptions` classes
- `DifferentFirstPageHeaderFooter`, `OddAndEvenPagesHeaderFooter`
- `get_Header`, `get_Footer`, `get_CenterHeader`, `get_LeftHeader`, `get_RightHeader`

**Missing:** How header/footer space affects printable area and coordinate normalization.

### Gap 6: Horizontal Alignment Compensation
**Status: UNKNOWN**

The `get_HorizontalAlignment` and `get_VerticalAlignment` methods exist in LibExcelController but we don't know how they affect coordinate calculation. Excel text alignment can shift the visual position of content within a cell.

## Page Setup Gaps

### Gap 7: Zoom/Scaling Interaction
**Status: KNOWN (these templates use 100%)**

All 3 templates use Zoom=100%, FitToPagesWide=1, FitToPagesTall=1. But the code detects page setup properties. How would different zoom settings affect coordinates?

### Gap 8: Page Size Detection
**Status: KNOWN**

The `GetClusterSize` method takes a PDF path. The PDF metadata contains the actual page dimensions. This is likely the authoritative page size source.

## Background Generation Gaps

### Gap 9: `CreateImageFromPdf` Internal Algorithm
**Status: UNKNOWN**

`ExcelControllerBase.CreateImageFromPdf(string pdfPath, int pageNumber, bool applyMorphology)` returns a `System.Drawing.Image`.

**Questions:**
- What DPI does it use for rendering? (DpiPdfToImage default?)
- What does `applyMorphology` do? (Image cleanup?)
- What image format is returned? (Bitmap? PNG?)
- How is the image scaled to canvas size?

### Gap 10: Morphology Operation
**Status: UNKNOWN**

The `applyMorphology` parameter in `CreateImageFromPdf` suggests image cleanup operations. `LibConMas.Imaging` has operations like:
- `CorrectDistortion`
- `CannyOperation`
- `MedianOperation`
- `ThresholdOperation`
- `SharpenOperation`
- `GrayscaleOperation`

These are likely used for barcode/QR/EdgeOCR processing, not background rendering. But this is unconfirmed.

## Designer UI Gaps

### Gap 11: Yellow Overlay Specifics
**Status: PARTIALLY KNOWN**

We know overlay colors exist (`__StandardClusterColor` = yellow, `__SelectedClusterColor`) and CanvasChild+Rect renders them. But:

- Exact opacity/alpha value of the yellow overlay
- Whether the overlay is semi-transparent fill or hashed pattern
- Exact border thickness and style
- How selection highlighting differs from default

### Gap 12: Grid Snapping Algorithm
**Status: UNKNOWN**

The `__AutoAlignmentThreshold` config key exists but the grid snap logic is not documented.
- Is it a fixed grid or cell-aligned?
- What is the default grid size?
- How does threshold affect snap behavior?

### Gap 13: Context Menu and Editing Behavior
**Status: UNKNOWN**

The Designer has extensive context menus and editing capabilities:
- `canvas_ContextMenuOpening` event handler
- Multiple xamDataGrid selection changed handlers
- Drag-and-drop cluster reordering
- Property editing panels

The exact interaction model is not documented.

### Gap 14: DPI Scaling for Overlays
**Status: PARTIALLY KNOWN**

We know `LOGPIXELSX`, `LOGPIXELSY`, `SetProcessDPIAWARE`, and `SnapsToDevicePixels` are used, but:
- How does overlay coordinate scaling interact with DPI?
- Are overlay coordinates in WPF device-independent pixels (DIPs) or physical pixels?
- How does zoom interact with the coordinate-to-pixel conversion?

## Database Gaps

### Gap 15: Additional Table Contents
**Status: PARTIALLY KNOWN**

We know the table structure but not all column values:
- `def_top.org_db_name` / `def_top.org_table_name` — external DB linkage
- `def_cluster.org_column_name` — external column mapping
- `def_cluster.default_value` / `def_cluster.data_type` — default/data type
- `def_grammar` contents (voice input grammar definitions)
- `def_layer` contents (drawing layer definitions)

## Config Gaps

### Gap 16: ApplicationConfig.xml Details
**Status: KNOWN**

The `iReporterExcelAddInCommon.ApplicationConfig.xml` defines all cluster types with platform availability and multilingual labels. The complete content is recovered.

### Gap 17: UserConfig Storage
**Status: KNOWN**

`UserConfig.xml` stores: language, device type, cluster colors, caption priority, AI usage, hints. Saved to `%APPDATA%/CIMTOPS/ConMasDesigner/UserConfig.xml`.

## Summary of Impact

| Gap | Impact | Can Work Around? |
|-----|--------|-----------------|
| G1: GetClusterSize algorithm | **HIGH** — Cannot compute exact coordinates | Need .NET 4.x decompilation |
| G2: Printed origin formula | **HIGH** — Offsets coordinates | Can reverse-engineer from DB |
| G3: Gap compensation | **HIGH** — ~0.1-4.4 pt errors | Might live in GetClusterSize |
| G4: Multi-page handling | **MEDIUM** — Affects only >1 page forms | Rare in practice |
| G5: Header/footer impact | **LOW** — Templates don't use them |
| G6: Alignment compensation | **LOW** — Likely no effect on coords |
| G7-G17: Other | **LOW-MEDIUM** | Documented or low impact |

**Bottom line:** Gaps 1-3 are the critical unknowns. They all converge on one question: what does `GetClusterSize()` actually calculate? This cannot be answered without either:
1. Decompiling `LibExcelController.dll` on .NET Framework 4.x
2. Running `GetClusterSize()` with known inputs and comparing outputs
3. Analyzing the PDF output geometry directly
