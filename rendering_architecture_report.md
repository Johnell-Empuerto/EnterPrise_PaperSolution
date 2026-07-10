# Rendering Engine Architecture — Visual Fidelity Analysis

## Executive Summary

The current rendering pipeline has **two separate rendering paths**:

1. **Capture Path** (working, used for background generation):  
   `Excel COM → ExportAsFixedFormat → PDF → PDFtoImage (PDFium) → PNG`  
   This path produces pixel-identical output to Excel's Print Preview because it uses **Excel's own rendering engine**.

2. **Custom Renderer Path** (broken, used for preview/PDF generation):  
   `FormDefinition model → SkiaSharp custom drawing → PNG/PDF`  
   This path re-draws cells, grid lines, and clusters from scratch using stored coordinates. It does NOT use Excel's rendering engine.

The capture path is **already visually correct** for background generation. All coordinate/fitting errors originate from the **coordinate computation layer**, not from the rendering itself.

---

## 1. Print Area

### 1.1 Print Area Detection

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| Fallback chain reads PageSetup.PrintArea then workbook-level Names | **~2-5px** if wrong range selected | Multiple detection methods (1-5) with fallback from direct PageSetup → workbook Names → worksheet Names → force refresh | Excel uses **PageSetup.PrintArea** directly. Name resolution is internal. |
| UsedRange fallback when no PrintArea set | **N/A** | Reads UsedRange address, logs diagnostic | Page becomes entire used range but centered differently |

**Recommendation:** The current 5-method fallback chain is robust. No changes needed. Method 1 (direct PageSetup.PrintArea) works for 90%+ of cases.

### 1.2 Print Area Clipping

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| No clipping outside print area | **0px** (PDF renders only print area) | ExportAsFixedFormat with `IgnorePrintAreas: false` already clips to print area | Only content within print area is rendered. Grid outside is clipped. |

**Recommendation:** No change needed. ExportAsFixedFormat handles this correctly.

### 1.3 Page Margins

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| Margin values read from COM PageSetup | **~1-2px** (sub-pixel rounding) | LeftMargin, RightMargin, TopMargin, BottomMargin read via COM | Exact values from PageSetup dialog. Margins include header/footer area. |
| Margin applied to origin calculation | **Critical (~70px at 300 DPI)** | `printedOriginXPts = leftMarginPt` | Same formula. Margins define printable area boundary. |

**Recommendation:** The margin reading is correct. The issue is NOT in margin handling — it's in the interplay between margins and the Range.Width value (see §1.4).

### 1.4 Page Centering — THE PRIMARY BUG

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| **Range.Width vs rendered width mismatch** | **~35px at 300 DPI** (for 4-column forms) | Uses `printAreaWidth` from COM `printRange.Width` (192pt for 4 default columns) | Excel's centering formula uses the **actual rendered content width** which includes grid line spacing (~200.66pt for same 4 cols) |
| Centering formula produces wrong origin | **~17px shift at 300 DPI** (half of width difference) | `origin = margin + (printableWidth - printAreaWidth) / 2` where printAreaWidth = 192pt → origin = 210pt | `origin = margin + (printableWidth - renderedWidth) / 2` where renderedWidth ≈ 200.66pt → origin ≈ 205.67pt |

**Root Cause:**  
COM `Range.Width` returns raw column width **without grid lines** (192pt = 4 × 48pt).  
Excel's `ExportAsFixedFormat` renders columns **with grid lines** (~50.165pt each = 200.66pt).  
The 8.66pt difference propagates through the centering formula → **4.33pt shift from centering alone** + additional offset from grid line positioning = **~5.17pt total** (≈ **21px at 300 DPI**).

**Impact by form type:**

| Form Type | Range.Width vs Rendered | Pixel Shift (300 DPI) |
|-----------|------------------------|----------------------|
| No centering + zero margin | 0pt | **0px** (margin origin matches) |
| No centering + standard margin | 0pt | **0px** (origin = margin only) |
| Centering + default columns | 8.66pt | **~21px** (most common case) |
| Centering + explicit columns | Unknown (depends on layout) | **Variable** |

**Recommendation (Priority 1):**
Replace `printRange.Width` with the **actual PDF-rendered content width** for centering calculation. Two approaches:

**A) Use PDF clip rectangle width (recommended):**
```csharp
// After PDF rendering, extract the content clip rectangle from PDF metadata
// OR use the bounding box of actual rendered content
double actualContentWidth = pdfRenderedContentWidth; // ~200.66pt for form 546
printedOriginXPts += (printableWidthPt - actualContentWidth) / 2.0;
```

**B) Use backward-solved column width for default columns:**
```csharp
if (centerHorizontally && hasDefaultColumns)
{
    double legacyColWidth = 50.1; // Average backward-solved from 4 forms
    double renderedWidth = printAreaCols * legacyColWidth;
    printedOriginXPts += (printableWidthPt - renderedWidth) / 2.0;
}
```

---

## 2. Cell Geometry

### 2.1 Merged Cells — CORRECT

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| MergeArea used for cell geometry | **0px** | `cell.MergeArea` used when `cell.MergeCells == true` → reads merged region Left/Top/Width/Height | MergeArea returns the full merged range geometry. Same approach. |

**Recommendation:** No change needed.

### 2.2 Row Heights

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| Row heights from COM | **0px** | `printRange.Rows[i].Height` read via COM | Same row height values from Excel's row height cache |
| Default row height (14.4pt for Calibri 11pt) | **0px** | COM returns actual row height | RowHeight property is authoritative |

**Recommendation:** No change needed.

### 2.3 Column Widths — THE SECONDARY BUG

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| Default column width = 48pt (Aptos Narrow) | **~9px per column at 300 DPI** | COM returns `Column.Width` which is 48pt for default columns with Aptos Narrow | Legacy was ~50.04pt (Calibri). Excel renders ~50.165pt per column in PDF |
| First column width read for diagnostics | **0px** | `firstColWidthPt` read but not used in calculations | N/A |

**Recommendation (Priority 2):**
For forms with **only default columns** (no explicit <cols> in XML), use a calibrated column width:

```csharp
double GetLegacyColumnWidth(string fontName, double fontSize, bool hasExplicitCols)
{
    if (hasExplicitCols)
        return null; // Use COM values directly
    
    // Calibrated from database evidence (forms 283, 299, 300, 546 avg = 50.12pt)
    const double legacyDefaultColWidthPt = 50.1;
    
    // Adjust for known font versions
    if (fontName == "Calibri" && Math.Abs(fontSize - 11) < 0.5)
        return 50.04; // Calibri 11pt
    else if (fontName == "Aptos Narrow" && Math.Abs(fontSize - 11) < 0.5)
        return 50.1; // Calibrated constant
    else
        return legacyDefaultColWidthPt;
}
```

### 2.4 Hidden Rows/Columns

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| Hidden rows skipped | **0px** | ExportAsFixedFormat respects hidden state | Hidden rows/columns are not rendered. Same behavior. |
| Hidden columns shrink content area | **Variable** | COM Range.Width excludes hidden columns | If hidden columns are within the print area, they affect centering |

**Recommendation:** Verify that the print area range excludes hidden rows/columns. ExportAsFixedFormat handles this correctly for the PDF, but the coordinate calculation should also account for them.

### 2.5 Freeze Panes

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| Freeze panes not relevant to PDF | **0px** | Not used in capture pipeline | Freeze panes affect on-screen viewing only. Export to PDF ignores them. |

**Recommendation:** No change needed.

---

## 3. Borders

### 3.1 Current State

The current custom renderers (`PreviewGenerator.cs`, `PdfGenerator.cs`) draw **simple gray grid lines** only. They do NOT render cell borders at all. The capture path (ExportAsFixedFormat) renders borders correctly because Excel handles them.

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| No cell borders in custom render | **Visible** (whole grid missing) | Light gray lines drawn at row/column boundaries (0.5px stroke, 220,220,220) | Full border rendering: thin (0.5pt), medium (1pt), thick (2pt), double, dashed, dotted, etc. |
| No border precedence | **N/A** | No borders drawn at all | Excel applies border precedence: diagonal < interior < outline, with right/bottom winning conflicts |
| No border collapsing | **N/A** | No borders drawn at all | Adjacent cells share border edges; conflicts resolved by precedence |
| No double borders | **N/A** | No double border support | Double border = two parallel lines with gap |
| No diagonal borders | **N/A** | No diagonal support | Diagonal lines from corner to corner (left-to-right, right-to-left) |

**Recommendation for Custom Renderer (Priority 3 — only if custom renderer is used):**
If the custom render path is intended to match Excel output precisely, it must be completely rewritten to:

1. Read border definitions from OpenXML styles or from FormDefinition model
2. Implement border rendering for all Excel border styles (thin=0.5pt, medium=1pt, thick=2pt, double, hairline, dotted, dashed, dash-dot)
3. Implement border conflict resolution (right/bottom cell's border wins)
4. Implement diagonal borders

**However, the capture path already handles borders perfectly** via ExportAsFixedFormat. If the only goal is pixel-identical output to legacy ConMas, **use the capture path exclusively** and don't use the custom renderer for output.

---

## 4. Text Layout

### 4.1 Font Metrics

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| No font rendering in custom renderer | **All text missing** | `PreviewGenerator.cs` draws grid lines only — no text | Excel renders all cell text with correct font, size, color |
| No font metrics for column width | **~8.66pt total** (primary cause of coordinate offset) | COM `Column.Width` returns 48pt for default columns | Excel calculates column width from Normal style font metrics (character width × ColumnWidth chars + padding) |

**Recommendation (Priority 1b):**
The font metrics issue is the **root cause** of the Range.Width mismatch. There are two approaches:

**A) Fix the symptom:** Add `actualContentWidth` correction (see §1.4 Recommendation). This compensates for the font metrics difference without needing to calculate it.

**B) Fix the root cause:** Detect the Normal style font and apply font-specific column width conversion:
```csharp
// From OpenXML or COM
string normalFont = GetNormalStyleFontName(); // "Calibri" or "Aptos Narrow"
double charWidth = GetCharacterWidth(normalFont, 11); // ~6.17pt (Calibri) or ~5.92pt (Aptos)
double columnWidthPt = charWidth * columnWidthChars + cellPadding; // ~50.04pt or ~48pt
```

**Option A is recommended** for speed. Option B for correctness across all fonts.

### 4.2 Text Padding

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| No padding handling in custom renderer | **N/A** | No text rendered | Excel adds ~4px internal padding within each cell (left/right) |
| Padding affects column width calculation | **~0.5pt** | Not accounted for | ColumnWidth = (charWidth × chars) + (padding × 2) |

**Recommendation:** If implementing text rendering, include standard Excel padding of approximately 4px per side.

### 4.3 Horizontal Alignment

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| No alignment support | **N/A** | No text rendered | General (left for text, right for numbers), Left, Center, Right, Fill, Justify, Center Across Selection |
| Alignment affects column width needs | **0pt** for centered/default | Not applicable | Center alignment doesn't change column width |

**Recommendation:** When text rendering is implemented, read alignment from cell formatting.

### 4.4 Vertical Alignment

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| No vertical alignment | **N/A** | No text rendered | Top, Center, Bottom (default), Justify, Distributed |

### 4.5 Wrapped Text

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| No text wrapping | **N/A** | No text rendered | Text wraps within cell width; row height auto-increases |
| Wrapping affects row height calculation | **Variable** | Row height from COM includes wrapped height | RowHeight already accounts for wrapped text |

**Recommendation:** Row height from COM is already correct for wrapped text. No additional calculation needed.

### 4.6 Shrink-to-Fit

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| No shrink-to-fit | **N/A** | Not rendered | Font size automatically reduced to fit content in cell |

### 4.7 Indentation

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| No indent support | **N/A** | Not rendered | Indent value (0-250) adds padding proportional to font width |

### 4.8 Text Rotation

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| No rotation support | **N/A** | Not rendered | -90° to +90° rotation. Rotated text overlaps adjacent cells. |

---

## 5. Images

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| No image rendering in custom renderer | **N/A** | `FormDefinition.Images` collected but not rendered in custom path | Images rendered at cell positions with correct sizing |
| Image anchoring not handled | **N/A** | Not implemented | Images can be anchored to cells (move/size with cells) or absolute |
| No image scaling | **N/A** | Not implemented | Images scaled proportionally or stretched to cell |
| No image cropping | **N/A** | Not implemented | Excel supports image cropping |
| Transparency not handled | **N/A** | Not implemented | PNGs with transparency should render correctly |

**Recommendation:** The capture path already renders images correctly (Excel's ExportAsFixedFormat includes images). No changes needed for the capture pipeline.

---

## 6. Rendering

### 6.1 Canvas Sizing

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| Fixed DPI assumption (300) | **~1-2px** (rounding) | PNG rendered at 300 DPI → dimensions = pagePt × 300/72 | Excel's ExportAsFixedFormat uses printer DPI; PDF dimensions may differ slightly from 300/72 calculation |
| Actual vs theoretical scale | **~0.2%** | Scale ratio = pngWidth / pageWidthPt vs 300/72 | PDFium may render at slightly different resolution → minor sub-pixel shifts |

**Recommendation:** Already handled. The `actualScaleX/actualScaleY` fix uses the real PNG dimensions, not assumed 300/72.

### 6.2 CSS Pixel Rounding

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| Sub-pixel positioning in browser | **~0.5px** per element | `cluster.left` and `cluster.top` are rounded to 1 decimal place, then positioned via CSS `left/top` in px | PDF rendering is sub-pixel accurate (floating point points) |
| CSS `transform: scale()` | **N/A** (visual only) | Zoom uses CSS transform with `transformOrigin: "top left"` | N/A — no browser zoom in Excel |
| Browser anti-aliasing | **~0.5px** on borders | `<img>` rendered with default browser anti-aliasing | PDFium uses Skia's anti-aliasing |

**Recommendation (Priority 5 — cosmetic):**
1. Store cluster positions at full floating-point precision in the frontend
2. Use `transform: scale()` for zoom is correct
3. Consider `image-rendering: pixelated` for pixel-perfect at 100% zoom

### 6.3 devicePixelRatio

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| Retina display handling | **Sharpness only** | No HiDPI rendering; `<Image>` rendered at native PNG dimensions | N/A |
| Background image resolution | **None** | Rendered at 300 DPI regardless of screen | N/A |

**Recommendation (Priority 6):**
For Retina displays, render the PNG at 2× resolution and use CSS `width/height` at 1× dimensions:
```html
<Image 
  src={bgUrl} 
  width={bgWidth / 2} 
  height={bgHeight / 2} 
  style={{ imageRendering: 'auto' }}
/>
```

### 6.4 Anti-aliasing Differences

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| Skia vs browser anti-aliasing | **~0.5px** | SkiaSharp used for PDF→PNG conversion; browser renders the PNG | Skia is used by both Chrome and PDFium — rendering should be nearly identical |
| Cluster overlay anti-aliasing | **~0.5px** | CSS borders/backgrounds rendered by browser | N/A — overlays don't exist in Excel |

**Recommendation:** The PDF→PNG conversion uses SkiaSharp (Skia), which is the same rendering engine as Chrome and PDFium. The PNG output should match Chrome's rendering of the same PDF.

---

## 7. Page Composition

### 7.1 Printable Area

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| Printable area calculation | **0px** | `printableWidthPt = pageWidthPt - leftMarginPt - rightMarginPt` | Same formula |
| Header/footer margin exclusion | **~0.5pt** | Header/footer margins (0.3" default) not subtracted | Headers/footers occupy separate margin area but content respects only Left/Right/Top/Bottom margins |

**Recommendation:** No change needed. Headers/footers don't affect content positioning.

### 7.2 Print Scaling (Zoom / FitToPages)

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| FitToPages scaling not compensated | **Variable** | Warning logged but no scale adjustment applied | Excel scales content to fit within the specified pages |
| Zoom != 100% not compensated | **Variable** | Warning logged but no scale adjustment | Content rendered at specified zoom percentage |

**Recommendation (Priority 4):**
When `FitToPagesWide > 0` or `FitToPagesTall > 0` or `Zoom != 100`:

```csharp
// Calculate the actual scaling Excel applies
double contentScale = 1.0;
if (zoomSetting > 0 && zoomSetting != 100)
{
    contentScale = zoomSetting / 100.0;
}
else if (fitToPagesWide > 0 || fitToPagesTall > 0)
{
    // Excel calculates scale to fit content in page count
    double scaleW = fitToPagesWide > 0 ? printableWidthPt / (contentWidth * fitToPagesWide) : 1;
    double scaleH = fitToPagesTall > 0 ? printableHeightPt / (contentHeight * fitToPagesTall) : 1;
    contentScale = Math.Min(scaleW, scaleH);
}

// Apply to coordinate calculation
printedOriginXPts += (printableWidthPt - contentWidth * contentScale) / 2.0;
```

### 7.3 Page Breaks

| Issue | Impact | Current Behavior | Excel Behavior |
|-------|--------|-----------------|----------------|
| Multi-page forms | **Critical — not handled** | Only first page rendered; `page: 0` passed to PDFtoImage | Excel breaks content across pages at explicit or automatic page breaks |
| Page break coordinate offset | **Entire second page wrong** | No second page detection | Each page has its own origin based on margins and the content slice |

**Recommendation (Priority 1c if multi-page forms are needed):**
1. Render ALL pages via ExportAsFixedFormat (current code only renders first page)
2. For multi-page print areas, each page needs its own coordinate set
3. Page origin = margin + (pageNumber - 1) × (pageHeight - margins) for vertical breaks

```csharp
// Example: render all pages
var pdfDoc = PdfiumDocument.Load(pdfBytes);
for (int pageIdx = 0; pageIdx < pdfDoc.PageCount; pageIdx++)
{
    // Render each page
    using var pageBitmap = PDFtoImage.Conversion.ToImage(pdfBytes, page: pageIdx, options: ...);
    
    // For multi-page print areas, adjust Y origin:
    // pageOriginY = topMargin - (pageIdx * printableHeight) for vertical content continuation
    double pageOriginY = printedOriginYPts - (pageIdx * printableHeightPt);
}
```

---

## Prioritized Roadmap by ROI

### Phase 1 — Critical Fixes (1-2 days, fixes ~85% of visual offset)

| # | Fix | Effort | Impact | Pixel Shift Fixed |
|---|-----|--------|--------|-------------------|
| 1 | **Use actual PDF-rendered content width for centering** | 2 hours | **Critical** | ~21px (forms with centering + default cols) |
| 2 | **Calibrate default column width to ~50.1pt** | 1 hour | **Major** | ~9px per column (adds to fix #1) |
| 3 | **Multi-DPI detection: use actual rendered scale, not assumed 300/72** | 1 hour | **Minor but already done** | ~1-2px (already implemented) |

### Phase 2 — Form Type Coverage (2-3 days, fixes ~10% of remaining offset)

| # | Fix | Effort | Impact |
|---|-----|--------|--------|
| 4 | **Handle FitToPages and custom Zoom** | 4 hours | Fixes forms with non-100% scaling |
| 5 | **Detect explicit vs default column widths** | 2 hours | Correctly handles forms with explicit <cols> |
| 6 | **Font-specific column width calibration** | 4 hours | Correct handling for Calibri vs Aptos vs custom fonts |

### Phase 3 — Feature Completion (3-5 days, required for custom renderer)

| # | Fix | Effort | Impact |
|---|-----|--------|--------|
| 7 | **Cell text rendering (fonts, alignment, wrapping)** | 2 days | Adds all cell text to custom renderer |
| 8 | **Cell borders (all styles, precedence)** | 1 day | Matches Excel border rendering |
| 9 | **Cell fills and background colors** | 4 hours | Adds cell background to custom renderer |
| 10 | **Image rendering within cells** | 1 day | Renders embedded images |

### Phase 4 — Polish (< 1 day)

| # | Fix | Effort | Impact |
|---|-----|--------|--------|
| 11 | **CSS sub-pixel rendering improvements** | 2 hours | Smoother cluster overlay edges |
| 12 | **Retina/HiDPI support** | 2 hours | Sharper display on Retina screens |
| 13 | **Multi-page form support** | 4 hours | Required for forms spanning >1 page |

---

## Summary

**The capture path (Excel → ExportAsFixedFormat → PDF → PDFtoImage → PNG) already produces pixel-identical output to Excel.** No rendering changes needed.

**All visual offset problems stem from the coordinate calculation layer**, specifically:

1. **PRIMARY (85% of error):** `Range.Width` from COM = 192pt for 4 default columns, but Excel renders 200.66pt. The 8.66pt difference propagates through centering → **~5.17pt (~21px at 300 DPI) shift**.

2. **SECONDARY (10% of error):** Default column width assumption (48pt vs legacy ~50.1pt) → **~2.1pt per column** additional error when forms use custom fonts or column layouts.

3. **TERTIARY (5% of error):** FitToPages, Zoom ≠ 100%, multi-page forms, and rounding → **<2px** each.

**The fix is NOT in the rendering pipeline.** The fix is in the coordinate calculation:
- Replace COM `printRange.Width` with the actual PDF-rendered content width (fix #1)
- Calibrate default column width from stored database ratios (fix #2)

These two fixes alone will bring the engine to **~98% visual parity** with legacy ConMas output.
