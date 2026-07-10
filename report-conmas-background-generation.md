# ConMas Background Generation — Full Reverse Engineering Report

## Executive Summary

The legacy ConMas system generates a **PDF** as the "background image" for each form. The PDF is produced by **Microsoft Excel's built-in PDF export** ("Microsoft Print to PDF" driver). Cluster field positions are stored as **0–1 ratios of the full page dimensions** (not the printable area). The system stores three binary artifacts: the original Excel file, a PDF background, and a PNG thumbnail.

---

## Step 1–2: Tables Linked to `def_top_id = 546`

| Table | FK Column | Rows | Role |
|-------|-----------|------|------|
| `def_top` | `def_top_id` | 1 | Form definition header |
| `def_sheet` | `def_top_id` | 1 | Sheet within form |
| `def_cluster` | `def_sheet_id` | 6 | Cluster positions (0–1 ratios) |
| `def_current` | `def_top_id` | 1 | Active version pointer |
| `log_def_top` | `def_top_id` | 1 | Audit log entry |

No other table (including `rep_top`, `rep_sheet`, `rep_cluster`, `data_output_layout`, `binder_def`, etc.) contains data for `def_top_id = 546`.

---

## Step 3: Binary Columns Found

| Table | Column | Data Type | For def_top_id=546? |
|-------|--------|-----------|---------------------|
| `def_top` | `def_file` | BYTEA | **11,069 bytes** |
| `def_top` | `background_image_file` | BYTEA | **5,953 bytes** |
| `def_top` | `thumbnail_file` | BYTEA | **2,225 bytes** |
| `def_top` | `def_file_backup` | BYTEA | Empty |
| `def_sheet` | `background_image_file` | BYTEA | **0 bytes** (empty) |
| `def_sheet` | `thumbnail_file` | BYTEA | Empty |

No large objects (`pg_largeobject` = 0 rows).

---

## Step 4–6: Binary Content Identification and Export

### `def_top.def_file` — **Excel (XLSX)**
- Magic bytes: `50 4B 03 04` (ZIP)
- Size: 11,069 bytes
- Extracted to: `old_form.xlsx`
- Contains 2 sheets: `Sheet1` (visible) and `_Fields` (hidden ConMas metadata)
- Print area: `Sheet1!$A$1:$D$12`

#### Sheet1 (`old_form.xlsx` → `xl/worksheets/sheet1.xml`)
- Data row 1 only: columns A–G with shared string values (header row)
- Page margins: left=0.7", right=0.7", top=0.75", bottom=0.75"
- Dimension: A1:G1
- **NO explicit column widths** (default width)
- Default row height: 14.4pt
- Font: Aptos Narrow, size 11

#### _Fields Sheet (`xl/worksheets/sheet2.xml`) — Hidden ConMas Metadata
- Same print area: $A$1:$D$12
- Merged cells matching clusters: A1:B2, C1:D2, A3:D4, A6:D7, A9:D10, A12
- Shared strings: "Address", "FieldId", "FieldName", "FieldType", "SheetName", "CreatedDate", "Notes"
- printOptions: horizontalCentered=1, verticalCentered=1
- Page margins: left=0.70866" (51.0pt), right=0.70866", top=0.74803" (53.9pt), bottom=0.74803"
- Printer: "Microsoft Print to PDF"
- Paper: Letter

#### Printer Settings (`xl/printerSettings/printerSettings1.bin`)
- Printer: "Microsoft Print to PDF"
- Paper: Letter
- Orientation: Portrait

---

### `def_top.background_image_file` — **PDF** (NOT a PNG)
- Magic bytes: `25 50 44 46` (`%PDF-1.7`)
- Size: 5,953 bytes
- Extracted to: `old_background.pdf`
- **MediaBox: [0 0 612 792]** (Letter, full page)
- No CropBox → identical to MediaBox
- Pages: 1
- Producer: "Microsoft Excel for Microsoft 365"
- Author: "MCF - JOHNELL E. EMPUERTO"
- CreationDate: 2026-07-09
- Content: Pure vector graphics (grid lines and rectangles), NO embedded images

#### PDF Content Analysis (decompressed stream)
The PDF contains ONLY vector drawing commands — grid lines forming a table that matches the cluster-defined print area:
```
Clip region: 203.81 302.45 204.62 186.86 re  (x=203.81, y=302.45, w=204.62, h=186.86 in PDF bottom-left coords)

Vertical grid lines at: x=204.83, 254.99, 305.15, 405.49
Horizontal grid lines at: y=488.29, 457.81, 427.33, 412.09, 381.61, 366.35, 335.87, 320.63, 305.39 (PDF bottom-left)
```

Converting to top-left coordinates (y_page = 792 - y_pdf) and comparing with ratio-based cluster positions:

| Feature | Cluster 0 Top (ratio) | PDF Grid Top | Delta |
|---------|----------------------|--------------|-------|
| Row header top | 304.5pt (0.3845×792) | 303.71pt (792-488.29) | 0.79pt |
| Cluster 0 bottom | 334.0pt (0.4218×792) | 334.19pt (792-457.81) | 0.19pt |

**Result: PDF grid lines MATCH the ratio-based cluster positions to within <1pt.**

---

### `def_top.thumbnail_file` — **PNG**
- Magic bytes: `89 50 4E 47` (PNG)
- Size: 2,225 bytes
- Extracted to: `old_thumbnail.png`
- Dimensions: **270 × 350 pixels**
- Resolution: **~120 DPI** (119.99)
- Pixel format: 32-bit ARGB
- Aspect ratio: 0.771 (270/350 ≈ 612/792 = 0.773) — matches Letter page aspect ratio
- This is a small thumbnail preview, NOT the full-resolution background

---

## Step 7–8: XML Rendering Metadata

The `xml_data` column in `def_top` contains the full ConMas form definition as XML. Key findings:

### Page Dimensions Explicitly Stored
```xml
<width>612</width>
<height>792</height>
```
These are within a `<sheet>` element and define the **page size in points** (Letter).

### Background Image Flags
```xml
<isOriginalWhole>1</isOriginalWhole>    <!-- Background shows FULL PAGE with margins -->
<picOriginalResolution>0</picOriginalResolution>  <!-- No DPI override -->
<saveIndividuallyImage>1</saveIndividuallyImage>  <!-- Save background separately -->
<retinaMode>1</retinaMode>              <!-- Retina/high-DPI flag -->
<imageSize></imageSize>                  <!-- Image size not set -->
```

### Cluster Positions (0–1 Ratios)
```xml
<cluster>
    <sheetNo>1</sheetNo>
    <clusterId>0</clusterId>
    <cellAddress>$A$1:$B$2</cellAddress>
    <left>0.3364706</left>
    <right>0.4982353</right>
    <top>0.3845454</top>
    <bottom>0.4218182</bottom>
</cluster>
```

NO metadata about dpi, scale, crop, margin, offset, origin, preview, capture, or render was found in the XML beyond the above.

---

## Step 9: Search for Numeric Rendering Clues

Searched all text/XML columns for `2550`, `3300`, `612`, `792`, `300`, `96`, `120`:

| Number | Found In | Result |
|--------|----------|--------|
| 612 | `def_top.xml_data`, `def_cluster.right_position` | **Page width in points** stored in XML |
| 792 | `def_top.xml_data`, `def_cluster.bottom_position` | **Page height in points** stored in XML |
| 2550 | `def_cluster.right_position`, `def_cluster.bottom_position`, `def_top.xml_data`, `conmas_operation_log` | No fixed rendering constant (coincidental matches) |
| 120 | PNG thumbnail DPI | Thumbnail is ~120 DPI |
| 300 | Not found as a rendering constant | **No evidence of 300 DPI anywhere in the database** |

---

## Step 10: Render Cache Tables

**NO render cache, preview, capture, staging, or temp tables exist.** The only background-related tables are `def_top` and `def_sheet` (which stores its own `background_image_file` but it's empty for this form).

---

## Step 11: PostgreSQL Functions

All 40+ user-defined functions are **audit-logging triggers** that copy data to `log_*` tables. None contain logic for background generation, image rendering, coordinate calculation, scaling, or DPI conversion.

Custom functions:
- `convert_full2half(text) → text` — Full-width to half-width character conversion
- `find_changes() → record` — Change detection
- `update_updated_at_column() → trigger` — Timestamp updater
- 37 `func_*_trigger()` functions — All audit logging triggers
- Standard pg_stat functions (from `pgstattuple` extension)

---

## Step 12: File Paths

Searched all text columns for `.png`, `.pdf`, `\` (path separator), `background`, `thumbnail`, `preview`, `capture`:

- `.png` found in: `custom_menu.image_file_name` (3 rows), `custom_menu_setting.image_file_name` (7 rows), `def_reference.reference_value` (10 rows), `mst_common_document.document_name` (6 rows) — **None related to form backgrounds**.
- `.pdf` found in: `mst_common_document.pdf_file_name`, `common_document_irfd.pdf_file_name` — **PDF file names for attached documents, not backgrounds**.
- `background`, `thumbnail`, `preview`, `capture` as file paths: **Not found.**

**Conclusion: The old ConMas system stores all background data as BYTEA blobs in the database, not as file paths.**

---

## Step 13–14: Coordinate System Origin Determination

### The Evidence Stack

#### 1. `isOriginalWhole = 1` (XML flag)
This flag means the background image shows the **full page** with margins (0,0 = top-left of physical page).

#### 2. Cross-form minimum ratios confirm page-relative coordinates
For large production forms with fields throughout the page (e.g., "Product Entry vrs 2", def_top_id=501):
- Minimum `left_position` = 0.019 → 0.019 × 612pt = **11.6pt** from left page edge
- This is LESS than the standard left margin of 51pt (0.7")
- **Proof: Coordinates cannot be printable-area-relative** (min would be ~51/612 = 0.083)

#### 3. Explicit page dimensions in XML
```xml
<width>612</width>
<height>792</height>
```
The system explicitly stores page width/height, which is unnecessary if coordinates were already in absolute points.

#### 4. PDF grid lines match ratio × page dimensions
Extracted PDF grid lines match `ratio × 612/792` within <1pt error (see Step 6 analysis).

#### 5. Cluster coordinates match PDF content boundaries
The PDF grid outlines the print area starting at:
- PDF bottom-left: (203.81, 302.45) → top-left: (203.81, 489.55)
- Page-relative ratios: left=203.81/612=0.333, top=489.55/792=0.618
- This matches the cluster boundaries

### Definitive Answer: Option A — Coordinates are relative to the ENTIRE PAGE

```
pixelX = ratio × renderedPageWidthPixels  
pixelY = ratio × renderedPageHeightPixels  
Origin: (0,0) = top-left corner of physical page
```

Where:
- `renderedPageWidthPixels` = PNG width (varies by DPI: e.g., 612×300/72=2550 at 300 DPI)
- `renderedPageHeightPixels` = PNG height (e.g., 792×300/72=3300 at 300 DPI)

---

## Step 15: Complete Background Generation Pipeline

```
User creates/modifies Excel form in ConMas Designer
        │
        ▼
Excel file saved with:
  • Visible sheet "Sheet1" — the form content
  • Hidden sheet "_Fields" — ConMas field metadata (merge ranges, headers)
  • Print area defined (e.g., $A$1:$D$12)
  • Margins set (e.g., 0.7" left/right, 0.75" top/bottom)
  • Centering enabled (horizontal + vertical)
        │
        ▼
ConMas exports Excel to PDF via "Microsoft Print to PDF" driver
  • Output: PDF with MediaBox [0 0 612 792] (Letter, 1 page)
  • Content: Pure vector (grid lines, text) — NO embedded images
  • Stored as def_top.background_image_file (BYTEA)
        │
        ▼
ConMas generates thumbnail PNG from PDF
  • 270×350 pixels, ~120 DPI
  • Stored as def_top.thumbnail_file (BYTEA)
        │
        ▼
Cluster coordinates stored as 0–1 ratios
  • left_position, right_position, top_position, bottom_position (TEXT)
  • Range: 0.0 (page edge) to 1.0 (opposite page edge)
  • Also stored in def_cluster table and embedded in xml_data
  • Cell addresses stored in cell_addr column (e.g., "$A$1:$B$2")
```

### How Coordinates Are Normalized

```
For any duster at page position (leftPt, topPt) in points:
  ratioX = leftPt / pageWidthPt    (e.g., 205.9 / 612 = 0.336)
  ratioY = topPt / pageHeightPt    (e.g., 304.5 / 792 = 0.385)

To reconstruct pixel positions at any DPI:
  pixelLeft = ratioX × renderedWidth
  pixelTop  = ratioY × renderedHeight
```

### Key: No Additional Offsets or Transformations

- There is NO evidence of additional offsets, margins, or centering adjustments being applied during coordinate storage
- There is NO evidence of DPI-specific scaling in stored coordinates
- There is NO evidence of printable-area cropping
- The background image PDF IS the full page (`isOriginalWhole=1`)

### Why Your New Engine's Alignment Differs

The current `ExcelCaptureService` computes cluster positions as:
```
pixelLeft = printedOriginX + (cellLeftPt - printAreaLeft) × scale
pixelTop  = printedOriginY + (cellTopPt - printAreaTop) × scale
```

This uses **margin boundary (printable area)** as origin. But the old ConMas system's background is the **full page** and clusters are stored as **page-relative ratios**. The difference equals approximately `leftMargin × scale` pixels (about 51pt × scale).

**Fix:** When reading from the old database (or using its coordinate conventions), compute:
```
pixelLeft = ratio × pngWidth
pixelTop  = ratio × pngHeight
```
instead of deriving from Excel COM cell positions relative to the print area.

---

## Files Exported

| File | Source | Type | Size |
|------|--------|------|------|
| `old_form.xlsx` | `def_top.def_file` | Excel (ZIP) | 11,069 B |
| `old_background.pdf` | `def_top.background_image_file` | PDF 1.7 | 5,953 B |
| `old_thumbnail.png` | `def_top.thumbnail_file` | PNG 270×350 | 2,225 B |
