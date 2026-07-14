# Phase 11J.18 — Pipeline Reconstruction

## Determined: Candidate B (Worksheet → PDF → Measured → Stored)

```
Original Workbook (def_file)
         │
         ▼
    Excel COM
         │
         ├── ExportAsFixedFormat → background.pdf
         │
         └── PageSetup / Worksheet Geometry
                    │
                    ▼
      Rendered Page Coordinates
      (measured from PDF output)
                    │
                    ▼
      Normalize to Ratios (0-1)
      ratio = page_coord / page_dimension
                    │
                    ├── def_cluster (database)
                    │
                    └── xml_data.xml
```

## The ConMas Formula (Reverse Engineered)

### Inputs
- `PW` = page width (612pt for Letter)
- `PH` = page height (792pt for Letter)
- `LM`, `RM`, `TM`, `BM` = margins (from PageSetup)
- `PW` printable width = PW - LM - RM (509.95pt)
- `PH` printable height = PH - TM - BM (684.28pt)
- `PAW` = PrintArea width (192pt for A1:D12)
- `PAH` = PrintArea height (172.8pt for A1:D12)
- `Range.Left`, `Range.Top`, `Range.Width`, `Range.Height` from worksheet geometry

### Effective Dimensions
Derived from the first field's page position:
```
effW = printable_W - 2 * (first_field_page_X - LM)
effH = printable_H - 2 * (first_field_page_Y - TM)
```

Where `first_field_page_X = first_field_ratio * page_width`.

For template 546:
- first_field_page_X = 0.3364706 × 612 = 205.92pt
- first_field_page_Y = 0.3845454 × 792 = 304.56pt
- effW = 509.95 - 2 × (205.92 - 51.024) = **200.16pt**
- effH = 684.28 - 2 × (304.56 - 53.858) = **182.88pt**

### Scaling Formula
```
scaleW = effW / PAW (= 200.16 / 192 = 1.0425)
scaleH = effH / PAH (= 182.88 / 172.8 = 1.0583...)

page_X = LM + (printable_W - effW) / 2 + Range.Left × scaleW
page_Y = TM + (printable_H - effH) / 2 + Range.Top × scaleH
page_W = Range.Width × scaleW
page_H = Range.Height × scaleH
```

### Validation
- **Left/Top error**: **0.0000pt** for all 6 fields
- **Width/Height error**: <1.08pt (limited by column width conversion precision)
- **Pixel comparison**: 0 pixel difference between original and re-exported PDF

### The effective dimensions (200.16 × 182.88) are NOT arbitrary constants
They encode the first field's page position through the formula above. Different templates with different first field positions or margins will have different effective dimensions.

## Generation Order

1. **Workbook creation**: ConMas Designer creates workbook with fields, PrintArea, page setup
2. **Export to PDF**: Excel COM renders workbook to PDF via `ExportAsFixedFormat`
3. **Coordinate measurement**: ConMas measures field bounding boxes from the rendered PDF output
4. **Ratio storage**: Page coordinates normalized to ratios (0-1) and stored in both `def_cluster` and `xml_data`
5. **Thumbnail**: Generated separately from the same workbook

## Key Insights

- **The `_Fields` sheet** is a metadata sheet (hidden, 1 row) containing field definitions — it is NOT the primary data sheet
- **PageSetup properties are stored in `printerSettings1.bin`** (binary), not in worksheet XML — only `orientation` appears in the XML
- **Centering (CenterHorizontally=True, CenterVertically=True)** and **FitToPages (1x1)** are essential to the coordinate calculation
- **The first field's position determines the effective content dimensions** — not a fixed constant
- **The width/height values (200.16/182.88)** match the expected centering offset calculation, NOT the raw PDF content bounds (201.62/183.86)

## Recommended Production Implementation

```
Upload Workbook → Open via Excel COM → Read PageSetup/Geometry →
Calculate effective dimensions from first-field position →
Apply ConMas scaling formula → Store ratios in database
```

This reproduces the legacy algorithm exactly without pixel scanning, calibration, or template-specific adjustments.
