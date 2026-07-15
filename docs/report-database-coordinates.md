# Database Coordinate Investigation — Report

## Summary

Connected to PostgreSQL `irepodb` (localhost:5432, user `postgres`) and reverse-engineered the ConMas coordinate storage for `def_top_id=546` ("FormTest - Copy"). This report confirms how the old system stores cluster coordinates and provides the evidence needed to fix the overlay alignment.

## Key Finding: Coordinates are 0–1 Ratios of the FULL PAGE

The `def_cluster` table stores positions as **decimal strings representing normalized ratios** (0.0 = left/top edge of page, 1.0 = right/bottom edge of page).

| Table | Column | Value (Cluster 0, $A$1:$B$2) |
|-------|--------|------------------------------|
| `def_cluster` | `left_position` | `0.3364706` |
| `def_cluster` | `right_position` | `0.4982353` |
| `def_cluster` | `top_position` | `0.3845454` |
| `def_cluster` | `bottom_position` | `0.4218182` |

### Evidence the Ratios Are PAGE-Relative (not printable-area relative):

1. **`isOriginalWhole=1`** in XML — The background image is the full page with margins.
2. **Min ratios < margin/page across all forms:**
   - "Product Entry vrs 2" (501–523): `min_left=0.019` → 11.6pt from left edge, less than standard 54pt margin
   - "Machine Production Report" (525–526): `min_top=0.003` → 2.4pt from top edge, essentially at page top
   - These values would be IMPOSSIBLE if ratios were printable-area-relative (min would be ~54/612=0.088)

## Data Model: def_top_id=546

```
def_top (id=546, name="FormTest - Copy")
├── def_top_option (546) — empty row
├── def_top_size (546) — empty row
├── def_sheet (id=1706, no=1, name="Sheet1")
│   ├── def_cluster (sheet_id=1706)
│   │   ├── id=0 left=0.3365 top=0.3845 right=0.4982 bottom=0.4218 cell=$A$1:$B$2  (sheets=KeyboardText, name=samples)
│   │   ├── id=1 left=0.5000 top=0.3845 right=0.6635 bottom=0.4218 cell=$C$1:$D$2
│   │   ├── id=2 left=0.3365 top=0.4232 right=0.6635 bottom=0.4605 cell=$A$3:$D$4
│   │   ├── id=3 left=0.3365 top=0.4809 right=0.6635 bottom=0.5182 cell=$A$6:$D$7
│   │   ├── id=4 left=0.3365 top=0.5386 right=0.6635 bottom=0.5759 cell=$A$9:$D$10
│   │   └── id=5 left=0.3365 top=0.5964 right=0.4165 bottom=0.6150 cell=$A$12
│   └── background_image_file — NOT set (0 bytes)
│   └── background_image — empty string
├── background_image_file — PNG, 5953 bytes (small thumbnail only)
└── def_file — Full .xlsx workbook (BYTEA)
```

## XML Coordinate Format

The `xml_data` column stores the full form definition in ConMas XML format. Cluster positions appear as `<left>`, `<right>`, `<top>`, `<bottom>` elements with the same 0–1 ratio values. Example:

```xml
<cluster>
    <clusterId>1</clusterId>
    <name>samples</name>
    <type>KeyboardText</type>
    <cellAddress>$A$1:$B$2</cellAddress>
    <left>0.3364706</left>
    <right>0.4982353</right>
    <top>0.3845454</top>
    <bottom>0.4218182</bottom>
</cluster>
```

## All Tables Involved for def_top_id=546

| Table | Purpose | Has def_top_id data? |
|-------|---------|---------------------|
| `def_top` | Form definition (header) | Yes, id=546 |
| `def_sheet` | Sheets in form | Yes, id=1706 |
| `def_cluster` | Clusters/fields with 0-1 position ratios | Yes, 6 rows |
| `def_current` | Points to active definition version | Yes → 546 |
| `def_top_option` | PDF font deduction setting | Empty |
| `def_top_size` | File sizes for def | Empty |
| `data_output_layout` | CSV/XML output layout config | No data |
| `rep_top` | Report instances (filled data) | None for this form |
| `rep_cluster` | Filled cluster data | None |
| `rep_sheet` | Report instance sheets | None |
| `history_rep_cluster` | Audit history for clusters | None for this form |
| `mst_common` | System constants (PAGE_SIZE settings only) | N/A |

## Coordinate Conversion Formula (Old System)

```
pixelX = ratioX * renderedPageWidthInPixels
pixelY = ratioY * renderedPageHeightInPixels
```

Where:
- `ratioX` = `left_position` or `right_position` (e.g., 0.3364706)
- `ratioY` = `top_position` or `bottom_position` (e.g., 0.3845454)
- `renderedPageWidthInPixels` = PNG width (e.g., 2550px at 300 DPI for Letter)
- `renderedPageHeightInPixels` = PNG height (e.g., 3300px at 300 DPI for Letter)
- Origin: (0,0) = **top-left corner of the physical page** (including margins)

This is proven by `isOriginalWhole=1` and cross-form ratio ranges.

## Implication for the Overlay Offset Bug

The current `ExcelCaptureService` computes positions differently:

```
pixelLeft = printedOriginX + (cellLeftPt - printAreaLeft) * scaleX
```

Where `printedOriginX` = leftMargin (in points, from Excel COM). This approach assumes the origin is at the **margin boundary** (printable area), NOT the page edge.

Since `isOriginalWhole=1` means the background PNG shows the FULL PAGE, the current code's positions would be off by exactly `leftMargin * scale` pixels — which is the constant translation hypothesized in the earlier investigation.

**The fix: If the background PNG shows the full page, the origin should be the PAGE edge (0,0 in image), not the printable area's margin boundary.** The ratios from the old system naturally account for this by being page-relative.

## Additional Notes

- `picOriginalResolution=0` — No DPI override
- `retinaMode=1` — Retina/high-DPI flag (possibly 2x rendering)
- The actual page dimensions come from Excel page setup (PaperSize, Orientation) — NOT stored in the database
- The `def_cluster` values match the `xml_data` values exactly (no additional transformation)
- No other columns across the entire database contain coordinate, scale, offset, or margin information
- The `rep_cluster`, `history_rep_cluster`, `log_def_cluster`, and `log_rep_cluster` tables all use the same left_position/right_position/top_position/bottom_position naming and format
