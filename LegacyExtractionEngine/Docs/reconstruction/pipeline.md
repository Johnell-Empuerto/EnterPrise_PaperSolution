# Legacy Extraction Engine — Pipeline Reconstruction Guide

## Overview

15-stage extraction pipeline that reproduces ConMas-compatible output from a PostgreSQL database and Excel workbook for `def_top_id = 546`.

## Pipeline Stages

| Stage | Component | Description |
|-------|-----------|-------------|
| 1 | DatabaseReader | Connect PostgreSQL, read metadata |
| 2 | DatabaseReader | Load `def_top` form definition |
| 3 | DatabaseReader | Load `def_sheet` sheet definitions |
| 4 | DatabaseReader | Load `def_cluster` cluster definitions |
| 5 | DatabaseReader | Load `def_top_size` size info |
| 6 | DatabaseReader | Load `def_cluster_refer`, `def_label`, `def_top_option` |
| 7 | DatabaseReader | Auto-discover FK relationships via `information_schema` |
| 8 | ExcelExtractionEngine | Open workbook, parse all raw XML parts |
| 9 | ExcelExtractionEngine | Extract sheets, cells, styles, merges, images, comments |
| 10 | ArtifactGenerator | Write `workbook.json`, `page.json`, `styles.json` |
| 11 | ArtifactGenerator | Write `cells.json`, `merged.json`, `xml_data.xml`, `def_cluster.json` |
| 12 | ComparisonEngine | Compare DB vs generated per section |
| 13 | DeepDiffEngine | Recursive deep diff of JSON dumps |
| 14 | ComparisonEngine | Write `comparison_report.md`, `difference_report.md` |
| 15 | Pipeline | Print summary with match percentages |

## Artifacts

- `workbook.json` — workbook properties + sheets metadata
- `page.json` — page setup, margins, paper size
- `styles.json` — fonts, fills, borders, alignment, number formats
- `cells.json` — cell references, values, styles
- `merged.json` — merged cell ranges
- `xml_data.xml` — ConMas-compatible XML with cluster positions
- `def_cluster.json` — computed cluster positions from sheet geometry
- `background_image.png` — extracted sheet background
- `database_dump.json` — full PostgreSQL dump
- `generated_dump.json` — full generated output dump
- `comparison_report.md` — section-by-section comparison
- `difference_report.md` — all mismatches
- `deep_diff.json` — recursive structural diff

## Key FK Discovery

```sql
SELECT
    kcu.column_name,
    ccu.table_name AS foreign_table_name,
    ccu.column_name AS foreign_column_name
FROM information_schema.table_constraints tc
JOIN information_schema.key_column_usage kcu ON tc.constraint_name = kcu.constraint_name
JOIN information_schema.constraint_column_usage ccu ON tc.constraint_name = ccu.constraint_name
WHERE tc.constraint_type = 'FOREIGN KEY' AND tc.table_name = @tableName;
```

## Configuration

```json
{
  "Database": "Host=localhost;Port=5432;Database=irepodb;Username=postgres;Password=cimtops",
  "def_top_id": 546,
  "Workbook": "C:\\Users\\MCF-JOHNELLEEMPUERTO\\Documents\\FormTest - Copy.xlsx",
  "Output": "LegacyExtractionEngine/Output/FormTest/"
}
```
