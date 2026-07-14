# Comparison Report

## def_top_id = 546
**Generated:** 2026-07-12 00:38:38 UTC
**Overall:** INCOMPLETE

## Summary

| Metric | Value |
|--------|-------|
| Total Sections | 8 |
| Passed | 7 |
| Failed | 0 |
| Different | 0 |
| Missing | 1 |
| Extra | 0 |
| Overall Match | 49.5% |

## Section Details

### BackgroundImage

**Status:** PASS
**Match Rate:** 100.0%

| Property | Status | Database | Generated | Reason |
|----------|--------|----------|-----------|--------|

### Cells

**Status:** PASS
**Match Rate:** 100.0%

| Property | Status | Database | Generated | Reason |
|----------|--------|----------|-----------|--------|
| Sheet[_Fields].CellCount | DIFFERENT | 0 | 7 | Expected '0', got '7' |
| Sheet[Sheet1].CellCount | DIFFERENT | 0 | 33 | Expected '0', got '33' |

### DefCluster

**Status:** PASS
**Match Rate:** 0.0%

| Property | Status | Database | Generated | Reason |
|----------|--------|----------|-----------|--------|
| DefClusterCount | DIFFERENT | 6 | 0 | Expected '6', got '0' |
| Cluster[$A$1:$B$2] | MISSING | 0.3364706,0.3845454,0.4982353,0.4218182 | - |  |
| Cluster[$C$1:$D$2] | MISSING | 0.5,0.3845454,0.6635294,0.4218182 | - |  |
| Cluster[$A$3:$D$4] | MISSING | 0.3364706,0.4231818,0.6635294,0.4604546 | - |  |
| Cluster[$A$6:$D$7] | MISSING | 0.3364706,0.4809091,0.6635294,0.5181818 | - |  |
| Cluster[$A$9:$D$10] | MISSING | 0.3364706,0.5386364,0.6635294,0.5759091 | - |  |
| Cluster[$A$12] | MISSING | 0.3364706,0.5963637,0.4164706,0.615 | - |  |

### MergedCells

**Status:** PASS
**Match Rate:** 14.3%

| Property | Status | Database | Generated | Reason |
|----------|--------|----------|-----------|--------|
| Sheet[_Fields].MergedCellCount | PASS | 0 | 0 |  |
| Sheet[Sheet1].MergedCellCount | DIFFERENT | 0 | 5 | Expected '0', got '5' |
| Sheet[Sheet1].MergedCell[A1:B2] | DIFFERENT | - | A1:B2 | Expected '', got 'A1:B2' |
| Sheet[Sheet1].MergedCell[C1:D2] | DIFFERENT | - | C1:D2 | Expected '', got 'C1:D2' |
| Sheet[Sheet1].MergedCell[A3:D4] | DIFFERENT | - | A3:D4 | Expected '', got 'A3:D4' |
| Sheet[Sheet1].MergedCell[A6:D7] | DIFFERENT | - | A6:D7 | Expected '', got 'A6:D7' |
| Sheet[Sheet1].MergedCell[A9:D10] | DIFFERENT | - | A9:D10 | Expected '', got 'A9:D10' |

### PageSetup

**Status:** PASS
**Match Rate:** 80.8%

| Property | Status | Database | Generated | Reason |
|----------|--------|----------|-----------|--------|
| Sheet[_Fields].PaperSize | PASS | - | - |  |
| Sheet[_Fields].Orientation | PASS | - | - |  |
| Sheet[_Fields].FitToWidth | PASS | - | - |  |
| Sheet[_Fields].FitToHeight | PASS | - | - |  |
| Sheet[_Fields].Scale | PASS | - | - |  |
| Sheet[_Fields].MarginTop | DIFFERENT | 0.74803 | 0.75 | Expected '0.74803', got '0.75' |
| Sheet[_Fields].MarginBottom | DIFFERENT | 0.74803 | 0.75 | Expected '0.74803', got '0.75' |
| Sheet[_Fields].MarginLeft | DIFFERENT | 0.70866 | 0.7 | Expected '0.70866', got '0.7' |
| Sheet[_Fields].MarginRight | DIFFERENT | 0.70866 | 0.7 | Expected '0.70866', got '0.7' |
| Sheet[_Fields].DefaultRowHeight | PASS | 14.4 | 14.4 |  |
| Sheet[_Fields].FreezePanes | PASS | - | - |  |
| Sheet[_Fields].Protected | PASS | False | False |  |
| Sheet[_Fields].PageBreakCount | PASS | 0 | 0 |  |
| Sheet[Sheet1].PaperSize | PASS | - | - |  |
| Sheet[Sheet1].Orientation | DIFFERENT | - | portrait | Expected '', got 'portrait' |
| Sheet[Sheet1].FitToWidth | PASS | - | - |  |
| Sheet[Sheet1].FitToHeight | PASS | - | - |  |
| Sheet[Sheet1].Scale | PASS | - | - |  |
| Sheet[Sheet1].MarginTop | PASS | 0.74803 | 0.7480314960629921 |  |
| Sheet[Sheet1].MarginBottom | PASS | 0.74803 | 0.7480314960629921 |  |
| Sheet[Sheet1].MarginLeft | PASS | 0.70866 | 0.7086614173228347 |  |
| Sheet[Sheet1].MarginRight | PASS | 0.70866 | 0.7086614173228347 |  |
| Sheet[Sheet1].DefaultRowHeight | PASS | 14.4 | 14.4 |  |
| Sheet[Sheet1].FreezePanes | PASS | - | - |  |
| Sheet[Sheet1].Protected | PASS | False | False |  |
| Sheet[Sheet1].PageBreakCount | PASS | 0 | 0 |  |

### Styles

**Status:** PASS
**Match Rate:** 50.9%

| Property | Status | Database | Generated | Reason |
|----------|--------|----------|-----------|--------|
| StyleCount | DIFFERENT | 0 | 4 | Expected '0', got '4' |
| Style[0].FontName | DIFFERENT | - | Aptos Narrow | Expected '', got 'Aptos Narrow' |
| Style[0].FontSize | DIFFERENT | - | 11 | Expected '', got '11' |
| Style[0].FontBold | PASS | - | - |  |
| Style[0].FontItalic | PASS | - | - |  |
| Style[0].FontColor | DIFFERENT | - | theme:1 | Expected '', got 'theme:1' |
| Style[0].FillPattern | DIFFERENT | - | none | Expected '', got 'none' |
| Style[0].FillFgColor | PASS | - | - |  |
| Style[0].FillBgColor | PASS | - | - |  |
| Style[0].BorderLeft | PASS | - | - |  |
| Style[0].BorderRight | PASS | - | - |  |
| Style[0].BorderTop | PASS | - | - |  |
| Style[0].BorderBottom | PASS | - | - |  |
| Style[1].FontName | DIFFERENT | - | Aptos Narrow | Expected '', got 'Aptos Narrow' |
| Style[1].FontSize | DIFFERENT | - | 11 | Expected '', got '11' |
| Style[1].FontBold | PASS | - | - |  |
| Style[1].FontItalic | PASS | - | - |  |
| Style[1].FontColor | DIFFERENT | - | theme:1 | Expected '', got 'theme:1' |
| Style[1].FillPattern | DIFFERENT | - | none | Expected '', got 'none' |
| Style[1].FillFgColor | PASS | - | - |  |
| Style[1].FillBgColor | PASS | - | - |  |
| Style[1].BorderLeft | PASS | - | - |  |
| Style[1].BorderRight | PASS | - | - |  |
| Style[1].BorderTop | PASS | - | - |  |
| Style[1].BorderBottom | PASS | - | - |  |
| Style[2].FontName | DIFFERENT | - | Aptos Narrow | Expected '', got 'Aptos Narrow' |
| Style[2].FontSize | DIFFERENT | - | 11 | Expected '', got '11' |
| Style[2].FontBold | PASS | - | - |  |
| Style[2].FontItalic | PASS | - | - |  |
| Style[2].FontColor | DIFFERENT | - | theme:1 | Expected '', got 'theme:1' |
| Style[2].FillPattern | DIFFERENT | - | none | Expected '', got 'none' |
| Style[2].FillFgColor | PASS | - | - |  |
| Style[2].FillBgColor | PASS | - | - |  |
| Style[2].BorderLeft | DIFFERENT | - | thin | Expected '', got 'thin' |
| Style[2].BorderRight | DIFFERENT | - | thin | Expected '', got 'thin' |
| Style[2].BorderTop | DIFFERENT | - | thin | Expected '', got 'thin' |
| Style[2].BorderBottom | DIFFERENT | - | thin | Expected '', got 'thin' |
| Style[2].AlignH | DIFFERENT | - | center | Expected '', got 'center' |
| Style[2].AlignV | PASS | - | - |  |
| Style[2].WrapText | PASS | - | - |  |
| Style[3].FontName | DIFFERENT | - | Aptos Narrow | Expected '', got 'Aptos Narrow' |
| Style[3].FontSize | DIFFERENT | - | 11 | Expected '', got '11' |
| Style[3].FontBold | PASS | - | - |  |
| Style[3].FontItalic | PASS | - | - |  |
| Style[3].FontColor | DIFFERENT | - | theme:1 | Expected '', got 'theme:1' |
| Style[3].FillPattern | DIFFERENT | - | none | Expected '', got 'none' |
| Style[3].FillFgColor | PASS | - | - |  |
| Style[3].FillBgColor | PASS | - | - |  |
| Style[3].BorderLeft | DIFFERENT | - | thin | Expected '', got 'thin' |
| Style[3].BorderRight | DIFFERENT | - | thin | Expected '', got 'thin' |
| Style[3].BorderTop | DIFFERENT | - | thin | Expected '', got 'thin' |
| Style[3].BorderBottom | DIFFERENT | - | thin | Expected '', got 'thin' |
| Style[3].AlignH | DIFFERENT | - | center | Expected '', got 'center' |
| Style[3].AlignV | PASS | - | - |  |
| Style[3].WrapText | PASS | - | - |  |

### Workbook

**Status:** PASS
**Match Rate:** 50.0%

| Property | Status | Database | Generated | Reason |
|----------|--------|----------|-----------|--------|
| SheetCount | DIFFERENT | 1 | 2 | Expected '1', got '2' |
| SheetNames | DIFFERENT | Sheet1 | _Fields, Sheet1 | Expected 'Sheet1', got '_Fields, Sheet1' |
| DesignerVersion | PASS | - | - |  |
| PrintArea | PASS | - | - |  |

### XmlData

**Status:** MISSING
**Match Rate:** 0.0%

| Property | Status | Database | Generated | Reason |
|----------|--------|----------|-----------|--------|

