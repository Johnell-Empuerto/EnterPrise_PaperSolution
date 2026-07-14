# Difference Report

## def_top_id = 546
**Generated:** 2026-07-12 00:38:38 UTC
**Total Differences:** 49

## Differences

| # | Section | Property | Database | Generated | Reason | Suggested Fix |
|---|---------|----------|----------|-----------|--------|---------------|
| 1 | Workbook | SheetCount | 1 | 2 | Expected '1', got '2' | Review extraction logic for this property |
| 2 | Workbook | SheetNames | Sheet1 | _Fields, Sheet1 | Expected 'Sheet1', got '_Fields, Sheet1' | Review extraction logic for this property |
| 3 | PageSetup | Sheet[_Fields].MarginTop | 0.74803 | 0.75 | Expected '0.74803', got '0.75' | Review extraction logic for this property |
| 4 | PageSetup | Sheet[_Fields].MarginBottom | 0.74803 | 0.75 | Expected '0.74803', got '0.75' | Review extraction logic for this property |
| 5 | PageSetup | Sheet[_Fields].MarginLeft | 0.70866 | 0.7 | Expected '0.70866', got '0.7' | Review extraction logic for this property |
| 6 | PageSetup | Sheet[_Fields].MarginRight | 0.70866 | 0.7 | Expected '0.70866', got '0.7' | Review extraction logic for this property |
| 7 | PageSetup | Sheet[Sheet1].Orientation | - | portrait | Expected '', got 'portrait' | Review extraction logic for this property |
| 8 | Styles | StyleCount | 0 | 4 | Expected '0', got '4' | Review extraction logic for this property |
| 9 | Styles | Style[0].FontName | - | Aptos Narrow | Expected '', got 'Aptos Narrow' | Review extraction logic for this property |
| 10 | Styles | Style[0].FontSize | - | 11 | Expected '', got '11' | Review extraction logic for this property |
| 11 | Styles | Style[0].FontColor | - | theme:1 | Expected '', got 'theme:1' | Review extraction logic for this property |
| 12 | Styles | Style[0].FillPattern | - | none | Expected '', got 'none' | Review extraction logic for this property |
| 13 | Styles | Style[1].FontName | - | Aptos Narrow | Expected '', got 'Aptos Narrow' | Review extraction logic for this property |
| 14 | Styles | Style[1].FontSize | - | 11 | Expected '', got '11' | Review extraction logic for this property |
| 15 | Styles | Style[1].FontColor | - | theme:1 | Expected '', got 'theme:1' | Review extraction logic for this property |
| 16 | Styles | Style[1].FillPattern | - | none | Expected '', got 'none' | Review extraction logic for this property |
| 17 | Styles | Style[2].FontName | - | Aptos Narrow | Expected '', got 'Aptos Narrow' | Review extraction logic for this property |
| 18 | Styles | Style[2].FontSize | - | 11 | Expected '', got '11' | Review extraction logic for this property |
| 19 | Styles | Style[2].FontColor | - | theme:1 | Expected '', got 'theme:1' | Review extraction logic for this property |
| 20 | Styles | Style[2].FillPattern | - | none | Expected '', got 'none' | Review extraction logic for this property |
| 21 | Styles | Style[2].BorderLeft | - | thin | Expected '', got 'thin' | Review extraction logic for this property |
| 22 | Styles | Style[2].BorderRight | - | thin | Expected '', got 'thin' | Review extraction logic for this property |
| 23 | Styles | Style[2].BorderTop | - | thin | Expected '', got 'thin' | Review extraction logic for this property |
| 24 | Styles | Style[2].BorderBottom | - | thin | Expected '', got 'thin' | Review extraction logic for this property |
| 25 | Styles | Style[2].AlignH | - | center | Expected '', got 'center' | Review extraction logic for this property |
| 26 | Styles | Style[3].FontName | - | Aptos Narrow | Expected '', got 'Aptos Narrow' | Review extraction logic for this property |
| 27 | Styles | Style[3].FontSize | - | 11 | Expected '', got '11' | Review extraction logic for this property |
| 28 | Styles | Style[3].FontColor | - | theme:1 | Expected '', got 'theme:1' | Review extraction logic for this property |
| 29 | Styles | Style[3].FillPattern | - | none | Expected '', got 'none' | Review extraction logic for this property |
| 30 | Styles | Style[3].BorderLeft | - | thin | Expected '', got 'thin' | Review extraction logic for this property |
| 31 | Styles | Style[3].BorderRight | - | thin | Expected '', got 'thin' | Review extraction logic for this property |
| 32 | Styles | Style[3].BorderTop | - | thin | Expected '', got 'thin' | Review extraction logic for this property |
| 33 | Styles | Style[3].BorderBottom | - | thin | Expected '', got 'thin' | Review extraction logic for this property |
| 34 | Styles | Style[3].AlignH | - | center | Expected '', got 'center' | Review extraction logic for this property |
| 35 | Cells | Sheet[_Fields].CellCount | 0 | 7 | Expected '0', got '7' | Review extraction logic for this property |
| 36 | Cells | Sheet[Sheet1].CellCount | 0 | 33 | Expected '0', got '33' | Review extraction logic for this property |
| 37 | MergedCells | Sheet[Sheet1].MergedCellCount | 0 | 5 | Expected '0', got '5' | Review extraction logic for this property |
| 38 | MergedCells | Sheet[Sheet1].MergedCell[A1:B2] | - | A1:B2 | Expected '', got 'A1:B2' | Review extraction logic for this property |
| 39 | MergedCells | Sheet[Sheet1].MergedCell[C1:D2] | - | C1:D2 | Expected '', got 'C1:D2' | Review extraction logic for this property |
| 40 | MergedCells | Sheet[Sheet1].MergedCell[A3:D4] | - | A3:D4 | Expected '', got 'A3:D4' | Review extraction logic for this property |
| 41 | MergedCells | Sheet[Sheet1].MergedCell[A6:D7] | - | A6:D7 | Expected '', got 'A6:D7' | Review extraction logic for this property |
| 42 | MergedCells | Sheet[Sheet1].MergedCell[A9:D10] | - | A9:D10 | Expected '', got 'A9:D10' | Review extraction logic for this property |
| 43 | DefCluster | DefClusterCount | 6 | 0 | Expected '6', got '0' | Review extraction logic for this property |
| 44 | DefCluster | Cluster[$A$1:$B$2] | 0.3364706,0.3845454,0.4982353,... | - | Mismatch detected | Review extraction logic for this property |
| 45 | DefCluster | Cluster[$C$1:$D$2] | 0.5,0.3845454,0.6635294,0.4218... | - | Mismatch detected | Review extraction logic for this property |
| 46 | DefCluster | Cluster[$A$3:$D$4] | 0.3364706,0.4231818,0.6635294,... | - | Mismatch detected | Review extraction logic for this property |
| 47 | DefCluster | Cluster[$A$6:$D$7] | 0.3364706,0.4809091,0.6635294,... | - | Mismatch detected | Review extraction logic for this property |
| 48 | DefCluster | Cluster[$A$9:$D$10] | 0.3364706,0.5386364,0.6635294,... | - | Mismatch detected | Review extraction logic for this property |
| 49 | DefCluster | Cluster[$A$12] | 0.3364706,0.5963637,0.4164706,... | - | Mismatch detected | Review extraction logic for this property |
