# Template 547 — Row Coordinate Investigation

## Row Dump

| Row | COM Top(A) | COM Ht(A) | COM Top(I) | COM Ht(I) | RowHeight | OpenXML Ht | CustomHt |
|-----|-----------|----------|-----------|----------|-----------|------------|----------|
| 1 | 0.00 | 13.20 | 0.00 | 13.20 | 13.20 | 98.25 | yes |
| 2 | 13.20 | 13.20 | 13.20 | 13.20 | 13.20 | 30.75 | yes |
| 3 | 26.40 | 13.20 | 26.40 | 13.20 | 13.20 | 16.50 | yes |
| 4 | 39.60 | 13.20 | 39.60 | 13.20 | 13.20 | 16.50 | yes |
| 5 | 52.80 | 13.20 | 52.80 | 13.20 | 13.20 | 16.50 | yes |
| 6 | 66.00 | 13.20 | 66.00 | 13.20 | 13.20 | 28.50 | yes |
| 7 | 79.20 | 13.20 | 79.20 | 13.20 | 13.20 | 28.50 | yes |
| 8 | 92.40 | 13.20 | 92.40 | 13.20 | 13.20 | 28.50 | yes |
| 9 | 105.60 | 13.20 | 105.60 | 13.20 | 13.20 | 28.50 | yes |
| 10 | 118.80 | 13.20 | 118.80 | 13.20 | 13.20 | 28.50 | yes |
| 11 | 132.00 | 13.20 | 132.00 | 13.20 | 13.20 | 28.50 | yes |
| 12 | 145.20 | 13.20 | 145.20 | 13.20 | 13.20 | 28.50 | yes |
| 13 | 158.40 | 13.20 | 158.40 | 13.20 | 13.20 | 15.75 | yes |
| 14 | 171.60 | 13.20 | 171.60 | 13.20 | 13.20 | 25.50 | yes |
| 15 | 184.80 | 13.20 | 184.80 | 13.20 | 13.20 | 35.25 | yes |

## Per-Row Difference Analysis

Comparing DB cluster positions vs COM Top:

| Row | COM Top | COM Top+Origin | DB Top(pts) | DB Top(pts) | Gap(pt) | Gap inc |
|-----|---------|----------------|-------------|-------------|---------|---------|
| 6 | 66.00 | 131.04 | 131.04 | 0.1654546 | +0.0000 | +0.0000 |
| 7 | 79.20 | 144.24 | 148.68 | 0.1877273 | +4.4400 | +4.4400 |
| 8 | 92.40 | 157.44 | 166.68 | 0.2104545 | +9.2400 | +4.7999 |
| 9 | 105.60 | 170.64 | 185.04 | 0.2336364 | +14.4000 | +5.1601 |
| 10 | 118.80 | 183.84 | 203.40 | 0.2568182 | +19.5600 | +5.1600 |

## Row Height vs Gap Analysis

Let's check if the gap increment equals something measurable:

Row 7: Gap=4.4400pt, RowHeight=13.2000pt, Gap/RowHeight=0.336365
Row 8: Gap=9.2400pt, RowHeight=13.2000pt, Gap/RowHeight=0.699997
Row 9: Gap=14.4000pt, RowHeight=13.2000pt, Gap/RowHeight=1.090911
Row 10: Gap=19.5600pt, RowHeight=13.2000pt, Gap/RowHeight=1.481819

Average gap (rows 7-10): 11.9100pt

## Comparison with Known Constants

PerRowExtraPt (from old engine): 0.36 pt
ClusterBaseGapPt (from old engine): 0.72 pt
Expected compounded gap: base + rowIdx * extra = 0.72 + N * 0.36

Row 6 height: 13.2000 pt (COM)
Row 6 height: 13.2000 pt (EntireRow)
If gap increment = defaultRowHeight * factor: 4.7520 pt
PerRowExtraPt / RowHeight = 0.027273
