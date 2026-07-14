# Phase 12 — Coordinate Source Investigation Report

**Generated:** 2026-07-12 11:43:56 UTC
**Templates:** 546, 547, 548

For each cluster, every available Excel coordinate source is dumped and compared against the DB value.
The goal: find a coordinate source that matches the database WITHOUT any correction formula.

## Template 546 — FormTest - Copy
- **Workbook:** `C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx`
- **DB Clusters:** 6

### Page Setup

| Property | Value |
|----------|-------|
| Page Size | 612.0 x 792.0 |
| Margins L,R,T,B | 50.4, 50.4, 54.0, 54.0 |
| Center H,V | False, False |
| Zoom | 100% |
| FitToPages | 1 x 1 |
| Printable Area | 511.2 x 684.0 |
| Printable Origin | (50.40, 54.00) |

### OpenXML Column Widths


### OpenXML Row Heights


### Per-Cluster Coordinate Source Comparison

For each cluster, the following sources are measured and compared to DB:

1. **Range.Left/Top** — our current source (COM Range position)
2. **MergeArea.Left/Top** — if merged, the merge area's origin
3. **MergeArea.Left + Width** — the right edge of merge area
4. **TopLeftCell.Left/Top** — first cell of the range
5. **Range(1).Left/Top** — first item within range
6. **Range.Cells(1).Left/Top** — Cells(1) = first cell of range
7. **Printable coords** — position relative to printable area
8. **Computed from Column widths** — sum of prior column widths
9. **Computed from Row heights** — sum of prior row heights
10. **PointsToScreenPixels** — screen pixel coordinates
11. **PageSetup coords** — position within page setup
12. **Center of range** — Left+Width/2, Top+Height/2

#### Cluster: `$A$1:$B$2`

**DB Values:** Left=0.3364706 (205.92pt)  Top=0.3845454 (304.56pt)  Right=0.4982353  Bottom=0.4218182

**Row:** 1, **Column:** 1

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 0.0000 | 0.0000 | 205.9200 | 304.5600 | ✗ |
| 0 | Cells(1) ($A$1) | 0.0000 | 0.0000 | 205.9200 | 304.5600 | ✗ |
| 0 | Range[1] (default index) | 0.0000 | 0.0000 | 205.9200 | 304.5600 | ✗ |
| 0 | OpenXML Column Width Sum | 0.0000 | N/A | 205.9200 | N/A | - |
| 0 |  | N/A | 0.0000 | N/A | 304.5600 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 0.0000 | 205.9200 | 304.5600 | ✗ |
| 0 | EntireColumn.Left/Top | 0.0000 | 0.0000 | 205.9200 | 304.5600 | ✗ |
| 0 | PointsToScreenPixels (34,149) | 25.5000 | 111.7500 | 180.4200 | 192.8100 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 48.0000 | 14.4000 | 157.9200 | 290.1600 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 0.0000 | N/A | 304.5600 | ✗ |
| 15 | Row# × 14.4pt | N/A | 0.0000 | N/A | 304.5600 | ✗ |
| 16 | Col# × 48pt | 0.0000 | N/A | 205.9200 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 0.0000 | N/A | 304.5600 | ✗ |
| 18 | COM Column Width Sum | 0.0000 | N/A | 205.9200 | N/A | ✗ |
| 19 | Printable Area Coords | 50.4000 | 54.0000 | 155.5200 | 250.5600 | ✗ |
| 20 | DB (expected) | 205.9200 | 304.5600 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin (205.92,304.56) | 205.9200 | 304.5600 | 0.0000 | 0.0000 | ✅ |


#### Cluster: `$C$1:$D$2`

**DB Values:** Left=0.5000000 (306.00pt)  Top=0.3845454 (304.56pt)  Right=0.6635294  Bottom=0.4218182

**Row:** 1, **Column:** 3

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 96.0000 | 0.0000 | 210.0000 | 304.5600 | ✗ |
| 0 | Cells(1) ($C$1) | 96.0000 | 0.0000 | 210.0000 | 304.5600 | ✗ |
| 0 | Range[1] (default index) | 96.0000 | 0.0000 | 210.0000 | 304.5600 | ✗ |
| 0 | OpenXML Column Width Sum | 96.0000 | N/A | 210.0000 | N/A | - |
| 0 |  | N/A | 0.0000 | N/A | 304.5600 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 0.0000 | 306.0000 | 304.5600 | ✗ |
| 0 | EntireColumn.Left/Top | 96.0000 | 0.0000 | 210.0000 | 304.5600 | ✗ |
| 0 | PointsToScreenPixels (130,149) | 97.5000 | 111.7500 | 208.5000 | 192.8100 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 144.0000 | 14.4000 | 162.0000 | 290.1600 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 0.0000 | N/A | 304.5600 | ✗ |
| 15 | Row# × 14.4pt | N/A | 0.0000 | N/A | 304.5600 | ✗ |
| 16 | Col# × 48pt | 96.0000 | N/A | 210.0000 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 0.0000 | N/A | 304.5600 | ✗ |
| 18 | COM Column Width Sum | 96.2810 | N/A | 209.7190 | N/A | ✗ |
| 19 | Printable Area Coords | 146.4000 | 54.0000 | 159.6000 | 250.5600 | ✗ |
| 20 | DB (expected) | 306.0000 | 304.5600 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin(205.92,304.56) | 301.9200 | 304.5600 | 4.0800 | 0.0000 | ✗ |


#### Cluster: `$A$3:$D$4`

**DB Values:** Left=0.3364706 (205.92pt)  Top=0.4231818 (335.16pt)  Right=0.6635294  Bottom=0.4604546

**Row:** 3, **Column:** 1

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 0.0000 | 28.8000 | 205.9200 | 306.3600 | ✗ |
| 0 | Cells(1) ($A$3) | 0.0000 | 28.8000 | 205.9200 | 306.3600 | ✗ |
| 0 | Range[1] (default index) | 0.0000 | 28.8000 | 205.9200 | 306.3600 | ✗ |
| 0 | OpenXML Column Width Sum | 0.0000 | N/A | 205.9200 | N/A | - |
| 0 |  | N/A | 28.8000 | N/A | 306.3600 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 28.8000 | 205.9200 | 306.3600 | ✗ |
| 0 | EntireColumn.Left/Top | 0.0000 | 0.0000 | 205.9200 | 335.1600 | ✗ |
| 0 | PointsToScreenPixels (34,177) | 25.5000 | 132.7500 | 180.4200 | 202.4100 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 96.0000 | 43.2000 | 109.9200 | 291.9600 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 28.8000 | N/A | 306.3600 | ✗ |
| 15 | Row# × 14.4pt | N/A | 28.8000 | N/A | 306.3600 | ✗ |
| 16 | Col# × 48pt | 0.0000 | N/A | 205.9200 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 28.8000 | N/A | 306.3600 | ✗ |
| 18 | COM Column Width Sum | 0.0000 | N/A | 205.9200 | N/A | ✗ |
| 19 | Printable Area Coords | 50.4000 | 82.8000 | 155.5200 | 252.3600 | ✗ |
| 20 | DB (expected) | 205.9200 | 335.1600 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin(205.92,304.56) | 205.9200 | 333.3600 | 0.0000 | 1.8000 | ✗ |


#### Cluster: `$A$6:$D$7`

**DB Values:** Left=0.3364706 (205.92pt)  Top=0.4809091 (380.88pt)  Right=0.6635294  Bottom=0.5181818

**Row:** 6, **Column:** 1

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 0.0000 | 72.0000 | 205.9200 | 308.8800 | ✗ |
| 0 | Cells(1) ($A$6) | 0.0000 | 72.0000 | 205.9200 | 308.8800 | ✗ |
| 0 | Range[1] (default index) | 0.0000 | 72.0000 | 205.9200 | 308.8800 | ✗ |
| 0 | OpenXML Column Width Sum | 0.0000 | N/A | 205.9200 | N/A | - |
| 0 |  | N/A | 72.0000 | N/A | 308.8800 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 72.0000 | 205.9200 | 308.8800 | ✗ |
| 0 | EntireColumn.Left/Top | 0.0000 | 0.0000 | 205.9200 | 380.8800 | ✗ |
| 0 | PointsToScreenPixels (34,221) | 25.5000 | 165.7500 | 180.4200 | 215.1300 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 96.0000 | 86.4000 | 109.9200 | 294.4800 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 72.0000 | N/A | 308.8800 | ✗ |
| 15 | Row# × 14.4pt | N/A | 72.0000 | N/A | 308.8800 | ✗ |
| 16 | Col# × 48pt | 0.0000 | N/A | 205.9200 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 72.0000 | N/A | 308.8800 | ✗ |
| 18 | COM Column Width Sum | 0.0000 | N/A | 205.9200 | N/A | ✗ |
| 19 | Printable Area Coords | 50.4000 | 126.0000 | 155.5200 | 254.8800 | ✗ |
| 20 | DB (expected) | 205.9200 | 380.8800 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin(205.92,304.56) | 205.9200 | 376.5600 | 0.0000 | 4.3200 | ✗ |


#### Cluster: `$A$9:$D$10`

**DB Values:** Left=0.3364706 (205.92pt)  Top=0.5386364 (426.60pt)  Right=0.6635294  Bottom=0.5759091

**Row:** 9, **Column:** 1

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 0.0000 | 115.2000 | 205.9200 | 311.4000 | ✗ |
| 0 | Cells(1) ($A$9) | 0.0000 | 115.2000 | 205.9200 | 311.4000 | ✗ |
| 0 | Range[1] (default index) | 0.0000 | 115.2000 | 205.9200 | 311.4000 | ✗ |
| 0 | OpenXML Column Width Sum | 0.0000 | N/A | 205.9200 | N/A | - |
| 0 |  | N/A | 115.2000 | N/A | 311.4000 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 115.2000 | 205.9200 | 311.4000 | ✗ |
| 0 | EntireColumn.Left/Top | 0.0000 | 0.0000 | 205.9200 | 426.6000 | ✗ |
| 0 | PointsToScreenPixels (34,264) | 25.5000 | 198.0000 | 180.4200 | 228.6000 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 96.0000 | 129.6000 | 109.9200 | 297.0000 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 115.2000 | N/A | 311.4000 | ✗ |
| 15 | Row# × 14.4pt | N/A | 115.2000 | N/A | 311.4000 | ✗ |
| 16 | Col# × 48pt | 0.0000 | N/A | 205.9200 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 115.2000 | N/A | 311.4000 | ✗ |
| 18 | COM Column Width Sum | 0.0000 | N/A | 205.9200 | N/A | ✗ |
| 19 | Printable Area Coords | 50.4000 | 169.2000 | 155.5200 | 257.4000 | ✗ |
| 20 | DB (expected) | 205.9200 | 426.6000 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin(205.92,304.56) | 205.9200 | 419.7600 | 0.0000 | 6.8401 | ✗ |


#### Cluster: `$A$12`

**DB Values:** Left=0.3364706 (205.92pt)  Top=0.5963637 (472.32pt)  Right=0.4164706  Bottom=0.6150000

**Row:** 12, **Column:** 1

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 0.0000 | 158.4000 | 205.9200 | 313.9201 | ✗ |
| 0 | MergeArea (range=$A$12) | 0.0000 | 158.4000 | 205.9200 | 313.9201 | ✗ |
| 0 | Cells(1) ($A$12) | 0.0000 | 158.4000 | 205.9200 | 313.9201 | ✗ |
| 0 | Range[1] (default index) | 0.0000 | 158.4000 | 205.9200 | 313.9201 | ✗ |
| 0 | OpenXML Column Width Sum | 0.0000 | N/A | 205.9200 | N/A | - |
| 0 |  | N/A | 158.4000 | N/A | 313.9201 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 158.4000 | 205.9200 | 313.9201 | ✗ |
| 0 | EntireColumn.Left/Top | 0.0000 | 0.0000 | 205.9200 | 472.3201 | ✗ |
| 0 | PointsToScreenPixels (34,307) | 25.5000 | 230.2500 | 180.4200 | 242.0701 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 24.0000 | 165.6000 | 181.9200 | 306.7201 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 158.4000 | N/A | 313.9201 | ✗ |
| 15 | Row# × 14.4pt | N/A | 158.4000 | N/A | 313.9201 | ✗ |
| 16 | Col# × 48pt | 0.0000 | N/A | 205.9200 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 158.4000 | N/A | 313.9201 | ✗ |
| 18 | COM Column Width Sum | 0.0000 | N/A | 205.9200 | N/A | ✗ |
| 19 | Printable Area Coords | 50.4000 | 212.4000 | 155.5200 | 259.9201 | ✗ |
| 20 | DB (expected) | 205.9200 | 472.3201 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin(205.92,304.56) | 205.9200 | 462.9600 | 0.0000 | 9.3601 | ✗ |


### Full Worksheet Range Dump (first 50 rows × first 20 columns)

For every cell position, dump the COM Left/Top values:

#### Row Heights

| Row | COM RowHeight(pt) | COM Left | COM Top |
|:---:|:-----------------:|:--------:|:-------:|
| 1 | 14.40 | 0.00 | 0.00 |

#### Column Widths

| Col | COM ColumnWidth(chars) | COM Left | COM Top |
|:---:|:----------------------:|:--------:|:-------:|
| 1(A) | 8.110 | 0.00 | 0.00 |
| 2(B) | 8.110 | 48.00 | 0.00 |
| 3(C) | 8.110 | 96.00 | 0.00 |
| 4(D) | 8.110 | 144.00 | 0.00 |
| 5(E) | 8.110 | 192.00 | 0.00 |
| 6(F) | 8.110 | 240.00 | 0.00 |
| 7(G) | 8.110 | 288.00 | 0.00 |

#### Every Cell Position Grid

| Row | Data |
|:---:|------|
| 1 | Ht=14.40 Top=0.00 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 

---

## Template 547 — [V3.1_Sample]アンケート用紙
- **Workbook:** `C:\Users\MCF-JOHNELLEEMPUERTO\Documents\[V3.1_Sample]アンケート用紙.xlsx`
- **DB Clusters:** 5

### Page Setup

| Property | Value |
|----------|-------|
| Page Size | 612.0 x 792.0 |
| Margins L,R,T,B | 50.4, 50.4, 54.0, 54.0 |
| Center H,V | False, False |
| Zoom | 100% |
| FitToPages | 1 x 1 |
| Printable Area | 511.2 x 684.0 |
| Printable Origin | (50.40, 54.00) |

### OpenXML Column Widths


### OpenXML Row Heights


### Per-Cluster Coordinate Source Comparison

For each cluster, the following sources are measured and compared to DB:

1. **Range.Left/Top** — our current source (COM Range position)
2. **MergeArea.Left/Top** — if merged, the merge area's origin
3. **MergeArea.Left + Width** — the right edge of merge area
4. **TopLeftCell.Left/Top** — first cell of the range
5. **Range(1).Left/Top** — first item within range
6. **Range.Cells(1).Left/Top** — Cells(1) = first cell of range
7. **Printable coords** — position relative to printable area
8. **Computed from Column widths** — sum of prior column widths
9. **Computed from Row heights** — sum of prior row heights
10. **PointsToScreenPixels** — screen pixel coordinates
11. **PageSetup coords** — position within page setup
12. **Center of range** — Left+Width/2, Top+Height/2

#### Cluster: `$I$6:$M$6`

**DB Values:** Left=0.5294118 (324.00pt)  Top=0.1654546 (131.04pt)  Right=0.9282353  Bottom=0.1868182

**Row:** 6, **Column:** 9

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 384.0000 | 66.0000 | 60.0000 | 65.0400 | ✗ |
| 0 | Cells(1) ($I$6) | 384.0000 | 66.0000 | 60.0000 | 65.0400 | ✗ |
| 0 | Range[1] (default index) | 384.0000 | 66.0000 | 60.0000 | 65.0400 | ✗ |
| 0 | OpenXML Column Width Sum | 384.0000 | N/A | 60.0000 | N/A | - |
| 0 |  | N/A | 72.0000 | N/A | 59.0400 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 66.0000 | 324.0000 | 65.0400 | ✗ |
| 0 | EntireColumn.Left/Top | 384.0000 | 0.0000 | 60.0000 | 131.0400 | ✗ |
| 0 | PointsToScreenPixels (427,219) | 320.2500 | 164.2500 | 3.7500 | 33.2100 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 504.0000 | 72.6000 | 180.0000 | 58.4400 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 72.0000 | N/A | 59.0400 | ✗ |
| 15 | Row# × 14.4pt | N/A | 72.0000 | N/A | 59.0400 | ✗ |
| 16 | Col# × 48pt | 384.0000 | N/A | 60.0000 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 66.0000 | N/A | 65.0400 | ✗ |
| 18 | COM Column Width Sum | 385.1240 | N/A | 61.1240 | N/A | ✗ |
| 19 | Printable Area Coords | 434.4000 | 120.0000 | 110.4000 | 11.0400 | ✗ |
| 20 | DB (expected) | 324.0000 | 131.0400 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin (-60.00,65.04) | 324.0000 | 131.0400 | 0.0000 | 0.0000 | ✅ |


#### Cluster: `$I$7:$M$7`

**DB Values:** Left=0.5294118 (324.00pt)  Top=0.1877273 (148.68pt)  Right=0.9282353  Bottom=0.2095454

**Row:** 7, **Column:** 9

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 384.0000 | 79.2000 | 60.0000 | 69.4800 | ✗ |
| 0 | Cells(1) ($I$7) | 384.0000 | 79.2000 | 60.0000 | 69.4800 | ✗ |
| 0 | Range[1] (default index) | 384.0000 | 79.2000 | 60.0000 | 69.4800 | ✗ |
| 0 | OpenXML Column Width Sum | 384.0000 | N/A | 60.0000 | N/A | - |
| 0 |  | N/A | 86.4000 | N/A | 62.2800 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 79.2000 | 324.0000 | 69.4800 | ✗ |
| 0 | EntireColumn.Left/Top | 384.0000 | 0.0000 | 60.0000 | 148.6800 | ✗ |
| 0 | PointsToScreenPixels (427,232) | 320.2500 | 174.0000 | 3.7500 | 25.3200 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 504.0000 | 85.8000 | 180.0000 | 62.8800 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 86.4000 | N/A | 62.2800 | ✗ |
| 15 | Row# × 14.4pt | N/A | 86.4000 | N/A | 62.2800 | ✗ |
| 16 | Col# × 48pt | 384.0000 | N/A | 60.0000 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 79.2000 | N/A | 69.4800 | ✗ |
| 18 | COM Column Width Sum | 385.1240 | N/A | 61.1240 | N/A | ✗ |
| 19 | Printable Area Coords | 434.4000 | 133.2000 | 110.4000 | 15.4800 | ✗ |
| 20 | DB (expected) | 324.0000 | 148.6800 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin(-60.00,65.04) | 324.0000 | 144.2400 | 0.0000 | 4.4400 | ✗ |


#### Cluster: `$I$8:$M$8`

**DB Values:** Left=0.5294118 (324.00pt)  Top=0.2104545 (166.68pt)  Right=0.9282353  Bottom=0.2327273

**Row:** 8, **Column:** 9

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 384.0000 | 92.4000 | 60.0000 | 74.2800 | ✗ |
| 0 | Cells(1) ($I$8) | 384.0000 | 92.4000 | 60.0000 | 74.2800 | ✗ |
| 0 | Range[1] (default index) | 384.0000 | 92.4000 | 60.0000 | 74.2800 | ✗ |
| 0 | OpenXML Column Width Sum | 384.0000 | N/A | 60.0000 | N/A | - |
| 0 |  | N/A | 100.8000 | N/A | 65.8800 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 92.4000 | 324.0000 | 74.2800 | ✗ |
| 0 | EntireColumn.Left/Top | 384.0000 | 0.0000 | 60.0000 | 166.6800 | ✗ |
| 0 | PointsToScreenPixels (427,245) | 320.2500 | 183.7500 | 3.7500 | 17.0700 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 504.0000 | 99.0000 | 180.0000 | 67.6800 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 100.8000 | N/A | 65.8800 | ✗ |
| 15 | Row# × 14.4pt | N/A | 100.8000 | N/A | 65.8800 | ✗ |
| 16 | Col# × 48pt | 384.0000 | N/A | 60.0000 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 92.4000 | N/A | 74.2800 | ✗ |
| 18 | COM Column Width Sum | 385.1240 | N/A | 61.1240 | N/A | ✗ |
| 19 | Printable Area Coords | 434.4000 | 146.4000 | 110.4000 | 20.2800 | ✗ |
| 20 | DB (expected) | 324.0000 | 166.6800 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin(-60.00,65.04) | 324.0000 | 157.4400 | 0.0000 | 9.2399 | ✗ |


#### Cluster: `$I$9:$M$9`

**DB Values:** Left=0.5294118 (324.00pt)  Top=0.2336364 (185.04pt)  Right=0.9282353  Bottom=0.2559091

**Row:** 9, **Column:** 9

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 384.0000 | 105.6000 | 60.0000 | 79.4400 | ✗ |
| 0 | Cells(1) ($I$9) | 384.0000 | 105.6000 | 60.0000 | 79.4400 | ✗ |
| 0 | Range[1] (default index) | 384.0000 | 105.6000 | 60.0000 | 79.4400 | ✗ |
| 0 | OpenXML Column Width Sum | 384.0000 | N/A | 60.0000 | N/A | - |
| 0 |  | N/A | 115.2000 | N/A | 69.8400 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 105.6000 | 324.0000 | 79.4400 | ✗ |
| 0 | EntireColumn.Left/Top | 384.0000 | 0.0000 | 60.0000 | 185.0400 | ✗ |
| 0 | PointsToScreenPixels (427,258) | 320.2500 | 193.5000 | 3.7500 | 8.4600 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 504.0000 | 112.2000 | 180.0000 | 72.8400 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 115.2000 | N/A | 69.8400 | ✗ |
| 15 | Row# × 14.4pt | N/A | 115.2000 | N/A | 69.8400 | ✗ |
| 16 | Col# × 48pt | 384.0000 | N/A | 60.0000 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 105.6000 | N/A | 79.4400 | ✗ |
| 18 | COM Column Width Sum | 385.1240 | N/A | 61.1240 | N/A | ✗ |
| 19 | Printable Area Coords | 434.4000 | 159.6000 | 110.4000 | 25.4400 | ✗ |
| 20 | DB (expected) | 324.0000 | 185.0400 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin(-60.00,65.04) | 324.0000 | 170.6400 | 0.0000 | 14.4000 | ✗ |


#### Cluster: `$I$10:$M$10`

**DB Values:** Left=0.5294118 (324.00pt)  Top=0.2568182 (203.40pt)  Right=0.9282353  Bottom=0.2786364

**Row:** 10, **Column:** 9

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 384.0000 | 118.8000 | 60.0000 | 84.6000 | ✗ |
| 0 | Cells(1) ($I$10) | 384.0000 | 118.8000 | 60.0000 | 84.6000 | ✗ |
| 0 | Range[1] (default index) | 384.0000 | 118.8000 | 60.0000 | 84.6000 | ✗ |
| 0 | OpenXML Column Width Sum | 384.0000 | N/A | 60.0000 | N/A | - |
| 0 |  | N/A | 129.6000 | N/A | 73.8000 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 118.8000 | 324.0000 | 84.6000 | ✗ |
| 0 | EntireColumn.Left/Top | 384.0000 | 0.0000 | 60.0000 | 203.4000 | ✗ |
| 0 | PointsToScreenPixels (427,271) | 320.2500 | 203.2500 | 3.7500 | 0.1500 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 504.0000 | 125.4000 | 180.0000 | 78.0000 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 129.6000 | N/A | 73.8000 | ✗ |
| 15 | Row# × 14.4pt | N/A | 129.6000 | N/A | 73.8000 | ✗ |
| 16 | Col# × 48pt | 384.0000 | N/A | 60.0000 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 118.8000 | N/A | 84.6000 | ✗ |
| 18 | COM Column Width Sum | 385.1240 | N/A | 61.1240 | N/A | ✗ |
| 19 | Printable Area Coords | 434.4000 | 172.8000 | 110.4000 | 30.6000 | ✗ |
| 20 | DB (expected) | 324.0000 | 203.4000 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin(-60.00,65.04) | 324.0000 | 183.8400 | 0.0000 | 19.5600 | ✗ |


### Full Worksheet Range Dump (first 50 rows × first 20 columns)

For every cell position, dump the COM Left/Top values:

#### Row Heights

| Row | COM RowHeight(pt) | COM Left | COM Top |
|:---:|:-----------------:|:--------:|:-------:|
| 1 | 13.20 | 0.00 | 0.00 |
| 2 | 13.20 | 0.00 | 13.20 |
| 3 | 13.20 | 0.00 | 26.40 |
| 4 | 13.20 | 0.00 | 39.60 |
| 5 | 13.20 | 0.00 | 52.80 |

#### Column Widths

| Col | COM ColumnWidth(chars) | COM Left | COM Top |
|:---:|:----------------------:|:--------:|:-------:|
| 1(A) | 8.110 | 0.00 | 0.00 |
| 2(B) | 8.110 | 48.00 | 0.00 |
| 3(C) | 8.110 | 96.00 | 0.00 |
| 4(D) | 8.110 | 144.00 | 0.00 |
| 5(E) | 8.110 | 192.00 | 0.00 |
| 6(F) | 8.110 | 240.00 | 0.00 |
| 7(G) | 8.110 | 288.00 | 0.00 |

#### Every Cell Position Grid

| Row | Data |
|:---:|------|
| 1 | Ht=13.20 Top=0.00 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 
| 2 | Ht=13.20 Top=13.20 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 
| 3 | Ht=13.20 Top=26.40 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 
| 4 | Ht=13.20 Top=39.60 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 
| 5 | Ht=13.20 Top=52.80 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 

---

## Template 548 — Sample A
- **Workbook:** `C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Sample A.xlsx`
- **DB Clusters:** 2

### Page Setup

| Property | Value |
|----------|-------|
| Page Size | 612.0 x 792.0 |
| Margins L,R,T,B | 51.0, 51.0, 53.9, 53.9 |
| Center H,V | False, False |
| Zoom | 100% |
| FitToPages | 1 x 1 |
| Printable Area | 510.0 x 684.3 |
| Printable Origin | (51.02, 53.86) |

### OpenXML Column Widths


### OpenXML Row Heights


### Per-Cluster Coordinate Source Comparison

For each cluster, the following sources are measured and compared to DB:

1. **Range.Left/Top** — our current source (COM Range position)
2. **MergeArea.Left/Top** — if merged, the merge area's origin
3. **MergeArea.Left + Width** — the right edge of merge area
4. **TopLeftCell.Left/Top** — first cell of the range
5. **Range(1).Left/Top** — first item within range
6. **Range.Cells(1).Left/Top** — Cells(1) = first cell of range
7. **Printable coords** — position relative to printable area
8. **Computed from Column widths** — sum of prior column widths
9. **Computed from Row heights** — sum of prior row heights
10. **PointsToScreenPixels** — screen pixel coordinates
11. **PageSetup coords** — position within page setup
12. **Center of range** — Left+Width/2, Top+Height/2

#### Cluster: `$A$10:$D$11`

**DB Values:** Left=0.0847059 (51.84pt)  Top=0.2040909 (161.64pt)  Right=0.4111765  Bottom=0.2418182

**Row:** 10, **Column:** 1

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 0.0000 | 129.6000 | 51.8400 | 32.0400 | ✗ |
| 0 | Cells(1) ($A$10) | 0.0000 | 129.6000 | 51.8400 | 32.0400 | ✗ |
| 0 | Range[1] (default index) | 0.0000 | 129.6000 | 51.8400 | 32.0400 | ✗ |
| 0 | OpenXML Column Width Sum | 0.0000 | N/A | 51.8400 | N/A | - |
| 0 |  | N/A | 129.6000 | N/A | 32.0400 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 129.6000 | 51.8400 | 32.0400 | ✗ |
| 0 | EntireColumn.Left/Top | 0.0000 | 0.0000 | 51.8400 | 161.6400 | ✗ |
| 0 | PointsToScreenPixels (34,278) | 25.5000 | 208.5000 | 26.3400 | 46.8600 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 96.0000 | 144.0000 | 44.1600 | 17.6400 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 129.6000 | N/A | 32.0400 | ✗ |
| 15 | Row# × 14.4pt | N/A | 129.6000 | N/A | 32.0400 | ✗ |
| 16 | Col# × 48pt | 0.0000 | N/A | 51.8400 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 129.6000 | N/A | 32.0400 | ✗ |
| 18 | COM Column Width Sum | 0.0000 | N/A | 51.8400 | N/A | ✗ |
| 19 | Printable Area Coords | 51.0236 | 183.4583 | 0.8164 | 21.8183 | ✗ |
| 20 | DB (expected) | 51.8400 | 161.6400 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin (51.84,32.04) | 51.8400 | 161.6400 | 0.0000 | 0.0000 | ✅ |


#### Cluster: `$E$10:$G$11`

**DB Values:** Left=0.4129412 (252.72pt)  Top=0.2040909 (161.64pt)  Right=0.6582353  Bottom=0.2418182

**Row:** 10, **Column:** 5

| # | Source | Left(pt) | Top(pt) | L-DB(pt) | T-DB(pt) | Match? |
|---|-------|:--------:|:-------:|:--------:|:--------:|:------:|
| 0 | Range.Left/Top | 192.0000 | 129.6000 | 60.7200 | 32.0400 | ✗ |
| 0 | Cells(1) ($E$10) | 192.0000 | 129.6000 | 60.7200 | 32.0400 | ✗ |
| 0 | Range[1] (default index) | 192.0000 | 129.6000 | 60.7200 | 32.0400 | ✗ |
| 0 | OpenXML Column Width Sum | 192.0000 | N/A | 60.7200 | N/A | - |
| 0 |  | N/A | 129.6000 | N/A | 32.0400 | - |
| 0 | EntireRow.Left/Top | 0.0000 | 129.6000 | 252.7200 | 32.0400 | ✗ |
| 0 | EntireColumn.Left/Top | 192.0000 | 0.0000 | 60.7200 | 161.6400 | ✗ |
| 0 | PointsToScreenPixels (226,278) | 169.5000 | 208.5000 | 83.2200 | 46.8600 | ✗ |
| 0 | Range Center (L+W/2, T+H/2) | 264.0000 | 144.0000 | 11.2800 | 17.6400 | ✗ |
| 14 | OpenXML Row Height Sum | N/A | 129.6000 | N/A | 32.0400 | ✗ |
| 15 | Row# × 14.4pt | N/A | 129.6000 | N/A | 32.0400 | ✗ |
| 16 | Col# × 48pt | 192.0000 | N/A | 60.7200 | N/A | ✗ |
| 17 | COM Row Height Sum | N/A | 129.6000 | N/A | 32.0400 | ✗ |
| 18 | COM Column Width Sum | 192.5620 | N/A | 60.1580 | N/A | ✗ |
| 19 | Printable Area Coords | 243.0236 | 183.4583 | 9.6964 | 21.8183 | ✗ |
| 20 | DB (expected) | 252.7200 | 161.6400 | 0.0000 | 0.0000 | REFERENCE |
| 21 | COM + Origin(51.84,32.04) | 243.8400 | 161.6400 | 8.8800 | 0.0000 | ✗ |


### Full Worksheet Range Dump (first 50 rows × first 20 columns)

For every cell position, dump the COM Left/Top values:

#### Row Heights

| Row | COM RowHeight(pt) | COM Left | COM Top |
|:---:|:-----------------:|:--------:|:-------:|
| 1 | 14.40 | 0.00 | 0.00 |
| 2 | 14.40 | 0.00 | 14.40 |
| 3 | 14.40 | 0.00 | 28.80 |
| 4 | 14.40 | 0.00 | 43.20 |
| 5 | 14.40 | 0.00 | 57.60 |
| 6 | 14.40 | 0.00 | 72.00 |
| 7 | 14.40 | 0.00 | 86.40 |
| 8 | 14.40 | 0.00 | 100.80 |
| 9 | 14.40 | 0.00 | 115.20 |

#### Column Widths

| Col | COM ColumnWidth(chars) | COM Left | COM Top |
|:---:|:----------------------:|:--------:|:-------:|
| 1(A) | 8.110 | 0.00 | 0.00 |
| 2(B) | 8.110 | 48.00 | 0.00 |
| 3(C) | 8.110 | 96.00 | 0.00 |
| 4(D) | 8.110 | 144.00 | 0.00 |
| 5(E) | 8.110 | 192.00 | 0.00 |
| 6(F) | 8.110 | 240.00 | 0.00 |
| 7(G) | 8.110 | 288.00 | 0.00 |

#### Every Cell Position Grid

| Row | Data |
|:---:|------|
| 1 | Ht=14.40 Top=0.00 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 
| 2 | Ht=14.40 Top=14.40 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 
| 3 | Ht=14.40 Top=28.80 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 
| 4 | Ht=14.40 Top=43.20 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 
| 5 | Ht=14.40 Top=57.60 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 
| 6 | Ht=14.40 Top=72.00 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 
| 7 | Ht=14.40 Top=86.40 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 
| 8 | Ht=14.40 Top=100.80 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 
| 9 | Ht=14.40 Top=115.20 | Cols: C1 L=0.0 W=48.0 C2 L=48.0 W=48.0 C3 L=96.0 W=48.0 C4 L=144.0 W=48.0 C5 L=192.0 W=48.0 C6 L=240.0 W=48.0 C7 L=288.0 W=48.0 

---

## Overall Results

| Template | Total Clusters | Any Source Matches? | Best Source | Best Accuracy |
|----------|:--------------:|:-------------------:|:------------|:-------------:|
| 546 | 6 | NO | Range.Left+Origin (source 21) | NONE |
| 547 | 5 | NO | Range.Left+Origin (source 21) | NONE |
| 548 | 2 | NO | Range.Left+Origin (source 21) | NONE |

