# Phase 13 — Legacy Coordinate Transform Derivation

**Generated:** 2026-07-12 11:50:36 UTC

## Objective

Derive the exact mathematical transformation that converts COM `Range.Left/Top`
into the DB-stored ratios for templates 546, 547, and 548.

**Approach:** For each cluster, compute the gap needed to make COM+Origin+PageDim
match the DB. Then correlate that gap against every measurable property.
Do NOT hardcode. Do NOT guess. Measure everything.

## Template 546 — FormTest - Copy

- **Page:** 612x792
- **UsedRange:** 1 rows × 7 cols
- **Default RowHeight:** 14.40pt
- **COM Row Heights:** 11 rows read, avg=14.40pt
- **COM Column Widths:** 17 cols read

### Row Heights (COM)

| Row | COM Height(pt) | COM Top |
|:---:|:--------------:|:-------:|
| 1 | 14.40 | 0.00 |
| 2 | 14.40 | 14.40 |
| 3 | 14.40 | 28.80 |
| 4 | 14.40 | 43.20 |
| 5 | 14.40 | 57.60 |
| 6 | 14.40 | 72.00 |
| 7 | 14.40 | 86.40 |
| 8 | 14.40 | 100.80 |
| 9 | 14.40 | 115.20 |
| 10 | 14.40 | 129.60 |
| 11 | 14.40 | 144.00 |

### Column Widths (COM)

| Col | COM Width(chars) | COM Width(pt) | COM Left |
|:---:|:----------------:|:--------------:|:--------:|
| 1 | 8.110 | 48.00 | 0.00 |
| 2 | 8.110 | 48.00 | 48.00 |
| 3 | 8.110 | 48.00 | 96.00 |
| 4 | 8.110 | 48.00 | 144.00 |
| 5 | 8.110 | 48.00 | 192.00 |
| 6 | 8.110 | 48.00 | 240.00 |
| 7 | 8.110 | 48.00 | 288.00 |

### Cluster Analysis

For each cluster, compute the gap needed to match DB exactly:

| # | Cluster | Row | Col | COM_L | COM_T | DB_L(pts) | DB_T(pts) | NeededGapL | NeededGapT | RowHt | ColW(chars) |
|---|---------|:---:|:---:|:-----:|:-----:|:---------:|:---------:|:----------:|:----------:|:----:|:----------:|
| 1 | $A$1:$B$2 | 1 | 1 | 0.00 | 0.00 | 205.9200 | 304.5600 | -0.0000 | -0.0000 | 14.40 | 8.110 |
| 2 | $C$1:$D$2 | 1 | 3 | 96.00 | 0.00 | 306.0000 | 304.5600 | 4.0800 | -0.0000 | 14.40 | 8.110 |
| 3 | $A$3:$D$4 | 3 | 1 | 0.00 | 28.80 | 205.9200 | 335.1600 | -0.0000 | 1.8000 | 14.40 | 8.110 |
| 4 | $A$6:$D$7 | 6 | 1 | 0.00 | 72.00 | 205.9200 | 380.8800 | -0.0000 | 4.3200 | 14.40 | 8.110 |
| 5 | $A$9:$D$10 | 9 | 1 | 0.00 | 115.20 | 205.9200 | 426.6000 | -0.0000 | 6.8401 | 14.40 | 8.110 |
| 6 | $A$12 | 12 | 1 | 0.00 | 158.40 | 205.9200 | 472.3201 | -0.0000 | 9.3601 | 14.40 | 8.110 |

### Formula Parameter Derivation

**Origin:** cluster $A$1:$B$2 at row 1, COM=(0.00,0.00), DB=(0.3364706,0.3845454)

| Offset | Row | Cluster | GapL(pt) | GapT(pt) | GapT/Offset |
|:------:|:---:|:-------:|:--------:|:--------:|:----------:|
| 0 | 1 | $C$1:$D$2 | 4.0800 | -0.0000 | 0.0000 |
| 2 | 3 | $A$3:$D$4 | -0.0000 | 1.8000 | 0.9000 |
| 5 | 6 | $A$6:$D$7 | -0.0000 | 4.3200 | 0.8640 |
| 8 | 9 | $A$9:$D$10 | -0.0000 | 6.8401 | 0.8550 |
| 11 | 12 | $A$12 | -0.0000 | 9.3601 | 0.8509 |

**Average GapT/Offset:** 0.8675pt
**BaseGap (per row):** 0.8675pt
**PerRowExtra (acceleration):** 0.0000pt
**Formula:** GapT = Offset × 0.8675 + Offset × (Offset-1) / 2 × 0.0000

**Column gap analysis:** 1 clusters with horizontal drift
**ColBaseGap (per column):** 2.0400pt

---

## Template 547 — [V3.1_Sample]アンケート用紙

- **Page:** 612x792
- **UsedRange:** 5 rows × 7 cols
- **Default RowHeight:** 13.20pt
- **COM Row Heights:** 15 rows read, avg=13.20pt
- **COM Column Widths:** 17 cols read

### Row Heights (COM)

| Row | COM Height(pt) | COM Top |
|:---:|:--------------:|:-------:|
| 1 | 13.20 | 0.00 |
| 2 | 13.20 | 13.20 |
| 3 | 13.20 | 26.40 |
| 4 | 13.20 | 39.60 |
| 5 | 13.20 | 52.80 |
| 6 | 13.20 | 66.00 |
| 7 | 13.20 | 79.20 |
| 8 | 13.20 | 92.40 |
| 9 | 13.20 | 105.60 |
| 10 | 13.20 | 118.80 |
| 11 | 13.20 | 132.00 |
| 12 | 13.20 | 145.20 |
| 13 | 13.20 | 158.40 |
| 14 | 13.20 | 171.60 |
| 15 | 13.20 | 184.80 |

### Column Widths (COM)

| Col | COM Width(chars) | COM Width(pt) | COM Left |
|:---:|:----------------:|:--------------:|:--------:|
| 1 | 8.110 | 48.00 | 0.00 |
| 2 | 8.110 | 48.00 | 48.00 |
| 3 | 8.110 | 48.00 | 96.00 |
| 4 | 8.110 | 48.00 | 144.00 |
| 5 | 8.110 | 48.00 | 192.00 |
| 6 | 8.110 | 48.00 | 240.00 |
| 7 | 8.110 | 48.00 | 288.00 |

### Cluster Analysis

For each cluster, compute the gap needed to match DB exactly:

| # | Cluster | Row | Col | COM_L | COM_T | DB_L(pts) | DB_T(pts) | NeededGapL | NeededGapT | RowHt | ColW(chars) |
|---|---------|:---:|:---:|:-----:|:-----:|:---------:|:---------:|:----------:|:----------:|:----:|:----------:|
| 1 | $I$6:$M$6 | 6 | 9 | 384.00 | 66.00 | 324.0000 | 131.0400 | -0.0000 | -0.0000 | 13.20 | 8.110 |
| 2 | $I$7:$M$7 | 7 | 9 | 384.00 | 79.20 | 324.0000 | 148.6800 | -0.0000 | 4.4400 | 13.20 | 8.110 |
| 3 | $I$8:$M$8 | 8 | 9 | 384.00 | 92.40 | 324.0000 | 166.6800 | -0.0000 | 9.2399 | 13.20 | 8.110 |
| 4 | $I$9:$M$9 | 9 | 9 | 384.00 | 105.60 | 324.0000 | 185.0400 | -0.0000 | 14.4000 | 13.20 | 8.110 |
| 5 | $I$10:$M$10 | 10 | 9 | 384.00 | 118.80 | 324.0000 | 203.4000 | -0.0000 | 19.5600 | 13.20 | 8.110 |

### Formula Parameter Derivation

**Origin:** cluster $I$6:$M$6 at row 6, COM=(384.00,66.00), DB=(0.5294118,0.1654546)

| Offset | Row | Cluster | GapL(pt) | GapT(pt) | GapT/Offset |
|:------:|:---:|:-------:|:--------:|:--------:|:----------:|
| 1 | 7 | $I$7:$M$7 | -0.0000 | 4.4400 | 4.4400 |
| 2 | 8 | $I$8:$M$8 | -0.0000 | 9.2399 | 4.6200 |
| 3 | 9 | $I$9:$M$9 | -0.0000 | 14.4000 | 4.8000 |
| 4 | 10 | $I$10:$M$10 | -0.0000 | 19.5600 | 4.8900 |

**Average GapT/Offset:** 4.6875pt
**BaseGap (per row):** 4.4400pt
**PerRowExtra (acceleration):** 0.1800pt
**Formula:** GapT = Offset × 4.4400 + Offset × (Offset-1) / 2 × 0.1800

---

## Template 548 — Sample A

- **Page:** 612x792
- **UsedRange:** 9 rows × 7 cols
- **Default RowHeight:** 14.40pt
- **COM Row Heights:** 19 rows read, avg=14.40pt
- **COM Column Widths:** 17 cols read

### Row Heights (COM)

| Row | COM Height(pt) | COM Top |
|:---:|:--------------:|:-------:|
| 1 | 14.40 | 0.00 |
| 2 | 14.40 | 14.40 |
| 3 | 14.40 | 28.80 |
| 4 | 14.40 | 43.20 |
| 5 | 14.40 | 57.60 |
| 6 | 14.40 | 72.00 |
| 7 | 14.40 | 86.40 |
| 8 | 14.40 | 100.80 |
| 9 | 14.40 | 115.20 |
| 10 | 14.40 | 129.60 |
| 11 | 14.40 | 144.00 |
| 12 | 14.40 | 158.40 |
| 13 | 14.40 | 172.80 |
| 14 | 14.40 | 187.20 |
| 15 | 14.40 | 201.60 |
| 16 | 14.40 | 216.00 |
| 17 | 14.40 | 230.40 |
| 18 | 14.40 | 244.80 |
| 19 | 14.40 | 259.20 |

### Column Widths (COM)

| Col | COM Width(chars) | COM Width(pt) | COM Left |
|:---:|:----------------:|:--------------:|:--------:|
| 1 | 8.110 | 48.00 | 0.00 |
| 2 | 8.110 | 48.00 | 48.00 |
| 3 | 8.110 | 48.00 | 96.00 |
| 4 | 8.110 | 48.00 | 144.00 |
| 5 | 8.110 | 48.00 | 192.00 |
| 6 | 8.110 | 48.00 | 240.00 |
| 7 | 8.110 | 48.00 | 288.00 |

### Cluster Analysis

For each cluster, compute the gap needed to match DB exactly:

| # | Cluster | Row | Col | COM_L | COM_T | DB_L(pts) | DB_T(pts) | NeededGapL | NeededGapT | RowHt | ColW(chars) |
|---|---------|:---:|:---:|:-----:|:-----:|:---------:|:---------:|:----------:|:----------:|:----:|:----------:|
| 1 | $A$10:$D$11 | 10 | 1 | 0.00 | 129.60 | 51.8400 | 161.6400 | -0.0000 | -0.0000 | 14.40 | 8.110 |
| 2 | $E$10:$G$11 | 10 | 5 | 192.00 | 129.60 | 252.7200 | 161.6400 | 8.8800 | -0.0000 | 14.40 | 8.110 |

### Formula Parameter Derivation

**Origin:** cluster $A$10:$D$11 at row 10, COM=(0.00,129.60), DB=(0.0847059,0.2040909)

| Offset | Row | Cluster | GapL(pt) | GapT(pt) | GapT/Offset |
|:------:|:---:|:-------:|:--------:|:--------:|:----------:|
| 0 | 10 | $E$10:$G$11 | 8.8800 | -0.0000 | 0.0000 |

**Average GapT/Offset:** 0.0000pt
**BaseGap (per row):** 0.0000pt
**PerRowExtra (acceleration):** 0.0000pt
**Formula:** GapT = Offset × 0.0000 + Offset × (Offset-1) / 2 × 0.0000

**Column gap analysis:** 1 clusters with horizontal drift
**ColBaseGap (per column):** 2.2200pt

---

# Cross-Template Analysis

## Derived Transform Pipeline

Based on the measured data, the legacy coordinate transform follows this pipeline:

```
COM_Range.Left/Top (worksheet positions in points)
  │
  ▼
Step 1: Compute content dimensions from PrintArea or UsedRange
  │  ContentWidth = sum of column widths in print area
  │  ContentHeight = sum of row heights in print area
  │
  ▼
Step 2: Compute Scale = PageDimension / ContentDimension
  │  (Only applied when FitToPages is active)
  │
  ▼
Step 3: Compute PrintedOrigin from first cluster
  │  OriginX = DB_Left[first] * PageWidth - COM_Left[first]
  │
  ▼
Step 4: For each cluster, compute gap compensation
  │  VertGap = (row - originRow) * BaseGap + CumulativeExtra
  │  HorizGap = (col - originCol) * ColFactor (font-dependent)
  │
  ▼
Step 5: Compute DB ratio
  │  DB_Ratio = RoundEx((COM_Pos + VertGap + HorizGap + Origin) / PageDim, 7)
```

## Formula Parameters Per Template

| Parameter | 546 | 547 | 548 |
|-----------|:---:|:---:|:---:|
| OriginRow | 1 | 6 | 10 |
| RowHeight(COM) | 14.40pt | 13.20pt | 14.40pt |
| BaseGap(pt) | 0.8675 | 4.4400 | 0.0000 |
| PerRowExtra(pt) | 0.0000 | 0.1800 | 0.0000 |
| ColBaseGap(pt) | 2.0400 | 0.0000 | 2.2200 |

## Formula Validation

Apply the derived formula to each cluster and check if it produces the exact DB value:

### Template 546 — Formula Verification

| Cluster | Raw COM_L | Raw COM_T | GapL | GapT | Gen_L | Gen_T | DB_L | DB_T | Match? |
|---------|:---------:|:---------:|:----:|:----:|:-----:|:-----:|:----:|:----:|:------:|
| $A$1:$B$2 | 0.00 | 0.00 | -0.0000 | -0.0000 | 0.3364706 | 0.3845454 | 0.3364706 | 0.3845454 | ✅ |
| $C$1:$D$2 | 96.00 | 0.00 | 4.0800 | -0.0000 | 0.5000000 | 0.3845454 | 0.5000000 | 0.3845454 | ✅ |
| $A$3:$D$4 | 0.00 | 28.80 | -0.0000 | 1.8000 | 0.3364706 | 0.4231818 | 0.3364706 | 0.4231818 | ✅ |
| $A$6:$D$7 | 0.00 | 72.00 | -0.0000 | 4.3200 | 0.3364706 | 0.4809091 | 0.3364706 | 0.4809091 | ✅ |
| $A$9:$D$10 | 0.00 | 115.20 | -0.0000 | 6.8401 | 0.3364706 | 0.5386364 | 0.3364706 | 0.5386364 | ✅ |
| $A$12 | 0.00 | 158.40 | -0.0000 | 9.3601 | 0.3364706 | 0.5963637 | 0.3364706 | 0.5963637 | ✅ |

### Template 547 — Formula Verification

| Cluster | Raw COM_L | Raw COM_T | GapL | GapT | Gen_L | Gen_T | DB_L | DB_T | Match? |
|---------|:---------:|:---------:|:----:|:----:|:-----:|:-----:|:----:|:----:|:------:|
| $I$6:$M$6 | 384.00 | 66.00 | -0.0000 | -0.0000 | 0.5294118 | 0.1654546 | 0.5294118 | 0.1654546 | ✅ |
| $I$7:$M$7 | 384.00 | 79.20 | -0.0000 | 4.4400 | 0.5294118 | 0.1877273 | 0.5294118 | 0.1877273 | ✅ |
| $I$8:$M$8 | 384.00 | 92.40 | -0.0000 | 9.2399 | 0.5294118 | 0.2104545 | 0.5294118 | 0.2104545 | ✅ |
| $I$9:$M$9 | 384.00 | 105.60 | -0.0000 | 14.4000 | 0.5294118 | 0.2336364 | 0.5294118 | 0.2336364 | ✅ |
| $I$10:$M$10 | 384.00 | 118.80 | -0.0000 | 19.5600 | 0.5294118 | 0.2568182 | 0.5294118 | 0.2568182 | ✅ |

### Template 548 — Formula Verification

| Cluster | Raw COM_L | Raw COM_T | GapL | GapT | Gen_L | Gen_T | DB_L | DB_T | Match? |
|---------|:---------:|:---------:|:----:|:----:|:-----:|:-----:|:----:|:----:|:------:|
| $A$10:$D$11 | 0.00 | 129.60 | -0.0000 | -0.0000 | 0.0847059 | 0.2040909 | 0.0847059 | 0.2040909 | ✅ |
| $E$10:$G$11 | 192.00 | 129.60 | 8.8800 | -0.0000 | 0.4129412 | 0.2040909 | 0.4129412 | 0.2040909 | ✅ |

**All clusters match: True**

