# Coordinate System Comparison

## Overview

This document compares three coordinate systems:
1. **OpenXML Calculated** — Our current engine (estimates from column widths)
2. **Excel COM** — Excel's own rendered positions (Range.Left/Top)
3. **Database** — The legacy system's stored ratios

## Per-Cluster Comparison

| Address | DB Left | COM Left | DB Top | COM Top | DB→COM x | DB→COM y |
|---------|---------|----------|--------|---------|----------|----------|
| $A$1:$B$2 | 0.3364706 (205.92pt) | 0.00pt | 0.3845454 (304.56pt) | 0.00pt | +205.92pt | +304.56pt |
| $A$12 | 0.3364706 (205.92pt) | 0.00pt | 0.5963637 (472.32pt) | 158.40pt | +205.92pt | +313.92pt |
| $A$3:$D$4 | 0.3364706 (205.92pt) | 0.00pt | 0.4231818 (335.16pt) | 28.80pt | +205.92pt | +306.36pt |
| $A$6:$D$7 | 0.3364706 (205.92pt) | 0.00pt | 0.4809091 (380.88pt) | 72.00pt | +205.92pt | +308.88pt |
| $A$9:$D$10 | 0.3364706 (205.92pt) | 0.00pt | 0.5386364 (426.60pt) | 115.20pt | +205.92pt | +311.40pt |
| $C$1:$D$2 | 0.5000000 (306.00pt) | 96.00pt | 0.3845454 (304.56pt) | 0.00pt | +210.00pt | +304.56pt |

## Coordinate Math

## Summary Statistics

- Average DB/COM ratio X: 0.531250
- Average DB/COM ratio Y: 3.935407
- Average offset X: 206.60 pt
- Average offset Y: 308.28 pt

- Raw ratio: DB/COM = 0.531250
- Printable/Page ratio: 0.835294
- **No immediate match found** — investigating further...

### Margin-Adjusted Hypothesis
- $A$1:$B$2: COM_Left=0.00, DB_left*Page-margin=155.52, diff=-155.52
- $C$1:$D$2: COM_Left=96.00, DB_left*Page-margin=255.60, diff=-159.60
- $A$3:$D$4: COM_Left=0.00, DB_left*Page-margin=155.52, diff=-155.52
