# Coordinate System Comparison

## Overview

This document compares three coordinate systems:
1. **OpenXML Calculated** — Our current engine (estimates from column widths)
2. **Excel COM** — Excel's own rendered positions (Range.Left/Top)
3. **Database** — The legacy system's stored ratios

## Per-Cluster Comparison

| Address | DB Left | COM Left | DB Top | COM Top | DB→COM x | DB→COM y |
|---------|---------|----------|--------|---------|----------|----------|
| $I$10:$M$10 | 0.5294118 (324.00pt) | 384.00pt | 0.2568182 (203.40pt) | 118.80pt | -60.00pt | +84.60pt |
| $I$6:$M$6 | 0.5294118 (324.00pt) | 384.00pt | 0.1654546 (131.04pt) | 66.00pt | -60.00pt | +65.04pt |
| $I$7:$M$7 | 0.5294118 (324.00pt) | 384.00pt | 0.1877273 (148.68pt) | 79.20pt | -60.00pt | +69.48pt |
| $I$8:$M$8 | 0.5294118 (324.00pt) | 384.00pt | 0.2104545 (166.68pt) | 92.40pt | -60.00pt | +74.28pt |
| $I$9:$M$9 | 0.5294118 (324.00pt) | 384.00pt | 0.2336364 (185.04pt) | 105.60pt | -60.00pt | +79.44pt |

## Coordinate Math

## Summary Statistics

- Average DB/COM ratio X: 0.843750
- Average DB/COM ratio Y: 1.826204
- Average offset X: -60.00 pt
- Average offset Y: 74.57 pt

- **Hypothesis: DB coordinates are relative to printable area (minus margins)**
  - Page width: 612.0 pt
  - Printable width: 511.2 pt
  - Printable/Page ratio: 0.835294
  - DB Left = (COM_Left - LeftMargin) / PrintableWidth

### Margin-Adjusted Hypothesis
- $I$6:$M$6: COM_Left=384.00, DB_left*Page-margin=273.60, diff=110.40
- $I$7:$M$7: COM_Left=384.00, DB_left*Page-margin=273.60, diff=110.40
- $I$8:$M$8: COM_Left=384.00, DB_left*Page-margin=273.60, diff=110.40
