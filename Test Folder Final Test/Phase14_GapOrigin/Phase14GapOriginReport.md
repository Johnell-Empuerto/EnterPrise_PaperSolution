# Phase 14 — Legacy Gap Origin Investigation Report

**Generated:** 2026-07-12 11:56:42 UTC

## Objective

Identify the measurable Excel property that explains why the vertical gap
parameters differ between templates 546 (BaseGap=0.87pt), 547 (BaseGap=4.44pt),
and 548 (BaseGap=0pt). No curve fitting — every parameter must be explained
by a measurable workbook property.

## Template 546

**PrintArea:** $A$1:$D$12
**Rows:** 1–12 (12 rows)
**Cols:** 1–4 (4 cols)
**BaseGap:** 0.9000pt
**PerRowExtra:** -0.0120pt
**ColGap:** 2.0400pt

### Print Settings

| Property | Value |
|----------|-------|
| Zoom | 100% |
| FitToPagesWide | 1 |
| FitToPagesTall | 1 |
| Orientation | 1 |
| Page Size | 612x792 |

### Content Dimensions

| Dimension | Value |
|-----------|-------|
| TotalContentHeight (COM) | 720.00pt |
| TotalContentWidth (COM) | 962.81pt |
| OpenXmlContentHeight | 172.80pt |
| OpenXmlContentWidth | 192.56pt |
| ContentHtRatio (COM/Page) | 0.9091 |
| ContentWdRatio (COM/Page) | 1.5732 |
| OpenXmlHtRatio | 0.2182 |
| OpenXmlWdRatio | 0.3146 |
| HScale (width constraint) | 1.0000 |
| VScale (height constraint) | 1.0000 |
| UniformScale | 1.0000 |
| MergedCells | 5 |
| StandardRowHeight | 14.40pt |
| MinRowHt | 14.40pt |
| MaxRowHt | 14.40pt |
| AvgRowHt | 14.40pt |
| AvgColWd | 48.14pt |

---

## Template 547

**PrintArea:** $B$1:$M$44
**Rows:** 1–44 (44 rows)
**Cols:** 2–13 (12 cols)
**BaseGap:** 4.4400pt
**PerRowExtra:** 0.1800pt
**ColGap:** 0.0000pt

### Print Settings

| Property | Value |
|----------|-------|
| Zoom | 100% |
| FitToPagesWide | 1 |
| FitToPagesTall | 1 |
| Orientation | 1 |
| Page Size | 612x792 |

### Content Dimensions

| Dimension | Value |
|-----------|-------|
| TotalContentHeight (COM) | 660.00pt |
| TotalContentWidth (COM) | 914.67pt |
| OpenXmlContentHeight | 1158.15pt |
| OpenXmlContentWidth | 817.75pt |
| ContentHtRatio (COM/Page) | 0.8333 |
| ContentWdRatio (COM/Page) | 1.4946 |
| OpenXmlHtRatio | 1.4623 |
| OpenXmlWdRatio | 1.3362 |
| HScale (width constraint) | 0.7484 |
| VScale (height constraint) | 0.6838 |
| UniformScale | 0.6838 |
| MergedCells | 42 |
| StandardRowHeight | 13.20pt |
| MinRowHt | 13.20pt |
| MaxRowHt | 13.20pt |
| AvgRowHt | 13.20pt |
| AvgColWd | 48.14pt |

---

## Template 548

**PrintArea:** $A$3:$G$11
**Rows:** 3–11 (9 rows)
**Cols:** 1–7 (7 cols)
**BaseGap:** 0.0000pt
**PerRowExtra:** 0.0000pt
**ColGap:** 2.2200pt

### Print Settings

| Property | Value |
|----------|-------|
| Zoom | 100% |
| FitToPagesWide | 1 |
| FitToPagesTall | 1 |
| Orientation | 1 |
| Page Size | 612x792 |

### Content Dimensions

| Dimension | Value |
|-----------|-------|
| TotalContentHeight (COM) | 691.20pt |
| TotalContentWidth (COM) | 962.81pt |
| OpenXmlContentHeight | 129.60pt |
| OpenXmlContentWidth | 336.98pt |
| ContentHtRatio (COM/Page) | 0.8727 |
| ContentWdRatio (COM/Page) | 1.5732 |
| OpenXmlHtRatio | 0.1636 |
| OpenXmlWdRatio | 0.5506 |
| HScale (width constraint) | 1.0000 |
| VScale (height constraint) | 1.0000 |
| UniformScale | 1.0000 |
| MergedCells | 5 |
| StandardRowHeight | 14.40pt |
| MinRowHt | 14.40pt |
| MaxRowHt | 14.40pt |
| AvgRowHt | 14.40pt |
| AvgColWd | 48.14pt |

---

# Cross-Template Correlation Analysis

| Property | 546 | 547 | 548 | Correlates with BaseGap? |
|----------|:---:|:---:|:---:|:------------------------:|
| AvgColWd | 48.14 | 48.14 | 48.14 | NONE |
| AvgRowHt | 14.4 | 13.2 | 14.4 | BaseGap(r=-0.98) PerRowExtra(r=-1.00) ColGap(r=1.00) |
| BottomMargin | 54 | 54 | 53.86 | NONE |
| ContentHtRatio | 0.9091 | 0.8333 | 0.8727 | BaseGap(r=-0.77) PerRowExtra(r=-0.90) ColGap(r=0.84) |
| ContentWdRatio | 1.573 | 1.495 | 1.573 | BaseGap(r=-0.98) PerRowExtra(r=-1.00) ColGap(r=1.00) |
| FitToPagesTall | 1 | 1 | 1 | NONE |
| FitToPagesWide | 1 | 1 | 1 | NONE |
| HScale | 1 | 0.7484 | 1 | BaseGap(r=-0.98) PerRowExtra(r=-1.00) ColGap(r=1.00) |
| LeftMargin | 50.4 | 50.4 | 51.02 | NONE |
| MaxColWd | 48.14 | 48.14 | 48.14 | NONE |
| MaxRowHt | 14.4 | 13.2 | 14.4 | BaseGap(r=-0.98) PerRowExtra(r=-1.00) ColGap(r=1.00) |
| MergedCells | 5 | 42 | 5 | BaseGap(r=0.98) PerRowExtra(r=1.00) ColGap(r=-1.00) |
| MinColWd | 48.14 | 48.14 | 48.14 | NONE |
| MinRowHt | 14.4 | 13.2 | 14.4 | BaseGap(r=-0.98) PerRowExtra(r=-1.00) ColGap(r=1.00) |
| OpenXmlContentHeight | 172.8 | 1158 | 129.6 | BaseGap(r=0.99) PerRowExtra(r=1.00) ColGap(r=-1.00) |
| OpenXmlContentWidth | 192.6 | 817.7 | 337 | BaseGap(r=0.91) PerRowExtra(r=0.99) ColGap(r=-0.96) |
| OpenXmlHtRatio | 0.2182 | 1.462 | 0.1636 | BaseGap(r=0.99) PerRowExtra(r=1.00) ColGap(r=-1.00) |
| OpenXmlWdRatio | 0.3146 | 1.336 | 0.5506 | BaseGap(r=0.91) PerRowExtra(r=0.99) ColGap(r=-0.96) |
| Orientation | 1 | 1 | 1 | NONE |
| PageHeight | 792 | 792 | 792 | NONE |
| PageWidth | 612 | 612 | 612 | NONE |
| PaperSize | 1 | 1 | 1 | NONE |
| PrintableHeight | 684 | 684 | 684.3 | NONE |
| PrintableWidth | 511.2 | 511.2 | 510 | NONE |
| PrintAreaCols | 4 | 12 | 7 | BaseGap(r=0.84) PerRowExtra(r=0.95) ColGap(r=-0.90) |
| PrintAreaRows | 12 | 44 | 9 | BaseGap(r=0.99) PerRowExtra(r=0.99) ColGap(r=-1.00) |
| RightMargin | 50.4 | 50.4 | 51.02 | NONE |
| StandardColumnWidth | 8.11 | 8.11 | 8.11 | NONE |
| StandardRowHeight | 14.4 | 13.2 | 14.4 | BaseGap(r=-0.98) PerRowExtra(r=-1.00) ColGap(r=1.00) |
| TopMargin | 54 | 54 | 53.86 | NONE |
| TotalContentHeight | 720 | 660 | 691.2 | BaseGap(r=-0.77) PerRowExtra(r=-0.90) ColGap(r=0.84) |
| TotalContentWidth | 962.8 | 914.7 | 962.8 | BaseGap(r=-0.98) PerRowExtra(r=-1.00) ColGap(r=1.00) |
| UniformScale | 1 | 0.6838 | 1 | BaseGap(r=-0.98) PerRowExtra(r=-1.00) ColGap(r=1.00) |
| VScale | 1 | 0.6838 | 1 | BaseGap(r=-0.98) PerRowExtra(r=-1.00) ColGap(r=1.00) |
| Zoom | 100 | 100 | 100 | NONE |

## Conclusions

The property with the strongest correlation to BaseGap is **PrintAreaRows** (r=0.9933).

However, with only 3 data points, statistical significance is limited.
Any correlation above r=0.997 would be needed for p<0.05 with n=3.

**STRONG CANDIDATE:** PrintAreaRows correlates with BaseGap at r=0.9933.
This suggests a measurable explanation exists.

