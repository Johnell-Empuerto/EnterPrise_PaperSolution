# Coordinate Math

## Exact Transformation Formulas

### Constants (from PageSetup)

- PageWidth = 612.0 pt
- PageHeight = 792.00 pt
- LeftMargin = 50.4000 pt
- RightMargin = 50.4000 pt
- TopMargin = 54.0000 pt
- BottomMargin = 54.0000 pt
- PrintableWidth = PageWidth - LeftMargin - RightMargin = 511.2 pt
- PrintableHeight = PageHeight - TopMargin - BottomMargin = 684.0 pt
- CenterHorizontally = False
- CenterVertically = False
- FitToPagesWide = 1
- FitToPagesTall = 1
- Zoom = 100

### Hypothesis 1: Printable Area Coordinates

```
DB_Left = (COM_Left - LeftMargin) / PrintableWidth
DB_Top  = (COM_Top - TopMargin) / PrintableHeight
```

| $A$1:$B$2 | -0.0985915 | 0.3364706 | 0.4350621 | -0.0789474 | 0.3845454 | 0.4634928 |
| $C$1:$D$2 | 0.0892019 | 0.5000000 | 0.4107981 | -0.0789474 | 0.3845454 | 0.4634928 |
| $A$3:$D$4 | -0.0985915 | 0.3364706 | 0.4350621 | -0.0368421 | 0.4231818 | 0.4600239 |
| $A$6:$D$7 | -0.0985915 | 0.3364706 | 0.4350621 | 0.0263158 | 0.4809091 | 0.4545933 |
| $A$9:$D$10 | -0.0985915 | 0.3364706 | 0.4350621 | 0.0894737 | 0.5386364 | 0.4491627 |
| $A$12 | -0.0985915 | 0.3364706 | 0.4350621 | 0.1526316 | 0.5963637 | 0.4437321 |

### Hypothesis 2: Centered Printable Coordinates

```
if CenterHorizontally:
  PrintedOriginX = (PageWidth - PrintableWidth) / 2
else:
  PrintedOriginX = LeftMargin
DB_Left = (COM_Left - PrintedOriginX) / PrintableWidth
```

| $A$1:$B$2 | -0.0985915 | 0.3364706 | 0.4350621 | -0.0789474 | 0.3845454 | 0.4634928 |
| $C$1:$D$2 | 0.0892019 | 0.5000000 | 0.4107981 | -0.0789474 | 0.3845454 | 0.4634928 |
| $A$3:$D$4 | -0.0985915 | 0.3364706 | 0.4350621 | -0.0368421 | 0.4231818 | 0.4600239 |
| $A$6:$D$7 | -0.0985915 | 0.3364706 | 0.4350621 | 0.0263158 | 0.4809091 | 0.4545933 |
| $A$9:$D$10 | -0.0985915 | 0.3364706 | 0.4350621 | 0.0894737 | 0.5386364 | 0.4491627 |
| $A$12 | -0.0985915 | 0.3364706 | 0.4350621 | 0.1526316 | 0.5963637 | 0.4437321 |

### Hypothesis 3: Page-Scaled Coordinates

```
DB_Left = COM_Left / PageWidth
DB_Top  = COM_Top / PageHeight
```

| $A$1:$B$2 | 0.0000000 | 0.3364706 | 0.3364706 | 0.0000000 | 0.3845454 | 0.3845454 |
| $C$1:$D$2 | 0.1568627 | 0.5000000 | 0.3431373 | 0.0000000 | 0.3845454 | 0.3845454 |
| $A$3:$D$4 | 0.0000000 | 0.3364706 | 0.3364706 | 0.0363636 | 0.4231818 | 0.3868182 |
| $A$6:$D$7 | 0.0000000 | 0.3364706 | 0.3364706 | 0.0909091 | 0.4809091 | 0.3900000 |
| $A$9:$D$10 | 0.0000000 | 0.3364706 | 0.3364706 | 0.1454545 | 0.5386364 | 0.3931819 |
| $A$12 | 0.0000000 | 0.3364706 | 0.3364706 | 0.2000000 | 0.5963637 | 0.3963637 |

### Hypothesis 4: Scaled Printable Coordinates

```
scale = Zoom / 100  (if Zoom is set)
scale = 1 / FitToPagesWide  (if FitToPages is set)
DB_Left = (COM_Left * scale - LeftMargin) / (PageWidth * scale - LeftMargin - RightMargin)
```

| $A$1:$B$2 | -0.0985915 | 0.3364706 | 0.4350621 | -0.0789474 | 0.3845454 | 0.4634928 |
| $C$1:$D$2 | 0.0892019 | 0.5000000 | 0.4107981 | -0.0789474 | 0.3845454 | 0.4634928 |
| $A$3:$D$4 | -0.0985915 | 0.3364706 | 0.4350621 | -0.0368421 | 0.4231818 | 0.4600239 |
| $A$6:$D$7 | -0.0985915 | 0.3364706 | 0.4350621 | 0.0263158 | 0.4809091 | 0.4545933 |
| $A$9:$D$10 | -0.0985915 | 0.3364706 | 0.4350621 | 0.0894737 | 0.5386364 | 0.4491627 |
| $A$12 | -0.0985915 | 0.3364706 | 0.4350621 | 0.1526316 | 0.5963637 | 0.4437321 |

### Hypothesis 5: Printed Page Coordinates (final rendered)

After all Excel layout is applied, Range.Left returns the position on the
printed page (in points). The ratio is simply:

```
DB_Left = COM_Left / PageWidth
DB_Top  = COM_Top / PageHeight
```

This is the simplest hypothesis and matches Template 546.


## Best Hypothesis Determination

- Printable: total error = 5.3206062
- CenteredPrintable: total error = 5.3206062
- PageScaled: total error = 4.3609449
- ScaledPrintable: total error = 5994.0000000
