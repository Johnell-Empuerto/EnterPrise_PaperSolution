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

| $I$6:$M$6 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0175439 | 0.1654546 | 0.1479107 |
| $I$7:$M$7 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0368421 | 0.1877273 | 0.1508852 |
| $I$8:$M$8 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0561404 | 0.2104545 | 0.1543141 |
| $I$9:$M$9 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0754386 | 0.2336364 | 0.1581978 |
| $I$10:$M$10 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0947368 | 0.2568182 | 0.1620814 |

### Hypothesis 2: Centered Printable Coordinates

```
if CenterHorizontally:
  PrintedOriginX = (PageWidth - PrintableWidth) / 2
else:
  PrintedOriginX = LeftMargin
DB_Left = (COM_Left - PrintedOriginX) / PrintableWidth
```

| $I$6:$M$6 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0175439 | 0.1654546 | 0.1479107 |
| $I$7:$M$7 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0368421 | 0.1877273 | 0.1508852 |
| $I$8:$M$8 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0561404 | 0.2104545 | 0.1543141 |
| $I$9:$M$9 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0754386 | 0.2336364 | 0.1581978 |
| $I$10:$M$10 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0947368 | 0.2568182 | 0.1620814 |

### Hypothesis 3: Page-Scaled Coordinates

```
DB_Left = COM_Left / PageWidth
DB_Top  = COM_Top / PageHeight
```

| $I$6:$M$6 | 0.6274510 | 0.5294118 | 0.0980392 | 0.0833333 | 0.1654546 | 0.0821213 |
| $I$7:$M$7 | 0.6274510 | 0.5294118 | 0.0980392 | 0.1000000 | 0.1877273 | 0.0877273 |
| $I$8:$M$8 | 0.6274510 | 0.5294118 | 0.0980392 | 0.1166667 | 0.2104545 | 0.0937878 |
| $I$9:$M$9 | 0.6274510 | 0.5294118 | 0.0980392 | 0.1333333 | 0.2336364 | 0.1003031 |
| $I$10:$M$10 | 0.6274510 | 0.5294118 | 0.0980392 | 0.1500000 | 0.2568182 | 0.1068182 |

### Hypothesis 4: Scaled Printable Coordinates

```
scale = Zoom / 100  (if Zoom is set)
scale = 1 / FitToPagesWide  (if FitToPages is set)
DB_Left = (COM_Left * scale - LeftMargin) / (PageWidth * scale - LeftMargin - RightMargin)
```

| $I$6:$M$6 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0175439 | 0.1654546 | 0.1479107 |
| $I$7:$M$7 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0368421 | 0.1877273 | 0.1508852 |
| $I$8:$M$8 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0561404 | 0.2104545 | 0.1543141 |
| $I$9:$M$9 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0754386 | 0.2336364 | 0.1581978 |
| $I$10:$M$10 | 0.6525822 | 0.5294118 | 0.1231704 | 0.0947368 | 0.2568182 | 0.1620814 |

### Hypothesis 5: Printed Page Coordinates (final rendered)

After all Excel layout is applied, Range.Left returns the position on the
printed page (in points). The ratio is simply:

```
DB_Left = COM_Left / PageWidth
DB_Top  = COM_Top / PageHeight
```

This is the simplest hypothesis and matches Template 546.


## Best Hypothesis Determination

- Printable: total error = 1.3892412
- CenteredPrintable: total error = 1.3892412
- PageScaled: total error = 0.9609537
- ScaledPrintable: total error = 4995.0000000
