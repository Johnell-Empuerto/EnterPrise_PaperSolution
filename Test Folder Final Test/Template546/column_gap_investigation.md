# Template 546 — Column Coordinate Investigation

## Column Dump

| Col | Letter | OpenXML Width(chars) | COM ColumnWidth(chars) | COM Range.Left(pt) | COM Range.Width(pt) | COM EntireCol.Left(pt) | COM EntireCol.Width(pt) |
|-----|--------|---------------------|----------------------|-------------------|--------------------|----------------------|-----------------------|
| 1 | A | (default 8.43) | 8.110000 | 0.00 | 48.00 | 0.00 | 48.00 |
| 2 | B | (default 8.43) | 8.110000 | 48.00 | 48.00 | 48.00 | 48.00 |
| 3 | C | (default 8.43) | 8.110000 | 96.00 | 48.00 | 96.00 | 48.00 |
| 4 | D | (default 8.43) | 8.110000 | 144.00 | 48.00 | 144.00 | 48.00 |
| 5 | E | (default 8.43) | 8.110000 | 192.00 | 48.00 | 192.00 | 48.00 |
| 6 | F | (default 8.43) | 8.110000 | 240.00 | 48.00 | 240.00 | 48.00 |
| 7 | G | (default 8.43) | 8.110000 | 288.00 | 48.00 | 288.00 | 48.00 |
| 8 | H | (default 8.43) | 8.110000 | 336.00 | 48.00 | 336.00 | 48.00 |
| 9 | I | (default 8.43) | 8.110000 | 384.00 | 48.00 | 384.00 | 48.00 |
| 10 | J | (default 8.43) | 8.110000 | 432.00 | 48.00 | 432.00 | 48.00 |

## Column Width Conversions

Excel's internal column width conversion formula:
- Character width → points: `points = (chars * 7 + 5) / 7 * font_width / 7 * 4`
- Or more simply: points = chars * (max_digit_width * 256 + 12) / 256

| Col | OpenXML chars | COM chars | COM Left | COM Width | OpenXML pts (50.04/8.43) | Derived factor |
|-----|--------------|-----------|----------|-----------|------------------------|----------------|
| A | 8.430000 | 8.110000 | 0.00 | 48.00 | 50.0400 | 5.918619 |
| B | 8.430000 | 8.110000 | 48.00 | 48.00 | 50.0400 | 5.918619 |
| C | 8.430000 | 8.110000 | 96.00 | 48.00 | 50.0400 | 5.918619 |
| D | 8.430000 | 8.110000 | 144.00 | 48.00 | 50.0400 | 5.918619 |
| E | 8.430000 | 8.110000 | 192.00 | 48.00 | 50.0400 | 5.918619 |
| F | 8.430000 | 8.110000 | 240.00 | 48.00 | 50.0400 | 5.918619 |
| G | 8.430000 | 8.110000 | 288.00 | 48.00 | 50.0400 | 5.918619 |
| H | 8.430000 | 8.110000 | 336.00 | 48.00 | 50.0400 | 5.918619 |
| I | 8.430000 | 8.110000 | 384.00 | 48.00 | 50.0400 | 5.918619 |
| J | 8.430000 | 8.110000 | 432.00 | 48.00 | 50.0400 | 5.918619 |

## Cumulative Left Position: COM vs OpenXML

| Col | COM Left | OpenXML Cumulative | Difference |
|-----|----------|-------------------|------------|
| A | 0.0000 | 0.0000 | +0.0000 |
| B | 48.0000 | 50.0400 | -2.0400 |
| C | 96.0000 | 100.0800 | -4.0800 |
| D | 144.0000 | 150.1200 | -6.1200 |
| E | 192.0000 | 200.1600 | -8.1600 |
| F | 240.0000 | 250.2000 | -10.2000 |
| G | 288.0000 | 300.2400 | -12.2400 |
| H | 336.0000 | 350.2800 | -14.2800 |
| I | 384.0000 | 400.3200 | -16.3200 |
| J | 432.0000 | 450.3600 | -18.3600 |

## Column C (Index 3) — Root Cause Analysis

COM Range.Left for Column C: 96.0000 pt
OpenXML cumulative (A+B): 100.0800 pt
Difference: -4.0800 pt

### Possible Sources
1. Gridline width: ~-2.0400 pt per gridline between A-B and B-C
2. Cell padding (left + right): -4.0800 pt
3. Border thickness: check if columns have borders
4. Font metric conversion: COM vs OpenXML character width formula difference
