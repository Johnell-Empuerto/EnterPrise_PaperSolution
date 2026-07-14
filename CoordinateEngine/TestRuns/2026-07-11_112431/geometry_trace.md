# Geometry Trace Report

## Summary
- Workbook: original
- Page: 612 × 792 pt
- Fields: 23
- DPI: 300

## Formula
```
range_left   = Σ column_widths(1 .. col-1)
range_top    = Σ row_heights(1 .. row-1)
range_width  = Σ column_widths(col .. col_end)
range_height = Σ row_heights(row .. row_end)

page_left    = origin_x + range_left  × scale_w
page_top     = origin_y + range_top   × scale_h
page_width   = range_width  × scale_w
page_height  = range_height × scale_h

pixel_left   = page_left   × (DPI / 72)
pixel_top    = page_top    × (DPI / 72)
pixel_width  = page_width  × (DPI / 72)
pixel_height = page_height × (DPI / 72)

left_ratio   = pixel_left   / page_width_px
top_ratio    = pixel_top    / page_height_px
width_ratio  = pixel_width  / page_width_px
height_ratio = pixel_height / page_height_px
```

## A1
**Merged:** A1:B2

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 0.0000 | Workbook XML → column width sum |
| Range Top | 0.0000 | Workbook XML → row height sum |
| Range Width | 100.1878 | Workbook XML → column width sum |
| Range Height | 30.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 204.7700 | origin_x + range_left × scale_w |
| Page Top | 303.6500 | origin_y + range_top × scale_h |
| Page Width | 101.2300 | range_width × scale_w |
| Page Height | 30.7833 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 853.21 | page_left × (DPI/72) |
| Pixel Top | 1265.21 | page_top × (DPI/72) |
| Pixel Width | 421.79 | page_width × (DPI/72) |
| Pixel Height | 128.26 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.334592 |
| Top Ratio | 0.383396 |
| Width Ratio | 0.165408 |
| Height Ratio | 0.038868 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field A1 | col=1, row=1, col_end=2, row_end=2 | A1 | cell | workbook XML |
| 2 | Col 1 width | col_widths_pt[1] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Col 2 width | col_widths_pt[2] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 4 | Row 1 height | row_heights_pt[1] = 15.0000 | 15.0 | pt | workbook XML |
| 5 | Row 2 height | row_heights_pt[2] = 15.0000 | 15.0 | pt | workbook XML |
| 6 | Range left = Σ col widths before field | Σ cols 1..0 = 0 (first column) | 0 | pt | workbook XML → computed |
| 7 | Range top = Σ row heights before field | Σ rows 1..0 = 0 (first row) | 0 | pt | workbook XML → computed |
| 8 | Range width = Σ col widths of field | Σ cols 1..2 = 50.0939 + 50.0939 | 100.18785 | pt | workbook XML → computed |
| 9 | Range height = Σ row heights of field | Σ rows 1..2 = 15.0000 + 15.0000 | 30.0 | pt | workbook XML → computed |
| 10 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 11 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 12 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 13 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 14 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 15 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 16 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 17 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 18 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 19 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 20 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 21 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 22 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 0.0000 * 1.010402 | 204.770004 | pt | computed |
| 23 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 0.0000 * 1.026111 | 303.649994 | pt | computed |
| 24 | Page width = range_width × scale_w | 100.1878 × 1.010402 | 101.229996 | pt | computed |
| 25 | Page height = range_height × scale_h | 30.0000 × 1.026111 | 30.783335 | pt | computed |
| 26 | Pixel left = page_left × (dpi/72) | 204.7700 × 4.166667 | 853.21 | px | computed |
| 27 | Pixel top = page_top × (dpi/72) | 303.6500 × 4.166667 | 1265.21 | px | computed |
| 28 | Pixel width = page_width × (dpi/72) | 101.2300 × 4.166667 | 421.79 | px | computed |
| 29 | Pixel height = page_height × (dpi/72) | 30.7833 × 4.166667 | 128.26 | px | computed |
| 30 | Left ratio = pixel_left / page_width_px | 853.21 / 2550.00 | 0.334592 |  | computed |
| 31 | Top ratio = pixel_top / page_height_px | 1265.21 / page_height | 0.383396 |  | computed |

## C1
**Merged:** C1:D2

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 100.1878 | Workbook XML → column width sum |
| Range Top | 0.0000 | Workbook XML → row height sum |
| Range Width | 100.1878 | Workbook XML → column width sum |
| Range Height | 30.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 306.0000 | origin_x + range_left × scale_w |
| Page Top | 303.6500 | origin_y + range_top × scale_h |
| Page Width | 101.2300 | range_width × scale_w |
| Page Height | 30.7833 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1275.00 | page_left × (DPI/72) |
| Pixel Top | 1265.21 | page_top × (DPI/72) |
| Pixel Width | 421.79 | page_width × (DPI/72) |
| Pixel Height | 128.26 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.500000 |
| Top Ratio | 0.383396 |
| Width Ratio | 0.165408 |
| Height Ratio | 0.038868 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field C1 | col=3, row=1, col_end=4, row_end=2 | C1 | cell | workbook XML |
| 2 | Col 3 width | col_widths_pt[3] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Col 4 width | col_widths_pt[4] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 4 | Row 1 height | row_heights_pt[1] = 15.0000 | 15.0 | pt | workbook XML |
| 5 | Row 2 height | row_heights_pt[2] = 15.0000 | 15.0 | pt | workbook XML |
| 6 | Range left = Σ col widths before field | Σ cols 1..2 = 50.0939 + 50.0939 | 100.18785 | pt | workbook XML → computed |
| 7 | Range top = Σ row heights before field | Σ rows 1..0 = 0 (first row) | 0 | pt | workbook XML → computed |
| 8 | Range width = Σ col widths of field | Σ cols 3..4 = 50.0939 + 50.0939 | 100.18785 | pt | workbook XML → computed |
| 9 | Range height = Σ row heights of field | Σ rows 1..2 = 15.0000 + 15.0000 | 30.0 | pt | workbook XML → computed |
| 10 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 11 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 12 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 13 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 14 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 15 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 16 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 17 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 18 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 19 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 20 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 21 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 22 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 100.1878 * 1.010402 | 306.0 | pt | computed |
| 23 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 0.0000 * 1.026111 | 303.649994 | pt | computed |
| 24 | Page width = range_width × scale_w | 100.1878 × 1.010402 | 101.229996 | pt | computed |
| 25 | Page height = range_height × scale_h | 30.0000 × 1.026111 | 30.783335 | pt | computed |
| 26 | Pixel left = page_left × (dpi/72) | 306.0000 × 4.166667 | 1275.0 | px | computed |
| 27 | Pixel top = page_top × (dpi/72) | 303.6500 × 4.166667 | 1265.21 | px | computed |
| 28 | Pixel width = page_width × (dpi/72) | 101.2300 × 4.166667 | 421.79 | px | computed |
| 29 | Pixel height = page_height × (dpi/72) | 30.7833 × 4.166667 | 128.26 | px | computed |
| 30 | Left ratio = pixel_left / page_width_px | 1275.00 / 2550.00 | 0.5 |  | computed |
| 31 | Top ratio = pixel_top / page_height_px | 1265.21 / page_height | 0.383396 |  | computed |

## E1

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 200.3757 | Workbook XML → column width sum |
| Range Top | 0.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 407.2300 | origin_x + range_left × scale_w |
| Page Top | 303.6500 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1696.79 | page_left × (DPI/72) |
| Pixel Top | 1265.21 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.665408 |
| Top Ratio | 0.383396 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field E1 | col=5, row=1, col_end=5, row_end=1 | E1 | cell | workbook XML |
| 2 | Col 5 width | col_widths_pt[5] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 1 height | row_heights_pt[1] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..4 = 50.0939 + 50.0939 + 50.0939 + 50.0939 | 200.3757 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..0 = 0 (first row) | 0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 5..5 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 1..1 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 200.3757 * 1.010402 | 407.229996 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 0.0000 * 1.026111 | 303.649994 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 407.2300 × 4.166667 | 1696.79 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 303.6500 × 4.166667 | 1265.21 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1696.79 / 2550.00 | 0.665408 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1265.21 / page_height | 0.383396 |  | computed |

## F1

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 250.4696 | Workbook XML → column width sum |
| Range Top | 0.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 457.8450 | origin_x + range_left × scale_w |
| Page Top | 303.6500 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1907.69 | page_left × (DPI/72) |
| Pixel Top | 1265.21 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.748113 |
| Top Ratio | 0.383396 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field F1 | col=6, row=1, col_end=6, row_end=1 | F1 | cell | workbook XML |
| 2 | Col 6 width | col_widths_pt[6] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 1 height | row_heights_pt[1] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..5 = 50.0939 + 50.0939 + 50.0939 + 50.0939 + 50.0939 | 250.469625 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..0 = 0 (first row) | 0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 6..6 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 1..1 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 250.4696 * 1.010402 | 457.844994 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 0.0000 * 1.026111 | 303.649994 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 457.8450 × 4.166667 | 1907.69 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 303.6500 × 4.166667 | 1265.21 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1907.69 / 2550.00 | 0.748113 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1265.21 / page_height | 0.383396 |  | computed |

## G1

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 300.5635 | Workbook XML → column width sum |
| Range Top | 0.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 508.4600 | origin_x + range_left × scale_w |
| Page Top | 303.6500 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 2118.58 | page_left × (DPI/72) |
| Pixel Top | 1265.21 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.830817 |
| Top Ratio | 0.383396 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field G1 | col=7, row=1, col_end=7, row_end=1 | G1 | cell | workbook XML |
| 2 | Col 7 width | col_widths_pt[7] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 1 height | row_heights_pt[1] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..6 = 50.0939 + 50.0939 + 50.0939 + 50.0939 + 50.0939 + 50.0939 | 300.56355 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..0 = 0 (first row) | 0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 7..7 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 1..1 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 300.5635 * 1.010402 | 508.459991 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 0.0000 * 1.026111 | 303.649994 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 508.4600 × 4.166667 | 2118.58 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 303.6500 × 4.166667 | 1265.21 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 2118.58 / 2550.00 | 0.830817 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1265.21 / page_height | 0.383396 |  | computed |

## E2

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 200.3757 | Workbook XML → column width sum |
| Range Top | 15.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 407.2300 | origin_x + range_left × scale_w |
| Page Top | 319.0417 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1696.79 | page_left × (DPI/72) |
| Pixel Top | 1329.34 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.665408 |
| Top Ratio | 0.402830 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field E2 | col=5, row=2, col_end=5, row_end=2 | E2 | cell | workbook XML |
| 2 | Col 5 width | col_widths_pt[5] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 2 height | row_heights_pt[2] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..4 = 50.0939 + 50.0939 + 50.0939 + 50.0939 | 200.3757 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..1 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 5..5 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 2..2 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 200.3757 * 1.010402 | 407.229996 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 15.0000 * 1.026111 | 319.041662 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 407.2300 × 4.166667 | 1696.79 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 319.0417 × 4.166667 | 1329.34 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1696.79 / 2550.00 | 0.665408 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1329.34 / page_height | 0.40283 |  | computed |

## F2

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 250.4696 | Workbook XML → column width sum |
| Range Top | 15.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 457.8450 | origin_x + range_left × scale_w |
| Page Top | 319.0417 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1907.69 | page_left × (DPI/72) |
| Pixel Top | 1329.34 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.748113 |
| Top Ratio | 0.402830 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field F2 | col=6, row=2, col_end=6, row_end=2 | F2 | cell | workbook XML |
| 2 | Col 6 width | col_widths_pt[6] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 2 height | row_heights_pt[2] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..5 = 50.0939 + 50.0939 + 50.0939 + 50.0939 + 50.0939 | 250.469625 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..1 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 6..6 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 2..2 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 250.4696 * 1.010402 | 457.844994 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 15.0000 * 1.026111 | 319.041662 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 457.8450 × 4.166667 | 1907.69 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 319.0417 × 4.166667 | 1329.34 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1907.69 / 2550.00 | 0.748113 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1329.34 / page_height | 0.40283 |  | computed |

## A3
**Merged:** A3:D4

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 0.0000 | Workbook XML → column width sum |
| Range Top | 30.0000 | Workbook XML → row height sum |
| Range Width | 200.3757 | Workbook XML → column width sum |
| Range Height | 30.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 204.7700 | origin_x + range_left × scale_w |
| Page Top | 334.4333 | origin_y + range_top × scale_h |
| Page Width | 202.4600 | range_width × scale_w |
| Page Height | 30.7833 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 853.21 | page_left × (DPI/72) |
| Pixel Top | 1393.47 | page_top × (DPI/72) |
| Pixel Width | 843.58 | page_width × (DPI/72) |
| Pixel Height | 128.26 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.334592 |
| Top Ratio | 0.422264 |
| Width Ratio | 0.330817 |
| Height Ratio | 0.038868 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field A3 | col=1, row=3, col_end=4, row_end=4 | A3 | cell | workbook XML |
| 2 | Col 1 width | col_widths_pt[1] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Col 2 width | col_widths_pt[2] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 4 | Col 3 width | col_widths_pt[3] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 5 | Col 4 width | col_widths_pt[4] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 6 | Row 3 height | row_heights_pt[3] = 15.0000 | 15.0 | pt | workbook XML |
| 7 | Row 4 height | row_heights_pt[4] = 15.0000 | 15.0 | pt | workbook XML |
| 8 | Range left = Σ col widths before field | Σ cols 1..0 = 0 (first column) | 0 | pt | workbook XML → computed |
| 9 | Range top = Σ row heights before field | Σ rows 1..2 = 15.0000 + 15.0000 | 30.0 | pt | workbook XML → computed |
| 10 | Range width = Σ col widths of field | Σ cols 1..4 = 50.0939 + 50.0939 + 50.0939 + 50.0939 | 200.3757 | pt | workbook XML → computed |
| 11 | Range height = Σ row heights of field | Σ rows 3..4 = 15.0000 + 15.0000 | 30.0 | pt | workbook XML → computed |
| 12 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 13 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 14 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 15 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 16 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 17 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 18 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 19 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 20 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 21 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 22 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 23 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 24 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 0.0000 * 1.010402 | 204.770004 | pt | computed |
| 25 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 30.0000 * 1.026111 | 334.433329 | pt | computed |
| 26 | Page width = range_width × scale_w | 200.3757 × 1.010402 | 202.459991 | pt | computed |
| 27 | Page height = range_height × scale_h | 30.0000 × 1.026111 | 30.783335 | pt | computed |
| 28 | Pixel left = page_left × (dpi/72) | 204.7700 × 4.166667 | 853.21 | px | computed |
| 29 | Pixel top = page_top × (dpi/72) | 334.4333 × 4.166667 | 1393.47 | px | computed |
| 30 | Pixel width = page_width × (dpi/72) | 202.4600 × 4.166667 | 843.58 | px | computed |
| 31 | Pixel height = page_height × (dpi/72) | 30.7833 × 4.166667 | 128.26 | px | computed |
| 32 | Left ratio = pixel_left / page_width_px | 853.21 / 2550.00 | 0.334592 |  | computed |
| 33 | Top ratio = pixel_top / page_height_px | 1393.47 / page_height | 0.422264 |  | computed |

## E3

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 200.3757 | Workbook XML → column width sum |
| Range Top | 30.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 407.2300 | origin_x + range_left × scale_w |
| Page Top | 334.4333 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1696.79 | page_left × (DPI/72) |
| Pixel Top | 1393.47 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.665408 |
| Top Ratio | 0.422264 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field E3 | col=5, row=3, col_end=5, row_end=3 | E3 | cell | workbook XML |
| 2 | Col 5 width | col_widths_pt[5] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 3 height | row_heights_pt[3] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..4 = 50.0939 + 50.0939 + 50.0939 + 50.0939 | 200.3757 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..2 = 15.0000 + 15.0000 | 30.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 5..5 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 3..3 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 200.3757 * 1.010402 | 407.229996 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 30.0000 * 1.026111 | 334.433329 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 407.2300 × 4.166667 | 1696.79 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 334.4333 × 4.166667 | 1393.47 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1696.79 / 2550.00 | 0.665408 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1393.47 / page_height | 0.422264 |  | computed |

## F3

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 250.4696 | Workbook XML → column width sum |
| Range Top | 30.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 457.8450 | origin_x + range_left × scale_w |
| Page Top | 334.4333 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1907.69 | page_left × (DPI/72) |
| Pixel Top | 1393.47 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.748113 |
| Top Ratio | 0.422264 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field F3 | col=6, row=3, col_end=6, row_end=3 | F3 | cell | workbook XML |
| 2 | Col 6 width | col_widths_pt[6] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 3 height | row_heights_pt[3] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..5 = 50.0939 + 50.0939 + 50.0939 + 50.0939 + 50.0939 | 250.469625 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..2 = 15.0000 + 15.0000 | 30.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 6..6 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 3..3 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 250.4696 * 1.010402 | 457.844994 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 30.0000 * 1.026111 | 334.433329 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 457.8450 × 4.166667 | 1907.69 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 334.4333 × 4.166667 | 1393.47 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1907.69 / 2550.00 | 0.748113 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1393.47 / page_height | 0.422264 |  | computed |

## E4

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 200.3757 | Workbook XML → column width sum |
| Range Top | 45.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 407.2300 | origin_x + range_left × scale_w |
| Page Top | 349.8250 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1696.79 | page_left × (DPI/72) |
| Pixel Top | 1457.60 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.665408 |
| Top Ratio | 0.441698 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field E4 | col=5, row=4, col_end=5, row_end=4 | E4 | cell | workbook XML |
| 2 | Col 5 width | col_widths_pt[5] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 4 height | row_heights_pt[4] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..4 = 50.0939 + 50.0939 + 50.0939 + 50.0939 | 200.3757 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..3 = 15.0000 + 15.0000 + 15.0000 | 45.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 5..5 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 4..4 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 200.3757 * 1.010402 | 407.229996 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 45.0000 * 1.026111 | 349.824997 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 407.2300 × 4.166667 | 1696.79 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 349.8250 × 4.166667 | 1457.6 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1696.79 / 2550.00 | 0.665408 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1457.60 / page_height | 0.441698 |  | computed |

## F4

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 250.4696 | Workbook XML → column width sum |
| Range Top | 45.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 457.8450 | origin_x + range_left × scale_w |
| Page Top | 349.8250 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1907.69 | page_left × (DPI/72) |
| Pixel Top | 1457.60 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.748113 |
| Top Ratio | 0.441698 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field F4 | col=6, row=4, col_end=6, row_end=4 | F4 | cell | workbook XML |
| 2 | Col 6 width | col_widths_pt[6] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 4 height | row_heights_pt[4] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..5 = 50.0939 + 50.0939 + 50.0939 + 50.0939 + 50.0939 | 250.469625 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..3 = 15.0000 + 15.0000 + 15.0000 | 45.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 6..6 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 4..4 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 250.4696 * 1.010402 | 457.844994 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 45.0000 * 1.026111 | 349.824997 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 457.8450 × 4.166667 | 1907.69 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 349.8250 × 4.166667 | 1457.6 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1907.69 / 2550.00 | 0.748113 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1457.60 / page_height | 0.441698 |  | computed |

## A5

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 0.0000 | Workbook XML → column width sum |
| Range Top | 60.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 204.7700 | origin_x + range_left × scale_w |
| Page Top | 365.2167 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 853.21 | page_left × (DPI/72) |
| Pixel Top | 1521.74 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.334592 |
| Top Ratio | 0.461132 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field A5 | col=1, row=5, col_end=1, row_end=5 | A5 | cell | workbook XML |
| 2 | Col 1 width | col_widths_pt[1] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 5 height | row_heights_pt[5] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..0 = 0 (first column) | 0 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..4 = 15.0000 + 15.0000 + 15.0000 + 15.0000 | 60.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 1..1 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 5..5 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 0.0000 * 1.010402 | 204.770004 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 60.0000 * 1.026111 | 365.216665 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 204.7700 × 4.166667 | 853.21 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 365.2167 × 4.166667 | 1521.74 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 853.21 / 2550.00 | 0.334592 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1521.74 / page_height | 0.461132 |  | computed |

## B5

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 50.0939 | Workbook XML → column width sum |
| Range Top | 60.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 255.3850 | origin_x + range_left × scale_w |
| Page Top | 365.2167 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1064.10 | page_left × (DPI/72) |
| Pixel Top | 1521.74 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.417296 |
| Top Ratio | 0.461132 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field B5 | col=2, row=5, col_end=2, row_end=5 | B5 | cell | workbook XML |
| 2 | Col 2 width | col_widths_pt[2] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 5 height | row_heights_pt[5] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..1 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..4 = 15.0000 + 15.0000 + 15.0000 + 15.0000 | 60.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 2..2 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 5..5 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 50.0939 * 1.010402 | 255.385002 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 60.0000 * 1.026111 | 365.216665 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 255.3850 × 4.166667 | 1064.1 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 365.2167 × 4.166667 | 1521.74 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1064.10 / 2549.99 | 0.417296 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1521.74 / page_height | 0.461132 |  | computed |

## C5

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 100.1878 | Workbook XML → column width sum |
| Range Top | 60.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 306.0000 | origin_x + range_left × scale_w |
| Page Top | 365.2167 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1275.00 | page_left × (DPI/72) |
| Pixel Top | 1521.74 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.500000 |
| Top Ratio | 0.461132 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field C5 | col=3, row=5, col_end=3, row_end=5 | C5 | cell | workbook XML |
| 2 | Col 3 width | col_widths_pt[3] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 5 height | row_heights_pt[5] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..2 = 50.0939 + 50.0939 | 100.18785 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..4 = 15.0000 + 15.0000 + 15.0000 + 15.0000 | 60.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 3..3 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 5..5 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 100.1878 * 1.010402 | 306.0 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 60.0000 * 1.026111 | 365.216665 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 306.0000 × 4.166667 | 1275.0 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 365.2167 × 4.166667 | 1521.74 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1275.00 / 2550.00 | 0.5 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1521.74 / page_height | 0.461132 |  | computed |

## D5

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 150.2818 | Workbook XML → column width sum |
| Range Top | 60.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 356.6150 | origin_x + range_left × scale_w |
| Page Top | 365.2167 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1485.90 | page_left × (DPI/72) |
| Pixel Top | 1521.74 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.582704 |
| Top Ratio | 0.461132 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field D5 | col=4, row=5, col_end=4, row_end=5 | D5 | cell | workbook XML |
| 2 | Col 4 width | col_widths_pt[4] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 5 height | row_heights_pt[5] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..3 = 50.0939 + 50.0939 + 50.0939 | 150.281775 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..4 = 15.0000 + 15.0000 + 15.0000 + 15.0000 | 60.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 4..4 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 5..5 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 150.2818 * 1.010402 | 356.614998 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 60.0000 * 1.026111 | 365.216665 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 356.6150 × 4.166667 | 1485.9 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 365.2167 × 4.166667 | 1521.74 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1485.90 / 2550.01 | 0.582704 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1521.74 / page_height | 0.461132 |  | computed |

## E5

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 200.3757 | Workbook XML → column width sum |
| Range Top | 60.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 407.2300 | origin_x + range_left × scale_w |
| Page Top | 365.2167 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1696.79 | page_left × (DPI/72) |
| Pixel Top | 1521.74 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.665408 |
| Top Ratio | 0.461132 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field E5 | col=5, row=5, col_end=5, row_end=5 | E5 | cell | workbook XML |
| 2 | Col 5 width | col_widths_pt[5] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 5 height | row_heights_pt[5] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..4 = 50.0939 + 50.0939 + 50.0939 + 50.0939 | 200.3757 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..4 = 15.0000 + 15.0000 + 15.0000 + 15.0000 | 60.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 5..5 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 5..5 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 200.3757 * 1.010402 | 407.229996 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 60.0000 * 1.026111 | 365.216665 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 407.2300 × 4.166667 | 1696.79 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 365.2167 × 4.166667 | 1521.74 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1696.79 / 2550.00 | 0.665408 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1521.74 / page_height | 0.461132 |  | computed |

## F5

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 250.4696 | Workbook XML → column width sum |
| Range Top | 60.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 457.8450 | origin_x + range_left × scale_w |
| Page Top | 365.2167 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1907.69 | page_left × (DPI/72) |
| Pixel Top | 1521.74 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.748113 |
| Top Ratio | 0.461132 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field F5 | col=6, row=5, col_end=6, row_end=5 | F5 | cell | workbook XML |
| 2 | Col 6 width | col_widths_pt[6] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 5 height | row_heights_pt[5] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..5 = 50.0939 + 50.0939 + 50.0939 + 50.0939 + 50.0939 | 250.469625 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..4 = 15.0000 + 15.0000 + 15.0000 + 15.0000 | 60.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 6..6 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 5..5 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 250.4696 * 1.010402 | 457.844994 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 60.0000 * 1.026111 | 365.216665 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 457.8450 × 4.166667 | 1907.69 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 365.2167 × 4.166667 | 1521.74 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1907.69 / 2550.00 | 0.748113 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1521.74 / page_height | 0.461132 |  | computed |

## A6
**Merged:** A6:D7

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 0.0000 | Workbook XML → column width sum |
| Range Top | 75.0000 | Workbook XML → row height sum |
| Range Width | 200.3757 | Workbook XML → column width sum |
| Range Height | 30.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 204.7700 | origin_x + range_left × scale_w |
| Page Top | 380.6083 | origin_y + range_top × scale_h |
| Page Width | 202.4600 | range_width × scale_w |
| Page Height | 30.7833 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 853.21 | page_left × (DPI/72) |
| Pixel Top | 1585.87 | page_top × (DPI/72) |
| Pixel Width | 843.58 | page_width × (DPI/72) |
| Pixel Height | 128.26 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.334592 |
| Top Ratio | 0.480566 |
| Width Ratio | 0.330817 |
| Height Ratio | 0.038868 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field A6 | col=1, row=6, col_end=4, row_end=7 | A6 | cell | workbook XML |
| 2 | Col 1 width | col_widths_pt[1] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Col 2 width | col_widths_pt[2] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 4 | Col 3 width | col_widths_pt[3] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 5 | Col 4 width | col_widths_pt[4] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 6 | Row 6 height | row_heights_pt[6] = 15.0000 | 15.0 | pt | workbook XML |
| 7 | Row 7 height | row_heights_pt[7] = 15.0000 | 15.0 | pt | workbook XML |
| 8 | Range left = Σ col widths before field | Σ cols 1..0 = 0 (first column) | 0 | pt | workbook XML → computed |
| 9 | Range top = Σ row heights before field | Σ rows 1..5 = 15.0000 + 15.0000 + 15.0000 + 15.0000 + 15.0000 | 75.0 | pt | workbook XML → computed |
| 10 | Range width = Σ col widths of field | Σ cols 1..4 = 50.0939 + 50.0939 + 50.0939 + 50.0939 | 200.3757 | pt | workbook XML → computed |
| 11 | Range height = Σ row heights of field | Σ rows 6..7 = 15.0000 + 15.0000 | 30.0 | pt | workbook XML → computed |
| 12 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 13 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 14 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 15 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 16 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 17 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 18 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 19 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 20 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 21 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 22 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 23 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 24 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 0.0000 * 1.010402 | 204.770004 | pt | computed |
| 25 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 75.0000 * 1.026111 | 380.608332 | pt | computed |
| 26 | Page width = range_width × scale_w | 200.3757 × 1.010402 | 202.459991 | pt | computed |
| 27 | Page height = range_height × scale_h | 30.0000 × 1.026111 | 30.783335 | pt | computed |
| 28 | Pixel left = page_left × (dpi/72) | 204.7700 × 4.166667 | 853.21 | px | computed |
| 29 | Pixel top = page_top × (dpi/72) | 380.6083 × 4.166667 | 1585.87 | px | computed |
| 30 | Pixel width = page_width × (dpi/72) | 202.4600 × 4.166667 | 843.58 | px | computed |
| 31 | Pixel height = page_height × (dpi/72) | 30.7833 × 4.166667 | 128.26 | px | computed |
| 32 | Left ratio = pixel_left / page_width_px | 853.21 / 2550.00 | 0.334592 |  | computed |
| 33 | Top ratio = pixel_top / page_height_px | 1585.87 / page_height | 0.480566 |  | computed |

## E6

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 200.3757 | Workbook XML → column width sum |
| Range Top | 75.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 407.2300 | origin_x + range_left × scale_w |
| Page Top | 380.6083 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1696.79 | page_left × (DPI/72) |
| Pixel Top | 1585.87 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.665408 |
| Top Ratio | 0.480566 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field E6 | col=5, row=6, col_end=5, row_end=6 | E6 | cell | workbook XML |
| 2 | Col 5 width | col_widths_pt[5] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 6 height | row_heights_pt[6] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..4 = 50.0939 + 50.0939 + 50.0939 + 50.0939 | 200.3757 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..5 = 15.0000 + 15.0000 + 15.0000 + 15.0000 + 15.0000 | 75.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 5..5 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 6..6 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 200.3757 * 1.010402 | 407.229996 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 75.0000 * 1.026111 | 380.608332 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 407.2300 × 4.166667 | 1696.79 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 380.6083 × 4.166667 | 1585.87 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1696.79 / 2550.00 | 0.665408 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1585.87 / page_height | 0.480566 |  | computed |

## F6

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 250.4696 | Workbook XML → column width sum |
| Range Top | 75.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 457.8450 | origin_x + range_left × scale_w |
| Page Top | 380.6083 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1907.69 | page_left × (DPI/72) |
| Pixel Top | 1585.87 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.748113 |
| Top Ratio | 0.480566 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field F6 | col=6, row=6, col_end=6, row_end=6 | F6 | cell | workbook XML |
| 2 | Col 6 width | col_widths_pt[6] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 6 height | row_heights_pt[6] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..5 = 50.0939 + 50.0939 + 50.0939 + 50.0939 + 50.0939 | 250.469625 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..5 = 15.0000 + 15.0000 + 15.0000 + 15.0000 + 15.0000 | 75.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 6..6 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 6..6 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 250.4696 * 1.010402 | 457.844994 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 75.0000 * 1.026111 | 380.608332 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 457.8450 × 4.166667 | 1907.69 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 380.6083 × 4.166667 | 1585.87 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1907.69 / 2550.00 | 0.748113 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1585.87 / page_height | 0.480566 |  | computed |

## E7

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 200.3757 | Workbook XML → column width sum |
| Range Top | 90.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 407.2300 | origin_x + range_left × scale_w |
| Page Top | 396.0000 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1696.79 | page_left × (DPI/72) |
| Pixel Top | 1650.00 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.665408 |
| Top Ratio | 0.500000 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field E7 | col=5, row=7, col_end=5, row_end=7 | E7 | cell | workbook XML |
| 2 | Col 5 width | col_widths_pt[5] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 7 height | row_heights_pt[7] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..4 = 50.0939 + 50.0939 + 50.0939 + 50.0939 | 200.3757 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..6 = 15.0000 + 15.0000 + 15.0000 + 15.0000 + 15.0000 + 15.0000 | 90.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 5..5 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 7..7 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 200.3757 * 1.010402 | 407.229996 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 90.0000 * 1.026111 | 396.0 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 407.2300 × 4.166667 | 1696.79 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 396.0000 × 4.166667 | 1650.0 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1696.79 / 2550.00 | 0.665408 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1650.00 / page_height | 0.5 |  | computed |

## F7

### Worksheet Position
| Property | Value (pt) | Source |
|----------|------------|--------|
| Range Left | 250.4696 | Workbook XML → column width sum |
| Range Top | 90.0000 | Workbook XML → row height sum |
| Range Width | 50.0939 | Workbook XML → column width sum |
| Range Height | 15.0000 | Workbook XML → row height sum |

### Page Position
| Property | Value (pt) | Formula |
|----------|------------|---------|
| Page Left | 457.8450 | origin_x + range_left × scale_w |
| Page Top | 396.0000 | origin_y + range_top × scale_h |
| Page Width | 50.6150 | range_width × scale_w |
| Page Height | 15.3917 | range_height × scale_h |

### Pixel Position
| Property | Value (px) | Formula |
|----------|------------|---------|
| Pixel Left | 1907.69 | page_left × (DPI/72) |
| Pixel Top | 1650.00 | page_top × (DPI/72) |
| Pixel Width | 210.90 | page_width × (DPI/72) |
| Pixel Height | 64.13 | page_height × (DPI/72) |

### Ratios
| Property | Value |
|----------|-------|
| Left Ratio | 0.748113 |
| Top Ratio | 0.500000 |
| Width Ratio | 0.082704 |
| Height Ratio | 0.019434 |

### Derivation Steps
| # | Description | Formula | Value | Unit | Source |
|---|-------------|---------|-------|------|--------|
| 1 | Field F7 | col=6, row=7, col_end=6, row_end=7 | F7 | cell | workbook XML |
| 2 | Col 6 width | col_widths_pt[6] = 50.0939 pt | 50.093925 | pt | workbook XML |
| 3 | Row 7 height | row_heights_pt[7] = 15.0000 | 15.0 | pt | workbook XML |
| 4 | Range left = Σ col widths before field | Σ cols 1..5 = 50.0939 + 50.0939 + 50.0939 + 50.0939 + 50.0939 | 250.469625 | pt | workbook XML → computed |
| 5 | Range top = Σ row heights before field | Σ rows 1..6 = 15.0000 + 15.0000 + 15.0000 + 15.0000 + 15.0000 + 15.0000 | 90.0 | pt | workbook XML → computed |
| 6 | Range width = Σ col widths of field | Σ cols 6..6 = 50.0939 | 50.093925 | pt | workbook XML → computed |
| 7 | Range height = Σ row heights of field | Σ rows 7..7 = 15.0000 | 15.0 | pt | workbook XML → computed |
| 8 | Page width | paper_size → page_width_pt | 612.0 | pt | Excel Page Setup |
| 9 | Page height | paper_size → page_height_pt | 792.0 | pt | Excel Page Setup |
| 10 | Printable width = page - Lmargin - Rmargin | 612.00 - 0.70 - 0.70 | 610.6 | pt | computed from page setup |
| 11 | Printable height = page - Tmargin - Bmargin | 792.00 - 0.75 - 0.75 | 790.5 | pt | computed from page setup |
| 12 | PA Width (PAW) = Σ col widths of print area | = 200.3757 | 200.3757 | pt | workbook XML |
| 13 | PA Height (PAH) = Σ row heights of print area | = 180.0000 | 180.0 | pt | workbook XML |
| 14 | Effective width (effW) from PDF content analysis | centered_h=True → effW=202.4600 | 202.459991 | pt | PDF content bounds |
| 15 | Effective height (effH) from PDF content analysis | centered_v=True → effH=184.7000 | 184.700012 | pt | PDF content bounds |
| 16 | Scale X = effW / PAW | 202.459991 / 200.375700 | 1.010402 |  | computed |
| 17 | Scale Y = effH / PAH | 184.700012 / 180.000000 | 1.026111 |  | computed |
| 18 | Origin X = Lmargin + (printable_w - effW)/2 | Lmargin + (printable_w - effW)/2 = 0.70 + (610.60 - 202.46)/2 | 204.770004 | pt | computed |
| 19 | Origin Y = Tmargin + (printable_h - effH)/2 | Tmargin + (printable_h - effH)/2 = 0.75 + (790.50 - 184.70)/2 | 303.649994 | pt | computed |
| 20 | Page left = origin_x + range_left × scale_w | origin_x + range_left * scale_w = 204.7700 + 250.4696 * 1.010402 | 457.844994 | pt | computed |
| 21 | Page top = origin_y + range_top × scale_h | origin_y + range_top * scale_h = 303.6500 + 90.0000 * 1.026111 | 396.0 | pt | computed |
| 22 | Page width = range_width × scale_w | 50.0939 × 1.010402 | 50.614998 | pt | computed |
| 23 | Page height = range_height × scale_h | 15.0000 × 1.026111 | 15.391668 | pt | computed |
| 24 | Pixel left = page_left × (dpi/72) | 457.8450 × 4.166667 | 1907.69 | px | computed |
| 25 | Pixel top = page_top × (dpi/72) | 396.0000 × 4.166667 | 1650.0 | px | computed |
| 26 | Pixel width = page_width × (dpi/72) | 50.6150 × 4.166667 | 210.9 | px | computed |
| 27 | Pixel height = page_height × (dpi/72) | 15.3917 × 4.166667 | 64.13 | px | computed |
| 28 | Left ratio = pixel_left / page_width_px | 1907.69 / 2550.00 | 0.748113 |  | computed |
| 29 | Top ratio = pixel_top / page_height_px | 1650.00 / page_height | 0.5 |  | computed |
