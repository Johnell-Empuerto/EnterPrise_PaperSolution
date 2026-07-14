# Coordinate Validation

Template 547: DB vs COM vs Generated

## Per-Cluster Detail

| Cell | DB Left | COM Left(pts) | Gen Left | DB Right | COM Right(pts) | Gen Right | DB Top | COM Top(pts) | Gen Top | DB Bottom | COM Bottom(pts) | Gen Bottom | Status |
|------|---------|--------------|---------|----------|---------------|----------|--------|-------------|---------|-----------|----------------|-----------|--------|
| $I$6:$M$6 | 0.5294118 | 384.00 | 0.5294118 | 0.9282353 | 624.00 | 0.9282353 | 0.1654546 | 66.00 | 0.1654546 | 0.1868182 | 79.20 | 0.1868182 | ✗ FAIL |
| $I$7:$M$7 | 0.5294118 | 384.00 | 0.5294118 | 0.9282353 | 624.00 | 0.9282353 | 0.1877273 | 79.20 | 0.1821213 | 0.2095454 | 92.40 | 0.2034849 | ✗ FAIL |
| $I$8:$M$8 | 0.5294118 | 384.00 | 0.5294118 | 0.9282353 | 624.00 | 0.9282353 | 0.2104545 | 92.40 | 0.1987879 | 0.2327273 | 105.60 | 0.2201515 | ✗ FAIL |
| $I$9:$M$9 | 0.5294118 | 384.00 | 0.5294118 | 0.9282353 | 624.00 | 0.9282353 | 0.2336364 | 105.60 | 0.2154546 | 0.2559091 | 118.80 | 0.2368182 | ✗ FAIL |
| $I$10:$M$10 | 0.5294118 | 384.00 | 0.5294118 | 0.9282353 | 624.00 | 0.9282353 | 0.2568182 | 118.80 | 0.2321213 | 0.2786364 | 132.00 | 0.2534849 | ✗ FAIL |

## Coordinate Parameters

| Parameter | Value |
|-----------|-------|
| PageWidth | 612.0 pt |
| PageHeight | 792.0 pt |
| ContentWidth | 0.0 pt |
| ContentHeight | 0.0 pt |
| PrintedOriginX | -60.0000 pt |
| PrintedOriginY | 65.0400 pt |
| PrintArea |  |
| CenterHorizontally | False |
| CenterVertically | False |
| LeftMargin | 50.4000 pt |
| RightMargin | 50.4000 pt |
| TopMargin | 54.0000 pt |
| BottomMargin | 54.0000 pt |
| UsedRangeWidth | 0.0 pt |
| UsedRangeHeight | 0.0 pt |

**Formula:** `DB_Ratio = RoundEx((COM_Position + PrintedOrigin) / PageDimension, 7)`

Where `RoundEx(x) = Math.Round((float)x, 7, MidpointRounding.AwayFromZero)`

**Content Width Derivation:**
- PrintedOriginX = (PageWidth - ContentWidth) / 2 = (612.0 - 0.0) / 2 = -60.0000 pt

## Result

**SOME CLUSTERS FAIL — see table above**
