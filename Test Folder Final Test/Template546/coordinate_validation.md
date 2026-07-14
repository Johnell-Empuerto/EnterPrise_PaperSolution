# Coordinate Validation

Template 546: DB vs COM vs Generated

## Per-Cluster Detail

| Cell | DB Left | COM Left(pts) | Gen Left | DB Right | COM Right(pts) | Gen Right | DB Top | COM Top(pts) | Gen Top | DB Bottom | COM Bottom(pts) | Gen Bottom | Status |
|------|---------|--------------|---------|----------|---------------|----------|--------|-------------|---------|-----------|----------------|-----------|--------|
| $A$1:$B$2 | 0.3364706 | 0.00 | 0.3364706 | 0.4982353 | 96.00 | 0.4982353 | 0.3845454 | 0.00 | 0.3845454 | 0.4218182 | 28.80 | 0.4218182 | ✗ FAIL |
| $C$1:$D$2 | 0.5000000 | 96.00 | 0.4933334 | 0.6635294 | 192.00 | 0.655098 | 0.3845454 | 0.00 | 0.3845454 | 0.4218182 | 28.80 | 0.4218182 | ✗ FAIL |
| $A$3:$D$4 | 0.3364706 | 0.00 | 0.3364706 | 0.6635294 | 192.00 | 0.655098 | 0.4231818 | 28.80 | 0.420909 | 0.4604546 | 57.60 | 0.4581818 | ✗ FAIL |
| $A$6:$D$7 | 0.3364706 | 0.00 | 0.3364706 | 0.6635294 | 192.00 | 0.655098 | 0.4809091 | 72.00 | 0.4754545 | 0.5181818 | 100.80 | 0.5127273 | ✗ FAIL |
| $A$9:$D$10 | 0.3364706 | 0.00 | 0.3364706 | 0.6635294 | 192.00 | 0.655098 | 0.5386364 | 115.20 | 0.53 | 0.5759091 | 144.00 | 0.5672727 | ✗ FAIL |
| $A$12 | 0.3364706 | 0.00 | 0.3364706 | 0.4164706 | 48.00 | 0.4198039 | 0.5963637 | 158.40 | 0.5845454 | 0.6150000 | 172.80 | 0.6036364 | ✗ FAIL |

## Coordinate Parameters

| Parameter | Value |
|-----------|-------|
| PageWidth | 612.0 pt |
| PageHeight | 792.0 pt |
| ContentWidth | 0.0 pt |
| ContentHeight | 0.0 pt |
| PrintedOriginX | 205.9200 pt |
| PrintedOriginY | 304.5600 pt |
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
- PrintedOriginX = (PageWidth - ContentWidth) / 2 = (612.0 - 0.0) / 2 = 205.9200 pt

## Result

**SOME CLUSTERS FAIL — see table above**
