# Legacy Coordinate Algorithm

## Reverse-Engineered Formula

### Step 1: Get workbook position from Excel COM

```vb
Dim leftPts As Double
Dim topPts As Double
Dim widthPts As Double
Dim heightPts As Double
leftPts = Range("$A$1:$B$2").Left
topPts = Range("$A$1:$B$2").Top
widthPts = Range("$A$1:$B$2").Width
heightPts = Range("$A$1:$B$2").Height
```

### Step 2: Read PageSetup

```vb
Dim pageWidth As Double: pageWidth = 612.0
Dim pageHeight As Double: pageHeight = 792.0
Dim leftMargin As Double: leftMargin = 50.4000
Dim rightMargin As Double: rightMargin = 50.4000
Dim topMargin As Double: topMargin = 54.0000
Dim bottomMargin As Double: bottomMargin = 54.0000
Dim centerH As Boolean: centerH = False
Dim centerV As Boolean: centerV = False
Dim zoom As Integer: zoom = 100
Dim fitW As Integer: fitW = 1
Dim fitH As Integer: fitH = 1
```

### Step 3: Compute printable area

```vb
Dim printableWidth As Double
printableWidth = pageWidth - leftMargin - rightMargin
' = 612.0 - 50.4000 - 50.4000 = 511.2
Dim printableHeight As Double
printableHeight = pageHeight - topMargin - bottomMargin
' = 792.0 - 54.0000 - 54.0000 = 684.0
```

### Best Fit: Raw page ratio: DB = COM / PageDimension

#### $A$1:$B$2

```vb
leftPts = 0.0000
topPts  = 0.0000
printedOriginX = 50.4000  ' leftMargin
printedOriginY = 54.0000  ' topMargin
leftRatio  = (0.0000 - 50.4000) / 511.2 = -0.0985915  (DB: 0.3364706) ✗ MISMATCH
topRatio   = (0.0000 - 54.0000) / 684.0 = -0.0789474  (DB: 0.3845454) ✗ MISMATCH
```

#### $C$1:$D$2

```vb
leftPts = 96.0000
topPts  = 0.0000
printedOriginX = 50.4000  ' leftMargin
printedOriginY = 54.0000  ' topMargin
leftRatio  = (96.0000 - 50.4000) / 511.2 = 0.0892019  (DB: 0.5000000) ✗ MISMATCH
topRatio   = (0.0000 - 54.0000) / 684.0 = -0.0789474  (DB: 0.3845454) ✗ MISMATCH
```

#### $A$3:$D$4

```vb
leftPts = 0.0000
topPts  = 28.8000
printedOriginX = 50.4000  ' leftMargin
printedOriginY = 54.0000  ' topMargin
leftRatio  = (0.0000 - 50.4000) / 511.2 = -0.0985915  (DB: 0.3364706) ✗ MISMATCH
topRatio   = (28.8000 - 54.0000) / 684.0 = -0.0368421  (DB: 0.4231818) ✗ MISMATCH
```

## Final Algorithm

```vb
' For each cluster cell address:
Function ComputeClusterRatio(cellAddress As String) As (left, top, right, bottom)
    Dim rng As Range
    Set rng = ws.Range(cellAddress)
    
    ' 1. Get COM positions
    Dim leftPts As Double: leftPts = rng.Left
    Dim topPts As Double: topPts = rng.Top
    Dim widthPts As Double: widthPts = rng.Width
    Dim heightPts As Double: heightPts = rng.Height
    
    ' 2. Get page dimensions
    Dim pw As Double: pw = ws.PageSetup.PageWidth
    Dim ph As Double: ph = ws.PageSetup.PageHeight
    Dim lm As Double: lm = ws.PageSetup.LeftMargin
    Dim rm As Double: rm = ws.PageSetup.RightMargin
    Dim tm As Double: tm = ws.PageSetup.TopMargin
    Dim bm As Double: bm = ws.PageSetup.BottomMargin
    
    ' 3. Compute printable area
    Dim printW As Double: printW = pw - lm - rm
    Dim printH As Double: printH = ph - tm - bm
    
    ' 4. Compute printed origin (accounts for centering)
    Dim originX As Double
    If ws.PageSetup.CenterHorizontally Then
        originX = (pw - printW) / 2
    Else
        originX = lm
    End If
    
    Dim originY As Double
    If ws.PageSetup.CenterVertically Then
        originY = (ph - printH) / 2
    Else
        originY = tm
    End If
    
    ' 5. Convert to ratios
    left   = RoundEx((leftPts - originX) / printW, 7)
    top    = RoundEx((topPts - originY) / printH, 7)
    right  = RoundEx((leftPts + widthPts - originX) / printW, 7)
    bottom = RoundEx((topPts + heightPts - originY) / printH, 7)
    
    Return (left, top, right, bottom)
End Function
```

Note: `RoundEx` uses VB6's round-half-away-from-zero with float (x87 FPU 80-bit precision).
In .NET, this is `Math.Round((float)value, 7, MidpointRounding.AwayFromZero)`.
