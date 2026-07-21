# Round-trip verification: Simulate the frontend's behavior
# The frontend always sends PaperLessConfig-derived fieldconfigs in the export payload.
# This test verifies that when the frontend does this correctly, config survives.

param(
    [string]$ApiBase = "http://localhost:5090",
    [string]$InputFile = "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx"
)

$ErrorActionPreference = "Stop"

function Upload-File($path) {
    $uploadUrl = "$ApiBase/api/form/upload-preview"
    $boundary = [System.Guid]::NewGuid().ToString()
    $fileBytes = [System.IO.File]::ReadAllBytes($path)
    $fileName = Split-Path $path -Leaf
    $sb = New-Object System.Text.StringBuilder
    $sb.Append("--$boundary`r`nContent-Disposition: form-data; name=`"file`"; filename=`"$fileName`"`r`nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`r`n`r`n") | Out-Null
    $headerStr = $sb.ToString()
    $footerStr = "`r`n--$boundary--`r`n"
    $encoding = [System.Text.Encoding]::UTF8
    $headerBytes = $encoding.GetBytes($headerStr)
    $footerBytes = $encoding.GetBytes($footerStr)
    $body = New-Object byte[] ($headerBytes.Length + $fileBytes.Length + $footerBytes.Length)
    [System.Buffer]::BlockCopy($headerBytes, 0, $body, 0, $headerBytes.Length)
    [System.Buffer]::BlockCopy($fileBytes, 0, $body, $headerBytes.Length, $fileBytes.Length)
    [System.Buffer]::BlockCopy($footerBytes, 0, $body, $headerBytes.Length + $fileBytes.Length, $footerBytes.Length)
    $response = Invoke-RestMethod -Uri $uploadUrl -Method Post -ContentType "multipart/form-data; boundary=$boundary" -Body $body
    return $response
}

function Export-Xlsx($sessionId, $pages, $fieldConfigs) {
    $exportUrl = "$ApiBase/api/form/save-edited"
    $sheetName = $pages[0].sheetName
    
    # fieldConfigs: array of objects with: id, cellAddr, type, value, placeholder, defaultValue, inputRestriction, required, maxLength, fontName, fontSize, fontColor, hAlign, vAlign
    $fieldJson = ($fieldConfigs | ForEach-Object {
        $sb = New-Object System.Text.StringBuilder
        $sb.Append('{"id":"') | Out-Null
        $sb.Append($_.id) | Out-Null
        $sb.Append('","cell":{"address":"') | Out-Null
        $sb.Append($_.cellAddr) | Out-Null
        $sb.Append('","rowIndex":1},"type":0,"value":"') | Out-Null
        $sb.Append(($_.value -replace '"','\"')) | Out-Null
        $sb.Append('"') | Out-Null
        
        if ($_.placeholder) { $sb.Append(',"placeholder":"'); $sb.Append($_.placeholder -replace '"','\"'); $sb.Append('"') }
        if ($_.defaultValue) { $sb.Append(',"defaultValue":"'); $sb.Append($_.defaultValue -replace '"','\"'); $sb.Append('"') }
        if ($_.inputRestriction) { $sb.Append(',"dataValidation":{"type":"'); $sb.Append($_.inputRestriction); $sb.Append('"}') }
        if ($_.required) { $sb.Append(',"required":true') }
        if ($_.maxLength -gt 0) { $sb.Append(',"maxLength":'); $sb.Append($_.maxLength) }
        if ($_.readOnly) { $sb.Append(',"locked":true') }
        if ($_.visible -eq $false) { $sb.Append(',"visible":false') }
        
        # Style
        $hasFont = $_.fontName -or $_.fontSize -or $_.fontColor
        $hasAlign = $_.hAlign -or $_.vAlign
        if ($hasFont -or $hasAlign) {
            $sb.Append(',"style":{')
            if ($hasFont) {
                $sb.Append('"font":{')
                $first = $true
                if ($_.fontName) { if (-not $first) { $sb.Append(',') }; $first = $false; $sb.Append('"name":"'); $sb.Append($_.fontName); $sb.Append('"') }
                if ($_.fontSize) { if (-not $first) { $sb.Append(',') }; $first = $false; $sb.Append('"sizePt":'); $sb.Append($_.fontSize) }
                if ($_.fontColor) { if (-not $first) { $sb.Append(',') }; $first = $false; $sb.Append('"colorArgb":"'); $sb.Append($_.fontColor); $sb.Append('"') }
                $sb.Append('}')
            }
            if ($hasAlign) {
                if ($hasFont) { $sb.Append(',') }
                $sb.Append('"alignment":{')
                $first2 = $true
                if ($_.hAlign) { if (-not $first2) { $sb.Append(',') }; $first2 = $false; $sb.Append('"horizontal":"'); $sb.Append($_.hAlign); $sb.Append('"') }
                if ($_.vAlign) { if (-not $first2) { $sb.Append(',') }; $first2 = $false; $sb.Append('"vertical":"'); $sb.Append($_.vAlign); $sb.Append('"') }
                $sb.Append('}')
            }
            $sb.Append('}')
        }
        
        $sb.Append('}') | Out-Null
        $sb.ToString()
    }) -join ","
    
    $json = '{"sessionId":"' + $sessionId + '","info":{"title":"RoundTrip"},"sheets":[{"name":"' + $sheetName + '","index":0,"fields":[' + $fieldJson + ']}]}'
    
    $response = Invoke-WebRequest -Uri $exportUrl -Method Post -ContentType "application/json" -Body $json -UseBasicParsing
    return $response.Content
}

function Dump-PLC($plc, $label) {
    if (-not $plc) { Write-Host "  ${label}: null"; return }
    foreach ($sheet in $plc.sheets) {
        foreach ($f in $sheet.fields) {
            $s = if ($f.style) { "font=$($f.style.font.name)/$($f.style.font.sizePt) color=$($f.style.font.colorArgb) hAlign=$($f.style.alignment.horizontal) vAlign=$($f.style.alignment.vertical)" } else { "no-style" }
            $c = if ($f.config) { "req=$($f.config.required) ml=$($f.config.maxLength) pl='$($f.config.placeholder)' dv='$($f.config.defaultValue)' rest=$($f.config.inputRestriction) ro=$($f.config.readOnly)" } else { "no-config" }
            Write-Host "  [$label] id=$($f.id) $s`n           $c"
        }
    }
}

# ════════════════════════════════════════
# TEST 1: Font color with R,G,B format
# ════════════════════════════════════════
Write-Host "=== TEST 1: Font color (R,G,B format) ===" -ForegroundColor Cyan
$r = Upload-File $InputFile
$fields = $r.pages[0].fields

# Simulate frontend setting font color as "255,0,0" (R,G,B string)
$cfg1 = @(
    @{ id = $fields[0].id; cellAddr = $fields[0].cellAddr; value = ""; fontColor = "255,0,0" },
    @{ id = $fields[1].id; cellAddr = $fields[1].cellAddr; value = ""; fontName = "Arial"; fontSize = 14 },
    @{ id = $fields[2].id; cellAddr = $fields[2].cellAddr; value = "" },
    @{ id = $fields[3].id; cellAddr = $fields[3].cellAddr; value = "" },
    @{ id = $fields[4].id; cellAddr = $fields[4].cellAddr; value = "" },
    @{ id = $fields[5].id; cellAddr = $fields[5].cellAddr; value = "" }
)

$b1 = Export-Xlsx $r.sessionId $r.pages $cfg1
Set-Content "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\test_t1.xlsx" $b1 -Encoding Byte
$r1 = Upload-File "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\test_t1.xlsx"
Dump-PLC $r1.paperLessConfig "T1"

# Check if R,G,B was converted to ARGB hex
if ($r1.paperLessConfig) {
    $f0 = $r1.paperLessConfig.sheets[0].fields[0]
    if ($f0.style -and $f0.style.font.colorArgb) {
        $color = $f0.style.font.colorArgb
        if ($color -eq "FFFF0000") {
            Write-Host "  PASS: Font color correctly converted to FFFF0000" -ForegroundColor Green
        } elseif ($color -eq "255,0,0") {
            Write-Host "  FAIL: Font color NOT converted (still '255,0,0')" -ForegroundColor Red
        } else {
            Write-Host "  CHECK: Font color = '$color' (unexpected format)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  FAIL: No font color in PLC" -ForegroundColor Red
    }
}

# ════════════════════════════════════════
# TEST 2: Alignment not leaking (Bug 1)
# ════════════════════════════════════════
Write-Host "=== TEST 2: Alignment leakage (Bug 1) ===" -ForegroundColor Cyan
$r = Upload-File $InputFile
$fields = $r.pages[0].fields

# Field 0: set ONLY font name, no alignment (simulate user only changing font)
# Field 1: set explicit alignment (center/center)
# Field 2: no changes
$cfg2 = @(
    @{ id = $fields[0].id; cellAddr = $fields[0].cellAddr; value = ""; fontName = "Arial"; fontSize = 14 },
    @{ id = $fields[1].id; cellAddr = $fields[1].cellAddr; value = ""; hAlign = "center"; vAlign = "center" },
    @{ id = $fields[2].id; cellAddr = $fields[2].cellAddr; value = "" },
    @{ id = $fields[3].id; cellAddr = $fields[3].cellAddr; value = "" },
    @{ id = $fields[4].id; cellAddr = $fields[4].cellAddr; value = "" },
    @{ id = $fields[5].id; cellAddr = $fields[5].cellAddr; value = "" }
)

$b2 = Export-Xlsx $r.sessionId $r.pages $cfg2
Set-Content "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\test_t2.xlsx" $b2 -Encoding Byte
$r2 = Upload-File "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\test_t2.xlsx"
Dump-PLC $r2.paperLessConfig "T2"

$plc2 = $r2.paperLessConfig
if ($plc2) {
    $f0 = $plc2.sheets[0].fields[0]
    $f1 = $plc2.sheets[0].fields[1]
    $f2 = $plc2.sheets[0].fields[2]
    
    $leak = $false
    
    # Field 0 should have Arial/14 and alignment should be from Excel (general/bottom)
    # NOT the center/center from Field 1
    if ($f0.style) {
        if ($f0.style.font.name -ne "Arial") { Write-Host "  FAIL T2-F0: Font name should be Arial, got '$($f0.style.font.name)'" -ForegroundColor Red; $leak = $true }
        # Check if alignment leaked from Field 1
        if ($f0.style.alignment -and $f0.style.alignment.horizontal -eq "center") { Write-Host "  FAIL T2-F0: hAlign=center leaked from Field 1!" -ForegroundColor Red; $leak = $true }
        if ($f0.style.alignment -and $f0.style.alignment.vertical -eq "center") { Write-Host "  FAIL T2-F0: vAlign=center leaked from Field 1!" -ForegroundColor Red; $leak = $true }
    }
    
    # Field 1 should have center/center alignment
    if ($f1.style) {
        if ($f1.style.alignment.horizontal -ne "center") { Write-Host "  FAIL T2-F1: hAlign should be center, got '$($f1.style.alignment.horizontal)'" -ForegroundColor Red; $leak = $true }
        if ($f1.style.alignment.vertical -ne "center") { Write-Host "  FAIL T2-F1: vAlign should be center, got '$($f1.style.alignment.vertical)'" -ForegroundColor Red; $leak = $true }
    }
    
    if (-not $leak) { Write-Host "  PASS: No alignment leakage!" -ForegroundColor Green }
}

# ════════════════════════════════════════
# TEST 3: InputRestriction persistence (Bug 2)
# ════════════════════════════════════════
Write-Host "=== TEST 3: InputRestriction persistence (Bug 2) ===" -ForegroundColor Cyan
$r = Upload-File $InputFile
$fields = $r.pages[0].fields

# Field 0: Set inputRestriction to "WholeNumber" (Numeric)
$cfg3 = @(
    @{ id = $fields[0].id; cellAddr = $fields[0].cellAddr; value = ""; inputRestriction = "WholeNumber" },
    @{ id = $fields[1].id; cellAddr = $fields[1].cellAddr; value = ""; inputRestriction = "Decimal" },
    @{ id = $fields[2].id; cellAddr = $fields[2].cellAddr; value = "" },
    @{ id = $fields[3].id; cellAddr = $fields[3].cellAddr; value = ""; inputRestriction = "Date" },
    @{ id = $fields[4].id; cellAddr = $fields[4].cellAddr; value = "" },
    @{ id = $fields[5].id; cellAddr = $fields[5].cellAddr; value = ""; inputRestriction = "TextLength" }
)

$b3 = Export-Xlsx $r.sessionId $r.pages $cfg3
Set-Content "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\test_t3.xlsx" $b3 -Encoding Byte
$r3 = Upload-File "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\test_t3.xlsx"
Dump-PLC $r3.paperLessConfig "T3"

$plc3 = $r3.paperLessConfig
if ($plc3) {
    $f0 = $plc3.sheets[0].fields[0]
    $f1 = $plc3.sheets[0].fields[1]
    $f2 = $plc3.sheets[0].fields[2]
    $f3 = $plc3.sheets[0].fields[3]
    $f5 = $plc3.sheets[0].fields[5]
    
    $restOk = $true
    if ($f0.config.inputRestriction -ne "WholeNumber") { Write-Host "  FAIL T3-F0: restriction should be 'WholeNumber', got '$($f0.config.inputRestriction)'" -ForegroundColor Red; $restOk = $false }
    if ($f1.config.inputRestriction -ne "Decimal") { Write-Host "  FAIL T3-F1: restriction should be 'Decimal', got '$($f1.config.inputRestriction)'" -ForegroundColor Red; $restOk = $false }
    if ($f2.config.inputRestriction -ne "None") { Write-Host "  FAIL T3-F2: restriction should be 'None', got '$($f2.config.inputRestriction)'" -ForegroundColor Red; $restOk = $false }
    if ($f3.config.inputRestriction -ne "Date") { Write-Host "  FAIL T3-F3: restriction should be 'Date', got '$($f3.config.inputRestriction)'" -ForegroundColor Red; $restOk = $false }
    if ($f5.config.inputRestriction -ne "TextLength") { Write-Host "  FAIL T3-F5: restriction should be 'TextLength', got '$($f5.config.inputRestriction)'" -ForegroundColor Red; $restOk = $false }
    
    if ($restOk) { Write-Host "  PASS: All input restrictions persisted!" -ForegroundColor Green }
}

# ════════════════════════════════════════
# TEST 4: Full round-trip (4 cycles) with placeholder/defaultValue/font
# ════════════════════════════════════════
Write-Host "=== TEST 4: 4-cycle round-trip with ALL config ===" -ForegroundColor Cyan
$r = Upload-File $InputFile
$fields = $r.pages[0].fields

$cfg4 = @(
    @{ id = $fields[0].id; cellAddr = $fields[0].cellAddr; value = ""; placeholder = "Name"; defaultValue = "John"; required = $true; maxLength = 50; fontName = "Arial"; fontSize = 12; fontColor = "FFFF0000"; hAlign = "left"; vAlign = "top" },
    @{ id = $fields[1].id; cellAddr = $fields[1].cellAddr; value = ""; inputRestriction = "WholeNumber"; required = $true; maxLength = 10; readOnly = $true },
    @{ id = $fields[2].id; cellAddr = $fields[2].cellAddr; value = ""; placeholder = "Notes"; defaultValue = "N/A"; fontName = "Times New Roman"; fontSize = 10; fontColor = "255,0,255" },
    @{ id = $fields[3].id; cellAddr = $fields[3].cellAddr; value = ""; hAlign = "center"; vAlign = "center" },
    @{ id = $fields[4].id; cellAddr = $fields[4].cellAddr; value = "" },
    @{ id = $fields[5].id; cellAddr = $fields[5].cellAddr; value = ""; inputRestriction = "Date"; required = $false; maxLength = 20; fontName = "Calibri"; fontSize = 11 }
)

$prevPlc = $null
$allOk = $true
for ($cycle = 1; $cycle -le 4; $cycle++) {
    $b = Export-Xlsx $r.sessionId $r.pages $cfg4
    $outFile = "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\test_t4_c$cycle.xlsx"
    Set-Content $outFile $b -Encoding Byte
    $r = Upload-File $outFile
    Write-Host "  Cycle ${cycle}: PLC=$(if ($r.paperLessConfig) { 'present' } else { 'null' })"
    
    if ($prevPlc -and $r.paperLessConfig) {
        $prevJson = ($prevPlc.sheets[0].fields | ConvertTo-Json -Depth 5 -Compress)
        $currJson = ($r.paperLessConfig.sheets[0].fields | ConvertTo-Json -Depth 5 -Compress)
        if ($prevJson -ne $currJson) {
            Write-Host "  FAIL Cycle ${cycle}: PaperLessConfig changed from previous!" -ForegroundColor Red
            $allOk = $false
        } else {
            Write-Host "    IDENTICAL to previous" -ForegroundColor Green
        }
    }
    $prevPlc = $r.paperLessConfig
}
if ($allOk) { Write-Host "  PASS: All 4 cycles identical!" -ForegroundColor Green }

Write-Host "=== DONE ===" -ForegroundColor Cyan
