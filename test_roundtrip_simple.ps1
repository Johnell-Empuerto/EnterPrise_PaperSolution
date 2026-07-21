param(
    [string]$ApiBase = "http://localhost:5090",
    [string]$InputFile = "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx",
    [string]$OutputDir = "C:\Users\MCF-JOHNELLEEMPUERTO\Documents"
)

$ErrorActionPreference = "Stop"

function Upload-File($path) {
    $uploadUrl = "$ApiBase/api/form/upload-preview"
    $boundary = [System.Guid]::NewGuid().ToString()
    $fileBytes = [System.IO.File]::ReadAllBytes($path)
    $fileName = Split-Path $path -Leaf
    
    $sb = New-Object System.Text.StringBuilder
    $sb.Append("--$boundary`r`n") | Out-Null
    $sb.Append("Content-Disposition: form-data; name=`"file`"; filename=`"$fileName`"`r`n") | Out-Null
    $sb.Append("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`r`n`r`n") | Out-Null
    $headerStr = $sb.ToString()
    $footerStr = "`r`n--$boundary--`r`n"
    
    $encoding = [System.Text.Encoding]::UTF8
    $headerBytes = $encoding.GetBytes($headerStr)
    $footerBytes = $encoding.GetBytes($footerStr)
    
    $body = New-Object byte[] ($headerBytes.Length + $fileBytes.Length + $footerBytes.Length)
    [System.Buffer]::BlockCopy($headerBytes, 0, $body, 0, $headerBytes.Length)
    [System.Buffer]::BlockCopy($fileBytes, 0, $body, $headerBytes.Length, $fileBytes.Length)
    [System.Buffer]::BlockCopy($footerBytes, 0, $body, $headerBytes.Length + $fileBytes.Length, $footerBytes.Length)
    
    $response = Invoke-RestMethod -Uri $uploadUrl -Method Post `
        -ContentType "multipart/form-data; boundary=$boundary" -Body $body
    return $response
}

function Export-Xlsx($sessionId, $pages) {
    $exportUrl = "$ApiBase/api/form/save-edited"
    $fields = $pages[0].fields
    $sheetName = $pages[0].sheetName
    
    $fieldJson = ($fields | ForEach-Object {
        '{"id":"' + $_.id + '","cell":{"address":"' + $_.cellAddr + '","rowIndex":1},"type":0,"value":""}'
    }) -join ","
    
    $json = '{"sessionId":"' + $sessionId + '","info":{"title":"RoundTrip"},"sheets":[{"name":"' + $sheetName + '","index":0,"fields":[' + $fieldJson + ']}]}'
    
    $response = Invoke-WebRequest -Uri $exportUrl -Method Post `
        -ContentType "application/json" -Body $json -UseBasicParsing
    return $response.Content
}

# ── Cycle 1: Upload fresh file ──
Write-Host "=== Cycle 1: Upload fresh file ===" -ForegroundColor Cyan
$r1 = Upload-File $InputFile
Write-Host "  Session: $($r1.sessionId)"
Write-Host "  Pages: $($r1.pages.Count)"
Write-Host "  Fields: $($r1.pages[0].fields.Count)"
Write-Host "  PaperLessConfig: $(if ($r1.paperLessConfig) { 'present' } else { 'null' })"

if (-not $r1.paperLessConfig) {
    Write-Host "  (No PaperLessConfig - expected for fresh upload)" -ForegroundColor Yellow
}

# ── Cycle 2: Export and re-upload ──
Write-Host "=== Cycle 2: Export and re-upload ===" -ForegroundColor Cyan
$c2_bytes = Export-Xlsx $r1.sessionId $r1.pages
Write-Host "  Export bytes: $($c2_bytes.Length)"
$c2_path = "$OutputDir\test_r2.xlsx"
[System.IO.File]::WriteAllBytes($c2_path, $c2_bytes)
Write-Host "  Written to: $c2_path"

$r2 = Upload-File $c2_path
Write-Host "  Session: $($r2.sessionId)"
Write-Host "  PaperLessConfig: $(if ($r2.paperLessConfig) { 'present' } else { 'null' })"

if ($r2.paperLessConfig) {
    $plc = $r2.paperLessConfig
    Write-Host "  Sheets: $($plc.sheets.Count)" -ForegroundColor Green
    foreach ($sheet in $plc.sheets) {
        Write-Host "  Sheet: $($sheet.name), Fields: $($sheet.fields.Count)"
        foreach ($f in $sheet.fields) {
            $s = if ($f.style) { "style[$($f.style.font.name)/$($f.style.font.sizePt)]" } else { "no-style" }
            $c = if ($f.config) { "config[pl=$($f.config.placeholder) dv=$($f.config.defaultValue)]" } else { "no-config" }
            Write-Host "    id=$($f.id) type=$($f.type) $s $c"
        }
    }
} else {
    Write-Host "  (No PaperLessConfig found)" -ForegroundColor Red
}

# ── Cycle 3: Export cycle 2 and re-upload ──
Write-Host "=== Cycle 3: Export cycle 2 and re-upload ===" -ForegroundColor Cyan
$c3_bytes = Export-Xlsx $r2.sessionId $r2.pages
Write-Host "  Export bytes: $($c3_bytes.Length)"
$c3_path = "$OutputDir\test_r3.xlsx"
[System.IO.File]::WriteAllBytes($c3_path, $c3_bytes)

$r3 = Upload-File $c3_path
Write-Host "  Session: $($r3.sessionId)"
Write-Host "  PaperLessConfig: $(if ($r3.paperLessConfig) { 'present' } else { 'null' })"

if ($r3.paperLessConfig) {
    $plc = $r3.paperLessConfig
    Write-Host "  Sheets: $($plc.sheets.Count)" -ForegroundColor Green
    foreach ($sheet in $plc.sheets) {
        Write-Host "  Sheet: $($sheet.name), Fields: $($sheet.fields.Count)"
        foreach ($f in $sheet.fields) {
            $sProps = if ($f.style) { "font=$($f.style.font.name)/$($f.style.font.sizePt) h=$($f.style.alignment.horizontal) v=$($f.style.alignment.vertical)" } else { "no-style" }
            $cProps = if ($f.config) { "req=$($f.config.required) ml=$($f.config.maxLength) pl='$($f.config.placeholder)' dv='$($f.config.defaultValue)' rest=$($f.config.inputRestriction)" } else { "no-config" }
            Write-Host "    id=$($f.id) type=$($f.type) $sProps | $cProps"
        }
    }
    # ── Compare cycle 2 vs cycle 3 ──
    Write-Host "  Comparing C2 vs C3 PaperLessConfig..." -ForegroundColor Yellow
    $identical = $true
    $sheets2 = $r2.paperLessConfig.sheets
    $sheets3 = $r3.paperLessConfig.sheets
    if ($sheets2.Count -ne $sheets3.Count) { Write-Host "  DIFF: Sheet count $($sheets2.Count) vs $($sheets3.Count)" -ForegroundColor Red; $identical = $false }
    for ($si = 0; $si -lt [Math]::Min($sheets2.Count, $sheets3.Count); $si++) {
        $f2 = $sheets2[$si].fields; $f3 = $sheets3[$si].fields
        if ($f2.Count -ne $f3.Count) { Write-Host "  DIFF: Field count $($f2.Count) vs $($f3.Count)" -ForegroundColor Red; $identical = $false }
        for ($fi = 0; $fi -lt [Math]::Min($f2.Count, $f3.Count); $fi++) {
            $aJson = ($f2[$fi] | ConvertTo-Json -Depth 5 -Compress)
            $bJson = ($f3[$fi] | ConvertTo-Json -Depth 5 -Compress)
            if ($aJson -ne $bJson) {
                Write-Host "  DIFF in field $($f2[$fi].id):" -ForegroundColor Red
                Write-Host "    C2: $aJson" -ForegroundColor Gray
                Write-Host "    C3: $bJson" -ForegroundColor Gray
                $identical = $false
            }
        }
    }
    if ($identical) { Write-Host "  C2 and C3 PaperLessConfig are IDENTICAL!" -ForegroundColor Green }
} else {
    Write-Host "  (No PaperLessConfig found)" -ForegroundColor Red
}

# ── Cycle 4: Export cycle 3 and re-upload ──
Write-Host "=== Cycle 4: Export cycle 3 and re-upload ===" -ForegroundColor Cyan
if ($r3.paperLessConfig) {
    $c4_bytes = Export-Xlsx $r3.sessionId $r3.pages
    Write-Host "  Export bytes: $($c4_bytes.Length)"
    $c4_path = "$OutputDir\test_r4.xlsx"
    [System.IO.File]::WriteAllBytes($c4_path, $c4_bytes)
    
    $r4 = Upload-File $c4_path
    Write-Host "  PaperLessConfig: $(if ($r4.paperLessConfig) { 'present' } else { 'null' })"
    if ($r4.paperLessConfig) {
        $identical2 = $true
        $sheets3b = $r3.paperLessConfig.sheets
        $sheets4 = $r4.paperLessConfig.sheets
        for ($si = 0; $si -lt [Math]::Min($sheets3b.Count, $sheets4.Count); $si++) {
            $f3b = $sheets3b[$si].fields; $f4 = $sheets4[$si].fields
            for ($fi = 0; $fi -lt [Math]::Min($f3b.Count, $f4.Count); $fi++) {
                $aJson = ($f3b[$fi] | ConvertTo-Json -Depth 5 -Compress)
                $bJson = ($f4[$fi] | ConvertTo-Json -Depth 5 -Compress)
                if ($aJson -ne $bJson) {
                    Write-Host "  DIFF in field $($f3b[$fi].id) (C3 vs C4):" -ForegroundColor Red
                    Write-Host "    C3: $aJson" -ForegroundColor Gray
                    Write-Host "    C4: $bJson" -ForegroundColor Gray
                    $identical2 = $false
                }
            }
        }
        if ($identical2) { Write-Host "  C3 and C4 PaperLessConfig are IDENTICAL!" -ForegroundColor Green }
    } else {
        Write-Host "  (No PaperLessConfig found)" -ForegroundColor Red
    }
} else {
    Write-Host "  Skipping (no PLC from C3)" -ForegroundColor Yellow
}

Write-Host "=== Done ===" -ForegroundColor Cyan
