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

function Export-Xlsx($sessionId, $pages, $configuredFields) {
    $exportUrl = "$ApiBase/api/form/save-edited"
    $sheetName = $pages[0].sheetName
    
    # Build field JSON. Use configured fields if provided
    if ($configuredFields) {
        $fieldJson = ($configuredFields | ForEach-Object {
            $json = '{"id":"' + $_.id + '","cell":{"address":"' + $_.cellAddr + '","rowIndex":1},"type":' + $_.type + ',"value":"' + ($_.value -replace '"','\"') + '"'
            if ($_.placeholder) { $json += ',"placeholder":"' + ($_.placeholder -replace '"','\"') + '"' }
            if ($_.defaultValue) { $json += ',"defaultValue":"' + ($_.defaultValue -replace '"','\"') + '"' }
            if ($_.inputRestriction) { $json += ',"dataValidation":{"type":"' + $_.inputRestriction + '"}' }
            if ($_.required -ne $null) { if ($_.required) { $json += ',"required":true' } else { $json += ',"required":false' } }
            if ($_.maxLength -gt 0) { $json += ',"maxLength":' + $_.maxLength }
            if ($_.fontName -or $_.fontSize -or $_.fontColor) {
                $json += ',"style":{"font":{'
                $first = $true
                if ($_.fontName) { if (-not $first) { $json += ',' }; $first = $false; $json += '"name":"' + $_.fontName + '"' }
                if ($_.fontSize) { if (-not $first) { $json += ',' }; $first = $false; $json += '"sizePt":' + $_.fontSize }
                if ($_.fontColor) { if (-not $first) { $json += ',' }; $first = $false; $json += '"colorArgb":"' + $_.fontColor + '"' }
                $json += '}'
                if ($_.hAlign -or $_.vAlign) {
                    $json += ',"alignment":{'
                    $first2 = $true
                    if ($_.hAlign) { if (-not $first2) { $json += ',' }; $first2 = $false; $json += '"horizontal":"' + $_.hAlign + '"' }
                    if ($_.vAlign) { if (-not $first2) { $json += ',' }; $first2 = $false; $json += '"vertical":"' + $_.vAlign + '"' }
                    $json += '}'
                }
                $json += '}'
            }
            $json += '}'
        }) -join ","
    } else {
        $fieldJson = ($pages[0].fields | ForEach-Object {
            '{"id":"' + $_.id + '","cell":{"address":"' + $_.cellAddr + '","rowIndex":1},"type":0,"value":""}'
        }) -join ","
    }
    
    $json = '{"sessionId":"' + $sessionId + '","info":{"title":"RoundTrip"},"sheets":[{"name":"' + $sheetName + '","index":0,"fields":[' + $fieldJson + ']}]}'
    
    $response = Invoke-WebRequest -Uri $exportUrl -Method Post `
        -ContentType "application/json" -Body $json -UseBasicParsing
    return $response.Content
}

# ── Cycle 1: Upload fresh file ──
Write-Host "=== Cycle 1: Upload fresh file ===" -ForegroundColor Cyan
$r = Upload-File $InputFile
Write-Host "  Fields: $($r.pages[0].fields.Count) | PaperLessConfig: $(if ($r.paperLessConfig) { 'present' } else { 'null' })"

# ── Build configured fields ──
$fields = $r.pages[0].fields
$cfgFields = @()

# Field 1: Set placeholder and defaultValue
$cfgFields += @{
    id = $fields[0].id
    cellAddr = $fields[0].cellAddr
    type = 0
    value = ""
    placeholder = "Enter name here"
    defaultValue = "John Doe"
    required = $true
    maxLength = 50
}

# Field 2: Set input restriction (custom data validation type)
$cfgFields += @{
    id = $fields[1].id
    cellAddr = $fields[1].cellAddr
    type = 0
    value = ""
    inputRestriction = "WholeNumber"
}

# Field 3: Set font properties - change font name and size
$cfgFields += @{
    id = $fields[2].id
    cellAddr = $fields[2].cellAddr
    type = 0
    value = ""
    fontName = "Arial"
    fontSize = 14
    fontColor = "FFFF0000"  # Red
}

# Field 4: Set alignment explicitly (should NOT leak to other fields)
$cfgFields += @{
    id = $fields[3].id
    cellAddr = $fields[3].cellAddr
    type = 0
    value = ""
    hAlign = "center"
    vAlign = "center"
}

# Field 5: No changes (default)
$cfgFields += @{
    id = $fields[4].id
    cellAddr = $fields[4].cellAddr
    type = 0
    value = ""
}

# Field 6: All properties
$cfgFields += @{
    id = $fields[5].id
    cellAddr = $fields[5].cellAddr
    type = 0
    value = ""
    placeholder = "Type here"
    defaultValue = "Default"
    required = $true
    maxLength = 100
    fontName = "Times New Roman"
    fontSize = 12
    fontColor = "FF0000FF"  # Blue
    hAlign = "right"
    vAlign = "top"
}

# ── Cycle 2: Export with configured fields, re-upload ──
Write-Host "=== Cycle 2: Export with config, re-upload ===" -ForegroundColor Cyan
$c2_bytes = Export-Xlsx $r.sessionId $r.pages $cfgFields
$c2_path = "$OutputDir\test_c2.xlsx"
[System.IO.File]::WriteAllBytes($c2_path, $c2_bytes)
$r2 = Upload-File $c2_path
Write-Host "  PaperLessConfig: $(if ($r2.paperLessConfig) { 'present' } else { 'null' })"

if ($r2.paperLessConfig) {
    foreach ($sheet in $r2.paperLessConfig.sheets) {
        foreach ($f in $sheet.fields) {
            $s = if ($f.style) { "font=$($f.style.font.name)/$($f.style.font.sizePt) color=$($f.style.font.colorArgb) hAlign=$($f.style.alignment.horizontal) vAlign=$($f.style.alignment.vertical)" } else { "no-style" }
            $c = if ($f.config) { "req=$($f.config.required) ml=$($f.config.maxLength) pl='$($f.config.placeholder)' dv='$($f.config.defaultValue)' rest=$($f.config.inputRestriction)" } else { "no-config" }
            Write-Host "    id=$($f.id) type=$($f.type) $s`n             $c"
        }
    }
} else {
    Write-Host "  FAIL: No PaperLessConfig!" -ForegroundColor Red
}

# ── Cycle 3: Export again, re-upload, verify config persisted ──
Write-Host "=== Cycle 3: Export again, verify config persisted ===" -ForegroundColor Cyan
$c3_bytes = Export-Xlsx $r2.sessionId $r2.pages $cfgFields  # Same configured fields
$c3_path = "$OutputDir\test_c3.xlsx"
[System.IO.File]::WriteAllBytes($c3_path, $c3_bytes)
$r3 = Upload-File $c3_path
Write-Host "  PaperLessConfig: $(if ($r3.paperLessConfig) { 'present' } else { 'null' })"

# Compare C2 vs C3 PaperLessConfig
if ($r2.paperLessConfig -and $r3.paperLessConfig) {
    $identical = $true
    $f2 = $r2.paperLessConfig.sheets[0].fields
    $f3 = $r3.paperLessConfig.sheets[0].fields
    
    for ($i = 0; $i -lt $f2.Count; $i++) {
        $aJson = ($f2[$i] | ConvertTo-Json -Depth 5 -Compress)
        $bJson = ($f3[$i] | ConvertTo-Json -Depth 5 -Compress)
        if ($aJson -ne $bJson) {
            Write-Host "  DIFF in field $($f2[$i].id) (C2 vs C3):" -ForegroundColor Red
            Write-Host "    C2: $aJson"
            Write-Host "    C3: $bJson"
            $identical = $false
        }
    }
    if ($identical) { Write-Host "  C2 and C3 PaperLessConfig IDENTICAL!" -ForegroundColor Green }
}

# ── Cycle 4: Export WITHOUT configured fields (simulate "no changes made"), verify config not lost ──
Write-Host "=== Cycle 4: Export WITHOUT field config, verify config not lost ===" -ForegroundColor Cyan
$c4_bytes = Export-Xlsx $r3.sessionId $r3.pages $null  # No configured fields
$c4_path = "$OutputDir\test_c4.xlsx"
[System.IO.File]::WriteAllBytes($c4_path, $c4_bytes)
$r4 = Upload-File $c4_path
Write-Host "  PaperLessConfig: $(if ($r4.paperLessConfig) { 'present' } else { 'null' })"

$allSurvived = $true
if ($r4.paperLessConfig) {
    foreach ($sheet in $r4.paperLessConfig.sheets) {
        foreach ($f in $sheet.fields) {
            $s = if ($f.style) { "font=$($f.style.font.name)/$($f.style.font.sizePt) color=$($f.style.font.colorArgb)" } else { "no-style" }
            $c = if ($f.config) { "req=$($f.config.required) ml=$($f.config.maxLength) pl='$($f.config.placeholder)' dv='$($f.config.defaultValue)' rest=$($f.config.inputRestriction)" } else { "no-config" }
            
            # Verify Field 1 had placeholder/defaultValue
            if ($f.id -eq $cfgFields[0].id) {
                if (-not $f.config -or $f.config.placeholder -ne "Enter name here") { Write-Host "  FAIL: Field 1 placeholder lost! Got: '$($f.config.placeholder)'" -ForegroundColor Red; $allSurvived = $false }
                if (-not $f.config -or $f.config.defaultValue -ne "John Doe") { Write-Host "  FAIL: Field 1 defaultValue lost! Got: '$($f.config.defaultValue)'" -ForegroundColor Red; $allSurvived = $false }
                if (-not $f.config -or $f.config.required -ne "True") { Write-Host "  FAIL: Field 1 required lost! Got: '$($f.config.required)'" -ForegroundColor Red; $allSurvived = $false }
            }
            
            # Verify Field 3 had font changes
            if ($f.id -eq $cfgFields[2].id) {
                if (-not $f.style -or $f.style.font.name -ne "Arial") { Write-Host "  FAIL: Field 3 font name lost! Got: '$($f.style.font.name)'" -ForegroundColor Red; $allSurvived = $false }
                if (-not $f.style -or $f.style.font.sizePt -ne 14) { Write-Host "  FAIL: Field 3 font size lost! Got: '$($f.style.font.sizePt)'" -ForegroundColor Red; $allSurvived = $false }
            }
            
            # Verify Field 4 alignment didn't leak to other fields
            if ($f.id -eq $cfgFields[3].id) {
                if (-not $f.style -or $f.style.alignment.horizontal -ne "center") { Write-Host "  FAIL: Field 4 hAlign lost! Got: '$($f.style.alignment.horizontal)'" -ForegroundColor Red; $allSurvived = $false }
            }
            if ($f.id -eq $cfgFields[4].id) {
                # Field 4's alignment should NOT appear on Field 5 (unconfigured)
                if ($f.style -and $f.style.alignment -and $f.style.alignment.horizontal -eq "center") {
                    Write-Host "  FAIL: Field 5 got alignment from Field 4 (leak)!" -ForegroundColor Red; $allSurvived = $false
                }
            }
            
            Write-Host "    id=$($f.id) $s`n             $c"
        }
    }
    if ($allSurvived) {
        Write-Host "  ALL CONFIG SURVIVED!" -ForegroundColor Green
    }
} else {
    Write-Host "  FAIL: No PaperLessConfig!" -ForegroundColor Red
}

# ── Dump PaperLessConfig XML from the final file ──
Write-Host "=== PaperLessConfig XML from final file ===" -ForegroundColor Cyan
try {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($c4_path)
    $entry = $zip.GetEntry("xl/worksheets/paperlessconfig.xml")
    if ($entry) {
        $reader = New-Object System.IO.StreamReader($entry.Open())
        $xml = $reader.ReadToEnd()
        $reader.Close()
        Write-Host $xml
    } else {
        Write-Host "  No paperlessconfig.xml entry found!"
    }
    $zip.Dispose()
} catch {
    Write-Host "  Could not read paperlessconfig.xml: $_"
}

Write-Host "=== Done ===" -ForegroundColor Cyan
