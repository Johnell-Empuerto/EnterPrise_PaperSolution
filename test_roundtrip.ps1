# PaperLessConfig Round-Trip Test Script (PowerShell 5.1 compatible)

param(
    [string]$ApiBase = "http://localhost:5090",
    [string]$InputFile = "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx",
    [int]$MaxCycles = 2
)

$ErrorActionPreference = "Stop"

Write-Host "=== PaperLessConfig Round-Trip Test ===" -ForegroundColor Cyan
Write-Host "API Base: $ApiBase"
Write-Host "Input File: $InputFile"
Write-Host "Max Cycles: $MaxCycles"
Write-Host ""

$timestamp = (Get-Date -Format "yyyyMMdd_HHmmss")
$cycleDir = "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\test_output_$timestamp"
New-Item -ItemType Directory -Path $cycleDir -Force | Out-Null

$uploadUrl = "$ApiBase/api/form/upload-preview"
$exportUrl = "$ApiBase/api/form/save-edited"

# ── Upload helper ──
function Upload-File {
    param([string]$FilePath)
    $boundary = [System.Guid]::NewGuid().ToString()
    $fileName = Split-Path $FilePath -Leaf
    $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    
    $sb = New-Object System.Text.StringBuilder
    $sb.Append("--$boundary`r`n") | Out-Null
    $sb.Append("Content-Disposition: form-data; name=`"file`"; filename=`"$fileName`"`r`n") | Out-Null
    $sb.Append("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`r`n`r`n") | Out-Null
    
    $headerPart = $sb.ToString()
    $footerPart = "`r`n--$boundary--`r`n"
    
    $headerBytes = [System.Text.Encoding]::UTF8.GetBytes($headerPart)
    $footerBytes = [System.Text.Encoding]::UTF8.GetBytes($footerPart)
    
    $fullBody = New-Object byte[] ($headerBytes.Length + $fileBytes.Length + $footerBytes.Length)
    [System.Buffer]::BlockCopy($headerBytes, 0, $fullBody, 0, $headerBytes.Length)
    [System.Buffer]::BlockCopy($fileBytes, 0, $fullBody, $headerBytes.Length, $fileBytes.Length)
    [System.Buffer]::BlockCopy($footerBytes, 0, $fullBody, $headerBytes.Length + $fileBytes.Length, $footerBytes.Length)
    
    $response = Invoke-RestMethod -Uri $uploadUrl -Method Post `
        -ContentType "multipart/form-data; boundary=$boundary" -Body $fullBody
    return $response
}

# ── Step 1: Initial Upload ──
Write-Host "--- Cycle 0: Initial Upload ---" -ForegroundColor Magenta
Write-Host "Uploading: $InputFile" -ForegroundColor Yellow

try {
    $response = Upload-File -FilePath $InputFile
} catch {
    Write-Host "Upload failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

$sessionId = ""
if ($response.sessionId) { $sessionId = $response.sessionId }
if (-not $sessionId -and $response.data -and $response.data.sessionId) { $sessionId = $response.data.sessionId }

if (-not $sessionId) {
    Write-Host "FAILED: No sessionId in response" -ForegroundColor Red
    Write-Host ($response | ConvertTo-Json -Depth 5) -ForegroundColor Gray
    exit 1
}

Write-Host "Session ID: $sessionId" -ForegroundColor Green

$pages = $response.pages
if (-not $pages) { $pages = @() }

# ── Cycle Loop ──
for ($cycle = 1; $cycle -le $MaxCycles; $cycle++) {
    Write-Host ""
    Write-Host "=== Cycle $cycle ===" -ForegroundColor Magenta
    
    # ── Inspect PaperLessConfig ──
    $plConfig = $response.paperLessConfig
    if ($plConfig -and $plConfig.sheets) {
        Write-Host "PaperLessConfig found:" -ForegroundColor Green
        foreach ($sheet in $plConfig.sheets) {
            Write-Host "  Sheet '$($sheet.name)': $($sheet.fields.Count) fields" -ForegroundColor Gray
            foreach ($field in $sheet.fields) {
                Write-Host "    Field id='$($field.id)' type='$($field.type)'" -ForegroundColor Gray
                $style = $field.style
                if ($style) {
                    $align = $style.alignment
                    $hasAlign = $align -and ($align.horizontal -or $align.vertical)
                    if ($hasAlign) {
                        Write-Host "      Alignment: h='$($align.horizontal)' v='$($align.vertical)'" -ForegroundColor Gray
                    }
                    $font = $style.font
                    if ($font) {
                        Write-Host "      Font: '$($font.name)' size=$($font.sizePt) bold=$($font.bold) color='$($font.colorArgb)'" -ForegroundColor Gray
                    }
                    $fill = $style.fill
                    if ($fill) {
                        Write-Host "      Fill: color='$($fill.colorArgb)'" -ForegroundColor Gray
                    }
                }
                $cfg = $field.config
                if ($cfg) {
                    Write-Host "      Config: required=$($cfg.required) maxLength=$($cfg.maxLength) readOnly=$($cfg.readOnly) hidden=$($cfg.hidden) placeholder='$($cfg.placeholder)' default='$($cfg.defaultValue)' validate=$($cfg.validateOnEditing) restriction='$($cfg.inputRestriction)'" -ForegroundColor Gray
                }
                if (-not $style -and -not $cfg) {
                    Write-Host "      (no style or config)" -ForegroundColor Gray
                }
            }
        }
    } else {
        Write-Host "No PaperLessConfig in response (fresh upload)" -ForegroundColor Yellow
    }
    
    # ── Build export payload ──
    $wbDef = @{
        info = @{ title = "RoundTrip Test" }
        sourceFileName = $sessionId
        sessionId = $sessionId
        sheets = @()
    }
    
    foreach ($page in $pages) {
        $sheet = @{
            name = $page.sheetName
            index = 0
            fields = @()
        }
        
        $pageFields = $page.fields
        if (-not $pageFields) { $pageFields = @() }
        
        foreach ($field in $pageFields) {
            $fieldId = $field.id
            $cellAddr = $field.cellAddr
            
            $wbField = @{
                id = $fieldId
                cell = @{ address = $cellAddr; rowIndex = 1 }
                name = if ($field.name) { $field.name } else { "" }
                type = 0
                value = ""
            }
            
            $sheet.fields += $wbField
        }
        
        $wbDef.sheets += $sheet
    }
    
    # wbDef IS the WorkbookDefinition object directly (no wrapping)
    $jsonBody = $wbDef | ConvertTo-Json -Depth 10
    Write-Host "Exporting... ($([Math]::Round($jsonBody.Length / 1024, 1)) KB)" -ForegroundColor Yellow
    
    $exportFile = "$cycleDir\cycle_$cycle.xlsx"
    
    try {
        $exportResponse = Invoke-WebRequest -Uri $exportUrl -Method Post `
            -ContentType "application/json" -Body $jsonBody
        
        [System.IO.File]::WriteAllBytes($exportFile, $exportResponse.Content)
        $fileSize = [Math]::Round((Get-Item $exportFile).Length / 1KB, 1)
        Write-Host "Exported to: $exportFile ($fileSize KB)" -ForegroundColor Green
        
        Write-Host "Diagnostic headers:" -ForegroundColor Gray
        foreach ($key in $exportResponse.Headers.Keys) {
            if ($key.StartsWith("X-")) {
                Write-Host "  $key = $($exportResponse.Headers[$key])" -ForegroundColor Gray
            }
        }
    }
    catch {
        Write-Host "Export failed: $($_.Exception.Message)" -ForegroundColor Red
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $errorBody = $reader.ReadToEnd()
            Write-Host "Error body: $errorBody" -ForegroundColor Red
        }
        break
    }
    
    # ── Upload exported file ──
    Write-Host "Uploading exported file for next cycle..." -ForegroundColor Yellow
    
    try {
        $response = Upload-File -FilePath $exportFile
        if ($response.sessionId) { $sessionId = $response.sessionId }
        Write-Host "New session ID: $sessionId" -ForegroundColor Green
        
        $pages = $response.pages
        if (-not $pages) { $pages = @() }
    }
    catch {
        Write-Host "Re-upload failed: $($_.Exception.Message)" -ForegroundColor Red
        break
    }
}

Write-Host ""
Write-Host "=== Test Complete ===" -ForegroundColor Cyan
Write-Host "Output directory: $cycleDir" -ForegroundColor Cyan
