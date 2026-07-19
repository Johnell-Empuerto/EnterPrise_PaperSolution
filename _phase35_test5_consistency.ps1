# Phase X.35 - Test 5: Print Engine Consistency
# Re-open Generated1 workbook and re-export to PDF
# Compare PDF hashes with first export

$ErrorActionPreference = "Stop"

Write-Host "=== Phase X.35 Test 5: Print Engine Consistency ==="
Write-Host ""

$workbooks = @(
    @{N="Generated1_R1"; P=(Resolve-Path "_x34_generated.xlsx").Path},
    @{N="Generated2_R1"; P=(Resolve-Path "_x34_generated2.xlsx").Path}
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($wbInfo in $workbooks) {
    $name = $wbInfo.N
    $path = $wbInfo.P
    Write-Host ("Opening: " + $name)
    Write-Host ("  Path: " + $path)
    
    $wb = $excel.Workbooks.Open($path)
    Start-Sleep -Milliseconds 300
    
    $pdfPath = (Join-Path $PWD.Path ("_x35_pdf_" + $name + ".pdf"))
    Write-Host ("  Exporting PDF: " + $pdfPath)
    
    $wb.ExportAsFixedFormat(0, $pdfPath)
    
    if (Test-Path $pdfPath) {
        $len = (Get-Item $pdfPath).Length
        $md5 = (Get-FileHash -LiteralPath $pdfPath -Algorithm MD5).Hash
        $sha = (Get-FileHash -LiteralPath $pdfPath -Algorithm SHA256).Hash
        Write-Host ("  PDF OK: " + $len + " bytes")
        Write-Host ("  MD5: " + $md5)
        Write-Host ("  SHA256: " + $sha)
    }
    
    $wb.Close($false)
    [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
}

$excel.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "=== Comparing with first-pass PDFs ==="
Write-Host ""

$firstPass = @(
    @{N="Generated1"; P=(Resolve-Path "_x35_pdf_Generated1.pdf").Path},
    @{N="Generated2"; P=(Resolve-Path "_x35_pdf_Generated2.pdf").Path}
)

foreach ($fp in $firstPass) {
    $name = $fp.N
    $rePath = (Resolve-Path ("_x35_pdf_" + $name + "_R1.pdf")).Path
    $firstPath = $fp.P
    
    Write-Host ("--- " + $name + " ---")
    $fpLen = (Get-Item $firstPath).Length
    $rpLen = (Get-Item $rePath).Length
    Write-Host ("  First: " + $fpLen + " bytes")
    Write-Host ("  Re-export: " + $rpLen + " bytes")
    
    $fpMd5 = (Get-FileHash -LiteralPath $firstPath -Algorithm MD5).Hash
    $rpMd5 = (Get-FileHash -LiteralPath $rePath -Algorithm MD5).Hash
    Write-Host ("  First MD5: " + $fpMd5)
    Write-Host ("  Re-export MD5: " + $rpMd5)
    
    if ($fpMd5 -eq $rpMd5) {
        Write-Host ("  >>> CONSISTENT: Hash match!")
    } else {
        Write-Host ("  >>> DIFFERENT: Hash mismatch - engine may be non-deterministic")
    }
}

Write-Host ""
Write-Host "Phase X.35 Test 5 COMPLETE"
