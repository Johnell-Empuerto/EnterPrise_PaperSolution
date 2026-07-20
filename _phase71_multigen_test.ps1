# Phase 7.1 - Multi-Generation Export Test via Excel COM Interop
$ErrorActionPreference = "Stop"
$ProjectDir = "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise"
$OutputDir = Join-Path $ProjectDir "_phase71_output"

if (!(Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$OriginalPath = Join-Path $ProjectDir "Investigation_546\original.xlsx"
$OriginalCopy = Join-Path $OutputDir "Original.xlsx"
$Export1Path = Join-Path $OutputDir "Export1.xlsx"
$Export2Path = Join-Path $OutputDir "Export2.xlsx"
$Export3Path = Join-Path $OutputDir "Export3.xlsx"

$paths = @(
    @{Label="Original"; Path=$OriginalCopy},
    @{Label="Export1";  Path=$Export1Path},
    @{Label="Export2";  Path=$Export2Path},
    @{Label="Export3";  Path=$Export3Path}
)

Write-Host "PHASE 7.1 - MULTI-GENERATION EXPORT TEST"
Write-Host "Using: Excel COM Interop (same engine as ConMas)"
Write-Host "Input: $OriginalPath"
Write-Host "Output: $OutputDir"

Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

Copy-Item -Path $OriginalPath -Destination $OriginalCopy -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $currentPath = $OriginalCopy
    
    for ($gen = 1; $gen -le 3; $gen++) {
        $outputPath = $paths[$gen].Path
        $label = $paths[$gen].Label
        
        Write-Host "Generation $gen => $label"
        Write-Host "  Input: $currentPath"
        Write-Host "  Output: $outputPath"
        
        $wb = $excel.Workbooks.Open($currentPath)
        Start-Sleep -Milliseconds 200
        
        Write-Host "  Sheets: " + $wb.Worksheets.Count
        for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
            $ws = $wb.Worksheets[$i]
            Write-Host "    [$i] " + $ws.Name + " visible=" + $ws.Visible
            [Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
        }
        
        $wb.SaveAs($outputPath, 51)  # xlOpenXMLWorkbook = 51
        $wb.Close($false)
        [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        
        Write-Host "  Output size: " + (Get-Item $outputPath).Length + " bytes"
        
        $currentPath = $outputPath
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Start-Sleep -Seconds 1
    }
}
catch {
    Write-Host "ERROR: $_"
}
finally {
    $excel.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "FILE HASHES"
foreach ($p in $paths) {
    if (Test-Path $p.Path) {
        $hash = (Get-FileHash -LiteralPath $p.Path -Algorithm SHA256).Hash
        $size = (Get-Item $p.Path).Length
        Write-Host ("  {0,-10} {1,8} bytes  {2}" -f $p.Label, $size, $hash)
    }
    else {
        Write-Host ("  {0,-10} NOT FOUND" -f $p.Label)
    }
}

Write-Host ""
Write-Host "Generation complete. Run forensic comparison next."
