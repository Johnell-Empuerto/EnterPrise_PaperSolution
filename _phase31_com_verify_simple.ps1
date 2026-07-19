# Phase X.31 Q2: COM PageSetup Verification
$ProjectDir = "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise"

$workbookPaths = @{
    "ConMas" = [System.IO.Path]::Combine($ProjectDir, "test_conmas_output.xlsx")
    "Our Output" = [System.IO.Path]::Combine($ProjectDir, "test_our_output.xlsx")
}

Write-Host "PHASE X.31 Q2 - COM PageSetup Verification"
Write-Host "============================================"

# Cleanup
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

foreach ($pair in $workbookPaths.GetEnumerator()) {
    $label = $pair.Key
    $path = $pair.Value
    Write-Host ""
    Write-Host "--- WORKBOOK: $label ---"
    Write-Host "  Path: $path"
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    try {
        $wb = $excel.Workbooks.Open($path)
        Write-Host "  ActivePrinter: $($excel.ActivePrinter)"
        Write-Host "  Sheet count: $($wb.Worksheets.Count)"
        
        for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
            $ws = $wb.Worksheets[$i]
            $ps = $ws.PageSetup
            
            Write-Host "`n  Sheet [$i]: $($ws.Name) (Visible=$($ws.Visible))"
            Write-Host "    PrintArea: $($ps.PrintArea)"
            Write-Host "    CenterH: $($ps.CenterHorizontally)"
            Write-Host "    CenterV: $($ps.CenterVertically)"
            
            $lm = [Math]::Round([double]$ps.LeftMargin, 4)
            $rm = [Math]::Round([double]$ps.RightMargin, 4)
            $tm = [Math]::Round([double]$ps.TopMargin, 4)
            $bm = [Math]::Round([double]$ps.BottomMargin, 4)
            $hm = [Math]::Round([double]$ps.HeaderMargin, 4)
            $fm = [Math]::Round([double]$ps.FooterMargin, 4)
            
            Write-Host "    LeftMargin: $lm pt ($([Math]::Round($lm/72*2.54, 2)) cm)"
            Write-Host "    RightMargin: $rm pt ($([Math]::Round($rm/72*2.54, 2)) cm)"
            Write-Host "    TopMargin: $tm pt ($([Math]::Round($tm/72*2.54, 2)) cm)"
            Write-Host "    BottomMargin: $bm pt ($([Math]::Round($bm/72*2.54, 2)) cm)"
            Write-Host "    HeaderMargin: $hm pt"
            Write-Host "    FooterMargin: $fm pt"
            Write-Host "    Orientation: $($ps.Orientation)"
            Write-Host "    Zoom: $($ps.Zoom)"
            Write-Host "    FitToPagesWide: $($ps.FitToPagesWide)"
            Write-Host "    FitToPagesTall: $($ps.FitToPagesTall)"
            
            [Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
        }
        $wb.Close($false)
    }
    catch {
        Write-Host "  ERROR: $($_.Exception.Message)"
    }
    finally {
        $excel.Quit()
        if ($wb) { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

Write-Host ""
Write-Host "COM verification complete."
