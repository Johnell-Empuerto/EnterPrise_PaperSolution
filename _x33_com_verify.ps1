
# Phase X.33 - COM Verification (fresh workbook)
$ProjectDir = "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise"
$genPath = [System.IO.Path]::Combine($ProjectDir, "_x33_generated_output.xlsx")
$conmasPath = [System.IO.Path]::Combine($ProjectDir, "test_conmas_output.xlsx")

Write-Host "Phase X.33 - STEP 3+4: COM Verification"
Write-Host ""

function Read-PageSetup($wb, $label) {
    Write-Host "--- $label ---"
    Write-Host "Sheet count: $($wb.Worksheets.Count)"
    for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
        $ws = $wb.Worksheets[$i]
        $ps = $ws.PageSetup
        Write-Host ""
        Write-Host "Sheet [$i]: $($ws.Name) (Visible=$($ws.Visible))"
        Write-Host "  PrintArea:            $($ps.PrintArea)"
        Write-Host "  CenterHorizontally:   $($ps.CenterHorizontally)"
        Write-Host "  CenterVertically:     $($ps.CenterVertically)"
        $lm = [Math]::Round([double]$ps.LeftMargin, 4)
        $rm = [Math]::Round([double]$ps.RightMargin, 4)
        $tm = [Math]::Round([double]$ps.TopMargin, 4)
        $bm = [Math]::Round([double]$ps.BottomMargin, 4)
        $hm = [Math]::Round([double]$ps.HeaderMargin, 4)
        $fm = [Math]::Round([double]$ps.FooterMargin, 4)
        Write-Host "  LeftMargin:           $lm pt  ($([Math]::Round($lm/72*2.54,2)) cm)"
        Write-Host "  RightMargin:          $rm pt  ($([Math]::Round($rm/72*2.54,2)) cm)"
        Write-Host "  TopMargin:            $tm pt  ($([Math]::Round($tm/72*2.54,2)) cm)"
        Write-Host "  BottomMargin:         $bm pt  ($([Math]::Round($bm/72*2.54,2)) cm)"
        Write-Host "  HeaderMargin:         $hm pt"
        Write-Host "  FooterMargin:         $fm pt"
        Write-Host "  Orientation:          $($ps.Orientation)"
        Write-Host "  Zoom:                 $($ps.Zoom)"
        Write-Host "  FitToPagesWide:       $($ps.FitToPagesWide)"
        Write-Host "  FitToPagesTall:       $($ps.FitToPagesTall)"
        [Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
    }
}

# First open - read generated workbook
Write-Host ""
Write-Host "=== FIRST OPEN ==="
$excel1 = New-Object -ComObject Excel.Application
$excel1.Visible = $false
$excel1.DisplayAlerts = $false
$wb1 = $excel1.Workbooks.Open($genPath)
Read-PageSetup $wb1 "GENERATED (first open)"
$wb1.Close($false)
$excel1.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($wb1) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($excel1) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Start-Sleep -Seconds 2

# Second open (reopen) - read generated workbook again
Write-Host ""
Write-Host "=== REOPEN ==="
$excel2 = New-Object -ComObject Excel.Application
$excel2.Visible = $false
$excel2.DisplayAlerts = $false
$wb2 = $excel2.Workbooks.Open($genPath)
Read-PageSetup $wb2 "GENERATED (reopen)"
$wb2.Close($false)
$excel2.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($wb2) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($excel2) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Start-Sleep -Seconds 2

# Third open - read ConMas workbook for comparison
Write-Host ""
Write-Host "=== CONMAS REFERENCE ==="
$excel3 = New-Object -ComObject Excel.Application
$excel3.Visible = $false
$excel3.DisplayAlerts = $false
$wb3 = $excel3.Workbooks.Open($conmasPath)
Read-PageSetup $wb3 "CONMAS (reference)"
$wb3.Close($false)
$excel3.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($wb3) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($excel3) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "COM verification complete."
