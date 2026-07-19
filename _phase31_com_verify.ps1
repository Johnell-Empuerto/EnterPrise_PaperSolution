<#
Phase X.31 Q2: COM Verification of PageSetup
Reads PageSetup properties from both workbooks via Excel COM Interop.
Then makes Excel visible for UI verification.
#>

$ProjectDir = "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise"

$Workbooks = @{
    "Our Output" = Join-Path $ProjectDir "test_our_output.xlsx"
    "ConMas" = Join-Path $ProjectDir "test_conmas_output.xlsx"
}

Write-Host "="*80
Write-Host "PHASE X.31 Q2 - COM PageSetup Verification"
Write-Host "="*80

# Cleanup any orphan Excel processes
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

foreach ($label in $Workbooks.Keys) {
    $path = $Workbooks[$label]
    Write-Host "`n" + ("─"*80)
    Write-Host "WORKBOOK: $label"
    Write-Host "  Path: $path"
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    try {
        $wb = $excel.Workbooks.Open($path.FullName)
        
        Write-Host "`n  ActivePrinter: '$($excel.ActivePrinter)'"
        Write-Host "  Workbook sheets: $($wb.Worksheets.Count)"
        
        for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
            $ws = $wb.Worksheets[$i]
            $ps = $ws.PageSetup
            
            Write-Host "`n  ── Sheet [$i]: $($ws.Name) (Visible=$($ws.Visible)) ──"
            Write-Host "    PrintArea:       '$($ps.PrintArea)'"
            Write-Host "    CenterH:          $($ps.CenterHorizontally)"
            Write-Host "    CenterV:          $($ps.CenterVertically)"
            Write-Host "    LeftMargin:       $([Math]::Round([double]$ps.LeftMargin, 4)) pt ($([Math]::Round([double]$ps.LeftMargin/72*2.54, 2)) cm)"
            Write-Host "    RightMargin:      $([Math]::Round([double]$ps.RightMargin, 4)) pt ($([Math]::Round([double]$ps.RightMargin/72*2.54, 2)) cm)"
            Write-Host "    TopMargin:        $([Math]::Round([double]$ps.TopMargin, 4)) pt ($([Math]::Round([double]$ps.TopMargin/72*2.54, 2)) cm)"
            Write-Host "    BottomMargin:     $([Math]::Round([double]$ps.BottomMargin, 4)) pt ($([Math]::Round([double]$ps.BottomMargin/72*2.54, 2)) cm)"
            Write-Host "    HeaderMargin:     $([Math]::Round([double]$ps.HeaderMargin, 4)) pt"
            Write-Host "    FooterMargin:     $([Math]::Round([double]$ps.FooterMargin, 4)) pt"
            Write-Host "    Orientation:      $($ps.Orientation)"
            Write-Host "    PaperSize:        $($ps.PaperSize)"
            Write-Host "    Zoom:             $($ps.Zoom)"
            Write-Host "    FitToPagesWide:   $($ps.FitToPagesWide)"
            Write-Host "    FitToPagesTall:   $($ps.FitToPagesTall)"
            Write-Host "    FirstPageNumber:  $($ps.FirstPageNumber)"
            Write-Host "    Order:            $($ps.Order)"
            
            [Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
        }
        
        $wb.Close($false)
    }
    catch {
        Write-Host "  ERROR: $($_.Exception.Message)"
    }
    finally {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

Write-Host "`n" + "="*80
Write-Host "COM verification complete."
Write-Host "="*80
Write-Host "`nNext: Open Excel VISIBLY for UI confirmation."

# ═══ UI VERIFICATION ═══
Write-Host "`n" + ("─"*80)
Write-Host "EXCEL UI VERIFICATION"
Write-Host ("─"*80)
Write-Host "Opening workbook for visual inspection..."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

# Open ConMas workbook (correct reference)
$conmasPath = $Workbooks["ConMas"]
$wbConMas = $excel.Workbooks.Open($conmasPath.FullName)
$wsConMas = $wbConMas.Worksheets(2)  # Sheet1
$wsConMas.Activate()

Write-Host "`n  Opened: ConMas workbook"
Write-Host "  Active sheet: $($excel.ActiveSheet.Name)"
Write-Host "  Please verify:"
Write-Host "    1. Page Layout -> Margins -> Custom Margins"
Write-Host "    2. Print Preview"
Write-Host "`n  Expected COM values (ConMas Sheet1):"
$ps = $wsConMas.PageSetup
Write-Host "    CenterH: $($ps.CenterHorizontally)"
Write-Host "    CenterV: $($ps.CenterVertically)"
Write-Host "    LeftMargin: $([Math]::Round([double]$ps.LeftMargin,4)) pt"
Write-Host "    TopMargin: $([Math]::Round([double]$ps.TopMargin,4)) pt"
Write-Host "    PrintArea: '$($ps.PrintArea)'"

Write-Host "`n  Press any key to open Our Output workbook..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# Now open our output
$ourPath = $Workbooks["Our Output"]
$wbOur = $excel.Workbooks.Open($ourPath.FullName)
$wsOur = $wbOur.Worksheets(1)  # Sheet1 (first sheet)
$wsOur.Activate()

Write-Host "`n  Now showing: Our Output workbook"
Write-Host "  Active sheet: $($excel.ActiveSheet.Name)"
Write-Host "  Please verify:"
Write-Host "    1. Page Layout -> Margins -> Custom Margins"
Write-Host "    2. Print Preview"
Write-Host "`n  Expected COM values (Our Output Sheet1):"
$ps2 = $wsOur.PageSetup
Write-Host "    CenterH: $($ps2.CenterHorizontally)"
Write-Host "    CenterV: $($ps2.CenterVertically)"
Write-Host "    LeftMargin: $([Math]::Round([double]$ps2.LeftMargin,4)) pt"
Write-Host "    TopMargin: $([Math]::Round([double]$ps2.TopMargin,4)) pt"
Write-Host "    PrintArea: '$($ps2.PrintArea)'"

Write-Host "`n  Press any key to close Excel..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

$wbConMas.Close($false)
$wbOur.Close($false)
$excel.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($wsConMas) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($wsOur) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($wbConMas) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($wbOur) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "`nDone."
