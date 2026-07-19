# Phase X.35 - Export all 4 workbooks to PDF via Excel COM
$ErrorActionPreference = "Stop"

$wbs = @(
    @{N="Original"; P="formtest.xlsx"},
    @{N="ConMas"; P="test_conmas_output.xlsx"},
    @{N="Generated1"; P="_x34_generated.xlsx"},
    @{N="Generated2"; P="_x34_generated2.xlsx"}
)

$results = @()
$pdfs = @{}

Write-Host "Starting Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
Write-Host ("Excel version: " + $excel.Version)

foreach ($wbInfo in $wbs) {
    $name = $wbInfo.N
    $path = $wbInfo.P
    $absPath = (Resolve-Path $path).Path
    Write-Host ""
    Write-Host "===== PROCESSING: $name ====="
    Write-Host ("File: " + $absPath)

    $wb = $excel.Workbooks.Open($absPath)
    Start-Sleep -Milliseconds 300

    Write-Host ("Sheets: " + $wb.Worksheets.Count)
    Write-Host ("Active: " + $wb.ActiveSheet.Name)

    # ActivePrinter
    try { Write-Host ("Printer: " + $excel.ActivePrinter) } catch { Write-Host "Printer: ERROR" }

    $sheets = @()
    for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
        $ws = $wb.Worksheets[$i]
        $wsName = $ws.Name
        $ps = $ws.PageSetup
        
        $pa = ""; try { $pa = $ps.PrintArea } catch { $pa = "ERR" }
        $ch = $false; try { $ch = [bool]$ps.CenterHorizontally } catch { }
        $cv = $false; try { $cv = [bool]$ps.CenterVertically } catch { }
        $lm = 0.0; try { $lm = [double]$ps.LeftMargin } catch { }
        $rm = 0.0; try { $rm = [double]$ps.RightMargin } catch { }
        $tm = 0.0; try { $tm = [double]$ps.TopMargin } catch { }
        $bm = 0.0; try { $bm = [double]$ps.BottomMargin } catch { }
        $hm = 0.0; try { $hm = [double]$ps.HeaderMargin } catch { }
        $fm = 0.0; try { $fm = [double]$ps.FooterMargin } catch { }
        $orient = ""; try { $orient = [string]$ps.Orientation } catch { }
        $psize = -1; try { $psize = [int]$ps.PaperSize } catch { }
        $zoom = -1; try { $zoom = [int]$ps.Zoom } catch { }
        $fw = -1; try { $fw = [int]$ps.FitToPagesWide } catch { }
        $fh = -1; try { $fh = [int]$ps.FitToPagesTall } catch { }

        Write-Host ("  Sheet[" + $i + "]='" + $wsName + "': PA=" + $pa + " CH=" + $ch + " CV=" + $cv)
        Write-Host ("    LM=" + $lm.ToString("F4") + " RM=" + $rm.ToString("F4") + " TM=" + $tm.ToString("F4") + " BM=" + $bm.ToString("F4") + " HM=" + $hm.ToString("F4") + " FM=" + $fm.ToString("F4"))
        Write-Host ("    O=" + $orient + " PS=" + $psize + " Z=" + $zoom + " FW=" + $fw + " FH=" + $fh)

        $sheets += @{
            Name = $wsName
            Idx = $i
            PA = $pa
            CH = $ch
            CV = $cv
            LM = $lm
            RM = $rm
            TM = $tm
            BM = $bm
            HM = $hm
            FM = $fm
            O = $orient
            PS = $psize
            Z = $zoom
            FW = $fw
            FH = $fh
        }
        [Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
    }

    # Export to PDF
    $pdfPath = (Join-Path $PWD.Path ("_x35_pdf_" + $name + ".pdf"))
    Write-Host ("  Exporting PDF: " + $pdfPath)
    try {
        $wb.ExportAsFixedFormat(0, $pdfPath)
        if (Test-Path $pdfPath) {
            $len = (Get-Item $pdfPath).Length
            Write-Host ("  PDF OK: " + $len + " bytes")
            $pdfs[$name] = $pdfPath
        }
    } catch {
        Write-Host ("  PDF FAILED: " + $_.Exception.Message)
    }

    $wb.Close($false)
    [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
    $results += @{
        Name = $name
        Path = $absPath
        Sheets = $sheets
    }
}

$excel.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "=============================="
Write-Host "COMPARISON TABLE - Print PageSetup (Sheet1 only)"
Write-Host "=============================="

$header = "{0,-12} {1,-12} {2,-6} {3,-6} {4,-10} {5,-10} {6,-10} {7,-10} {8,-10} {9,-10} {10,-10} {11,-6} {12,-6} {13,-6}"
$fmt = "{0,-12} {1,-12} {2,-6} {3,-6} {4,-10:F4} {5,-10:F4} {6,-10:F4} {7,-10:F4} {8,-10:F4} {9,-10:F4} {10,-10} {11,-6} {12,-6} {13,-6}"

Write-Host ($header -f "Workbook", "PrintArea", "CH", "CV", "LM(pt)", "RM(pt)", "TM(pt)", "BM(pt)", "HM(pt)", "FM(pt)", "Orient", "PS", "Zoom", "FW")
Write-Host ("{0,-12} {1,-12} {2,-6} {3,-6} {4,-10} {5,-10} {6,-10} {7,-10} {8,-10} {9,-10} {10,-10} {11,-6} {12,-6} {13,-6}" -f "--------", "--------", "--", "--", "------", "------", "------", "------", "------", "------", "------", "--", "----", "--")

foreach ($r in $results) {
    # Find Sheet1 (content sheet, usually index 2)
    $s1 = $null
    foreach ($s in $r.Sheets) {
        if ($s.Name -eq "Sheet1" -or $s.Idx -eq 2) { $s1 = $s; break }
    }
    if (-not $s1) { $s1 = $r.Sheets[0] }
    
    $orientShort = $s1.O
    if ($orientShort.Length -gt 8) { $orientShort = $orientShort.Substring(0,8) }
    
    Write-Host ($fmt -f $r.Name, $s1.PA, $s1.CH, $s1.CV, $s1.LM, $s1.RM, $s1.TM, $s1.BM, $s1.HM, $s1.FM, $orientShort, $s1.PS, $s1.Z, $s1.FW)
}

Write-Host ""
Write-Host "=== PDFs ==="
foreach ($kv in $pdfs.GetEnumerator()) {
    Write-Host ("  " + $kv.Key + ": " + $kv.Value)
}
Write-Host ""
Write-Host "Phase X.35 Tests 1+2 COMPLETE"
