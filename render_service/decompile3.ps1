# Extract full IL for key methods and the Image type
$asm = [System.Reflection.Assembly]::LoadFrom('C:/Program Files (x86)/CIMTOPS CORPORATION/ConMas Designer/bin/LibExcelController.dll')

function Dump-MethodIL($method) {
    Write-Host "`nMethod: $($method.Name)"
    $pars = $method.GetParameters()
    foreach ($p in $pars) {
        Write-Host "  Param: $($p.ParameterType.FullName) $($p.Name)"
    }
    try {
        $body = $method.GetMethodBody()
        if ($body) {
            $il = $body.GetILAsByteArray()
            Write-Host "  IL size: $($il.Length) bytes"
            Write-Host "  MaxStack: $($body.MaxStackSize)"
            $locals = $body.LocalVariables
            Write-Host "  Locals: $($locals.Count)"
            $i_loc = 0
            foreach ($lv in $locals) {
                Write-Host "    [$i_loc] $($lv.LocalType.FullName)"
                $i_loc++
            }
            Write-Host "  IL bytes (full):"
            for ($i = 0; $i -lt $il.Length; $i += 16) {
                $end = [Math]::Min($i + 16, $il.Length)
                $hex = ($il[$i..($end-1)] | ForEach-Object { $_.ToString('X2') }) -join ' '
                Write-Host "    $($i.ToString('X4')): $hex"
            }
            # Check for exception handling clauses
            $exHandlers = $body.ExceptionHandlingClauses
            if ($exHandlers -and $exHandlers.Count -gt 0) {
                Write-Host "  Exception handlers: $($exHandlers.Count)"
                foreach ($eh in $exHandlers) {
                    Write-Host "    Flags=$($eh.Flags) TryOffset=$($eh.TryOffset) TryLength=$($eh.TryLength) HandlerOffset=$($eh.HandlerOffset) HandlerLength=$($eh.HandlerLength)"
                }
            }
        } else {
            Write-Host "  No method body"
        }
    } catch {
        Write-Host "  Error: $_"
    }
}

# === ExcelControllerInterop.CalcClusterSize (FULL) ===
$t = $asm.GetType('LibExcelController.ExcelControllerInterop')
Write-Host "==============================================="
Write-Host "CALC CLUSTER SIZE"
Write-Host "==============================================="
$m = $t.GetMethod('CalcClusterSize')
Dump-MethodIL $m

# === CalcClusterSize in the base class (ExcelControllerBase) if exists ===
$t_base = $asm.GetType('LibExcelController.ExcelControllerBase')
if ($t_base) {
    Write-Host "`n==============================================="
    Write-Host "EXCEL CONTROLLER BASE - ALL METHODS"
    Write-Host "==============================================="
    foreach ($m in $t_base.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance,DeclaredOnly')) {
        $name = $m.Name
        $pars = ($m.GetParameters() | ForEach-Object { "$($_.ParameterType.Name) $($_.Name)" }) -join ', '
        Write-Host "  [$($m.ReturnType.Name)] $($name)($pars)"
    }
}

# === GetFirstPagePrintArea (FULL) ===
$t3 = $asm.GetType('LibExcelController.ExcelLib.ExcelWorksheetCom')
Write-Host "`n==============================================="
Write-Host "GET FIRST PAGE PRINT AREA"
Write-Host "==============================================="
$m = $t3.GetMethod('GetFirstPagePrintArea')
Dump-MethodIL $m

# === ExistsPrintAreaOnePage (FULL) ===
Write-Host "`n==============================================="
Write-Host "EXISTS PRINT AREA ONE PAGE"
Write-Host "==============================================="
$m = $t3.GetMethod('ExistsPrintAreaOnePage')
Dump-MethodIL $m

# === ExportPdf2010 (FULL) ===
$t_img = $asm.GetType('LibExcelController.Lib.ImageUtility')
Write-Host "`n==============================================="
Write-Host "EXPORT PDF 2010"
Write-Host "==============================================="
$m = $t_img.GetMethod('ExportPdf2010')
Dump-MethodIL $m

# === ExportPdf2007 (FULL) ===
Write-Host "`n==============================================="
Write-Host "EXPORT PDF 2007"
Write-Host "==============================================="
$m = $t_img.GetMethod('ExportPdf2007')
Dump-MethodIL $m

# === ExportPdf (delegate) ===
Write-Host "`n==============================================="
Write-Host "EXPORT PDF"
Write-Host "==============================================="
$m = $t_img.GetMethod('ExportPdf')
Dump-MethodIL $m

# === GetClusterSize (FULL) ===
Write-Host "`n==============================================="
Write-Host "GET CLUSTER SIZE"
Write-Host "==============================================="
$m = $t_img.GetMethod('GetClusterSize')
Dump-MethodIL $m

# === SortClusters (FULL) ===
Write-Host "`n==============================================="
Write-Host "SORT CLUSTERS"
Write-Host "==============================================="
$m = $t_img.GetMethod('SortClusters')
Dump-MethodIL $m

# === MakeCluster method to see how ranges are accessed ===
Write-Host "`n==============================================="
Write-Host "MAKE CLUSTER"
Write-Host "==============================================="
$m = $t.GetMethod('MakeCluster')
Dump-MethodIL $m

# === Check ExcelWorksheetBase for page/print area methods ===
$t_ws = $asm.GetType('LibExcelController.ExcelLib.ExcelWorksheetBase')
Write-Host "`n==============================================="
Write-Host "EXCEL WORKSHEET BASE (all methods)"
Write-Host "==============================================="
if ($t_ws) {
    foreach ($m in $t_ws.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance,DeclaredOnly')) {
        $name = $m.Name
        $pars = ($m.GetParameters() | ForEach-Object { "$($_.ParameterType.Name) $($_.Name)" }) -join ', '
        Write-Host "  [$($m.ReturnType.Name)] $($name)($pars)"
    }
}
