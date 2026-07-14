# Extract IL bytecode from ImageUtility methods
$asm = [System.Reflection.Assembly]::LoadFrom('C:/Program Files (x86)/CIMTOPS CORPORATION/ConMas Designer/bin/LibExcelController.dll')

$t = $asm.GetType('LibExcelController.Lib.ImageUtility')
Write-Host '=== ImageUtility ==='
foreach ($m in $t.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance,DeclaredOnly')) {
    Write-Host "`nMethod: $($m.Name)"
    $pars = $m.GetParameters()
    foreach ($p in $pars) {
        Write-Host "  Param: $($p.ParameterType.Name) $($p.Name)"
    }
    try {
        $body = $m.GetMethodBody()
        if ($body) {
            $il = $body.GetILAsByteArray()
            Write-Host "  IL size: $($il.Length) bytes"
            # Print IL bytes as hex in groups of 16
            for ($i = 0; $i -lt [Math]::Min(512, $il.Length); $i += 16) {
                $end = [Math]::Min($i + 16, $il.Length)
                $hex = ($il[$i..($end-1)] | ForEach-Object { $_.ToString('X2') }) -join ' '
                Write-Host "    $($i.ToString('X4')): $hex"
            }
            $locals = $body.LocalVariables
            Write-Host "  Locals: $($locals.Count)"
            foreach ($lv in $locals) {
                Write-Host "    $($lv.LocalType.FullName)"
            }
            # Check for exception handling clauses
            $exHandlers = $body.ExceptionHandlingClauses
            if ($exHandlers) {
                Write-Host "  Exception handlers: $($exHandlers.Count)"
            }
        } else {
            Write-Host "  No method body (abstract/external)"
        }
    } catch {
        Write-Host "  Error: $_"
    }
}

# Also check the ExcelControllerInterop methods
$t2 = $asm.GetType('LibExcelController.ExcelControllerInterop')
Write-Host "`n=== ExcelControllerInterop ==="
foreach ($m in $t2.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance,DeclaredOnly')) {
    Write-Host "`nMethod: $($m.Name)"
    $pars = $m.GetParameters()
    foreach ($p in $pars) {
        Write-Host "  Param: $($p.ParameterType.Name) $($p.Name)"
    }
    try {
        $body = $m.GetMethodBody()
        if ($body) {
            $il = $body.GetILAsByteArray()
            Write-Host "  IL size: $($il.Length) bytes"
            for ($i = 0; $i -lt [Math]::Min(128, $il.Length); $i += 16) {
                $end = [Math]::Min($i + 16, $il.Length)
                $hex = ($il[$i..($end-1)] | ForEach-Object { $_.ToString('X2') }) -join ' '
                Write-Host "    $($i.ToString('X4')): $hex"
            }
            $locals = $body.LocalVariables
            Write-Host "  Locals: $($locals.Count)"
            foreach ($lv in $locals) {
                Write-Host "    $($lv.LocalType.FullName)"
            }
        }
    } catch {
        Write-Host "  Error: $_"
    }
}

# Check ExcelWorksheetCom for GetFirstPagePrintArea
$t3 = $asm.GetType('LibExcelController.ExcelLib.ExcelWorksheetCom')
Write-Host "`n=== ExcelWorksheetCom ==="
foreach ($m in $t3.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance,DeclaredOnly')) {
    if ($m.Name -match 'GetFirst|Exists|PrintArea') {
        Write-Host "`nMethod: $($m.Name)"
        try {
            $body = $m.GetMethodBody()
            if ($body) {
                $il = $body.GetILAsByteArray()
                Write-Host "  IL size: $($il.Length) bytes"
                for ($i = 0; $i -lt [Math]::Min(128, $il.Length); $i += 16) {
                    $end = [Math]::Min($i + 16, $il.Length)
                    $hex = ($il[$i..($end-1)] | ForEach-Object { $_.ToString('X2') }) -join ' '
                    Write-Host "    $($i.ToString('X4')): $hex"
                }
                $locals = $body.LocalVariables
                Write-Host "  Locals: $($locals.Count)"
                foreach ($lv in $locals) {
                    Write-Host "    $($lv.LocalType.FullName)"
                }
            }
        } catch {
            Write-Host "  Error: $_"
        }
    }
}
