# Extract ClusterRect structure and ExcelWorksheetCom coordinate methods

$asm = [System.Reflection.Assembly]::LoadFrom('C:/Program Files (x86)/CIMTOPS CORPORATION/ConMas Designer/bin/LibExcelController.dll')

# === ClusterRect structure ===
$t = $asm.GetType('LibExcelController.Lib.ClusterRect')
Write-Host '=== ClusterRect ==='
if ($t) {
    Write-Host "Type: $($t.FullName)"
    Write-Host "IsValueType: $($t.IsValueType)"
    Write-Host "Fields:"
    foreach ($f in $t.GetFields([System.Reflection.BindingFlags]'Public,NonPublic,Instance,Static')) {
        Write-Host "  $($f.FieldType.Name) $($f.Name)"
    }
    Write-Host "Properties:"
    foreach ($p in $t.GetProperties([System.Reflection.BindingFlags]'Public,NonPublic,Instance,Static')) {
        Write-Host "  $($p.PropertyType.Name) $($p.Name)"
    }
    Write-Host "Methods:"
    foreach ($m in $t.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Instance,Static,DeclaredOnly')) {
        Write-Host "  $($m.Name)"
    }
} else {
    Write-Host "NOT FOUND"
}

# === ConMasExcelClient.dll - ClusterRect and ExcelWorksheetCom ===
$asm2 = [System.Reflection.Assembly]::LoadFrom('C:/Program Files (x86)/CIMTOPS CORPORATION/ConMas Designer/bin/ConMasExcelClient.dll')

# Check for similar types
$types = $asm2.GetTypes()
foreach ($t2 in $types) {
    $n = $t2.FullName
    if ($n -match 'Print|Cluster|Rect|Position|Coordinate|Page|CalcSize') {
        Write-Host "`n=== $n ==="
        Write-Host "Base: $($t2.BaseType)"
        Write-Host "Methods:"
        foreach ($m in $t2.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Instance,Static,DeclaredOnly')) {
            if (-not $m.Name.StartsWith('get_') -and -not $m.Name.StartsWith('set_')) {
                $pars = ($m.GetParameters() | ForEach-Object { "$($_.ParameterType.Name) $($_.Name)" }) -join ', '
                Write-Host "  [$($m.ReturnType.Name)] $($m.Name)($pars)"
            }
        }
    }
}

# === ExcelWorksheetCom in LibExcelController - get print area methods ===
$t3 = $asm.GetType('LibExcelController.ExcelLib.ExcelWorksheetCom')
Write-Host "`n=== ExcelWorksheetCom Methods ==="
foreach ($m in $t3.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance,DeclaredOnly')) {
    $pars = ($m.GetParameters() | ForEach-Object { "$($_.ParameterType.Name) $($_.Name)" }) -join ', '
    Write-Host "  [$($m.ReturnType.Name)] $($m.Name)($pars)"
}

# === Check for any type with "Ratio" or "Normalize" in name ===
$allTypes = $asm.GetTypes()
foreach ($t4 in $allTypes) {
    $n = $t4.FullName
    if ($n -match 'Ratio|Normaliz|Denominator|Denom|Reference|PageSize|PaperSize') {
        Write-Host "`nFound: $n"
    }
}

# === ExcelRangeBase - check for coordinate methods ===
$t_range = $asm.GetType('LibExcelController.ExcelLib.ExcelRangeBase')
Write-Host "`n=== ExcelRangeBase Methods (coordinate-related) ==="
foreach ($m in $t_range.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Instance,Static,DeclaredOnly')) {
    $name = $m.Name
    if ($name -match 'Left|Top|Width|Height|GetRange|Point|Pixel|Position|Address|Row|Column|Merge|Print') {
        $pars = ($m.GetParameters() | ForEach-Object { "$($_.ParameterType.Name) $($_.Name)" }) -join ', '
        Write-Host "  [$($m.ReturnType.Name)] $($name)($pars)"
    }
}

# === ExcelWorkbookBase ===
$t_wb = $asm.GetType('LibExcelController.ExcelLib.ExcelWorkbookBase')
Write-Host "`n=== ExcelWorkbookBase Methods (coordinate/print-related) ==="
foreach ($m in $t_wb.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Instance,Static,DeclaredOnly')) {
    $name = $m.Name
    if ($name -match 'Page|Print|Ratio|Export|GetDef|Cluster|Coordinate') {
        $pars = ($m.GetParameters() | ForEach-Object { "$($_.ParameterType.Name) $($_.Name)" }) -join ', '
        Write-Host "  [$($m.ReturnType.Name)] $($name)($pars)"
    }
}

# === Check the static fields that appear as metadata (defined in the class) ===
$t_imgutil = $asm.GetType('LibExcelController.Lib.ImageUtility')
Write-Host "`n=== ImageUtility Fields/Properties ==="
foreach ($f in $t_imgutil.GetFields([System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance')) {
    Write-Host "  Field: $($f.FieldType.Name) $($f.Name)"
}
foreach ($p in $t_imgutil.GetProperties([System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance')) {
    Write-Host "  Property: $($p.PropertyType.Name) $($p.Name)"
}
foreach ($m in $t_imgutil.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance,DeclaredOnly')) {
    $name = $m.Name
    if ($name -match 'get_|set_') {
        continue
    }
    if ($name -match 'Create|Load|Bit|Image|Pdf|Crop|Resize|Thresh|Morph|Dilat|Erod|Contour|Bound|Box|Rect') {
        $pars = ($m.GetParameters() | ForEach-Object { "$($_.ParameterType.Name) $($_.Name)" }) -join ', '
        Write-Host "  [$($m.ReturnType.Name)] $($name)($pars)"
    }
}

# === GetImagePixelsFromPdf if it exists ===
foreach ($m in $t_imgutil.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance,DeclaredOnly')) {
    $name = $m.Name
    if ($name -notmatch 'get_|set_') {
        $pars = ($m.GetParameters() | ForEach-Object { "$($_.ParameterType.Name) $($_.Name)" }) -join ', '
        Write-Host "  [$($m.ReturnType.Name)] $($name)($pars)"
    }
}
