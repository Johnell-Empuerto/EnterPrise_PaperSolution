"""
Phase X.34 - Behavioral Acceptance Test (Round-Trip Parity Certification)
No code modifications. No added logging. Fresh runtime evidence only.
"""
import os, sys, json, subprocess, time, hashlib, shutil, zipfile
from xml.etree import ElementTree as ET
from datetime import datetime

PROJECT = os.path.dirname(os.path.abspath(__file__))
ORIGINAL_PATH = os.path.join(PROJECT, "formtest.xlsx")
CONMAS_PATH = os.path.join(PROJECT, "test_conmas_output.xlsx")
API_URL = "http://localhost:5090"
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
SEP = "=" * 70
SEP2 = "-" * 70

print(SEP)
print("PHASE X.34 - BEHAVIORAL ACCEPTANCE TEST")
print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print(SEP)

# Verify API is running
import urllib.request
try:
    resp = urllib.request.urlopen(f'{API_URL}/api/health', timeout=5)
    health = resp.read().decode()
    print(f"\nAPI status: RUNNING - {health[:150]}")
except Exception as e:
    print(f"\nAPI status: NOT RUNNING - {e}")
    sys.exit(1)

# ============================================================
# TEST B: ConMas.xlsx verification (COM + OOXML)
# ============================================================
print(f"\n{SEP2}")
print("TEST B: ConMas.xlsx - OOXML Verification")
print(SEP2)

print(f"\n  File: {CONMAS_PATH}")
conmas_size = os.path.getsize(CONMAS_PATH)
conmas_md5 = hashlib.md5(open(CONMAS_PATH, 'rb').read()).hexdigest()[:12]
print(f"  Size: {conmas_size:,} bytes, MD5: {conmas_md5}")

with zipfile.ZipFile(CONMAS_PATH) as z:
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    sheets = wb.findall('.//s:sheet', NS)
    dns = wb.findall('.//s:definedName', NS)
    
    print(f"  Sheets: {len(sheets)}")
    for s in sheets:
        print(f"    [{s.get('state','visible')}] {s.get('name')}")
    
    print(f"  Defined names: {len(dns)}")
    for dn in dns:
        print(f"    _xlnm.{dn.get('name')} = '{dn.text or ''}'")
    
    # Sheet1 OOXML
    for s in sheets:
        sname = s.get('name', '')
        if sname != 'Sheet1': continue
        rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
        target = ''
        for r in rels:
            if r.get('Id') == rid: target = r.get('Target', ''); break
        spath = f"xl/{target}" if target else ''
        if spath not in z.namelist(): continue
        
        ws = ET.fromstring(z.read(spath))
        dim = ws.find('.//s:dimension', NS)
        print(f"\n  Sheet1 OOXML:")
        print(f"    dimension: {dim.get('ref','') if dim is not None else 'N/A'}")
        
        po = ws.find('.//s:printOptions', NS)
        if po is not None:
            print(f"    printOptions: H={po.get('horizontalCentered','?')} V={po.get('verticalCentered','?')}")
        else:
            print(f"    printOptions: MISSING")
        
        pm = ws.find('.//s:pageMargins', NS)
        if pm is not None:
            print(f"    pageMargins:")
            for k, v in pm.attrib.items():
                v_pt = float(v) * 72
                print(f"      {k}: {v} in = {v_pt:.4f} pt")
        
        ps = ws.find('.//s:pageSetup', NS)
        if ps is not None:
            clean = {k.split('}')[-1]: v for k,v in ps.attrib.items()}
            print(f"    pageSetup: {clean}")

# ============================================================
# TEST C: First Generation via Current API
# ============================================================
print(f"\n{SEP2}")
print("TEST C: First Generation - via Current API")
print(SEP2)

# Upload ConMas to get FormDefinition
upload_path = os.path.join(PROJECT, "_x34_upload.xlsx")
shutil.copy2(CONMAS_PATH, upload_path)

print(f"  Uploading ConMas.xlsx to get FormDefinition...")
result = subprocess.run([
    'curl', '-s', '-X', 'POST',
    f'{API_URL}/api/form/upload-excel',
    '-F', f'file=@{upload_path};filename=ConMas.xlsx',
    '--connect-timeout', '30', '--max-time', '60'
], capture_output=True, text=True)

if result.returncode != 0:
    print(f"  Upload FAILED"); sys.exit(1)

response = json.loads(result.stdout)
data = response.get('data', response.get('Data', {}))
if isinstance(data, dict):
    form_def = data.get('formDefinition', data.get('FormDefinition', {})) or data
else:
    print(f"  Unexpected data format"); sys.exit(1)

print(f"  FormDefinition received: {len(form_def.get('sheets', form_def.get('Sheets', [])))} sheet(s)")

# Generate output
gen1_path = os.path.join(PROJECT, "_x34_generated.xlsx")
print(f"  Generating workbook 1...")
result2 = subprocess.run([
    'curl', '-s', '-X', 'POST',
    f'{API_URL}/api/form/output-excel',
    '-H', 'Content-Type: application/json',
    '-d', json.dumps(form_def),
    '--connect-timeout', '60', '--max-time', '120',
    '-o', gen1_path
], capture_output=True, text=True)

if not os.path.exists(gen1_path) or os.path.getsize(gen1_path) == 0:
    print(f"  Generation FAILED"); sys.exit(1)

g1_size = os.path.getsize(gen1_path)
g1_md5 = hashlib.md5(open(gen1_path, 'rb').read()).hexdigest()[:12]
g1_ts = os.path.getmtime(gen1_path)
print(f"  Generated1: {gen1_path}")
print(f"  Size: {g1_size:,} bytes, MD5: {g1_md5}")
print(f"  Time: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(g1_ts))}")

# OOXML verification of Generated1
print(f"\n  Generated1 OOXML:")
with zipfile.ZipFile(gen1_path) as z:
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    sheets = wb.findall('.//s:sheet', NS)
    dns = wb.findall('.//s:definedName', NS)
    
    print(f"    Sheets: {len(sheets)}")
    for s in sheets:
        print(f"      [{s.get('state','visible')}] {s.get('name')}")
    
    print(f"    Defined names: {len(dns)}")
    for dn in dns:
        print(f"      _xlnm.{dn.get('name')} = '{dn.text or ''}'")
    
    for s in sheets:
        sname = s.get('name', '')
        if sname != 'Sheet1': continue
        rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
        target = ''
        for r in rels:
            if r.get('Id') == rid: target = r.get('Target', ''); break
        spath = f"xl/{target}" if target else ''
        if spath not in z.namelist(): continue
        
        ws = ET.fromstring(z.read(spath))
        dim = ws.find('.//s:dimension', NS)
        print(f"    dimension: {dim.get('ref','') if dim is not None else 'N/A'}")
        
        po = ws.find('.//s:printOptions', NS)
        if po is not None:
            print(f"    printOptions: H={po.get('horizontalCentered','?')} V={po.get('verticalCentered','?')}")
        else:
            print(f"    printOptions: MISSING")
        
        pm = ws.find('.//s:pageMargins', NS)
        if pm is not None:
            print(f"    pageMargins:")
            for k, v in pm.attrib.items():
                v_pt = float(v) * 72
                print(f"      {k}: {v} in = {v_pt:.4f} pt")

# ============================================================
# TEST D: Second Generation (Idempotency)
# ============================================================
print(f"\n{SEP2}")
print("TEST D: Second Generation - Upload Generated1 back into API")
print(SEP2)

# Upload Generated1.xlsx to get new FormDefinition
shutil.copy2(gen1_path, upload_path)
print(f"  Uploading Generated1.xlsx back into API...")

result3 = subprocess.run([
    'curl', '-s', '-X', 'POST',
    f'{API_URL}/api/form/upload-excel',
    '-F', f'file=@{upload_path};filename=Generated1.xlsx',
    '--connect-timeout', '30', '--max-time', '60'
], capture_output=True, text=True)

if result3.returncode != 0:
    print(f"  Upload FAILED"); sys.exit(1)

response3 = json.loads(result3.stdout)
data3 = response3.get('data', response3.get('Data', {}))
form_def2 = data3.get('formDefinition', data3.get('FormDefinition', {})) or data3

print(f"  FormDefinition2 received: {len(form_def2.get('sheets', form_def2.get('Sheets', [])))} sheet(s)")

# Generate Generated2
gen2_path = os.path.join(PROJECT, "_x34_generated2.xlsx")
print(f"  Generating workbook 2...")

result4 = subprocess.run([
    'curl', '-s', '-X', 'POST',
    f'{API_URL}/api/form/output-excel',
    '-H', 'Content-Type: application/json',
    '-d', json.dumps(form_def2),
    '--connect-timeout', '60', '--max-time', '120',
    '-o', gen2_path
], capture_output=True, text=True)

if not os.path.exists(gen2_path) or os.path.getsize(gen2_path) == 0:
    print(f"  Generation FAILED"); sys.exit(1)

g2_size = os.path.getsize(gen2_path)
g2_md5 = hashlib.md5(open(gen2_path, 'rb').read()).hexdigest()[:12]
g2_ts = os.path.getmtime(gen2_path)
print(f"  Generated2: {gen2_path}")
print(f"  Size: {g2_size:,} bytes, MD5: {g2_md5}")
print(f"  Time: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(g2_ts))}")

# OOXML verification of Generated2
print(f"\n  Generated2 OOXML:")
with zipfile.ZipFile(gen2_path) as z:
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    sheets = wb.findall('.//s:sheet', NS)
    dns = wb.findall('.//s:definedName', NS)
    
    print(f"    Sheets: {len(sheets)}")
    for s in sheets:
        print(f"      [{s.get('state','visible')}] {s.get('name')}")
    
    print(f"    Defined names: {len(dns)}")
    for dn in dns:
        print(f"      _xlnm.{dn.get('name')} = '{dn.text or ''}'")
    
    for s in sheets:
        sname = s.get('name', '')
        if sname != 'Sheet1': continue
        rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
        target = ''
        for r in rels:
            if r.get('Id') == rid: target = r.get('Target', ''); break
        spath = f"xl/{target}" if target else ''
        if spath not in z.namelist(): continue
        
        ws = ET.fromstring(z.read(spath))
        dim = ws.find('.//s:dimension', NS)
        print(f"    dimension: {dim.get('ref','') if dim is not None else 'N/A'}")
        
        po = ws.find('.//s:printOptions', NS)
        if po is not None:
            print(f"    printOptions: H={po.get('horizontalCentered','?')} V={po.get('verticalCentered','?')}")
        else:
            print(f"    printOptions: MISSING")
        
        pm = ws.find('.//s:pageMargins', NS)
        if pm is not None:
            print(f"    pageMargins:")
            for k, v in pm.attrib.items():
                v_pt = float(v) * 72
                print(f"      {k}: {v} in = {v_pt:.4f} pt")

# ============================================================
# TEST F: Binary Stability - Generated1 vs Generated2
# ============================================================
print(f"\n{SEP2}")
print("TEST F: Binary Stability - Generated1 vs Generated2")
print(SEP2)

# Compare OOXML print-relevant elements only
def get_print_relevant(path):
    """Extract print-relevant OOXML values from a workbook."""
    result = {}
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        dns = wb.findall('.//s:definedName', NS)
        result['defined_names'] = {}
        for dn in dns:
            result['defined_names'][dn.get('name')] = dn.text or ''
        
        for s in wb.findall('.//s:sheet', NS):
            sname = s.get('name', '')
            if sname != 'Sheet1': continue
            rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            target = ''
            for r in rels:
                if r.get('Id') == rid: target = r.get('Target', ''); break
            spath = f"xl/{target}" if target else ''
            if spath not in z.namelist(): continue
            
            ws = ET.fromstring(z.read(spath))
            
            po = ws.find('.//s:printOptions', NS)
            result['printOptions'] = dict(po.attrib) if po is not None else {}
            
            pm = ws.find('.//s:pageMargins', NS)
            result['pageMargins'] = dict(pm.attrib) if pm is not None else {}
            
            dim = ws.find('.//s:dimension', NS)
            result['dimension'] = dim.get('ref', '') if dim is not None else ''
    
    return result

g1_pr = get_print_relevant(gen1_path)
g2_pr = get_print_relevant(gen2_path)

print(f"\n  {'Property':<25} {'Generated1':<25} {'Generated2':<25} {'Match?'}")
print(f"  {'─'*25} {'─'*25} {'─'*25} {'─'*8}")

all_match = True

# Compare defined names
g1_dn = json.dumps(g1_pr.get('defined_names', {}), sort_keys=True)
g2_dn = json.dumps(g2_pr.get('defined_names', {}), sort_keys=True)
match = g1_dn == g2_dn
if not match: all_match = False
print(f"  {'Defined names':<25} {str(list(g1_pr['defined_names'].keys())):<25} {str(list(g2_pr['defined_names'].keys())):<25} {'OK' if match else 'DIFF'}")

# Compare printOptions
g1_po = json.dumps(g1_pr.get('printOptions', {}), sort_keys=True)
g2_po = json.dumps(g2_pr.get('printOptions', {}), sort_keys=True)
match = g1_po == g2_po
if not match: all_match = False
print(f"  {'printOptions':<25} {g1_po:<25} {g2_po:<25} {'OK' if match else 'DIFF'}")

# Compare pageMargins (in inches)
for key in ['left', 'right', 'top', 'bottom', 'header', 'footer']:
    g1v = g1_pr['pageMargins'].get(key, 'N/A')
    g2v = g2_pr['pageMargins'].get(key, 'N/A')
    match = abs(float(g1v) - float(g2v)) < 0.0001 if g1v != 'N/A' and g2v != 'N/A' else g1v == g2v
    if not match: all_match = False
    print(f"  {'pageMargin ' + key:<25} {g1v:<25} {g2v:<25} {'OK' if match else 'DIFF'}")

# Compare dimension
match = g1_pr.get('dimension', '') == g2_pr.get('dimension', '')
if not match: all_match = False
print(f"  {'dimension':<25} {g1_pr.get('dimension',''):<25} {g2_pr.get('dimension',''):<25} {'OK' if match else 'DIFF'}")

print(f"\n  Binary stability: {'PASS - All print-relevant properties match' if all_match else 'FAIL - Differences found'}")

# ============================================================
# COM Verification Script
# ============================================================
print(f"\n{SEP2}")
print("COM Verification (via PowerShell)")
print(SEP2)

ps_path = os.path.join(PROJECT, "_x34_com_verify.ps1")
ps_content = f'''
$ProjectDir = "{PROJECT}"

$files = @{{
    "Generated1" = Join-Path $ProjectDir "_x34_generated.xlsx"
    "Generated2" = Join-Path $ProjectDir "_x34_generated2.xlsx"
    "ConMas" = Join-Path $ProjectDir "test_conmas_output.xlsx"
}}

Write-Host "PHASE X.34 - COM VERIFICATION"
Write-Host "============================"

Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

foreach ($pair in $files.GetEnumerator()) {{
    $label = $pair.Key
    $path = $pair.Value

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {{
        $wb = $excel.Workbooks.Open($path)
        Write-Host ""
        Write-Host "--- $label ---"
        Write-Host "Sheets: $($wb.Worksheets.Count)"

        for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {{
            $ws = $wb.Worksheets[$i]
            $ps = $ws.PageSetup
            Write-Host ""
            Write-Host "Sheet [$i]: $($ws.Name)"
            Write-Host "  PrintArea:            $($ps.PrintArea)"
            Write-Host "  CenterHorizontally:   $($ps.CenterHorizontally)"
            Write-Host "  CenterVertically:     $($ps.CenterVertically)"
            $lm = [Math]::Round([double]$ps.LeftMargin, 4)
            $rm = [Math]::Round([double]$ps.RightMargin, 4)
            $tm = [Math]::Round([double]$ps.TopMargin, 4)
            $bm = [Math]::Round([double]$ps.BottomMargin, 4)
            $hm = [Math]::Round([double]$ps.HeaderMargin, 4)
            $fm = [Math]::Round([double]$ps.FooterMargin, 4)
            Write-Host "  LeftMargin:           $lm pt ($([Math]::Round($lm/72*2.54,2)) cm)"
            Write-Host "  RightMargin:          $rm pt ($([Math]::Round($rm/72*2.54,2)) cm)"
            Write-Host "  TopMargin:            $tm pt ($([Math]::Round($tm/72*2.54,2)) cm)"
            Write-Host "  BottomMargin:         $bm pt ($([Math]::Round($bm/72*2.54,2)) cm)"
            Write-Host "  HeaderMargin:         $hm pt"
            Write-Host "  FooterMargin:         $fm pt"
            Write-Host "  Orientation:          $($ps.Orientation)"
            Write-Host "  Zoom:                 $($ps.Zoom)"
            Write-Host "  FitToPagesWide:       $($ps.FitToPagesWide)"
            Write-Host "  FitToPagesTall:       $($ps.FitToPagesTall)"
            [Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
        }}
        $wb.Close($false)
    }}
    catch {{
        Write-Host "  ERROR: $($_.Exception.Message)"
    }}
    finally {{
        $excel.Quit()
        if ($wb) {{ [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }}
}}

Write-Host ""
Write-Host "COM verification complete."
'''

with open(ps_path, 'w') as f:
    f.write(ps_content)

# ============================================================
# TEST E: Excel UI verification script
# ============================================================
print(f"\n{SEP2}")
print("TEST E: Excel UI (make visible)")
print(SEP2)

ui_path = os.path.join(PROJECT, "_x34_ui_verify.ps1")
ui_content = f'''
$ProjectDir = "{PROJECT}"

Write-Host "PHASE X.34 - EXCEL UI VERIFICATION"
Write-Host "==================================="
Write-Host ""
Write-Host "This will open all four workbooks in Excel."
Write-Host "Please verify Print Preview and Page Setup."
Write-Host ""

Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

# Open ConMas
$conmasPath = Join-Path $ProjectDir "test_conmas_output.xlsx"
$wb1 = $excel.Workbooks.Open($conmasPath)
$ws1 = $wb1.Worksheets("Sheet1")
$ws1.Activate()
Write-Host ""
Write-Host "=== CONMAS.xlsx - Sheet1 ==="
$ps = $ws1.PageSetup
$lm = [Math]::Round([double]$ps.LeftMargin, 2)
$tm = [Math]::Round([double]$ps.TopMargin, 2)
Write-Host "CenterH: $($ps.CenterHorizontally)"
Write-Host "CenterV: $($ps.CenterVertically)"
Write-Host "LeftMargin: $lm pt ($([Math]::Round($lm/72*2.54,2)) cm)"
Write-Host "TopMargin: $tm pt ($([Math]::Round($tm/72*2.54,2)) cm)"
Write-Host "PrintArea: '$($ps.PrintArea)'"
Write-Host ""
Write-Host "Press any key to load Generated1..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# Open Generated1 (new window)
$gen1Path = Join-Path $ProjectDir "_x34_generated.xlsx"
$wb2 = $excel.Workbooks.Open($gen1Path)
Write-Host ""
Write-Host "=== GENERATED1.xlsx ==="
$ws2 = $wb2.Worksheets(1)
$ws2.Activate()
$ps2 = $ws2.PageSetup
$lm2 = [Math]::Round([double]$ps2.LeftMargin, 2)
$tm2 = [Math]::Round([double]$ps2.TopMargin, 2)
Write-Host "CenterH: $($ps2.CenterHorizontally)"
Write-Host "CenterV: $($ps2.CenterVertically)"
Write-Host "LeftMargin: $lm2 pt ($([Math]::Round($lm2/72*2.54,2)) cm)"
Write-Host "TopMargin: $tm2 pt ($([Math]::Round($tm2/72*2.54,2)) cm)"
Write-Host "PrintArea: '$($ps2.PrintArea)'"
Write-Host ""
Write-Host "Press any key to load Generated2..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

$gen2Path = Join-Path $ProjectDir "_x34_generated2.xlsx"
$wb3 = $excel.Workbooks.Open($gen2Path)
$ws3 = $wb3.Worksheets(1)
$ws3.Activate()
$ps3 = $ws3.PageSetup
$lm3 = [Math]::Round([double]$ps3.LeftMargin, 2)
$tm3 = [Math]::Round([double]$ps3.TopMargin, 2)
Write-Host ""
Write-Host "=== GENERATED2.xlsx ==="
Write-Host "CenterH: $($ps3.CenterHorizontally)"
Write-Host "CenterV: $($ps3.CenterVertically)"
Write-Host "LeftMargin: $lm3 pt ($([Math]::Round($lm3/72*2.54,2)) cm)"
Write-Host "TopMargin: $tm3 pt ($([Math]::Round($tm3/72*2.54,2)) cm)"
Write-Host "PrintArea: '$($ps3.PrintArea)'"
Write-Host ""
Write-Host "Three workbooks are now visible:"
Write-Host "  1. ConMas.xlsx (Sheet1)"
Write-Host "  2. Generated1.xlsx"
Write-Host "  3. Generated2.xlsx"
Write-Host ""
Write-Host "Please compare Print Preview and Page Setup dialogs."
Write-Host "Press any key to close all..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

$wb1.Close($false)
$wb2.Close($false)
$wb3.Close($false)
$excel.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($wb1) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($wb2) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($wb3) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "Excel closed."
'''

with open(ui_path, 'w') as f:
    f.write(ui_content)

print(f"  COM script: {ps_path}")
print(f"  UI script:  {ui_path}")
print(f"\n  PowerShell commands:")
print(f"  powershell -ExecutionPolicy Bypass -File {ps_path}")
print(f"  powershell -ExecutionPolicy Bypass -File {ui_path}")

# ============================================================
# FINAL ACCEPTANCE MATRIX
# ============================================================
print(f"\n{SEP2}")
print("FINAL ACCEPTANCE MATRIX (OOXML evidence)")
print(SEP2)

print(f"\n  {'Test':<20} {'Property':<20} {'ConMas':<12} {'Gen1':<12} {'Gen2':<12} {'Pass?'}")
print(f"  {'─'*20} {'─'*20} {'─'*12} {'─'*12} {'─'*12} {'─'*8}")

# Compare all three workbooks
def compare_print_props(path, label):
    props = {}
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        dns = wb.findall('.//s:definedName', NS)
        props['definedNames'] = len(dns)
        props['hasPrintArea'] = any('Print_Area' in dn.get('name','') for dn in dns)
        
        for s in wb.findall('.//s:sheet', NS):
            sname = s.get('name', '')
            if sname != 'Sheet1': continue
            rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            target = ''
            for r in rels:
                if r.get('Id') == rid: target = r.get('Target', ''); break
            spath = f"xl/{target}" if target else ''
            if spath not in z.namelist(): continue
            
            ws = ET.fromstring(z.read(spath))
            props['dimension'] = ws.find('.//s:dimension', NS).get('ref','') if ws.find('.//s:dimension', NS) else ''
            po = ws.find('.//s:printOptions', NS)
            props['poH'] = po.get('horizontalCentered','') if po is not None else 'MISSING'
            props['poV'] = po.get('verticalCentered','') if po is not None else 'MISSING'
            pm = ws.find('.//s:pageMargins', NS)
            if pm is not None:
                for k in ['left','right','top','bottom']:
                    props[f'pm_{k}'] = pm.get(k, '')
    return props

props_con = compare_print_props(CONMAS_PATH, "ConMas")
props_g1 = compare_print_props(gen1_path, "Gen1")
props_g2 = compare_print_props(gen2_path, "Gen2")

all_pass = True
for test_name, key, label in [
    ("B/C/D", "Print Area", "hasPrintArea"),
    ("B/C/D", "printOptions H", "poH"),
    ("B/C/D", "printOptions V", "poV"),
    ("B/C/D", "LeftMargin", "pm_left"),
    ("B/C/D", "RightMargin", "pm_right"),
    ("B/C/D", "TopMargin", "pm_top"),
    ("B/C/D", "BottomMargin", "pm_bottom"),
    ("F", "Dimension", "dimension"),
]:
    v_con = props_con.get(label, '')
    v_g1 = props_g1.get(label, '')
    v_g2 = props_g2.get(label, '')
    
    pass_c = str(v_con == v_g1).upper()[:2]
    pass_f = str(v_g1 == v_g2).upper()[:2]
    
    if v_con != v_g1 or v_g1 != v_g2:
        all_pass = False
    
    print(f"  {test_name:<20} {key:<20} {str(v_con):<12} {str(v_g1):<12} {str(v_g2):<12} {pass_c}/{pass_f}")

print(f"\n{'─'*70}")
if all_pass:
    print("OVERALL: ALL TESTS PASS - Pipeline certified for print-layout preservation")
else:
    print("OVERALL: SOME TESTS FAILED - See details above")
print(f"{'─'*70}")

# Cleanup
for f in ['_x34_upload.xlsx']:
    try: os.remove(os.path.join(PROJECT, f))
    except: pass

print(f"\n{SEP}")
print("PHASE X.34 ACCEPTANCE TEST COMPLETE")
print(SEP)
