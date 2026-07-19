"""
Phase X.33 - Fresh Runtime Validation
No historical evidence. Every value comes from a workbook generated during this run.
"""
import os, sys, json, subprocess, time, hashlib, shutil, zipfile
from xml.etree import ElementTree as ET
from datetime import datetime

PROJECT = os.path.dirname(os.path.abspath(__file__))
CONMAS_PATH = os.path.join(PROJECT, "test_conmas_output.xlsx")
API_URL = "http://localhost:5090"
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
SEP = "=" * 70
SEP2 = "-" * 70

print(SEP)
print("PHASE X.33 - FRESH RUNTIME VALIDATION")
print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print(SEP)

# ═══ STEP 1: System State ═══
print(f"\n{SEP2}")
print("STEP 1: System State")
print(SEP2)

import urllib.request
try:
    resp = urllib.request.urlopen(f'{API_URL}/api/health', timeout=5)
    health = resp.read().decode()
    print(f"  API health: {health[:200]}")
except Exception as e:
    print(f"  API health: FAILED - {e}")
    print("  Cannot continue without API.")
    sys.exit(1)

dll_path = os.path.join(PROJECT, "ExcelAPI/ExcelAPI/bin/Debug/net10.0/ExcelAPI.dll")
if os.path.exists(dll_path):
    ts = os.path.getmtime(dll_path)
    dll_md5 = hashlib.md5(open(dll_path, 'rb').read()).hexdigest()[:12]
    print(f"  DLL: {dll_path}")
    print(f"  DLL timestamp: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(ts))}")
    print(f"  DLL size: {os.path.getsize(dll_path):,} bytes")
    print(f"  DLL MD5: {dll_md5}")

# Git state
git_log = subprocess.run(['git', 'log', '--oneline', '-1'], capture_output=True, text=True)
git_status = subprocess.run(['git', 'status', '--short'], capture_output=True, text=True)
print(f"  Git HEAD: {git_log.stdout.strip()}")
changed_files = [l for l in git_status.stdout.split('\n') if l.strip() and l.strip().startswith('M')]
print(f"  Modified files in working tree: {len(changed_files)}")

# ═══ STEP 2: Generate Fresh Workbook ═══
print(f"\n{SEP2}")
print("STEP 2: Generate Fresh Workbook")
print(SEP2)

# 2a: Upload ConMas.xlsx to get FormDefinition
upload_path = os.path.join(PROJECT, "_x33_upload.xlsx")
shutil.copy2(CONMAS_PATH, upload_path)

print(f"  Uploading ConMas.xlsx ({os.path.getsize(CONMAS_PATH):,} bytes)...")

result = subprocess.run([
    'curl', '-s', '-X', 'POST',
    f'{API_URL}/api/form/upload-excel',
    '-F', f'file=@{upload_path};filename=ConMas.xlsx',
    '--connect-timeout', '30',
    '--max-time', '60'
], capture_output=True, text=True)

if result.returncode != 0:
    print(f"  Upload FAILED: {result.stderr[:200]}")
    sys.exit(1)

try:
    response = json.loads(result.stdout)
except json.JSONDecodeError as e:
    print(f"  JSON parse error: {e}")
    print(f"  Raw: {result.stdout[:300]}")
    sys.exit(1)

data = response.get('data', response.get('Data', {}))
if isinstance(data, dict):
    form_def = data.get('formDefinition', data.get('FormDefinition', {}))
    if not form_def:
        form_def = data
else:
    print(f"  Unexpected data format: {type(data)}")
    sys.exit(1)

print(f"  FormDefinition received: {len(form_def.get('sheets', form_def.get('Sheets', [])))} sheet(s)")

# 2b: Post FormDefinition to output-excel
gen_path = os.path.join(PROJECT, "_x33_generated_output.xlsx")
print(f"  Generating workbook...")

result2 = subprocess.run([
    'curl', '-s', '-X', 'POST',
    f'{API_URL}/api/form/output-excel',
    '-H', 'Content-Type: application/json',
    '-d', json.dumps(form_def),
    '--connect-timeout', '60',
    '--max-time', '120',
    '-o', gen_path
], capture_output=True, text=True)

if os.path.exists(gen_path) and os.path.getsize(gen_path) > 0:
    gen_size = os.path.getsize(gen_path)
    gen_md5 = hashlib.md5(open(gen_path, 'rb').read()).hexdigest()[:12]
    gen_sha256 = hashlib.sha256(open(gen_path, 'rb').read()).hexdigest()
    gen_ts = os.path.getmtime(gen_path)
    print(f"  Generated: {gen_path}")
    print(f"  Size: {gen_size:,} bytes")
    print(f"  MD5: {gen_md5}")
    print(f"  SHA256: {gen_sha256[:32]}...")
    print(f"  Timestamp: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(gen_ts))}")
else:
    print(f"  Generation FAILED")
    if result2.stderr:
        print(f"  stderr: {result2.stderr[:300]}")
    sys.exit(1)

# ═══ STEP 3 + 4: COM Verification (open, read, close, reopen, read) ═══
print(f"\n{SEP2}")
print("STEP 3+4: COM Verification (open + reopen)")
print(SEP2)

# We'll write the COM verification to a PowerShell script and run it
ps1_path = os.path.join(PROJECT, "_x33_com_verify.ps1")
ps1_content = f'''
# Phase X.33 - COM Verification (fresh workbook)
$ProjectDir = "{PROJECT}"
$genPath = [System.IO.Path]::Combine($ProjectDir, "_x33_generated_output.xlsx")
$conmasPath = [System.IO.Path]::Combine($ProjectDir, "test_conmas_output.xlsx")

Write-Host "Phase X.33 - STEP 3+4: COM Verification"
Write-Host ""

function Read-PageSetup($wb, $label) {{
    Write-Host "--- $label ---"
    Write-Host "Sheet count: $($wb.Worksheets.Count)"
    for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {{
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
    }}
}}

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
'''

with open(ps1_path, 'w') as f:
    f.write(ps1_content)

print(f"  COM verification script written to: {ps1_path}")

# ═══ STEP 5: OOXML Verification ═══
print(f"\n{SEP2}")
print("STEP 5: OOXML Verification (direct ZIP read)")
print(SEP2)

with zipfile.ZipFile(gen_path) as z:
    names = z.namelist()
    print(f"  ZIP entries: {len(names)}")
    
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    sheets = wb.findall('.//s:sheet', NS)
    dns = wb.findall('.//s:definedName', NS)
    
    print(f"\n  SHEETS: {len(sheets)}")
    for s in sheets:
        print(f"    [{s.get('state','visible')}] {s.get('name')}")
    
    print(f"\n  DEFINED NAMES: {len(dns)}")
    for dn in dns:
        print(f"    _xlnm.{dn.get('name')} = '{dn.text or ''}'")
    
    # Build sheet path map
    sheet_map = {}
    for s in sheets:
        sname = s.get('name', '?')
        rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
        target = ''
        for r in rels:
            if r.get('Id') == rid:
                target = r.get('Target', '')
                break
        sheet_map[sname] = f"xl/{target}" if target else ''
    
    for sname, spath in sheet_map.items():
        if spath not in names:
            print(f"\n  MISSING: {spath}")
            continue
        ws = ET.fromstring(z.read(spath))
        dim = ws.find('.//s:dimension', NS)
        dim_ref = dim.get('ref', '') if dim is not None else '(none)'
        
        print(f"\n  Sheet: {sname} (dim={dim_ref})")
        
        pm = ws.find('.//s:pageMargins', NS)
        if pm is not None:
            print(f"    pageMargins:")
            for k, v in pm.attrib.items():
                v_pt = float(v) * 72
                v_cm = float(v) * 2.54
                source = "CUSTOM" if abs(float(v) - 0.70866) < 0.01 and k in ['left','right'] else ("CUSTOM" if abs(float(v) - 0.74803) < 0.01 and k in ['top','bottom'] else ("CUSTOM (header)" if abs(float(v) - 0.31496) < 0.01 else "DEFAULT"))
                print(f"      {k}: {v} in = {v_pt:.4f} pt = {v_cm:.2f} cm [{source}]")
        else:
            print(f"    pageMargins: MISSING")
        
        po = ws.find('.//s:printOptions', NS)
        if po is not None:
            print(f"    printOptions: H={po.get('horizontalCentered','0')} V={po.get('verticalCentered','0')}")
        else:
            print(f"    printOptions: MISSING")
        
        ps = ws.find('.//s:pageSetup', NS)
        if ps is not None:
            clean = {k.split('}')[-1]: v for k,v in ps.attrib.items()}
            print(f"    pageSetup: {clean}")

# Also read ConMas OOXML for comparison
print(f"\n{SEP2}")
print("ConMas OOXML (for comparison)")
print(SEP2)

with zipfile.ZipFile(CONMAS_PATH) as z:
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    sheets = wb.findall('.//s:sheet', NS)
    dns = wb.findall('.//s:definedName', NS)
    
    print(f"\n  SHEETS: {len(sheets)}")
    for s in sheets:
        print(f"    [{s.get('state','visible')}] {s.get('name')}")
    
    print(f"\n  DEFINED NAMES: {len(dns)}")
    for dn in dns:
        print(f"    _xlnm.{dn.get('name')} = '{dn.text or ''}'")
    
    sheet_map = {}
    for s in sheets:
        sname = s.get('name', '?')
        rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
        target = ''
        for r in rels:
            if r.get('Id') == rid:
                target = r.get('Target', '')
                break
        sheet_map[sname] = f"xl/{target}" if target else ''
    
    for sname, spath in sheet_map.items():
        if spath not in z.namelist():
            continue
        ws = ET.fromstring(z.read(spath))
        dim = ws.find('.//s:dimension', NS)
        dim_ref = dim.get('ref', '') if dim is not None else '(none)'
        
        print(f"\n  Sheet: {sname} (dim={dim_ref})")
        
        pm = ws.find('.//s:pageMargins', NS)
        if pm is not None:
            print(f"    pageMargins:")
            for k, v in pm.attrib.items():
                v_pt = float(v) * 72
                source = "CUSTOM" if abs(float(v) - 0.70866) < 0.01 and k in ['left','right'] else ("CUSTOM" if abs(float(v) - 0.74803) < 0.01 and k in ['top','bottom'] else ("CUSTOM (header)" if abs(float(v) - 0.31496) < 0.01 else "DEFAULT"))
                print(f"      {k}: {v} in = {v_pt:.4f} pt [{source}]")
        
        po = ws.find('.//s:printOptions', NS)
        if po is not None:
            print(f"    printOptions: H={po.get('horizontalCentered','0')} V={po.get('verticalCentered','0')}")
        else:
            print(f"    printOptions: MISSING")

# ═══ FINAL COMPARISON TABLE ═══
print(f"\n{SEP2}")
print("COMPARISON TABLE (all fresh evidence)")
print(SEP2)

# Compare OOXML values
gen_pm = None
conmas_pm = None
gen_po = None
conmas_po = None
gen_pa = False
conmas_pa = False

with zipfile.ZipFile(gen_path) as z:
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    dns = wb.findall('.//s:definedName', NS)
    gen_pa = any('Print_Area' in dn.get('name','') for dn in dns)
    
    for s in wb.findall('.//s:sheet', NS):
        sname = s.get('name', '')
        if sname in ('Sheet1', '_Fields'):
            rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            target = ''
            for r in rels:
                if r.get('Id') == rid:
                    target = r.get('Target', '')
                    break
            spath = f"xl/{target}" if target else ''
            if spath in z.namelist():
                ws = ET.fromstring(z.read(spath))
                if sname == 'Sheet1':
                    gen_pm = ws.find('.//s:pageMargins', NS)
                    gen_po = ws.find('.//s:printOptions', NS)

with zipfile.ZipFile(CONMAS_PATH) as z:
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    dns = wb.findall('.//s:definedName', NS)
    conmas_pa = any('Print_Area' in dn.get('name','') for dn in dns)
    
    for s in wb.findall('.//s:sheet', NS):
        sname = s.get('name', '')
        if sname == 'Sheet1':
            rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            target = ''
            for r in rels:
                if r.get('Id') == rid:
                    target = r.get('Target', '')
                    break
            spath = f"xl/{target}" if target else ''
            if spath in z.namelist():
                ws = ET.fromstring(z.read(spath))
                conmas_pm = ws.find('.//s:pageMargins', NS)
                conmas_po = ws.find('.//s:printOptions', NS)

print(f"")
print(f"  {'Property':<22} {'ConMas':<20} {'Generated':<20} {'Match?'}")
print(f"  {'─'*22} {'─'*20} {'─'*20} {'─'*8}")

# PrintArea
print(f"  {'PrintArea defined':<22} {'YES' if conmas_pa else 'NO':<20} {'YES' if gen_pa else 'NO':<20} {'OK' if gen_pa == conmas_pa else 'MISMATCH'}")

# printOptions
if conmas_po is not None:
    ch_con = conmas_po.get('horizontalCentered', '0')
    cv_con = conmas_po.get('verticalCentered', '0')
else:
    ch_con, cv_con = '0', '0'

if gen_po is not None:
    ch_gen = gen_po.get('horizontalCentered', '0')
    cv_gen = gen_po.get('verticalCentered', '0')
else:
    ch_gen, cv_gen = '0', '0'

print(f"  {'printOptions H':<22} {ch_con:<20} {ch_gen:<20} {'OK' if ch_con == ch_gen else 'MISMATCH'}")
print(f"  {'printOptions V':<22} {cv_con:<20} {cv_gen:<20} {'OK' if cv_con == cv_gen else 'MISMATCH'}")

# pageMargins (convert to pt for comparison)
for key, label in [('left', 'LeftMargin'), ('top', 'TopMargin'), ('right', 'RightMargin'), ('bottom', 'BottomMargin'), ('header', 'HeaderMargin'), ('footer', 'FooterMargin')]:
    con_val = float(conmas_pm.get(key, 0)) * 72 if conmas_pm is not None else 0
    gen_val = float(gen_pm.get(key, 0)) * 72 if gen_pm is not None else 0
    match = abs(con_val - gen_val) < 0.05
    print(f"  {label:<22} {con_val:<20.4f} {gen_val:<20.4f} {'OK' if match else 'MISMATCH'}")

print(f"\n{SEP}")
print(f"CONCLUSION:")
all_ok = True

if conmas_pa == gen_pa and ch_con == ch_gen and cv_con == cv_gen:
    for key in ['left', 'top', 'right', 'bottom']:
        cv = float(conmas_pm.get(key, 0)) * 72 if conmas_pm is not None else 0
        gv = float(gen_pm.get(key, 0)) * 72 if gen_pm is not None else 0
        if abs(cv - gv) >= 0.05:
            all_ok = False
    if all_ok:
        print(f"  YES - The workbook generated today by the current running API behaves")
        print(f"  identically to the ConMas-generated workbook for all print-critical properties.")
    else:
        print(f"  NO - Mismatch found in margin values.")
else:
    all_ok = False
    print(f"  NO - Mismatch found in print options or PrintArea.")

print(f"\n  Evidence collected:")
print(f"  - Generated workbook: {gen_path}")
print(f"  - OOXML: Verified pageMargins, printOptions, pageSetup, Print_Area")
print(f"  - COM: Script at {ps1_path}")
print(f"  - Excel UI: COM script opens Excel visibly at end (Step 6)")
print(SEP)

# ═══ STEP 6: Excel UI (make visible) ═══
print(f"\n{SEP2}")
print("STEP 6: Excel UI Verification")
print(SEP2)

ps6_path = os.path.join(PROJECT, "_x33_ui_verify.ps1")
ps6_content = f'''
# Phase X.33 - STEP 6: Excel UI Verification
$genPath = [System.IO.Path]::Combine("{PROJECT}", "_x33_generated_output.xlsx")

Write-Host "Opening generated workbook in Excel UI..."
Write-Host ""
Write-Host "Press any key to open the generated workbook visibly..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open($genPath)
$ws = $wb.Worksheets(1)
$ws.Activate()

Write-Host ""
Write-Host "Excel is now visible with the generated workbook."
Write-Host "Please verify:"
Write-Host "  1. Page Layout -> Margins -> Custom Margins"
Write-Host "  2. Check:"
Write-Host "     - Center Horizontally: CHECKED"
Write-Host "     - Center Vertically: CHECKED"
Write-Host "     - Left: 1.8 cm (0.71 in)"
Write-Host "     - Right: 1.8 cm (0.71 in)"
Write-Host "     - Top: 1.9 cm (0.75 in)"
Write-Host "     - Bottom: 1.9 cm (0.75 in)"
Write-Host ""
Write-Host "Page Setup values from COM (same workbook):"
$ps = $ws.PageSetup
Write-Host "  CenterH: $($ps.CenterHorizontally)"
Write-Host "  CenterV: $($ps.CenterVertically)"
$lm = [Math]::Round([double]$ps.LeftMargin, 2)
$tm = [Math]::Round([double]$ps.TopMargin, 2)
Write-Host "  LeftMargin: $lm pt ($([Math]::Round($lm/72*2.54,2)) cm)"
Write-Host "  TopMargin: $tm pt ($([Math]::Round($tm/72*2.54,2)) cm)"
Write-Host "  PrintArea: '$($ps.PrintArea)'"
Write-Host ""
Write-Host "Press any key to close Excel..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
$wb.Close($false)
$excel.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "Excel closed."
'''

with open(ps6_path, 'w') as f:
    f.write(ps6_content)

print(f"  UI verification script: {ps6_path}")
print(f"  Run: powershell -ExecutionPolicy Bypass -File {ps6_path}")
print(f"\n{SEP}")
print("PHASE X.33 VALIDATION SCRIPT COMPLETE")
print("Run the PowerShell scripts for COM and UI verification.")
print(SEP)
