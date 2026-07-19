"""
Phase X.32: Automated pipeline trace with ConMas.xlsx
Stages A-H: Trace PageSetup values through the entire pipeline.
No code modifications - only investigation.
Uses only safe ASCII characters.
"""
import os, sys, json, subprocess, time, hashlib, zipfile
from xml.etree import ElementTree as ET

PROJECT = os.path.dirname(os.path.abspath(__file__))
CONMAS_PATH = os.path.join(PROJECT, "test_conmas_output.xlsx")
API_URL = "http://localhost:5090"
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

SEP = "=" * 70
SEP2 = "-" * 70

print(SEP)
print("PHASE X.32 - End-to-End Pipeline Trace with ConMas.xlsx")
print(SEP)

assert os.path.exists(CONMAS_PATH), "ConMas.xlsx not found: " + CONMAS_PATH
md5 = hashlib.md5(open(CONMAS_PATH,'rb').read()).hexdigest()[:12]
print(f"\nConMas.xlsx: {os.path.getsize(CONMAS_PATH):,} bytes, MD5={md5}")

# STAGE A: Ground Truth (from Phase X.31 COM verification)
print("\n" + SEP2)
print("STAGE A: Ground Truth (from Phase X.31 COM verification)")
print(SEP2)
print("""
  Property                    ConMas (COM)
  PrintArea                   $A$1:$D$12
  CenterHorizontally          True
  CenterVertically            True
  LeftMargin                  51.0236 pt (1.80 cm)
  RightMargin                 51.0236 pt (1.80 cm)
  TopMargin                   53.8583 pt (1.90 cm)
  BottomMargin                53.8583 pt (1.90 cm)
  Orientation                 portrait
  Zoom                        100
  FitToPagesWide              1
  FitToPagesTall              1
""")

# STAGE B: Upload ConMas.xlsx to upload-excel
print(SEP2)
print("STAGE B: Upload ConMas.xlsx to upload-excel endpoint")
print(SEP2)

upload_path = os.path.join(PROJECT, "_temp_upload.xlsx")
import shutil
shutil.copy2(CONMAS_PATH, upload_path)

result = subprocess.run([
    'curl', '-s', '-X', 'POST',
    f'{API_URL}/api/form/upload-excel',
    '-F', f'file=@{upload_path};filename=ConMas.xlsx',
    '--connect-timeout', '30',
    '--max-time', '60'
], capture_output=True, text=True)

form_def = None
if result.returncode == 0:
    try:
        response = json.loads(result.stdout)
        data = response.get('data', response.get('Data', {}))
        if isinstance(data, dict):
            form_def = data.get('formDefinition', data.get('FormDefinition', {}))
            if not form_def:
                form_def = data
            
            success = response.get('success', response.get('Success', False))
            msg = response.get('message', response.get('Message', 'N/A'))
            print(f"\n  Upload response: success={success}")
            print(f"  Message: {msg}")
            
            # STAGE C: Inspect FormDefinition PageSettings
            print("\n" + SEP2)
            print("STAGE C: FormDefinition (after WorkbookReader + JSON serialization)")
            print(SEP2)
            
            # Check FieldsPageSettings
            fld_ps = form_def.get('fieldsPageSettings', form_def.get('FieldsPageSettings', None))
            print(f"\n  FieldsPageSettings on FormDefinition: {'PRESENT' if fld_ps else 'NOT PRESENT (good - merge removed)'}")
            
            sheets = form_def.get('sheets', form_def.get('Sheets', []))
            print(f"\n  Sheets in FormDefinition: {len(sheets)}")
            
            for i, sheet in enumerate(sheets):
                sname = sheet.get('name', sheet.get('Name', '?'))
                ps = sheet.get('pageSettings', sheet.get('PageSettings', {}))
                
                print(f"\n  Sheet [{i}]: {sname}")
                if ps:
                    print(f"    CenterHorizontally: {ps.get('centerHorizontally', ps.get('CenterHorizontally', 'N/A'))}")
                    print(f"    CenterVertically:   {ps.get('centerVertically', ps.get('CenterVertically', 'N/A'))}")
                    print(f"    LeftMargin:         {ps.get('leftMargin', ps.get('LeftMargin', 'N/A'))}")
                    print(f"    RightMargin:        {ps.get('rightMargin', ps.get('RightMargin', 'N/A'))}")
                    print(f"    TopMargin:          {ps.get('topMargin', ps.get('TopMargin', 'N/A'))}")
                    print(f"    BottomMargin:       {ps.get('bottomMargin', ps.get('BottomMargin', 'N/A'))}")
                    print(f"    Orientation:        {ps.get('orientation', ps.get('Orientation', 'N/A'))}")
                    print(f"    Zoom:               {ps.get('zoom', ps.get('Zoom', 'N/A'))}")
                    print(f"    FitToPagesWide:     {ps.get('fitToPagesWide', ps.get('FitToPagesWide', 'N/A'))}")
                    print(f"    FitToPagesTall:     {ps.get('fitToPagesTall', ps.get('FitToPagesTall', 'N/A'))}")
                else:
                    print(f"    PageSettings: MISSING - will use ConvertCaptureToForm defaults (70pt margins, no centering)")
                
                pa = sheet.get('printArea', sheet.get('PrintArea', {}))
                print(f"    PrintArea Address: {pa.get('address', pa.get('Address', 'N/A'))}")
            
            # Compare against Stage A (Ground Truth)
            print("\n" + SEP2)
            print("COMPARISON: Stage A (Ground Truth) vs Stage C (FormDefinition)")
            print(SEP2)
            
            if sheets and len(sheets) > 0:
                ps = sheets[0].get('pageSettings', sheets[0].get('PageSettings', {}))
                ch = ps.get('centerHorizontally', ps.get('CenterHorizontally', 'MISSING'))
                cv = ps.get('centerVertically', ps.get('CenterVertically', 'MISSING'))
                lm = ps.get('leftMargin', ps.get('LeftMargin', 'MISSING'))
                tm = ps.get('topMargin', ps.get('TopMargin', 'MISSING'))
                
                print(f"")
                print(f"  Property          | Ground Truth (A)  | FormDefinition (C) | Match?")
                print(f"  ------------------+--------------------+--------------------+--------")
                print(f"  CenterH           | True              | {str(ch):<18} | {'OK' if ch == True else 'MISMATCH'}")
                print(f"  CenterV           | True              | {str(cv):<18} | {'OK' if cv == True else 'MISMATCH'}")
                print(f"  LeftMargin        | 51.0236           | {str(lm):<18} | {'OK' if lm and float(str(lm)) == 51.0236 else 'MISMATCH'}")
                print(f"  TopMargin         | 53.8583           | {str(tm):<18} | {'OK' if tm and float(str(tm)) == 53.8583 else 'MISMATCH'}")
            else:
                print(f"  No sheets found in FormDefinition - cannot compare")
            
            # STAGE D+E+F+G: Post to output-excel
            print("\n" + SEP2)
            print("STAGE D+E+F+G: Post FormDefinition to output-excel")
            print(SEP2)
            
            gen_path = os.path.join(PROJECT, "_temp_generated_output.xlsx")
            result2 = subprocess.run([
                'curl', '-s', '-X', 'POST',
                f'{API_URL}/api/form/output-excel',
                '-H', 'Content-Type: application/json',
                '-d', json.dumps(form_def),
                '--connect-timeout', '60',
                '--max-time', '120',
                '-o', gen_path
            ], capture_output=True, text=True)
            
            print(f"\n  curl exit code: {result2.returncode}")
            
            if os.path.exists(gen_path) and os.path.getsize(gen_path) > 0:
                gen_size = os.path.getsize(gen_path)
                gen_md5 = hashlib.md5(open(gen_path,'rb').read()).hexdigest()[:12]
                print(f"  Generated output: {gen_size:,} bytes, MD5={gen_md5}")
                
                # STAGE H: OOXML verification
                print("\n" + SEP2)
                print("STAGE H: OOXML Verification of Generated Output")
                print(SEP2)
                
                with zipfile.ZipFile(gen_path) as z:
                    names = z.namelist()
                    print(f"\n  ZIP entries: {len(names)}")
                    
                    wb = ET.fromstring(z.read('xl/workbook.xml'))
                    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
                    sheets = wb.findall('.//s:sheet', NS)
                    dns = wb.findall('.//s:definedName', NS)
                    
                    print(f"  Sheets: {len(sheets)}")
                    for s in sheets:
                        sname = s.get('name', '?')
                        state = s.get('state', 'visible')
                        print(f"    [{state}] {sname}")
                    
                    print(f"  Defined names: {len(dns)}")
                    for dn in dns:
                        print(f"    _xlnm.{dn.get('name')} = '{dn.text or ''}'")
                    
                    # Build sheet->path map
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
                                print(f"      {k}: {v} in = {v_pt:.4f} pt")
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
                        
                        # Compare with Stage A
                        print(f"\n    VS Ground Truth (Stage A):")
                        if pm is not None:
                            lm_val = float(pm.get('left', 0))
                            tm_val = float(pm.get('top', 0))
                            lm_match = abs(lm_val * 72 - 51.0236) < 0.1
                            tm_match = abs(tm_val * 72 - 53.8583) < 0.1
                            print(f"      LeftMargin: {lm_val*72:.4f}pt vs 51.0236pt -> {'OK' if lm_match else 'MISMATCH'}")
                            print(f"      TopMargin:  {tm_val*72:.4f}pt vs 53.8583pt -> {'OK' if tm_match else 'MISMATCH'}")
                        if po is not None:
                            h = po.get('horizontalCentered', '0')
                            v = po.get('verticalCentered', '0')
                            print(f"      CenterH: {h} vs True -> {'OK' if h == '1' else 'MISMATCH'}")
                            print(f"      CenterV: {v} vs True -> {'OK' if v == '1' else 'MISMATCH'}")
                        
                        print(f"      PrintArea defined: {'YES' if any('Print_Area' in dn.get('name','') for dn in dns) else 'MISSING'} -> {'OK' if any('Print_Area' in dn.get('name','') for dn in dns) else 'MISMATCH'}")
                    
                    # Print comparison table
                    print("\n" + SEP2)
                    print("COMPARISON TABLE: All Stages")
                    print(SEP2)
                    print()
                    print("  Stage    | PrintArea | CenterH | CenterV | LeftMargin | TopMargin | Status")
                    print("  ---------+-----------+---------+---------+------------+-----------+--------")
                    print("  A (COM)  | $A$1:$D$12| True    | True    | 51.0236pt  | 53.8583pt | GROUND TRUTH")
                    
                    # Stage B/C we computed above                        if sheets and len(sheets) > 0 and sheets[0].get('pageSettings', sheets[0].get('PageSettings', {})):
                            ps0 = sheets[0].get('pageSettings', sheets[0].get('PageSettings', {}))
                            pa0 = sheets[0].get('printArea', sheets[0].get('PrintArea', {}) or {})
                            pa_addr = pa0.get('address', pa0.get('Address', 'N/A')) if pa0 else 'N/A'
                            ch = ps0.get('centerHorizontally', ps0.get('CenterHorizontally', '?'))
                            cv = ps0.get('centerVertically', ps0.get('CenterVertically', '?'))
                            lm = ps0.get('leftMargin', ps0.get('LeftMargin', '?'))
                            tm = ps0.get('topMargin', ps0.get('TopMargin', '?'))
                            status = "OK" if ch == True and cv == True and abs(float(str(lm)) - 51.0236) < 1 else "MISMATCH"
                            print(f"  C (JSON) | {str(pa_addr):<9} | {str(ch):<7} | {str(cv):<7} | {str(lm):<10} | {str(tm):<9} | {status}")
                        else:
                            print(f"  C (JSON) | N/A        | N/A     | N/A     | N/A        | N/A       | NO SHEET DATA")
                    
                    # Stage H
                    for sname, spath in sheet_map.items():
                        if spath not in names:
                            continue
                        ws = ET.fromstring(z.read(spath))
                        pm = ws.find('.//s:pageMargins', NS)
                        po = ws.find('.//s:printOptions', NS)
                        
                        has_pa = any('Print_Area' in dn.get('name','') for dn in dns)
                        lm_val = float(pm.get('left', 0)) * 72 if pm is not None else 0
                        tm_val = float(pm.get('top', 0)) * 72 if pm is not None else 0
                        ch_val = po.get('horizontalCentered', '0') if po is not None else 'MISSING'
                        cv_val = po.get('verticalCentered', '0') if po is not None else 'MISSING'
                        
                        lm_match = abs(lm_val - 51.0236) < 0.1
                        tm_match = abs(tm_val - 53.8583) < 0.1
                        h_match = ch_val == '1'
                        v_match = cv_val == '1'
                        
                        status = "OK" if (has_pa and h_match and v_match and lm_match and tm_match) else "MISMATCH"
                        print(f"  H (OOXML)| {'YES' if has_pa else 'NO':<9} | {ch_val:<7} | {cv_val:<7} | {lm_val:<10.4f} | {tm_val:<9.4f} | {status}")
                    
                    if dns:
                        print(f"  Note: {len(dns)} defined name(s) found")
                    else:
                        print(f"  Note: NO defined names found")
                        
            else:
                print(f"  Failed: Generated output is empty or missing")
                if result2.stderr:
                    print(f"  stderr: {result2.stderr[:500]}")
        else:
            print(f"  Unexpected data type: {type(data)}")
    except json.JSONDecodeError as e:
        print(f"  JSON parse error: {e}")
        print(f"  Raw response: {result.stdout[:500]}")
else:
    print(f"  Upload failed (exit code {result.returncode})")
    print(f"  stderr: {result.stderr[:500]}")
    print(f"  stdout: {result.stdout[:500]}")

# Also read the C# API log for CP1-CP6 diagnostic output
api_log = os.path.join(PROJECT, "api_trace.log")
if os.path.exists(api_log):
    print("\n" + SEP2)
    print("C# API LOG (checkpoints CP1-CP6)")
    print(SEP2)
    with open(api_log, 'r', encoding='utf-8', errors='replace') as f:
        for line in f:
            if any(tag in line for tag in ['[X28]', '[FORENSIC]', 'PrintArea', 'PageSetup', 'CenterH', 'CenterV', 'LeftMargin', 'TopMargin', 'Print_Area', 'CP1', 'CP2', 'CP3', 'CP4', 'CP5', 'CP6']):
                cleaned = line.strip()[:200]
                print(f"  {cleaned}")

# Cleanup temp files
for f in ['_temp_upload.xlsx', '_temp_form_definition.json', '_temp_generated_output.xlsx']:
    try: os.remove(os.path.join(PROJECT, f))
    except: pass

print("\n" + SEP)
print("PHASE X.32 TRACE COMPLETE")
print(SEP)
