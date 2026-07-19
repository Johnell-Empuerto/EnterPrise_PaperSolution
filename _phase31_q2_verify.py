"""
Phase X.31 Q2: Verify PageSetup from 3 independent sources
1. COM (Excel Interop)
2. OOXML (direct ZIP read)
3. Excel UI (make visible, user confirms)
"""
import os, sys, zipfile, json, hashlib
from xml.etree import ElementTree as ET
from datetime import datetime

PROJECT = os.path.dirname(os.path.abspath(__file__))
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

workbooks = {
    "Our Output": os.path.join(PROJECT, "test_our_output.xlsx"),
    "ConMas": os.path.join(PROJECT, "test_conmas_output.xlsx"),
}

print("=" * 80)
print("PHASE X.31 — Q2: PageSetup Verification (3 sources)")
print("=" * 80)

for label, path in workbooks.items():
    assert os.path.exists(path), f"Missing: {path}"
    print(f"\n{'─'*80}")
    print(f"WORKBOOK: {label}")
    print(f"  Path: {path}")
    print(f"  Size: {os.path.getsize(path):,} bytes")
    print(f"  MD5: {hashlib.md5(open(path,'rb').read()).hexdigest()[:12]}")
    print(f"  Modified: {datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S')}")
    
    # ── SOURCE 1: OOXML Analysis ──
    print(f"\n  ── SOURCE 1: OOXML (direct ZIP read) ──")
    
    with zipfile.ZipFile(path) as z:
        namelist = z.namelist()
        
        # Parse workbook.xml for sheets and defined names
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        sheets = wb.findall('.//s:sheet', NS)
        dns = wb.findall('.//s:definedName', NS)
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        
        # Build sheet map
        sheet_map = {}
        for s in sheets:
            sname = s.get('name', '?')
            state = s.get('state', 'visible')
            rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            target = ''
            for r in rels:
                if r.get('Id') == rid:
                    target = r.get('Target', '')
                    break
            sheet_path = f"xl/{target}" if target else ''
            sheet_map[sname] = (state, sheet_path)
        
        print(f"    Sheets: {len(sheets)}")
        for sname, (state, spath) in sheet_map.items():
            print(f"      {sname} ({state}) -> {spath}")
        
        print(f"    Defined names: {len(dns)}")
        for dn in dns:
            print(f"      _xlnm.{dn.get('name')} = {dn.text or ''}")
        
        # Analyze each sheet XML
        for sname, (state, spath) in sheet_map.items():
            if spath not in z.namelist():
                continue
            ws = ET.fromstring(z.read(spath))
            
            print(f"\n    Sheet: {sname}")
            
            dim = ws.find('.//s:dimension', NS)
            if dim is not None:
                print(f"      dimension: {dim.get('ref', '')}")
            
            pm = ws.find('.//s:pageMargins', NS)
            if pm is not None:
                margins = dict(pm.attrib)
                print(f"      pageMargins (in inches):")
                for k, v in margins.items():
                    v_pt = float(v) * 72
                    v_cm = float(v) * 2.54
                    source = "CUSTOM" if abs(float(v) - 0.31496) > 0.001 and abs(float(v) - 0.7) > 0.005 else "EXCEL DEFAULT"
                    print(f"        {k}: {v} in = {v_pt:.4f} pt = {v_cm:.2f} cm [{source}]")
            else:
                print(f"      pageMargins: MISSING")
            
            po = ws.find('.//s:printOptions', NS)
            if po is not None:
                h = po.get('horizontalCentered', '0')
                v = po.get('verticalCentered', '0')
                print(f"      printOptions: H={h} V={v}")
            else:
                print(f"      printOptions: MISSING")
            
            ps = ws.find('.//s:pageSetup', NS)
            if ps is not None:
                print(f"      pageSetup: {dict((k,v) for k,v in ps.attrib.items() if '}' not in k)}")
            else:
                print(f"      pageSetup: MISSING")
    
    # ── SOURCE 2: COM (will be executed separately) ──
    print(f"\n  ── SOURCE 2: COM (separate PowerShell script) ──")
    print(f"      See _phase31_com_verify.ps1")

print(f"\n{'='*80}")
print("Q2 OOXML analysis complete. COM verification follows separately.")
print("=" * 80)
