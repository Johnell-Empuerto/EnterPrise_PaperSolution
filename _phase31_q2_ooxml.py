"""
Phase X.31 Q2: OOXML verification (simple output, no Unicode)
"""
import os, sys, zipfile, json, hashlib
from xml.etree import ElementTree as ET

PROJECT = os.path.dirname(os.path.abspath(__file__))
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

workbooks = {
    "Our Output": os.path.join(PROJECT, "test_our_output.xlsx"),
    "ConMas": os.path.join(PROJECT, "test_conmas_output.xlsx"),
}

print("=" * 70)
print("PHASE X.31 Q2 - OOXML PageSetup Verification")
print("=" * 70)

for label, path in workbooks.items():
    print(f"\n--- WORKBOOK: {label} ---")
    print(f"  Size: {os.path.getsize(path):,} bytes")
    print(f"  MD5: {hashlib.md5(open(path,'rb').read()).hexdigest()[:12]}")
    
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        
        sheets = wb.findall('.//s:sheet', NS)
        dns = wb.findall('.//s:definedName', NS)
        
        print(f"\n  SHEETS: {len(sheets)}")
        for s in sheets:
            sname = s.get('name', '?')
            state = s.get('state', 'visible')
            rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            target = ''
            for r in rels:
                if r.get('Id') == rid:
                    target = r.get('Target', '')
                    break
            spath = f"xl/{target}" if target else ''
            print(f"    [{state}] {sname} -> {spath}")
        
        print(f"\n  DEFINED NAMES: {len(dns)}")
        for dn in dns:
            print(f"    _xlnm.{dn.get('name')} = '{dn.text or ''}'")
        
        for s in sheets:
            sname = s.get('name', '?')
            rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            target = ''
            for r in rels:
                if r.get('Id') == rid:
                    target = r.get('Target', '')
                    break
            spath = f"xl/{target}" if target else ''
            
            if spath not in z.namelist():
                continue
            
            ws = ET.fromstring(z.read(spath))
            
            print(f"\n  --- Sheet: {sname} ---")
            
            dim = ws.find('.//s:dimension', NS)
            if dim is not None:
                print(f"    dimension: {dim.get('ref', '')}")
            
            pm = ws.find('.//s:pageMargins', NS)
            if pm is not None:
                print(f"    pageMargins:")
                for k, v in pm.attrib.items():
                    v_pt = float(v) * 72
                    v_cm = float(v) * 2.54
                    is_default = abs(float(v) - 0.7) < 0.01 or abs(float(v) - 0.75) < 0.01 or abs(float(v) - 0.3) < 0.01
                    source = "CUSTOM" if not is_default and abs(float(v) - 0.70866) < 0.01 else ("CUSTOM (header)" if abs(float(v) - 0.31496) < 0.01 else "EXCEL DEFAULT")
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
                attrs = dict(ps.attrib)
                # Remove namespace attributes
                clean = {k.split('}')[-1]: v for k,v in attrs.items()}
                print(f"    pageSetup: {clean}")
            else:
                print(f"    pageSetup: MISSING")

print(f"\n{'='*70}")
print("OOXML analysis complete.")
