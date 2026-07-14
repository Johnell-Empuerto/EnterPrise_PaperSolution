"""
Investigate How ConMas Stores Cluster Metadata Inside Excel Workbooks

Checks: Comments, Notes, Hidden Sheets, Defined Names,
        Custom XML Parts, Custom Properties, VBA, embedded metadata.
"""

import os, sys, zipfile, json
from pathlib import Path
from xml.etree import ElementTree as ET

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent

# Workbooks to inspect
WORKBOOKS = {
    "547": os.path.join(_PROJECT_ROOT, "[V3.1_Sample]アンケート用紙.xlsx"),
    "546": os.path.join(_PROJECT_ROOT, "Investigation_546", "original.xlsx"),
    "old_form": os.path.join(_PROJECT_ROOT, "old_form.xlsx"),
}


def inspect_wb(label, wb_path):
    if not os.path.isfile(wb_path):
        print(f"\n{'='*70}")
        print(f"  {label}: FILE NOT FOUND: {wb_path}")
        return

    print(f"\n{'='*70}")
    print(f"  {label}: {os.path.basename(wb_path)}")
    print(f"{'='*70}")

    try:
        with zipfile.ZipFile(wb_path, 'r') as z:
            names = z.namelist()
    except Exception as e:
        print(f"  ERROR opening XLSX: {e}")
        return

    # ── 1. List all files ──────────────────────────────────────────────
    print(f"\n  [1] TOTAL FILES IN ARCHIVE: {len(names)}")
    for n in sorted(names):
        print(f"      {n}")

    # ── 2. Content Types ───────────────────────────────────────────────
    print(f"\n  [2] CONTENT TYPES (Custom XML parts)")
    try:
        with zipfile.ZipFile(wb_path, 'r') as z:
            ct = z.read('[Content_Types].xml').decode('utf-8')
        root = ET.fromstring(ct)
        ns = {'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'}
        overrides = root.findall('ct:Override', ns)
        for ov in overrides:
            pn = ov.get('PartName', '')
            ct_name = ov.get('ContentType', '')
            # Filter for non-standard content types
            if 'custom' in ct_name.lower() or 'vba' in ct_name.lower() or 'macro' in ct_name.lower():
                print(f"      Custom: PartName={pn}, ContentType={ct_name}")
    except Exception as e:
        print(f"      (Error: {e})")

    # ── 3. workbook.xml — sheets, defined names ────────────────────────
    print(f"\n  [3] WORKBOOK.XML (Sheets & Defined Names)")
    try:
        with zipfile.ZipFile(wb_path, 'r') as z:
            wbxml = z.read('xl/workbook.xml').decode('utf-8')
        root = ET.fromstring(wbxml)
        ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
              'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
        # Sheets
        for sheet in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
            name = sheet.get('name', '')
            state = sheet.get('state', 'visible')
            rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            print(f"      Sheet: name={name}, state={state}, rId={rid}")
        # Defined names
        for dn in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}definedName'):
            name = dn.get('name', '')
            text = dn.text or ''
            print(f"      DefinedName: {name} = {text[:200]}")
    except Exception as e:
        print(f"      (Error: {e})")

    # ── 4. Comments ────────────────────────────────────────────────────
    print(f"\n  [4] COMMENTS")
    try:
        with zipfile.ZipFile(wb_path, 'r') as z:
            comment_files = [n for n in names if 'comments' in n.lower()]
            for cf in comment_files:
                content = z.read(cf).decode('utf-8', errors='replace')
                root = ET.fromstring(content)
                ns_comm = {'cm': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for auth in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}author'):
                    print(f"      Author: {auth.text}")
                for comment in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}comment'):
                    ref = comment.get('ref', '')
                    # Get text
                    text_el = comment.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                    text = text_el.text if text_el is not None else ''
                    print(f"      Cell {ref}: text='{text[:300]}'")
    except Exception as e:
        print(f"      (Error: {e})")

    # ── 5. Custom XML Parts (xl/customXml) ────────────────────────────
    print(f"\n  [5] CUSTOM XML PARTS")
    custom_xml = [n for n in names if 'customXml' in n or 'custom' in n.lower()]
    for cx in custom_xml:
        try:
            with zipfile.ZipFile(wb_path, 'r') as z:
                content = z.read(cx).decode('utf-8', errors='replace')
            print(f"      {cx}: {content[:500]}")
        except Exception as e:
            print(f"      {cx}: (Error: {e})")

    # ── 6. Hidden / VeryHidden sheets via sheet XML visibility ─────────
    print(f"\n  [6] SHEET HIDING - Checking sheet files for visibility overrides")
    try:
        with zipfile.ZipFile(wb_path, 'r') as z:
            sheet_files = [n for n in names if n.startswith('xl/worksheets/') and n.endswith('.xml') and 'rels' not in n]
            for sf in sheet_files:
                content = z.read(sf).decode('utf-8', errors='replace')
                root = ET.fromstring(content)
                sheet_el = root.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetPr')
                visibility = 'visible'
                if sheet_el is not None:
                    visibility = sheet_el.get('outlinePr', None)
                # Check for hidden rows/cols
                print(f"      {sf}: size={len(content)} bytes")
    except Exception as e:
        print(f"      (Error: {e})")

    # ── 7. VBA / Macros ───────────────────────────────────────────────
    print(f"\n  [7] VBA / MACROS")
    vba_files = [n for n in names if 'vba' in n.lower() or 'macro' in n.lower()]
    if vba_files:
        for vf in vba_files:
            print(f"      FOUND: {vf}")
    else:
        print(f"      None found")

    # ── 8. Custom Document Properties ──────────────────────────────────
    print(f"\n  [8] CUSTOM DOCUMENT PROPERTIES")
    try:
        with zipfile.ZipFile(wb_path, 'r') as z:
            if 'docProps/custom.xml' in names:
                content = z.read('docProps/custom.xml').decode('utf-8', errors='replace')
                print(f"      {content[:1000]}")
            elif 'docProps/app.xml' in names:
                content = z.read('docProps/app.xml').decode('utf-8', errors='replace')
                print(f"      (app.xml): {content[:500]}")
            else:
                print(f"      No custom properties found")
    except Exception as e:
        print(f"      (Error: {e})")

    # ── 9. Drawing / Shapes ────────────────────────────────────────────
    print(f"\n  [9] DRAWINGS / SHAPES")
    drawing_files = [n for n in names if 'drawing' in n.lower() or 'shape' in n.lower()]
    for df in drawing_files:
        try:
            with zipfile.ZipFile(wb_path, 'r') as z:
                content = z.read(df).decode('utf-8', errors='replace')
            print(f"      {df}: {len(content)} bytes")
            if len(content) < 5000:
                print(f"      Content: {content[:500]}")
        except Exception as e:
            print(f"      {df}: (Error: {e})")

    # ── 10. Summary ────────────────────────────────────────────────────
    print(f"\n  [10] SUMMARY")
    keywords_found = []
    interesting_parts = ['comments', 'customXml', 'vba', 'drawing', 'metadata', 'sharedStrings',
                         'calcChain', 'volatileDependencies', 'pivotCache',
                         'queryTable', 'oleObject', 'activeX', 'ctrlProp']
    for part in interesting_parts:
        matched = [n for n in names if part in n.lower()]
        if matched:
            keywords_found.append(f"{part}: {len(matched)} files")
    print(f"      Keywords found: {', '.join(keywords_found)}")


# ── Main ──────────────────────────────────────────────────────────────────────
for label, path in WORKBOOKS.items():
    inspect_wb(label, path)

print(f"\n{'='*70}")
print("  INVESTIGATION COMPLETE")
print(f"{'='*70}")
