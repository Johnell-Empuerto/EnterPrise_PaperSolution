import os, sys, zipfile
from xml.etree import ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

legacy_path = r'C:\Users\MCF-JOHNELLEEMPUERTO\Downloads\FormTest - Copy.xlsx'

ns = {
    's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

with zipfile.ZipFile(legacy_path, 'r') as z:
    all_names = sorted(z.namelist())
    
    # ========== 1. Content Types ==========
    print("=" * 70)
    print("1. CONTENT_TYPES.xml")
    print("=" * 70)
    ct = ET.fromstring(z.read('[Content_Types].xml'))
    ct_ns = {'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'}
    for override in ct.findall('.//ct:Override', ct_ns):
        print(f"  Override: {override.get('PartName', ''):60s} → {override.get('ContentType', '')}")
    for default in ct.findall('.//ct:Default', ct_ns):
        print(f"  Default:  .{default.get('Extension', ''):10s} → {default.get('ContentType', '')}")

    # ========== 2. Relationships ==========
    for relfile in sorted([n for n in all_names if n.endswith('.rels')]):
        print(f"\n{'=' * 70}")
        print(f"2. RELATIONSHIPS: {relfile}")
        print("=" * 70)
        rel_xml = ET.fromstring(z.read(relfile))
        rel_ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        for rel in rel_xml.findall('.//r:Relationship', rel_ns):
            print(f"  {rel.get('Id', ''):10s} Type={rel.get('Type', '').split('/')[-1]:50s} Target={rel.get('Target', '')}")

    # ========== 3. Sheet3 (ExcelOutputSetting) ==========
    print(f"\n{'=' * 70}")
    print("3. EXCELOUTPUTSETTING SHEET (sheet3.xml)")
    print("=" * 70)
    ws3 = ET.fromstring(z.read('xl/worksheets/sheet3.xml'))
    rows3 = ws3.findall('.//s:row', ns)
    print(f"  Total rows: {len(rows3)}")
    print(f"  Sheet dimension: {ws3.find('.//s:dimension', ns).get('ref', '') if ws3.find('.//s:dimension', ns) else 'unknown'}")
    for row in rows3[:5]:  # Only first 5 rows, the content is sharedStrings
        r_num = row.get('r', '?')
        cells = row.findall('s:c', ns)
        for cell in cells:
            ref = cell.get('r', '')
            t = cell.get('t', '')
            v = cell.find('s:v', ns)
            val = v.text if v is not None and v.text is not None else ''
            print(f"  Row {r_num}, {ref}: type={t} val_idx={val}")

    # ========== 4. Sheet1 (content sheet) cell values ==========
    print(f"\n{'=' * 70}")
    print("4. SHEET1 CELL VALUES (via shared strings)")
    print("=" * 70)
    
    # Read shared strings
    ss_xml = ET.fromstring(z.read('xl/sharedStrings.xml'))
    ss_items = []
    for si in ss_xml.findall('.//s:si', ns):
        t = si.find('.//s:t', ns)
        text = t.text if t is not None and t.text is not None else ''
        ss_items.append(text)
    
    # Read sheet2.xml (which has the merged cells and cell data)
    ws2 = ET.fromstring(z.read('xl/worksheets/sheet2.xml'))
    merges = ws2.findall('.//s:mergeCells/s:mergeCell', ns)
    print(f"  Merged regions ({len(merges)}):")
    for m in merges:
        print(f"    {m.get('ref', '')}")
    
    rows2 = ws2.findall('.//s:row', ns)
    print(f"\n  Cells by shared string index:")
    for row in rows2:
        r_num = row.get('r', '?')
        cells = row.findall('s:c', ns)
        for cell in cells:
            ref = cell.get('r', '')
            t = cell.get('t', '')
            v = cell.find('s:v', ns)
            if v is not None and v.text is not None:
                if t == 's':
                    idx = int(v.text)
                    val = ss_items[idx] if idx < len(ss_items) else f'[OUT_OF_RANGE:{idx}]'
                    print(f"    {ref}: {repr(val[:100])}")
                else:
                    print(f"    {ref}: (inline) {v.text}")

    # ========== 5. VML Drawing ==========
    print(f"\n{'=' * 70}")
    print("5. VML DRAWING (vmlDrawing1.vml)")
    print("=" * 70)
    vml = z.read('xl/drawings/vmlDrawing1.vml').decode('utf-8', errors='replace')
    # Count comment shapes
    import re
    shape_count = len(re.findall(r'<v:shape', vml))
    print(f"  VML shapes: {shape_count}")
    # Show first shape as example
    shapes = re.findall(r'<v:shape[^>]*>.*?</v:shape>', vml, re.DOTALL)
    for i, shape in enumerate(shapes):
        # Extract key attributes
        row_val = re.search(r'row="(\d+)"', shape)
        col_val = re.search(r'col="(\d+)"', shape)
        row_s = row_val.group(1) if row_val else '?'
        col_s = col_val.group(1) if col_val else '?'
        print(f"\n  Shape #{i}: row={row_s} col={col_s}")
        # Extract text
        text_match = re.search(r'<!--(.*?)-->', shape, re.DOTALL)
        if text_match:
            print(f"    Comment text: {repr(text_match.group(1)[:200])}")

    # ========== 6. Workbook properties ==========
    print(f"\n{'=' * 70}")
    print("6. WORKBOOK PROPERTIES")
    print("=" * 70)
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    # Defined names
    defined_names = wb.findall('.//s:definedName', ns)
    if defined_names:
        print(f"  Defined names ({len(defined_names)}):")
        for dn in defined_names:
            name = dn.get('name', '')
            text = dn.text if dn.text else ''
            print(f"    {name} = {text}")
    else:
        print("  No defined names found")

    # ========== 7. Document properties ==========
    print(f"\n{'=' * 70}")
    print("7. DOCUMENT PROPERTIES")
    print("=" * 70)
    for docprop in ['docProps/core.xml', 'docProps/app.xml']:
        if docprop in z.namelist():
            print(f"\n  --- {docprop} ---")
            dp = ET.fromstring(z.read(docprop))
            def print_elem(e, indent=4):
                for child in e:
                    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    text = (child.text or '').strip()
                    if text or len(child) == 0:
                        print(f"{' ' * indent}{tag} = {text}")
                    print_elem(child, indent + 2)
            print_elem(dp)
