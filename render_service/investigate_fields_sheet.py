"""
Deep-dive into the _Fields hidden sheet found in all ConMas templates.
Extract the cell contents and compare across templates.
Also deep-dive into comments in template 547.
"""

import os, sys, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent

WORKBOOKS = {
    "547 (アンケート用紙)": os.path.join(_PROJECT_ROOT, "[V3.1_Sample]アンケート用紙.xlsx"),
    "546 (Investigation)": os.path.join(_PROJECT_ROOT, "Investigation_546", "original.xlsx"),
    "old_form": os.path.join(_PROJECT_ROOT, "old_form.xlsx"),
}

SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def get_sheet_map(wb_path):
    """Map rId to sheet name and state from workbook.xml"""
    with zipfile.ZipFile(wb_path, 'r') as z:
        wbxml = z.read('xl/workbook.xml').decode('utf-8')
    root = ET.fromstring(wbxml)
    sheets = {}
    for sheet in root.iter(f'{{{SHEET_NS}}}sheet'):
        name = sheet.get('name', '')
        state = sheet.get('state', 'visible')
        rid = sheet.get(f'{{{R_NS}}}id', '')
        # Get the relationship target
        sheets[rid] = {'name': name, 'state': state}
    # Get relationships
    with zipfile.ZipFile(wb_path, 'r') as z:
        rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rels_root = ET.fromstring(rels_xml)
    rels_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    for rel in rels_root.iter(f'{{{rels_ns}}}Relationship'):
        rid = rel.get('Id', '')
        target = rel.get('Target', '')
        if rid in sheets:
            sheets[rid]['target'] = target
    return sheets


def dump_sheet(wb_path, sheet_target, label):
    """Dump all cell contents from a worksheet XML file."""
    with zipfile.ZipFile(wb_path, 'r') as z:
        # The target is relative to xl/ directory
        full_path = f'xl/{sheet_target}'
        if full_path not in z.namelist():
            print(f"      NOT FOUND: {full_path}")
            return
        content = z.read(full_path).decode('utf-8', errors='replace')
    
    root = ET.fromstring(content)
    
    # Get shared strings if available
    sst = {}
    try:
        with zipfile.ZipFile(wb_path, 'r') as z:
            if 'xl/sharedStrings.xml' in z.namelist():
                sst_content = z.read('xl/sharedStrings.xml').decode('utf-8', errors='replace')
                sst_root = ET.fromstring(sst_content)
                for i, si in enumerate(sst_root.iter(f'{{{SHEET_NS}}}si')):
                    texts = []
                    for t in si.iter(f'{{{SHEET_NS}}}t'):
                        if t.text:
                            texts.append(t.text)
                    sst[i] = ' '.join(texts)
    except:
        pass
    
    print(f"\n    [{label}] Sheet contents:")
    
    # Dump the sheet data rows
    sheet_data = root.find(f'{{{SHEET_NS}}}sheetData')
    if sheet_data is None:
        print("      No sheet data found")
        return
    
    row_count = 0
    for row in sheet_data.iter(f'{{{SHEET_NS}}}row'):
        r = row.get('r', '')
        # Skip empty rows
        has_content = False
        for c in row.iter(f'{{{SHEET_NS}}}c'):
            r_ref = c.get('r', '')
            t = c.get('t', '')  # type: 's' = shared string, 'str' = inline, '' = number
            v_el = c.find(f'{{{SHEET_NS}}}v')
            v = v_el.text if v_el is not None else ''
            
            if t == 's' and v and v.isdigit():
                v = sst.get(int(v), v)
            
            if v.strip():
                has_content = True
                print(f"      Cell {r_ref}: type={t}, value='{v[:200]}'")
        
        if has_content:
            row_count += 1
    
    if row_count == 0:
        print("      (empty sheet - no cell data)")
    else:
        print(f"      Total non-empty rows: {row_count}")
    
    # Check for merge cells
    merge_cells = root.find(f'{{{SHEET_NS}}}mergeCells')
    if merge_cells is not None:
        print(f"\n    Merge cells:")
        for mc in merge_cells.iter(f'{{{SHEET_NS}}}mergeCell'):
            ref = mc.get('ref', '')
            print(f"      {ref}")
    
    # Check drawing reference
    drawing = root.find(f'{{{SHEET_NS}}}drawing')
    if drawing is not None:
        rid = drawing.get(f'{{{R_NS}}}id', '')
        print(f"\n    Drawing reference: rId={rid}")


def dump_comments(wb_path):
    """Dump the full comments XML for all comment files."""
    with zipfile.ZipFile(wb_path, 'r') as z:
        names = z.namelist()
        comment_files = [n for n in names if 'comments' in n.lower()]
        
        for cf in comment_files:
            content = z.read(cf).decode('utf-8', errors='replace')
            root = ET.fromstring(content)
            
            # Get authors
            authors = []
            for auth in root.iter(f'{{{SHEET_NS}}}author'):
                authors.append(auth.text or '')
            print(f"\n    Authors: {', '.join(authors)}")
            
            # Get comments with full XML
            for comment_list in root.iter(f'{{{SHEET_NS}}}commentList'):
                for comment in comment_list.iter(f'{{{SHEET_NS}}}comment'):
                    ref = comment.get('ref', '')
                    # Get the full text run
                    texts = []
                    for t in comment.iter(f'{{{SHEET_NS}}}t'):
                        if t.text:
                            texts.append(t.text)
                    full_text = '\n'.join(texts)
                    print(f"\n    Cell {ref} - FULL COMMENT:")
                    print(f"      {full_text}")


def dump_drawing(wb_path):
    """Dump the drawing XML to understand shapes/controls."""
    with zipfile.ZipFile(wb_path, 'r') as z:
        names = z.namelist()
        drawing_files = [n for n in names if 'drawing' in n.lower() and n.endswith('.xml') and 'vml' not in n.lower()]
        for df in drawing_files:
            content = z.read(df).decode('utf-8', errors='replace')
            print(f"\n    Drawing file: {df}")
            print(f"      Content:\n{content[:3000]}")
        
        # Also dump VML drawing
        vml_files = [n for n in names if 'vml' in n.lower()]
        for vf in vml_files:
            content = z.read(vf).decode('utf-8', errors='replace')
            print(f"\n    VML file: {vf}")
            print(f"      Content:\n{content[:2000]}")


# =============================================================================
print("=" * 70)
print("INVESTIGATION: How ConMas Stores Cluster Metadata")
print("=" * 70)

for label, wb_path in WORKBOOKS.items():
    if not os.path.isfile(wb_path):
        print(f"\n--- {label} ---")
        print(f"  FILE NOT FOUND: {wb_path}")
        continue
    
    print(f"\n{'='*70}")
    print(f"  {label}")
    print(f"{'='*70}")
    
    sheets = get_sheet_map(wb_path)
    
    # Find the _Fields sheet
    fields_rid = None
    for rid, info in sheets.items():
        print(f"  Sheet: {info['name']} (state={info['state']}, target={info.get('target','')})")
        if info['name'] == '_Fields':
            fields_rid = rid
    
    # Dump the _Fields sheet
    if fields_rid and sheets[fields_rid].get('target'):
        dump_sheet(wb_path, sheets[fields_rid]['target'], '_Fields')
    
    # Dump visible sheet too for comparison
    for rid, info in sheets.items():
        if info['name'] != '_Fields' and info.get('target'):
            print(f"\n  --- Visible Sheet: {info['name']} ---")
            dump_sheet(wb_path, info['target'], info['name'])
    
    # Dump comments
    print(f"\n  --- COMMENTS ---")
    dump_comments(wb_path)
    
    # Dump drawings
    print(f"\n  --- DRAWINGS ---")
    dump_drawing(wb_path)

print(f"\n{'='*70}")
print("INVESTIGATION COMPLETE")
print(f"{'='*70}")
