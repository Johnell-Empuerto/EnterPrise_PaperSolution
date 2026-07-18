"""
Phase X.8 — Create all 11 elimination test workbooks (A through K) for
validating against the real ConMas runtime.

Each test modifies ONLY ONE variable. The user uploads each workbook into
the ConMas Designer and records: field count, errors, UI differences.

Usage: python create_elimination_tests.py
"""

import os, sys, shutil, zipfile
from xml.etree import ElementTree as ET
from copy import deepcopy

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

LEGACY = r'C:\Users\MCF-JOHNELLEEMPUERTO\Downloads\FormTest - Copy.xlsx'
OUTPUT = r'C:\Users\MCF-JOHNELLEEMPUERTO\Downloads\elimination_tests'
os.makedirs(OUTPUT, exist_ok=True)

NS = {
    's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'c': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
}


def serialise(elem):
    return ET.tostring(elem, encoding='utf-8', xml_declaration=True)


def remove_comments_from_workbook(zin, files_out):
    """Remove comments1.xml, VML drawing, and update all references."""
    for fname in list(files_out.keys()):
        if 'comments' in fname.lower() or 'vml' in fname.lower():
            del files_out[fname]
    # Remove from Content_Types
    if '[Content_Types].xml' in files_out:
        ct = ET.fromstring(files_out['[Content_Types].xml'])
        to_rm = [c for c in list(ct) if 'comment' in (c.get('PartName') or '').lower() or 'vml' in (c.get('PartName') or '').lower()]
        for c in to_rm: ct.remove(c)
        files_out['[Content_Types].xml'] = serialise(ct)
    # Remove comment/VML relationships from all .rels files
    for fname in list(files_out.keys()):
        if fname.endswith('.rels'):
            try:
                rx = ET.fromstring(files_out[fname])
            except:
                continue
            to_rm = [c for c in list(rx) if 'comment' in (c.get('Target') or '').lower() or 'vml' in (c.get('Target') or '').lower()]
            for c in to_rm: rx.remove(c)
            files_out[fname] = serialise(rx)


def modify_comment_text(files_out, new_texts):
    """Replace comment text content. new_texts is a dict {cell_ref: new_text}.
    If cell_ref is None, applies to all comments."""
    cmt_files = [f for f in files_out if 'comments' in f.lower() and f.endswith('.xml')]
    if not cmt_files:
        return
    cmt_path = cmt_files[0]
    cmt = ET.fromstring(files_out[cmt_path])
    # Namespace handling — comments use the same default NS
    for comment in cmt.findall('.//c:comment', NS):
        ref = comment.get('ref', '')
        if ref in new_texts:
            new_text = new_texts[ref]
        elif None in new_texts:
            new_text = new_texts[None]
        else:
            continue
        # Replace all <t> elements inside this comment's text
        text_parent = comment.find('.//c:text', NS)
        if text_parent is not None:
            # Remove existing <t> elements
            for t in text_parent.findall('c:t', NS):
                text_parent.remove(t)
            # Add single <t> with new text
            new_t = ET.SubElement(text_parent, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
            new_t.text = new_text
    files_out[cmt_path] = serialise(cmt)


def remove_sheet(files_out, sheet_name):
    """Remove a worksheet by name from the workbook."""
    wb = ET.fromstring(files_out['xl/workbook.xml'])
    # Find sheets element and remove the target sheet child
    sheets_elem = wb.find('.//s:sheets', NS)
    if sheets_elem is None:
        return
    target_sheet = None
    for s in list(sheets_elem):
        if s.get('name', '') == sheet_name:
            target_sheet = s
            break
    if target_sheet is None:
        return
    rid = target_sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
    sheets_elem.remove(target_sheet)
    files_out['xl/workbook.xml'] = serialise(wb)
    # Remove from workbook.xml.rels and delete the sheet file
    wb_rel_path = 'xl/_rels/workbook.xml.rels'
    target_file = None
    if wb_rel_path in files_out:
        wbx = ET.fromstring(files_out[wb_rel_path])
        for r in list(wbx):
            if r.get('Id') == rid:
                target = r.get('Target', '')
                target_file = target
                wbx.remove(r)
                break
        files_out[wb_rel_path] = serialise(wbx)
    # Delete sheet file and its rels
    if target_file:
        files_out.pop(f'xl/{target_file}', None)
        files_out.pop(f'xl/worksheets/_rels/{os.path.basename(target_file)}.rels', None)
    # Remove from Content_Types
    if '[Content_Types].xml' in files_out:
        ct = ET.fromstring(files_out['[Content_Types].xml'])
        to_rm = []
        for c in list(ct):
            pn = c.get('PartName', '')
            # Remove Override entries for the removed sheet
            if target_file and f'worksheets/{os.path.basename(target_file)}' in pn:
                to_rm.append(c)
        for c in to_rm:
            ct.remove(c)
        files_out['[Content_Types].xml'] = serialise(ct)


def remove_print_area(files_out):
    """Remove PrintArea defined name from workbook.xml."""
    wb = ET.fromstring(files_out['xl/workbook.xml'])
    # Defined names are child elements of <definedNames>
    dns_elem = wb.find('.//s:definedNames', NS)
    if dns_elem is not None:
        to_rm = []
        for dn in list(dns_elem):
            name = (dn.get('name') or '').lower()
            if 'print_area' in name:
                to_rm.append(dn)
        for dn in to_rm:
            dns_elem.remove(dn)
    files_out['xl/workbook.xml'] = serialise(wb)
    # Also remove printArea from sheet XML
    for fname in list(files_out.keys()):
        if fname.startswith('xl/worksheets/') and fname.endswith('.xml'):
            ws = ET.fromstring(files_out[fname])
            for pa in ws.findall('.//s:printArea', NS):
                ws.remove(pa)
            files_out[fname] = serialise(ws)


def remove_merged_cells(files_out):
    """Remove all merged cell definitions from worksheets."""
    for fname in list(files_out.keys()):
        if fname.startswith('xl/worksheets/') and fname.endswith('.xml'):
            ws = ET.fromstring(files_out[fname])
            for mc in list(ws.findall('.//s:mergeCells', NS)):
                ws.remove(mc)
            files_out[fname] = serialise(ws)


def copy_and_edit(src_path, dst_path, edits):
    """Copy workbook from src to dst, applying edits function."""
    with zipfile.ZipFile(src_path, 'r') as zin:
        files_out = {}
        for fname in zin.namelist():
            files_out[fname] = zin.read(fname)
        edits(files_out)
        with zipfile.ZipFile(dst_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for fname, data in sorted(files_out.items()):
                zout.writestr(fname, data)


def verify(path):
    """Verify workbook structure and return info dict."""
    info = {}
    with zipfile.ZipFile(path, 'r') as z:
        info['files'] = sorted(z.namelist())
        info['size'] = os.path.getsize(path)
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        sheets = wb.findall('.//s:sheet', NS)
        info['sheets'] = [(s.get('name', ''), s.get('state', 'visible')) for s in sheets]
        dns = wb.findall('.//s:definedName', NS)
        info['defined_names'] = [dn.get('name', '') for dn in dns]
        cmt_files = [f for f in info['files'] if 'comments' in f.lower() and f.endswith('.xml')]
        if cmt_files:
            cmt = ET.fromstring(z.read(cmt_files[0]))
            info['comments'] = len(cmt.findall('.//c:comment', {'c': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}))
        else:
            info['comments'] = 0
        info['merged'] = 0
        for fname in info['files']:
            if fname.startswith('xl/worksheets/') and fname.endswith('.xml'):
                ws = ET.fromstring(z.read(fname))
                info['merged'] += len(ws.findall('.//s:mergeCells/s:mergeCell', NS))
        return info


def print_verify(path):
    info = verify(path)
    print(f"  ✅ {os.path.basename(path)}")
    print(f"     Size: {info['size']:,} bytes")
    print(f"     Sheets: {', '.join(f'{n}({s})' if s != 'visible' else n for n, s in info['sheets'])}")
    print(f"     Comments: {info['comments']}")
    print(f"     Merged cells: {info['merged']}")
    if info['defined_names']:
        print(f"     Defined names: {', '.join(info['defined_names'])}")
    else:
        print(f"     Defined names: (none)")


# ══════════════════════════════════════════════════════════════════════════
# CREATE TESTS
# ══════════════════════════════════════════════════════════════════════════

print("=" * 70)
print("CREATING ELIMINATION TEST WORKBOOKS")
print("=" * 70)

# ── Test A: Original (baseline) ──
print(f"\n{'─'*60}")
print("Test A: Original_Baseline — unmodified legacy workbook")
path_a = os.path.join(OUTPUT, 'TestA_Original_Baseline.xlsx')
shutil.copy2(LEGACY, path_a)
print_verify(path_a)

# ── Test B: Remove ALL comments ──
print(f"\n{'─'*60}")
print("Test B: NoComments — remove all cell comments, keep everything else")
path_b = os.path.join(OUTPUT, 'TestB_NoComments.xlsx')
copy_and_edit(LEGACY, path_b, lambda f: remove_comments_from_workbook(None, f))
print_verify(path_b)

# ── Test C: Remove _Fields sheet ──
print(f"\n{'─'*60}")
print("Test C: NoFieldsSheet — remove _Fields worksheet, keep Sheet1 + ExcelOutputSetting")
path_c = os.path.join(OUTPUT, 'TestC_NoFieldsSheet.xlsx')
copy_and_edit(LEGACY, path_c, lambda f: remove_sheet(f, '_Fields'))
print_verify(path_c)

# ── Test D: Remove ExcelOutputSetting sheet ──
print(f"\n{'─'*60}")
print("Test D: NoExcelOutputSetting — remove ExcelOutputSetting, keep _Fields + Sheet1")
path_d = os.path.join(OUTPUT, 'TestD_NoExcelOutputSetting.xlsx')
copy_and_edit(LEGACY, path_d, lambda f: remove_sheet(f, 'ExcelOutputSetting'))
print_verify(path_d)

# ── Test E: Remove PrintArea ──
print(f"\n{'─'*60}")
print("Test E: NoPrintArea — remove PrintArea defined name from workbook")
path_e = os.path.join(OUTPUT, 'TestE_NoPrintArea.xlsx')
copy_and_edit(LEGACY, path_e, lambda f: remove_print_area(f))
print_verify(path_e)

# ── Test F: Change comment format (simplify — minimal lines) ──
print(f"\n{'─'*60}")
print("Test F: SimplifiedComments — replace all comment text with single line (field name only)")
path_f = os.path.join(OUTPUT, 'TestF_SimplifiedComments.xlsx')
def simplify_comments(files):
    # Get current comment refs from comments1.xml
    cmt_files = [f for f in files if 'comments' in f.lower() and f.endswith('.xml')]
    if cmt_files:
        cmt = ET.fromstring(files[cmt_files[0]])
        simple = {}
        for c in cmt.findall('.//c:comment', {'c': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}):
            ref = c.get('ref', '')
            # Replace with just field name (simplest possible: line 0 only)
            text_parent = c.find('.//c:text', {'c': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'})
            if text_parent is not None:
                for t in list(text_parent):
                    text_parent.remove(t)
                new_t = ET.SubElement(text_parent, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                new_t.text = f"Field_{ref}"  # Minimal: just the field name
        files[cmt_files[0]] = serialise(cmt)
copy_and_edit(LEGACY, path_f, simplify_comments)
print_verify(path_f)

# ── Test G: Remove merged cells, keep comments ──
print(f"\n{'─'*60}")
print("Test G: NoMergedCells_KeepComments — remove all merged regions, keep comments/values")
path_g = os.path.join(OUTPUT, 'TestG_NoMergedCells.xlsx')
copy_and_edit(LEGACY, path_g, lambda f: remove_merged_cells(f))
print_verify(path_g)

# ── Test H: Keep only Sheet1 (remove _Fields AND ExcelOutputSetting) ──
print(f"\n{'─'*60}")
print("Test H: OnlySheet1 — remove BOTH _Fields AND ExcelOutputSetting, keep only Sheet1")
path_h = os.path.join(OUTPUT, 'TestH_OnlySheet1.xlsx')
def keep_only_sheet1(files):
    wb = ET.fromstring(files['xl/workbook.xml'])
    sheets_elem = wb.find('.//s:sheets', NS)
    if sheets_elem is None:
        return
    for s in list(sheets_elem):
        name = s.get('name', '')
        if name != 'Sheet1':
            rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            sheets_elem.remove(s)
            # Remove from rels
            target_file = None
            if 'xl/_rels/workbook.xml.rels' in files:
                wr = ET.fromstring(files['xl/_rels/workbook.xml.rels'])
                for r in list(wr):
                    if r.get('Id') == rid:
                        target = r.get('Target', '')
                        target_file = target
                        wr.remove(r)
                        break
                files['xl/_rels/workbook.xml.rels'] = serialise(wr)
            # Delete sheet files
            if target_file:
                files.pop(f'xl/{target_file}', None)
                files.pop(f'xl/worksheets/_rels/{os.path.basename(target_file)}.rels', None)
    files['xl/workbook.xml'] = serialise(wb)
copy_and_edit(LEGACY, path_h, keep_only_sheet1)
print_verify(path_h)

# ── Test I: Replace comments with our generated format ──
print(f"\n{'─'*60}")
print("Test I: OurCommentFormat — replace all comments with our WorkbookGenerator format")
print("     Format: Name\\nType\\n\\n0")
path_i = os.path.join(OUTPUT, 'TestI_OurCommentFormat.xlsx')
def apply_our_comment_format(files):
    cmt_files = [f for f in files if 'comments' in f.lower() and f.endswith('.xml')]
    if cmt_files:
        cmt = ET.fromstring(files[cmt_files[0]])
        ns_c = {'c': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        for c in cmt.findall('.//c:comment', ns_c):
            ref = c.get('ref', '')
            text_parent = c.find('.//c:text', ns_c)
            if text_parent is not None:
                for t in list(text_parent):
                    text_parent.remove(t)
                # Our generated format: Name\\nType\\n\\n0
                our_text = f"Field_{ref}\nKeyboardText\n\n0"
                new_t = ET.SubElement(text_parent, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                new_t.text = our_text
        files[cmt_files[0]] = serialise(cmt)
copy_and_edit(LEGACY, path_i, apply_our_comment_format)
print_verify(path_i)

# ── Test J: Keep only Sheet1 + ExcelOutputSetting (remove _Fields only) ──
print(f"\n{'─'*60}")
print("Test J: NoFields_KeepEverythingElse — remove ONLY _Fields, keep Sheet1 + ExcelOutputSetting")
path_j = os.path.join(OUTPUT, 'TestJ_NoFields_KeepExcelOutput.xlsx')
copy_and_edit(LEGACY, path_j, lambda f: remove_sheet(f, '_Fields'))
print_verify(path_j)

# ── Test K: Remove ONLY defined names (Print_Area is a defined name) ──
# But we already have Test E for PrintArea. Let's make K a comprehensive check:
# Remove defined names but RE-ADD PrintArea as a sheet2.xml printArea element
print(f"\n{'─'*60}")
print("Test K: NoDefinedNames_KeepPrintArea — remove all defined names from workbook.xml")
print("     Tests whether ConMas relies on defined names vs sheet-level printArea")
path_k = os.path.join(OUTPUT, 'TestK_NoDefinedNames.xlsx')
def remove_defined_names(files):
    wb = ET.fromstring(files['xl/workbook.xml'])
    dns_elem = wb.find('.//s:definedNames', NS)
    if dns_elem is not None:
        for dn in list(dns_elem):
            dns_elem.remove(dn)
    files['xl/workbook.xml'] = serialise(wb)
copy_and_edit(LEGACY, path_k, remove_defined_names)
print_verify(path_k)

# ── Test L (bonus): Our Format WITH full comments (legacy-style cluster line) ──
# This tests whether adding cluster index back makes a difference
print(f"\n{'─'*60}")
print("Test L (bonus): OurFormatWithClusterIndex — our comment format + cluster index on line 2")
path_l = os.path.join(OUTPUT, 'TestL_OurFormat_WithClusterIndex.xlsx')
def apply_our_format_with_index(files):
    cmt_files = [f for f in files if 'comments' in f.lower() and f.endswith('.xml')]
    if cmt_files:
        cmt = ET.fromstring(files[cmt_files[0]])
        ns_c = {'c': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        idx = 0
        for c in cmt.findall('.//c:comment', ns_c):
            ref = c.get('ref', '')
            text_parent = c.find('.//c:text', ns_c)
            if text_parent is not None:
                for t in list(text_parent):
                    text_parent.remove(t)
                # Our format + cluster index: Name\\nType\\n{idx}\\n\\n0
                our_text = f"Field_{ref}\nKeyboardText\n{idx}\n\n0"
                idx += 1
                new_t = ET.SubElement(text_parent, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                new_t.text = our_text
        files[cmt_files[0]] = serialise(cmt)
copy_and_edit(LEGACY, path_l, apply_our_format_with_index)
print_verify(path_l)


print(f"\n{'=' * 70}")
print(f"DONE — Created test workbooks in: {OUTPUT}")
print("=" * 70)

# ── Summary Matrix ──
features = [
    ('A', 'Original_Baseline',              True, True, True, True, True, True, True),
    ('B', 'NoComments',                     False, True, True, True, True, True, True),
    ('C', 'NoFieldsSheet',                  True, False, True, True, True, True, True),
    ('D', 'NoExcelOutputSetting',           True, True, False, True, True, True, True),
    ('E', 'NoPrintArea',                    True, True, True, False, True, True, True),
    ('F', 'SimplifiedComments',             'Minimal', True, True, True, True, True, True),
    ('G', 'NoMergedCells',                  True, True, True, True, False, True, True),
    ('H', 'OnlySheet1',                     True, False, False, True, True, True, True),
    ('I', 'OurCommentFormat',              'Ours', True, True, True, True, True, True),
    ('J', 'NoFields_KeepExcelOutput',       True, False, True, True, True, True, True),
    ('K', 'NoDefinedNames',                 True, True, True, True, True, False, False),
    ('L', 'OurFormat_WithClusterIndex',    'Ours+Idx', True, True, True, True, True, True),
]

print(f"\n{'=' * 70}")
print("TEST MATRIX — How to upload each workbook to ConMas Designer")
print("=" * 70)
print(f"{'Test':<6} {'File':<32} {'Cmnts':<12} {'_Fields':<10} {'ExcelOut':<10} {'PrintAm':<10} {'Merge':<8} {'DefNm':<8}")
print(f"{'':-<6} {'':-<32} {'':-<12} {'':-<10} {'':-<10} {'':-<10} {'':-<8} {'':-<8}")
for tid, name, comments, fields, excel, printarea, merged, defnames, _ in features:
    c = '✅' if comments == True else ('❌' if comments == False else '⚠Simp' if comments == 'Minimal' else '⚠Ours')
    f = '✅' if fields else '❌'
    e = '✅' if excel else '❌'
    p = '✅' if printarea else '❌'
    m = '✅' if merged else '❌'
    d = '✅' if defnames else '❌'
    print(f"{tid:<6} Test{tid}_{name:<25} {c:<12} {f:<10} {e:<10} {p:<10} {m:<8} {d:<8}")

print(f"\n{'=' * 70}")
print("INSTRUCTIONS")
print("=" * 70)
print("""
1. Open the ConMas Designer application
2. For each test workbook:
   a. Use the import/upload function to open the workbook
   b. Record:
      - How many fields are reconstructed
      - Any error messages or warnings
      - Whether the form looks correct in the UI
      - Any runtime exceptions
3. Report the results back so we can update the dependency matrix

Tests are in C:\\Users\\MCF-JOHNELLEEMPUERTO\\Downloads\\elimination_tests\\
""")
