"""
Phase X.30 — Behavioral Parity Forensic Investigation
Compares Original, ConMas, and Our generated workbooks at every OOXML structural level.
Ranks each difference by likelihood of affecting Excel's print engine.
Identifies the single most likely remaining cause of behavioral mismatch.

Usage:
    python phase30_forensic_diff.py <original.xlsx> <conmas.xlsx> <our_output.xlsx>
"""

import os, sys, zipfile, json, hashlib
from xml.etree import ElementTree as ET
from collections import defaultdict
from difflib import unified_diff

sys.stdout.reconfigure(encoding='utf-8')

NS = {
    's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
}

# ── Constants ────────────────────────────────────────────────────────────────

KNOWN_IRRELEVANT = {
    'docProps/core.xml': 'Document metadata (creator, timestamps) does not affect Excel print engine.',
    'docProps/app.xml': 'Application metadata (app name, version) does not affect Excel print engine.',
    'xl/sharedStrings.xml': 'String values do not affect print geometry unless they change cell content layout.',
    'xl/styles.xml': 'Fonts/colors do not affect print geometry unless they affect column widths/row heights.',
    'xl/theme/': 'Theme colors do not affect print geometry.',
}

PRINT_RELEVANT_SCORE = {
    'pageMargins': 10,      # DIRECTLY affects print position
    'pageSetup': 10,        # DIRECTLY affects print behavior
    'printOptions': 10,     # DIRECTLY affects centering
    'definedName': 9,       # Print_Area, Print_Titles DIRECTLY affect print range
    'sheetPr': 8,           # pageSetUpPr, tabColor affect print behavior
    'sheetView': 6,         # showFormulas, showGridLines affect print
    'rowBreaks': 8,         # Page breaks DIRECTLY affect print pagination
    'colBreaks': 8,         # Page breaks DIRECTLY affect print pagination
    'mergeCells': 7,        # Affects cell layout and print area calculation
    'cols': 7,              # Column widths DIRECTLY affect print layout
    'row': 7,               # Row heights DIRECTLY affect print layout
    'sheetFormatPr': 5,     # Default row height/column width
    'extLst': 6,            # May contain print extensions
    'printerSettings': 9,   # Printer config DIRECTLY affects margins/rendering
    'vmlDrawing': 3,        # VML shapes for comments - don't affect print of content
    'legacyDrawing': 3,     # Legacy drawing - doesn't affect print
    'drawing': 4,           # Drawing - may affect content layout
    'oleObjects': 5,        # OLE objects may affect layout
    'printTitles': 9,       # Print_Titles DIRECTLY affects print layout
    'workbookProtection': 2, # Protection doesn't affect print geometry
    'ignoredError': 1,      # Does not affect print
    'selection': 1,         # Does not affect print
    'activeCell': 1,        # Does not affect print
    'tabSelected': 1,       # Does not affect print
    'calcPr': 3,            # Calculation properties - only affects values
    'vbaProject': 5,        # VBA may affect print via macros
    'pageSetUpPr': 9,       # fitToPage DIRECTLY affects print scaling
    'AlternateContent': 4,  # Compatibility - may affect rendering
    'sheetPr': 8,           # Sheet properties affect print
}


# ── Core Helpers ─────────────────────────────────────────────────────────────

def load_entries(path):
    with zipfile.ZipFile(path, 'r') as z:
        return {n: z.read(n) for n in z.namelist()}


def xml_elem(data):
    try:
        return ET.fromstring(data)
    except:
        return None


def local_tag(elem):
    if elem is None: return ''
    t = elem.tag
    return t.split('}')[-1] if '}' in t else t


def get_attr(elem, key, ns=None):
    """Get attribute with or without namespace."""
    if elem is None: return None
    # Try direct
    v = elem.get(key)
    if v is not None: return v
    # Try with namespace
    if ns:
        v = elem.get(f'{{{ns}}}{key}')
        if v is not None: return v
    return None


def elem_path(elem, parent_path=''):
    """Get path of element."""
    tag = local_tag(elem)
    path = f"{parent_path}/{tag}" if parent_path else tag
    return path


def all_attributes(elem, prefix=''):
    """Recursively get all attribute key-value pairs."""
    attrs = {}
    tag = local_tag(elem)
    for k, v in elem.attrib.items():
        lk = k.split('}')[-1] if '}' in k else k
        attrs[f"{prefix}{tag}.@{lk}"] = v
    for child in elem:
        attrs.update(all_attributes(child, f"{prefix}{tag}/"))
    return attrs


def get_entry_hash(data):
    return hashlib.md5(data).hexdigest()[:12]


def get_value_safe(elem, xpath):
    """Get text of first matching element, or None."""
    if elem is None: return None
    try:
        e = elem.find(xpath, NS)
        return e.text if e is not None else None
    except:
        return None


# ── Print Relevance Scoring ──────────────────────────────────────────────────

def score_difference(name, diff_detail):
    """Score a difference by likelihood of affecting Excel's print engine. 1-10."""
    # Check for known irrelevant
    for prefix, reason in KNOWN_IRRELEVANT.items():
        if name.startswith(prefix):
            return 0, reason
    
    name_lower = name.lower()
    diff_lower = diff_detail.lower()
    
    # High impact (9-10)
    if 'pageMargins' in diff_lower or 'pageMargins' in name_lower:
        return 10, "Margins DIRECTLY control print position on page."
    if 'pageSetup' in diff_lower or 'pageSetup' in name_lower:
        return 10, "Page setup DIRECTLY controls orientation, scaling, paper size."
    if 'printOptions' in diff_lower or 'printOptions' in name_lower:
        return 10, "Print options DIRECTLY control centering (horizontal/vertical)."
    if 'Print_Area' in diff_lower or 'Print_Titles' in diff_lower:
        return 9, "Print area/titles DIRECTLY control which cells are printed."
    if 'printerSettings' in name_lower:
        return 9, "Printer settings DIRECTLY affect margin interpretation and rendering."
    
    # Medium-high impact (7-8)
    if 'sheetPr' in diff_lower or 'pageSetUpPr' in diff_lower:
        return 8, "Sheet properties (fitToPage) affect print scaling."
    if 'rowBreaks' in diff_lower or 'colBreaks' in diff_lower:
        return 8, "Page breaks DIRECTLY control pagination."
    if 'mergeCells' in diff_lower or 'mergeCell' in diff_lower:
        return 7, "Merge cells affect cell layout and print area."
    if 'cols' in name_lower and 'width' in diff_lower:
        return 7, "Column widths DIRECTLY control horizontal print layout."
    if 'row' in name_lower and 'height' in diff_lower:
        return 7, "Row heights DIRECTLY control vertical print layout."
    if 'definedName' in diff_lower and ('print' in diff_lower):
        return 9, "Print-related defined names DIRECTLY affect print range."
    
    # Medium impact (5-6)
    if 'extLst' in name_lower:
        return 6, "Extension list may contain print-related extensions."
    if 'AlternateContent' in diff_lower:
        return 4, "Compatibility markup may affect how Excel interprets the file."
    if 'vbaProject' in name_lower:
        return 5, "VBA project may contain print macros."
    if 'oleObjects' in name_lower:
        return 5, "OLE objects may affect content layout and print area."
    if 'sheetFormatPr' in name_lower:
        return 5, "Default row height/column width affects overall layout."
    
    # Low impact (1-3)
    if 'comment' in name_lower:
        return 2, "Comments do not appear in print unless explicitly set."
    if 'vmlDrawing' in name_lower:
        return 2, "VML drawings (comment shapes) do not affect print of content."
    if 'styles' in name_lower:
        return 1, "Font/color differences do not affect print geometry (unless they affect width/height)."
    if 'sharedStrings' in name_lower:
        return 1, "String values only affect layout if they change cell content width."
    if 'selection' in diff_lower or 'activeCell' in diff_lower or 'tabSelected' in diff_lower:
        return 1, "UI state selections do not affect print."
    
    return 3, "Unknown impact - defaulting to low-medium."


# ── Structural Comparison Functions ──────────────────────────────────────────

def compare_entries(names1, entries1, entries2, label1, label2):
    """Compare ZIP entries between two workbooks."""
    differences = []
    
    for name in sorted(names1):
        in1 = name in entries1
        in2 = name in entries2
        
        if not in2:
            score, reason = score_difference(name, f"ENTRY ONLY IN {label1}")
            differences.append({
                'entry': name,
                'type': 'entry_only_in_1',
                'detail': f"Entry present in {label1} but MISSING in {label2}",
                'size_1': len(entries1[name]),
                'size_2': 0,
                'score': score,
                'reason': reason,
            })
            continue
        
        data1 = entries1[name]
        data2 = entries2[name]
        
        # Compare sizes
        if len(data1) != len(data2):
            hash1 = get_entry_hash(data1)
            hash2 = get_entry_hash(data2)
            
            if hash1 != hash2:
                score, reason = score_difference(name, f"SIZE/MD5 DIFFERENCE")
                differences.append({
                    'entry': name,
                    'type': 'content_diff',
                    'detail': f"Content differs: {len(data1)} vs {len(data2)} bytes (MD5: {hash1} vs {hash2})",
                    'size_1': len(data1),
                    'size_2': len(data2),
                    'score': score,
                    'reason': reason,
                })
    
    # Check entries only in workbook 2
    for name in sorted(entries2):
        if name not in entries1:
            score, reason = score_difference(name, f"ENTRY ONLY IN {label2}")
            differences.append({
                'entry': name,
                'type': 'entry_only_in_2',
                'detail': f"Entry MISSING in {label1} but present in {label2}",
                'size_1': 0,
                'size_2': len(entries2[name]),
                'score': score,
                'reason': reason,
            })
    
    return differences


def compare_xml_structures(elem1, elem2, path='', depth=0, max_depth=15):
    """Deep structural comparison of two XML elements. Returns list of diffs."""
    diffs = []
    
    if elem1 is None and elem2 is None:
        return diffs
    if elem1 is None:
        return [(path, 'ELEMENT_MISSING_IN_1', 'Element missing in workbook 1')]
    if elem2 is None:
        return [(path, 'ELEMENT_MISSING_IN_2', 'Element missing in workbook 2')]
    
    tag1 = local_tag(elem1)
    tag2 = local_tag(elem2)
    
    if tag1 != tag2:
        return [(path, 'TAG_MISMATCH', f"<{tag1}> vs <{tag2}>")]
    
    if depth > max_depth:
        return diffs
    
    # Compare attributes
    attrs1 = dict(elem1.attrib)
    attrs2 = dict(elem2.attrib)
    
    # Normalize namespace keys
    attrs1_norm = {k.split('}')[-1] if '}' in k else k: v for k, v in attrs1.items()}
    attrs2_norm = {k.split('}')[-1] if '}' in k else k: v for k, v in attrs2.items()}
    
    for k in sorted(set(list(attrs1_norm.keys()) + list(attrs2_norm.keys()))):
        v1 = attrs1_norm.get(k)
        v2 = attrs2_norm.get(k)
        if v1 is None:
            diffs.append((f"{path}@{k}", 'ATTR_MISSING_IN_1', f"@{k}={v2} missing in workbook 1"))
        elif v2 is None:
            diffs.append((f"{path}@{k}", 'ATTR_MISSING_IN_2', f"@{k}={v1} missing in workbook 2"))
        elif v1 != v2:
            # For print-relevant attrs, show the diff
            diffs.append((f"{path}@{k}", 'ATTR_VALUE', f"'{v1}' vs '{v2}'"))
    
    # Compare text content
    text1 = (elem1.text or '').strip()
    text2 = (elem2.text or '').strip()
    if text1 != text2:
        if tag1 in ('definedName', 'sheet', 'mergeCell'):
            diffs.append((f"{path}", 'TEXT_VALUE', f"'{text1}' vs '{text2}'"))
    
    # Compare children by tag
    children1 = list(elem1)
    children2 = list(elem2)
    
    by_tag1 = defaultdict(list)
    by_tag2 = defaultdict(list)
    for c in children1:
        by_tag1[local_tag(c)].append(c)
    for c in children2:
        by_tag2[local_tag(c)].append(c)
    
    all_tags = set(list(by_tag1.keys()) + list(by_tag2.keys()))
    for tag in sorted(all_tags):
        list1 = by_tag1.get(tag, [])
        list2 = by_tag2.get(tag, [])
        
        if len(list1) != len(list2):
            diffs.append((f"{path}/{tag}", 'COUNT', f"{len(list1)} vs {len(list2)}"))
        
        # Compare element by element (up to limit)
        for i in range(min(len(list1), len(list2))):
            sub_diffs = compare_xml_structures(list1[i], list2[i], f"{path}/{tag}[{i}]", depth + 1, max_depth)
            diffs.extend(sub_diffs)
    
    return diffs


def analyze_workbook_structure(entries, label):
    """Extract all print-relevant structures from a workbook."""
    info = {
        'label': label,
        'sheets': [],
        'defined_names': [],
        'print_relevant': {},
        'all_elements': {},
        'binary_hashes': {},
        'extensions': [],
        'has_vba': 'xl/vbaProject.bin' in entries,
        'has_printer_settings': any('printerSettings' in n for n in entries),
    }
    
    # Analyze workbook.xml
    wb = xml_elem(entries.get('xl/workbook.xml', b''))
    if wb is not None:
        # Sheets
        for s in wb.findall('.//s:sheet', NS):
            info['sheets'].append({
                'name': s.get('name', '?'),
                'id': s.get('sheetId', '?'),
                'state': s.get('state', 'visible'),
                'rid': get_attr(s, 'id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
            })
        
        # Defined names
        for dn in wb.findall('.//s:definedName', NS):
            info['defined_names'].append({
                'name': dn.get('name', '?'),
                'localSheetId': dn.get('localSheetId', ''),
                'text': (dn.text or '').strip(),
            })
    
    # Analyze each sheet
    for name, data in entries.items():
        if not name.endswith('.xml') and not name.endswith('.rels'):
            continue
        
        elem = xml_elem(data)
        if elem is None:
            continue
        
        # Collect all element counts
        counts = defaultdict(int)
        for e in elem.iter():
            counts[local_tag(e)] += 1
        
        info['all_elements'][name] = dict(counts)
        
        # Extract print-relevant
        pr = {}
        for tag in ['pageMargins', 'pageSetup', 'printOptions', 'sheetPr', 
                     'sheetView', 'mergeCells', 'rowBreaks', 'colBreaks',
                     'row', 'col', 'sheetFormatPr', 'extLst',
                     'ignoredError', 'oleObjects', 'selection',
                     'AlternateContent', 'legacyDrawing', 'drawing',
                     'pageSetUpPr', 'customSheetView']:
            if tag in counts:
                pr[tag] = counts[tag]
        
        if pr:
            info['print_relevant'][name] = pr
    
    # Hash binary files
    for name, data in entries.items():
        if name.endswith('.bin') or name.endswith('.vml') or 'printerSettings' in name:
            info['binary_hashes'][name] = {
                'size': len(data),
                'md5': hashlib.md5(data).hexdigest()[:12],
            }
    
    return info


def compare_workbooks_structure(s1, s2, label1, label2):
    """Compare two workbook structure analyses, return scored differences."""
    differences = []
    
    # Compare sheet definitions
    sheets1 = {s['name']: s for s in s1['sheets']}
    sheets2 = {s['name']: s for s in s2['sheets']}
    
    all_sheet_names = set(list(sheets1.keys()) + list(sheets2.keys()))
    for name in sorted(all_sheet_names):
        sh1 = sheets1.get(name)
        sh2 = sheets2.get(name)
        if sh1 and not sh2:
            differences.append({
                'entry': f'workbook/sheet/{name}',
                'type': 'sheet_only_in_1',
                'detail': f"Sheet '{name}' present in {label1} but MISSING in {label2}",
                'score': 7,
                'reason': "Missing sheet can affect print layout.",
            })
        elif sh2 and not sh1:
            differences.append({
                'entry': f'workbook/sheet/{name}',
                'type': 'sheet_only_in_2',
                'detail': f"Sheet '{name}' MISSING in {label1} but present in {label2}",
                'score': 7,
                'reason': "Extra sheet can affect print layout.",
            })
        elif sh1 and sh2:
            if sh1.get('state') != sh2.get('state'):
                differences.append({
                    'entry': f'workbook/sheet/{name}/state',
                    'type': 'sheet_state',
                    'detail': f"Sheet '{name}' state: '{sh1['state']}' vs '{sh2['state']}'",
                    'score': 6,
                    'reason': "Hidden vs visible sheet state can affect print behavior.",
                })
    
    # Compare defined names
    dn1 = {d['name']: d for d in s1['defined_names']}
    dn2 = {d['name']: d for d in s2['defined_names']}
    
    all_dn = set(list(dn1.keys()) + list(dn2.keys()))
    for name in sorted(all_dn):
        d1 = dn1.get(name)
        d2 = dn2.get(name)
        if d1 and not d2:
            is_print = 'print' in name.lower()
            differences.append({
                'entry': f'definedName/{name}',
                'type': 'dn_only_in_1',
                'detail': f"Defined name '{name}'={d1['text']} present in {label1} but MISSING in {label2}",
                'score': 9 if is_print else 4,
                'reason': "Print-related defined names DIRECTLY affect print range." if is_print else "Non-print defined name.",
            })
        elif d2 and not d1:
            is_print = 'print' in name.lower()
            differences.append({
                'entry': f'definedName/{name}',
                'type': 'dn_only_in_2',
                'detail': f"Defined name '{name}'={d2['text']} MISSING in {label1} but present in {label2}",
                'score': 9 if is_print else 4,
                'reason': "Print-related defined names DIRECTLY affect print range." if is_print else "Non-print defined name.",
            })
        elif d1 and d2:
            if d1['text'] != d2['text']:
                is_print = 'print' in name.lower()
                differences.append({
                    'entry': f'definedName/{name}',
                    'type': 'dn_value',
                    'detail': f"Defined name '{name}': '{d1['text']}' vs '{d2['text']}'",
                    'score': 9 if is_print else 4,
                    'reason': "Print-related defined names DIRECTLY affect print range." if is_print else "Non-print defined name.",
                })
    
    # Compare binary hashes
    all_bin = set(list(s1['binary_hashes'].keys()) + list(s2['binary_hashes'].keys()))
    for name in sorted(all_bin):
        b1 = s1['binary_hashes'].get(name)
        b2 = s2['binary_hashes'].get(name)
        if b1 and not b2:
            is_printer = 'printer' in name.lower()
            differences.append({
                'entry': name,
                'type': 'binary_only_in_1',
                'detail': f"Binary file present in {label1} but MISSING in {label2} ({b1['size']} bytes, MD5: {b1['md5']})",
                'score': 9 if is_printer else 3,
                'reason': "Printer settings DIRECTLY affect margin interpretation" if is_printer else "Binary file.",
            })
        elif b2 and not b1:
            is_printer = 'printer' in name.lower()
            differences.append({
                'entry': name,
                'type': 'binary_only_in_2',
                'detail': f"Binary file MISSING in {label1} but present in {label2} ({b2['size']} bytes, MD5: {b2['md5']})",
                'score': 9 if is_printer else 3,
                'reason': "Printer settings DIRECTLY affect margin interpretation" if is_printer else "Binary file.",
            })
        elif b1 and b2:
            if b1['md5'] != b2['md5']:
                is_printer = 'printer' in name.lower()
                differences.append({
                    'entry': name,
                    'type': 'binary_diff',
                    'detail': f"Binary content differs: {b1['size']}b (MD5:{b1['md5']}) vs {b2['size']}b (MD5:{b2['md5']})",
                    'score': 9 if is_printer else 3,
                    'reason': "Printer settings DIRECTLY affect margin interpretation" if is_printer else "Binary file.",
                })
    
    # Compare print-relevant elements
    all_pr_files = set(list(s1['print_relevant'].keys()) + list(s2['print_relevant'].keys()))
    for name in sorted(all_pr_files):
        pr1 = s1['print_relevant'].get(name, {})
        pr2 = s2['print_relevant'].get(name, {})
        
        all_pr_tags = set(list(pr1.keys()) + list(pr2.keys()))
        for tag in sorted(all_pr_tags):
            c1 = pr1.get(tag, 0)
            c2 = pr2.get(tag, 0)
            if c1 != c2:
                score = PRINT_RELEVANT_SCORE.get(tag, 3)
                differences.append({
                    'entry': f"{name}/{tag}",
                    'type': 'print_element_count',
                    'detail': f"<{tag}> x{c1} vs x{c2} in {name}",
                    'score': score,
                    'reason': f"{tag} {'DIRECTLY' if score >= 8 else ''} affects print behavior.",
                })
    
    # Compare VBA
    if s1['has_vba'] != s2['has_vba']:
        differences.append({
            'entry': 'xl/vbaProject.bin',
            'type': 'vba_presence',
            'detail': f"VBA project: {'present' if s1['has_vba'] else 'MISSING'} in {label1} vs {'present' if s2['has_vba'] else 'MISSING'} in {label2}",
            'score': 5,
            'reason': "VBA may contain print macros that affect behavior.",
        })
    
    return differences


def deep_diff_xml_entries(entries1, entries2, label1, label2):
    """Deep XML structure comparison for all XML entries."""
    differences = []
    
    xml_names = sorted(set(
        [n for n in entries1 if n.endswith('.xml') or n.endswith('.rels')] +
        [n for n in entries2 if n.endswith('.xml') or n.endswith('.rels')]
    ))
    
    for name in xml_names:
        if name not in entries1 or name not in entries2:
            continue
        
        elem1 = xml_elem(entries1[name])
        elem2 = xml_elem(entries2[name])
        
        if elem1 is None or elem2 is None:
            continue
        
        struct_diffs = compare_xml_structures(elem1, elem2, name)
        for path, dtype, detail in struct_diffs:
            score, reason = score_difference(name, f"{dtype}: {detail}")
            differences.append({
                'entry': name,
                'type': f'xml_{dtype}',
                'detail': f"[{path}] {detail}",
                'score': score,
                'reason': reason,
            })
    
    return differences


# ── Report Generation ────────────────────────────────────────────────────────

def generate_report(path_orig, path_conmas, path_ours):
    """Generate full Phase X.30 forensic report."""
    
    # Load all three workbooks
    e_orig = load_entries(path_orig)
    e_conmas = load_entries(path_conmas)
    e_ours = load_entries(path_ours)
    
    # Structural analysis
    s_orig = analyze_workbook_structure(e_orig, "Original")
    s_conmas = analyze_workbook_structure(e_conmas, "ConMas")
    s_ours = analyze_workbook_structure(e_ours, "Our output")
    
    # ========== COMPARISONS ==========
    
    # Comparison 1: Original vs ConMas
    print("\n" + "=" * 80)
    print("COMPARISON 1: ORIGINAL vs CONMAS")
    print("=" * 80)
    
    diffs_orig_vs_conmas = []
    
    # ZIP entry comparison
    print("\n--- ZIP Entry Differences ---")
    entry_diffs = compare_entries(
        set(list(e_orig.keys()) + list(e_conmas.keys())),
        e_orig, e_conmas, "Original", "ConMas"
    )
    for d in sorted(entry_diffs, key=lambda x: -x['score']):
        if d['score'] >= 5:
            print(f"  [{d['score']}/10] {d['entry']}: {d['detail']}")
            print(f"    Reason: {d['reason']}")
            diffs_orig_vs_conmas.append(d)
    
    # Structural comparison
    print("\n--- Structural Differences ---")
    struct_diffs = compare_workbooks_structure(s_orig, s_conmas, "Original", "ConMas")
    for d in sorted(struct_diffs, key=lambda x: -x['score']):
        if d['score'] >= 5:
            print(f"  [{d['score']}/10] {d['entry']}: {d['detail']}")
            print(f"    Reason: {d['reason']}")
            diffs_orig_vs_conmas.append(d)
    
    # Deep XML diff
    print("\n--- Deep XML Structure Differences ---")
    xml_diffs = deep_diff_xml_entries(e_orig, e_conmas, "Original", "ConMas")
    for d in sorted(xml_diffs, key=lambda x: -x['score']):
        if d['score'] >= 6:
            print(f"  [{d['score']}/10] {d['entry']}: {d['detail']}")
            print(f"    Reason: {d['reason']}")
            diffs_orig_vs_conmas.append(d)
    
    # ========== COMPARISON 2: ConMas vs Our ==========
    
    print("\n" + "=" * 80)
    print("COMPARISON 2: CONMAS vs OUR OUTPUT")
    print("=" * 80)
    
    diffs_conmas_vs_ours = []
    
    print("\n--- ZIP Entry Differences ---")
    entry_diffs = compare_entries(
        set(list(e_conmas.keys()) + list(e_ours.keys())),
        e_conmas, e_ours, "ConMas", "Our output"
    )
    for d in sorted(entry_diffs, key=lambda x: -x['score']):
        if d['score'] >= 3:
            print(f"  [{d['score']}/10] {d['entry']}: {d['detail']}")
            print(f"    Reason: {d['reason']}")
            diffs_conmas_vs_ours.append(d)
    
    print("\n--- Structural Differences ---")
    struct_diffs = compare_workbooks_structure(s_conmas, s_ours, "ConMas", "Our output")
    for d in sorted(struct_diffs, key=lambda x: -x['score']):
        if d['score'] >= 3:
            print(f"  [{d['score']}/10] {d['entry']}: {d['detail']}")
            print(f"    Reason: {d['reason']}")
            diffs_conmas_vs_ours.append(d)
    
    print("\n--- Deep XML Structure Differences ---")
    xml_diffs = deep_diff_xml_entries(e_conmas, e_ours, "ConMas", "Our output")
    for d in sorted(xml_diffs, key=lambda x: -x['score']):
        if d['score'] >= 5:
            print(f"  [{d['score']}/10] {d['entry']}: {d['detail']}")
            print(f"    Reason: {d['reason']}")
            diffs_conmas_vs_ours.append(d)
    
    # ========== SUMMARY ==========
    
    print("\n" + "=" * 80)
    print("FINAL SUMMARY: ALL DIFFERENCES RANKED")
    print("=" * 80)
    
    # Combine ConMas vs Our differences
    all_diffs = diffs_conmas_vs_ours
    all_diffs.sort(key=lambda x: -x['score'])
    
    print("\nAll differences between ConMas and Our output (sorted by impact):")
    print()
    
    high_impact = [d for d in all_diffs if d['score'] >= 8]
    med_impact = [d for d in all_diffs if 5 <= d['score'] < 8]
    low_impact = [d for d in all_diffs if d['score'] < 5]
    
    print(f"  🔴 HIGH IMPACT (score 8-10): {len(high_impact)}")
    for d in high_impact:
        print(f"    [{d['score']}/10] {d['entry']}")
        print(f"      {d['detail']}")
        print(f"      {d['reason']}")
    
    print(f"\n  🟡 MEDIUM IMPACT (score 5-7): {len(med_impact)}")
    for d in med_impact:
        print(f"    [{d['score']}/10] {d['entry']}: {d['detail']}")
    
    print(f"\n  🟢 LOW IMPACT (score 1-4): {len(low_impact)}")
    for d in low_impact:
        print(f"    [{d['score']}/10] {d['entry']}: {d['detail']}")
    
    # ========== FINAL ANSWER ==========
    print("\n" + "=" * 80)
    print("ROOT CAUSE ANALYSIS")
    print("=" * 80)
    
    if high_impact:
        top = high_impact[0]
        print(f"\n  Single most likely remaining cause of behavioral mismatch:")
        print(f"\n  🔴 [{top['score']}/10] {top['entry']}")
        print(f"     {top['detail']}")
        print(f"     {top['reason']}")
        print()
    
    print("  Top 3 differences by print-engine impact:")
    for i, d in enumerate(high_impact[:3]):
        print(f"    {i+1}. [{d['score']}/10] {d['entry']}")
        print(f"       {d['detail']}")
    
    # Already proven irrelevant
    print(f"\n  Already proven irrelevant (from prior phases):")
    print(f"    - HeaderMargin/FooterMargin: ~0.015in delta (X.28 confirmed does not affect content position)")
    print(f"    - Document metadata (creator, timestamps)")
    print(f"    - Font/color/style differences (unless they affect column/row sizing)")
    print(f"    - Selection state, active cell, tab selection")
    
    print("\n" + "=" * 80)
    print("RECOMMENDED NEXT STEPS")
    print("=" * 80)
    print("""
    1. Fix the highest-impact differences found above (starting with score 9-10)
    2. Generate a new output with fixes
    3. Re-run this comparison to verify differences are resolved
    4. Perform acceptance test (Tests A-D from Phase X.30 specification)
    """)


# ── Main ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 4:
        print("Usage: python phase30_forensic_diff.py <original.xlsx> <conmas.xlsx> <our_output.xlsx>")
        sys.exit(1)
    
    generate_report(sys.argv[1], sys.argv[2], sys.argv[3])
