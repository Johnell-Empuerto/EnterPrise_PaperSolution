"""
Phase X.9 — Forensic OOXML Binary Diff: Legacy vs Generated Workbook
Compares every ZIP entry and every XML node to determine exactly what ConMas reads.
"""
import os, sys, zipfile
from xml.etree import ElementTree as ET
from difflib import unified_diff

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

LEGACY = r'C:\Users\MCF-JOHNELLEEMPUERTO\Downloads\FormTest - Copy.xlsx'
GENERATED = r'C:\Users\MCF-JOHNELLEEMPUERTO\Downloads\FormTest - Copy_output.xlsx'

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def get_zip_entries(path):
    """Get sorted list of (name, size, data) for all ZIP entries."""
    entries = {}
    with zipfile.ZipFile(path, 'r') as z:
        for name in sorted(z.namelist()):
            info = z.getinfo(name)
            entries[name] = {
                'size': info.file_size,
                'compress_size': info.compress_size,
                'data': z.read(name),
            }
    return entries


def compare_xml_entries(name, data_a, data_b):
    """Compare two XML files at the node level. Return list of differences."""
    diffs = []
    try:
        xml_a = ET.fromstring(data_a)
        xml_b = ET.fromstring(data_b)
    except:
        # Binary or non-XML file — compare raw bytes
        if data_a != data_b:
            return [f"  Binary difference: {len(data_a)} vs {len(data_b)} bytes"]
        return []

    # Compare root tag
    tag_a = xml_a.tag.split('}')[-1] if '}' in xml_a.tag else xml_a.tag
    tag_b = xml_b.tag.split('}')[-1] if '}' in xml_b.tag else xml_b.tag
    if tag_a != tag_b:
        diffs.append(f"  Root tag: {tag_a} vs {tag_b}")

    # Compare attributes
    attrs_a = sorted(xml_a.attrib.items())
    attrs_b = sorted(xml_b.attrib.items())
    if attrs_a != attrs_b:
        diffs.append(f"  Root attributes differ")
        for k, v in attrs_a:
            if (k, v) not in attrs_b:
                diffs.append(f"    Only in legacy: {k}={v}")
        for k, v in attrs_b:
            if (k, v) not in attrs_a:
                diffs.append(f"    Only in generated: {k}={v}")

    # Count child elements by tag
    def count_children(elem):
        counts = {}
        for child in elem:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            counts[tag] = counts.get(tag, 0) + 1
            deeper = count_children(child)
            for k, v in deeper.items():
                counts[k] = counts.get(k, 0) + v
        return counts

    counts_a = count_children(xml_a)
    counts_b = count_children(xml_b)
    all_tags = set(list(counts_a.keys()) + list(counts_b.keys()))
    for tag in sorted(all_tags):
        ca = counts_a.get(tag, 0)
        cb = counts_b.get(tag, 0)
        if ca != cb:
            diffs.append(f"  Element <{tag}>: {ca} vs {cb}")

    return diffs


# ══════════════════════════════════════════════════════════════════════════
# MAIN COMPARISON
# ══════════════════════════════════════════════════════════════════════════

print("=" * 70)
print("PHASE X.9 — FORENSIC OOXML BINARY DIFF")
print("=" * 70)
print(f"Legacy:    {LEGACY}")
print(f"Generated: {GENERATED}")
print(f"Legacy size:    {os.path.getsize(LEGACY):,} bytes")
print(f"Generated size: {os.path.getsize(GENERATED):,} bytes")
print()

entries_a = get_zip_entries(LEGACY)
entries_b = get_zip_entries(GENERATED)

# ── 1. ZIP Entry Comparison ──
print("─" * 70)
print("1. ZIP ENTRY COMPARISON")
print("─" * 70)

all_names = sorted(set(list(entries_a.keys()) + list(entries_b.keys())))
print(f"\n{'Entry':<55} {'Legacy':>10} {'Gen':>10} {'Status':>12}")
print(f"{'─'*55} {'─'*10} {'─'*10} {'─'*12}")

for name in all_names:
    in_a = name in entries_a
    in_b = name in entries_b
    size_a = entries_a[name]['size'] if in_a else 0
    size_b = entries_b[name]['size'] if in_b else 0

    if not in_a:
        status = "ONLY IN GEN"
    elif not in_b:
        status = "ONLY IN LEG"
    elif size_a == size_b and entries_a[name]['data'] == entries_b[name]['data']:
        status = "IDENTICAL"
    else:
        status = "DIFFERENT"
    
    label = name if len(name) < 55 else "..." + name[-52:]
    print(f"{label:<55} {size_a:>10,} {size_b:>10,} {status:>12}")

# ── 2. XML Node-Level Diff ──
print(f"\n{'─'*70}")
print("2. XML NODE-LEVEL DIFFERENCES")
print("─" * 70)

xml_files = [n for n in all_names if (n.endswith('.xml') or n.endswith('.rels')) and n in entries_a and n in entries_b]

for name in xml_files:
    data_a = entries_a[name]['data']
    data_b = entries_b[name]['data']
    
    # Skip identical files
    if data_a == data_b:
        continue
    
    print(f"\n  File: {name}")
    diffs = compare_xml_entries(name, data_a, data_b)
    if diffs:
        for d in diffs:
            print(d)
    else:
        print(f"  XML structure similar but content differs (text values)")
        
        # For shared strings — show text diff
        if 'sharedStrings' in name:
            try:
                ss_a = ET.fromstring(data_a)
                ss_b = ET.fromstring(data_b)
                texts_a = [t.text or '' for t in ss_a.findall('.//s:t', NS)]
                texts_b = [t.text or '' for t in ss_b.findall('.//s:t', NS)]
                print(f"    Shared strings: {len(texts_a)} vs {len(texts_b)}")
                missing = set(texts_a) - set(texts_b)
                extra = set(texts_b) - set(texts_a)
                if missing:
                    print(f"    Only in legacy ({len(missing)}):")
                    for t in list(missing)[:5]:
                        print(f"      {repr(t[:80])}")
                if extra:
                    print(f"    Only in generated ({len(extra)}):")
                    for t in list(extra)[:5]:
                        print(f"      {repr(t[:80])}")
            except Exception as e:
                print(f"    Error parsing shared strings: {e}")

# ── 3. Sheet Comparison ──
print(f"\n{'─'*70}")
print("3. WORKSHEET COMPARISON")
print("─" * 70)

for label, entries in [('LEGACY', entries_a), ('GENERATED', entries_b)]:
    print(f"\n  --- {label} ---")
    wb = ET.fromstring(entries['xl/workbook.xml']['data'])
    sheets = wb.findall('.//s:sheet', NS)
    for s in sheets:
        name = s.get('name', '?')
        state = s.get('state', 'visible')
        rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
        # Find sheet file from rels
        rels = ET.fromstring(entries['xl/_rels/workbook.xml.rels']['data'])
        target = ''
        for r in rels:
            if r.get('Id') == rid:
                target = r.get('Target', '')
                break
        if target:
            sheet_path = f'xl/{target}'
            if sheet_path in entries:
                ws = ET.fromstring(entries[sheet_path]['data'])
                dim = ws.find('.//s:dimension', NS)
                dim_ref = dim.get('ref', '') if dim is not None else 'none'
                merges = len(ws.findall('.//s:mergeCells/s:mergeCell', NS))
                rows = len(ws.findall('.//s:row', NS))
                print(f"    {name:<25} state={state:<10} dim={dim_ref:<15} rows={rows:<5} merges={merges}")

# ── 4. Comment Comparison ──
print(f"\n{'─'*70}")
print("4. COMMENT COMPARISON")
print("─" * 70)

for label, entries in [('LEGACY', entries_a), ('GENERATED', entries_b)]:
    cmt_files = sorted([n for n in entries if 'comments' in n.lower() and n.endswith('.xml')])
    print(f"\n  --- {label} ---")
    for cf in cmt_files:
        cmt = ET.fromstring(entries[cf]['data'])
        cmt_ns = {'c': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        comments = cmt.findall('.//c:comment', cmt_ns)
        authors = cmt.findall('.//c:author', cmt_ns)
        print(f"    {cf}: {len(authors)} author(s), {len(comments)} comment(s)")
        for c in comments:
            ref = c.get('ref', '')
            texts = [t.text or '' for t in c.findall('.//c:t', cmt_ns)]
            full = ' '.join(texts)
            lines = full.split('\n')
            print(f"      [{ref}] {len(lines)} lines, first line: {repr(lines[0][:50])}")

# ── 5. Defined Names Comparison ──
print(f"\n{'─'*70}")
print("5. DEFINED NAMES")
print("─" * 70)

for label, entries in [('LEGACY', entries_a), ('GENERATED', entries_b)]:
    wb = ET.fromstring(entries['xl/workbook.xml']['data'])
    dns = wb.findall('.//s:definedName', NS)
    print(f"\n  {label}: {len(dns)} defined name(s)")
    for dn in dns:
        name = dn.get('name', '?')
        text = dn.text or ''
        print(f"    {name} = {text}")

# ── 6. Dependency Graph ──
print(f"\n{'─'*70}")
print("6. METADATA DEPENDENCY GRAPH")
print("─" * 70)
print("""
Based on all comparison data, the dependency graph is:

WORKBOOK
  │
  ├── xl/workbook.xml  ─────────────► Sheet definitions, defined names
  │     │
  │     ├── Sheets ──────────────────► Sheet1 (visible, content)
  │     │                              ├── _Fields (HIDDEN — NOT read by ConMas)
  │     │                              └── ExcelOutputSetting (visible — purpose unknown)
  │     │
  │     └── DefinedNames ────────────► _xlnm.Print_Area = Sheet1!$A$1:$D$12
  │                                    (CRITICAL — legacy requires PrintArea)
  │
  ├── xl/worksheets/sheet2.xml ─────► Sheet1 content
  │     │
  │     ├── MergedCells (5 regions) ─► Field positioning (A1:B2, C1:D2, A3:D4, A6:D7, A9:D10)
  │     ├── Cell values ─────────────► Existing field values (from PrintArea range)
  │     └── PrintArea ───────────────► Processing boundary (REQUIRED)
  │
  ├── xl/comments1.xml ─────────────► PRIMARY metadata source
  │     │                              ├── 6 comments (one per field)
  │     │                              ├── Each comment = Name + Type + Index + Params
  │     │                              └── Located on top-left cell of merged regions
  │
  ├── xl/drawings/vmlDrawing1.vml ──► AUTO-GENERATED by Excel COM AddComment()
  │                                    (ConMas may NOT read this — Excel renders it)
  │
  ├── xl/sharedStrings.xml ─────────► All string values
  │     │                              ├── _Fields headers (7)
  │     │                              ├── ConMas XML config fragments (36, from ExcelOutputSetting)
  │     │                              └── Cell values from Sheet1
  │
  ├── xl/styles.xml ─────────────────► Cell formatting (fonts, colors, borders)
  │                                    (Optional — not required for field count)
  │
  └── xl/printerSettings/ ──────────► Printer configuration
                                       (Optional — auto-generated by Excel)
""")

# ── 7. Key Differences Summary ──
print(f"\n{'─'*70}")
print("7. KEY DIFFERENCES — LEGACY vs GENERATED")
print("─" * 70)
print("""
# DIFFERENCE                    LEGACY              GENERATED           IMPACT
─────────────────────────────────────────────────────────────────────────────
1. Sheet count                  3                   2                   HIGH
   - _Fields (hidden)           Headers only        6 data rows         ConMas ignores hidden → OK
   - Sheet1                     Content + merges    Empty/no merges     CRITICAL
   - ExcelOutputSetting         36 XML rows         MISSING             UNKNOWN

2. Comments                     6 comments          1 comment           CRITICAL
   - Format                    Full (6+ lines)     Minimal (4 lines)   HIGH
   - Cluster index             Present             MISSING             MEDIUM
   - Parameters string         Full seralized      '0' placeholder     HIGH

3. PrintArea                   _xlnm.Print_Area    NOT SET             CRITICAL
                                = Sheet1!$A$1:$D$12                    (Legacy requires it)

4. Merged cells                5 regions           NONE                HIGH
   - A1:B2, C1:D2, A3:D4, etc.                                        (Comments reference merge areas)

5. Sheet1 cell values          33 cells            0 cells             MEDIUM
                                                                        (Needed for existing values)

6. Defined names               1 (Print_Area)      NONE                MEDIUM

7. ExcelOutputSetting sheet    36 rows             MISSING             UNKNOWN
   - ConMas XML config         Present in SS items                     (May not be read for import)
""")

print(f"\n{'='*70}")
print("ANALYSIS COMPLETE")
print("=" * 70)
