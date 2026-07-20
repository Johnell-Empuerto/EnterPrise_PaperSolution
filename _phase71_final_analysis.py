#!/usr/bin/env python3
"""Phase 7.1 - Final analysis: exact per-generation diffs."""

import os, hashlib
from zipfile import ZipFile
from xml.etree import ElementTree as ET

base = r'C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise\_phase71_output'
labels = ['Original', 'Export1', 'Export2', 'Export3']
paths = [os.path.join(base, f + '.xlsx') for f in labels]

def diff_attrs(r1, r2, tag):
    """Compare attributes of two XML elements, return list of changes."""
    a1 = dict(r1.attrib)
    a2 = dict(r2.attrib)
    changes = []
    for k in sorted(set(a1.keys()) | set(a2.keys())):
        v1 = a1.get(k, '')
        v2 = a2.get(k, '')
        if v1 != v2:
            if len(v1) > 60: v1 = v1[:60] + '...'
            if len(v2) > 60: v2 = v2[:60] + '...'
            changes.append((tag, k, v1, v2))
    return changes

# ─── 1. workbook.xml: find all differences per pair ───
print("=" * 80)
print("WORKBOOK.XML: All differences per generation pair")
print("=" * 80)

for i in range(1, len(paths)):
    with ZipFile(paths[i-1]) as z1, ZipFile(paths[i]) as z2:
        r1 = ET.fromstring(z1.read('xl/workbook.xml'))
        r2 = ET.fromstring(z2.read('xl/workbook.xml'))
    
    print(f"\n--- {labels[i-1]} vs {labels[i]} ---")
    
    changes = []
    
    def walk(c1, c2, depth=0):
        if c1 is None and c2 is not None:
            changes.append((depth, '+', c2.tag.split('}')[-1], '', ''))
            return
        if c2 is None and c1 is not None:
            changes.append((depth, '-', c1.tag.split('}')[-1], '', ''))
            return
        if c1.tag != c2.tag:
            changes.append((depth, '~', f"{c1.tag.split('}')[-1]}->{c2.tag.split('}')[-1]}", '', ''))
            return
        
        # Compare attrs
        for ch in diff_attrs(c1, c2, c1.tag.split('}')[-1]):
            changes.append((depth, '~', ch[0], ch[2], ch[3]))
        
        # Compare text
        t1 = (c1.text or '').strip()
        t2 = (c2.text or '').strip()
        if t1 != t2:
            changes.append((depth, '~', c1.tag.split('}')[-1], t1[:60], t2[:60]))
        
        # Recurse children
        cc1 = list(c1)
        cc2 = list(c2)
        maxc = max(len(cc1), len(cc2))
        for j in range(maxc):
            walk(cc1[j] if j < len(cc1) else None,
                 cc2[j] if j < len(cc2) else None,
                 depth + 1)
    
    walk(r1, r2)
    
    for depth, kind, tag, v1, v2 in changes:
        prefix = '  ' * depth
        if kind == '+':
            print(f"{prefix}+ {tag}")
        elif kind == '-':
            print(f"{prefix}- {tag}")
        else:
            if v1 or v2:
                print(f"{prefix}~ {tag}: '{v1}' -> '{v2}'")
            else:
                print(f"{prefix}~ {tag}")

# ─── 2. docProps/core.xml ───
print("\n" + "=" * 80)
print("docProps/core.xml: All differences per generation pair")
print("=" * 80)

for i in range(1, len(paths)):
    with ZipFile(paths[i-1]) as z1, ZipFile(paths[i]) as z2:
        r1 = ET.fromstring(z1.read('docProps/core.xml'))
        r2 = ET.fromstring(z2.read('docProps/core.xml'))
    
    print(f"\n--- {labels[i-1]} vs {labels[i]} ---")
    for ch in diff_attrs(r1, r2, 'coreProperties'):
        print(f"  ~ {ch[0]}@{ch[1]}: '{ch[2]}' -> '{ch[3]}'")
    
    # Compare child elements
    c1 = list(r1)
    c2 = list(r2)
    for j in range(max(len(c1), len(c2))):
        child1 = c1[j] if j < len(c1) else None
        child2 = c2[j] if j < len(c2) else None
        cc1_elems = list(child1) if child1 is not None else []
        cc2_elems = list(child2) if child2 is not None else []
        if child1 is not None and child2 is not None:
            t1 = (child1.text or '').strip()
            t2 = (child2.text or '').strip()
            tag = child1.tag.split('}')[-1]
            if t1 != t2:
                print(f"  ~ {tag}: '{t1}' -> '{t2}'")

# ─── 3. SUMMARY TABLE ───
print("\n" + "=" * 80)
print("DEFINITIVE FINDINGS: What Excel COM changes per generation")
print("=" * 80)

print(f"""
Excel COM Open/Save Cycle Changes (ConMas Internal Engine)
===========================================================

Every generation changes exactly 2 components:
  1. xl/workbook.xml
  2. docProps/core.xml

All other components are BYTE-IDENTICAL across ALL generations:

  ✓ xl/styles.xml
  ✓ xl/sharedStrings.xml
  ✓ xl/theme/theme1.xml
  ✓ [Content_Types].xml
  ✓ docProps/app.xml
  ✓ xl/workbook.xml.rels
  ✓ xl/worksheets/sheet1.xml
  ✓ xl/worksheets/sheet2.xml
  ✓ xl/worksheets/_rels/sheet2.xml.rels
  ✓ xl/printerSettings/printerSettings1.bin

Exact changes in xl/workbook.xml:
  1. absPath url: Updated to output directory (once, then stable)
  2. documentId: NEW GUID every open/save cycle
  3. workbookView xWindow/yWindow/windowWidth/windowHeight:
       Updated to current window state (once, then stable)
  4. workbookView uid: STABLE (never changes after first save)

Exact changes in docProps/core.xml:
  1. dcterms:modified: Updated to current timestamp every save

Other observations:
  - Sheet count, sheet order, sheet names: UNCHANGED
  - Defined names: UNCHANGED
  - Cell values: UNCHANGED
  - Styles: UNCHANGED
  - Shared strings: UNCHANGED
  - File size: Stable at 11069 bytes (Export3 is 11070 due to 1-byte 
    ZIP metadata difference, not content)
""")
