#!/usr/bin/env python3
"""Phase 7.1 - Detailed XML analysis of what changes per generation."""

import os, hashlib
from zipfile import ZipFile
from xml.etree import ElementTree as ET

base = r'C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise\_phase71_output'
files = ['Original.xlsx', 'Export1.xlsx', 'Export2.xlsx', 'Export3.xlsx']
labels = ['Original', 'Export1', 'Export2', 'Export3']
paths = [os.path.join(base, f) for f in files]

# ─── 1. workbook.xml across all generations ───
print("=" * 90)
print("1. WORKBOOK.XML - Full content per generation")
print("=" * 90)

wb_parts = ['']

for idx, (label, path) in enumerate(zip(labels, paths)):
    with ZipFile(path) as z:
        data = z.read('xl/workbook.xml').decode('utf-8')
    print(f"\n--- {label} ({len(data)} bytes, hash={hashlib.sha256(data.encode()).hexdigest()[:16]}) ---")
    print(data)

# ─── 2. docProps/core.xml across all generations ───
print("\n" + "=" * 90)
print("2. docProps/core.xml - Full content per generation")
print("=" * 90)

for idx, (label, path) in enumerate(zip(labels, paths)):
    with ZipFile(path) as z:
        data = z.read('docProps/core.xml').decode('utf-8')
    print(f"\n--- {label} ({len(data)} bytes, hash={hashlib.sha256(data.encode()).hexdigest()[:16]}) ---")
    print(data)

# ─── 3. Is the pattern consistent? ───
print("\n" + "=" * 90)
print("3. ANALYSIS: What changes each generation?")
print("=" * 90)

# Compare generations pairwise
for i in range(len(paths) - 1):
    f1, f2 = paths[i], paths[i+1]
    l1, l2 = labels[i], labels[i+1]
    
    print(f"\n--- {l1} vs {l2} ---")
    
    with ZipFile(f1) as z1, ZipFile(f2) as z2:
        all_names = sorted(set(z1.namelist()) | set(z2.namelist()))
        changes = []
        for name in all_names:
            d1 = z1.read(name) if name in z1.namelist() else None
            d2 = z2.read(name) if name in z2.namelist() else None
            if d1 != d2:
                changes.append((name, len(d1) if d1 else 0, len(d2) if d2 else 0))
        
        print(f"  Changed components: {len(changes)}")
        for name, sz1, sz2 in changes:
            status = "SAME SIZE" if sz1 == sz2 else f"{sz1}->{sz2} bytes"
            print(f"    {name}: {status}")

# ─── 4. What about the actual diff in workbook.xml per generation? ───
print("\n" + "=" * 90)
print("4. WORKBOOK.XML - Element-by-element diff per generation pair")
print("=" * 90)

ns = {
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'x15ac': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac',
    'xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
    'xr6': 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision6',
    'xr10': 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10',
    'xr2': 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2',
    's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
}

for i in range(1, len(paths)):
    f1, f2 = paths[i-1], paths[i]
    l1, l2 = labels[i-1], labels[i]
    
    with ZipFile(f1) as z1, ZipFile(f2) as z2:
        r1 = ET.fromstring(z1.read('xl/workbook.xml'))
        r2 = ET.fromstring(z2.read('xl/workbook.xml'))
    
    print(f"\n--- {l1} vs {l2} ---")
    diff_elements(r1, r2, '')

# ─── 5. docProps/core.xml diff per generation ───
print("\n" + "=" * 90)
print("5. docProps/core.xml - Element-by-element diff per generation pair")
print("=" * 90)

for i in range(1, len(paths)):
    f1, f2 = paths[i-1], paths[i]
    l1, l2 = labels[i-1], labels[i]
    
    with ZipFile(f1) as z1, ZipFile(f2) as z2:
        r1 = ET.fromstring(z1.read('docProps/core.xml'))
        r2 = ET.fromstring(z2.read('docProps/core.xml'))
    
    print(f"\n--- {l1} vs {l2} ---")
    diff_elements(r1, r2, '')

# ─── 6. SUMMARY ───
print("\n" + "=" * 90)
print("6. SUMMARY OF FINDINGS")
print("=" * 90)
print("""
Multi-Generation COM Open/Save Test Results
============================================
Input: Investigation_546/original.xlsx (2-sheet ConMas form template)
Excel: Microsoft Excel COM Interop (same engine ConMas uses internally)

KEY FINDINGS:
""")

# Check which components changed
changes_by_gen = []
for i in range(len(paths) - 1):
    with ZipFile(paths[i]) as z1, ZipFile(paths[i+1]) as z2:
        changed = []
        for name in sorted(set(z1.namelist()) | set(z2.namelist())):
            d1 = z1.read(name) if name in z1.namelist() else None
            d2 = z2.read(name) if name in z2.namelist() else None
            if d1 != d2:
                changed.append(name)
        changes_by_gen.append(changed)

for idx, (l1, l2, changed) in enumerate(zip(labels[:-1], labels[1:], changes_by_gen)):
    print(f"  {l1} -> {l2}: {len(changed)} component(s) changed")
    for name in changed:
        print(f"    - {name}")


def diff_elements(r1, r2, path):
    """Show element-by-element differences."""
    a1 = dict(r1.attrib)
    a2 = dict(r2.attrib)
    
    all_keys = set(a1.keys()) | set(a2.keys())
    for k in all_keys:
        v1 = a1.get(k, '')
        v2 = a2.get(k, '')
        if v1 != v2:
            # Try to simplify long values
            if len(v1) > 60: v1 = v1[:60] + '...'
            if len(v2) > 60: v2 = v2[:60] + '...'
            tag = r1.tag.split('}')[-1] if '}' in r1.tag else r1.tag
            print(f"  @{tag}/{k}: '{v1}' -> '{v2}'")
    
    c1 = list(r1)
    c2 = list(r2)
    max_c = max(len(c1), len(c2))
    for i in range(max_c):
        child1 = c1[i] if i < len(c1) else None
        child2 = c2[i] if i < len(c2) else None
        if child1 is None and child2 is not None:
            tag2 = child2.tag.split('}')[-1]
            print(f"  + New: [{i}] {tag2}")
        elif child2 is None and child1 is not None:
            tag1 = child1.tag.split('}')[-1]
            print(f"  - Removed: [{i}] {tag1}")
        elif ET.tostring(child1) != ET.tostring(child2):
            tag = child1.tag.split('}')[-1]
            a1c = dict(child1.attrib)
            a2c = dict(child2.attrib)
            for k in set(a1c.keys()) | set(a2c.keys()):
                v1 = a1c.get(k, '')
                v2 = a2c.get(k, '')
                if v1 != v2:
                    if len(v1) > 60: v1 = v1[:60] + '...'
                    if len(v2) > 60: v2 = v2[:60] + '...'
                    print(f"  [{i}] {tag}/@{k}: '{v1}' -> '{v2}'")
            diff_elements(child1, child2, f"{path}/{tag}")
