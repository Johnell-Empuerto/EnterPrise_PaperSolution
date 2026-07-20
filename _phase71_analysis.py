#!/usr/bin/env python3
"""Phase 7.1 - Multi-Generation Workbook Comparison across all 4 generations."""

import os, sys, hashlib
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from collections import defaultdict

base = r'C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise\_phase71_output'
files = ['Original.xlsx', 'Export1.xlsx', 'Export2.xlsx', 'Export3.xlsx']

# ─── 1. ZIP Entry Hash Table ───
print("=" * 90)
print("PHASE 7.1 - MULTI-GENERATION COMPONENT HASH TABLE")
print("=" * 90)

entries = {}
for f in files:
    path = os.path.join(base, f)
    with ZipFile(path) as z:
        for e in z.infolist():
            data = z.read(e.filename)
            h = hashlib.sha256(data).hexdigest()
            entries.setdefault(e.filename, {})[f] = (e.file_size, h)

header = f"{'Component':<45} {'Orig hash':<44} {'E1 hash':<44} {'E2 hash':<44} {'E3 hash':<44} {'Stable?'}"
print(header)
print("-" * len(header))
for name in sorted(entries.keys()):
    row = f"{name:<45}"
    hashes = []
    for f in files:
        if f in entries[name]:
            sz, h = entries[name][f]
            row += f"{h[:12]:<44}"
            hashes.append(h)
        else:
            row += f"{'---':<44}"
            hashes.append(None)
    # Check stability: all present and identical
    non_none = [h for h in hashes if h is not None]
    stable = "YES" if len(set(non_none)) == 1 else "NO"
    row += f"{stable}"
    print(row)

# ─── 2. Detailed Diff: Export1 vs Export2 ───
print("\n" + "=" * 90)
print("DETAILED COMPARISON: Export1 vs Export2")
print("=" * 90)

with ZipFile(os.path.join(base, files[1])) as z1, ZipFile(os.path.join(base, files[2])) as z2:
    for name in sorted(set(z1.namelist()) | set(z2.namelist())):
        d1 = z1.read(name) if name in z1.namelist() else None
        d2 = z2.read(name) if name in z2.namelist() else None
        if d1 == d2:
            continue
        h1 = hashlib.sha256(d1).hexdigest()[:16] if d1 else "MISSING"
        h2 = hashlib.sha256(d2).hexdigest()[:16] if d2 else "MISSING"
        print(f"\n  [{name}]")
        print(f"    Export1: {len(d1) if d1 else 0} bytes, hash={h1}")
        print(f"    Export2: {len(d2) if d2 else 0} bytes, hash={h2}")
        
        # For XML, show the actual diff
        if d1 and d2 and name.endswith('.xml'):
            try:
                r1 = ET.fromstring(d1)
                r2 = ET.fromstring(d2)
                show_xml_diff(r1, r2, name, indent=4)
            except Exception as ex:
                print(f"    XML parse error: {ex}")
        elif d1 and d2 and name.endswith('.bin'):
            print(f"    Binary diff: first 20 bytes differ")
            if d1[:20] != d2[:20]:
                print(f"      Export1: {d1[:20].hex()}")
                print(f"      Export2: {d2[:20].hex()}")

# ─── 3. Detailed Diff: Export2 vs Export3 ───
print("\n" + "=" * 90)
print("DETAILED COMPARISON: Export2 vs Export3")
print("=" * 90)

with ZipFile(os.path.join(base, files[2])) as z1, ZipFile(os.path.join(base, files[3])) as z2:
    for name in sorted(set(z1.namelist()) | set(z2.namelist())):
        d1 = z1.read(name) if name in z1.namelist() else None
        d2 = z2.read(name) if name in z2.namelist() else None
        if d1 == d2:
            continue
        h1 = hashlib.sha256(d1).hexdigest()[:16] if d1 else "MISSING"
        h2 = hashlib.sha256(d2).hexdigest()[:16] if d2 else "MISSING"
        print(f"\n  [{name}]")
        print(f"    Export2: {len(d1) if d1 else 0} bytes, hash={h1}")
        print(f"    Export3: {len(d2) if d2 else 0} bytes, hash={h2}")
        
        if d1 and d2 and name.endswith('.xml'):
            try:
                r1 = ET.fromstring(d1)
                r2 = ET.fromstring(d2)
                show_xml_diff(r1, r2, name, indent=4)
            except Exception as ex:
                print(f"    XML parse error: {ex}")
        elif d1 and d2 and name.endswith('.bin'):
            print(f"    Binary diff: first 20 bytes differ")
            if d1[:20] != d2[:20]:
                print(f"      Export2: {d1[:20].hex()}")
                print(f"      Export3: {d2[:20].hex()}")

# ─── 4. Structural Stability Report ───
print("\n" + "=" * 90)
print("STRUCTURAL STABILITY REPORT")
print("=" * 90)
print(f"{'Component':<45} {'Orig→E1':<15} {'E1→E2':<15} {'E2→E3':<15}")
print("-" * 90)

S_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

parts = [
    "xl/workbook.xml",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "xl/theme/theme1.xml",
    "[Content_Types].xml",
    "docProps/core.xml",
    "docProps/app.xml",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/_rels/sheet2.xml.rels",
    "xl/printerSettings/printerSettings1.bin",
]

pairs = [("Original", "Export1"), ("Export1", "Export2"), ("Export2", "Export3")]

for part in parts:
    row = f"{part:<45}"
    for f1, f2 in pairs:
        p1 = os.path.join(base, f1 + ".xlsx")
        p2 = os.path.join(base, f2 + ".xlsx")
        try:
            with ZipFile(p1) as z1, ZipFile(p2) as z2:
                d1 = z1.read(part) if part in z1.namelist() else None
                d2 = z2.read(part) if part in z2.namelist() else None
                if d1 is None and d2 is None:
                    row += f"{'N/A':<15}"
                elif d1 is None or d2 is None:
                    row += f"{'ADDED/REMOVED':<15}"
                elif d1 == d2:
                    row += f"{'SAME':<15}"
                else:
                    row += f"{'DIFFERENT':<15}"
        except:
            row += f"{'ERROR':<15}"
    print(row)


def show_xml_diff(r1, r2, path, indent=0):
    """Show differences between two XML trees."""
    sp = " " * indent
    
    # Compare attributes
    a1 = set(r1.attrib.items())
    a2 = set(r2.attrib.items())
    if a1 != a2:
        for k, v in sorted(a1 - a2):
            print(f"{sp}Orig attr: {k}={v}")
        for k, v in sorted(a2 - a1):
            print(f"{sp}Edit attr: {k}={v}")
    
    # Compare text
    t1 = (r1.text or "").strip()
    t2 = (r2.text or "").strip()
    if t1 != t2:
        print(f"{sp}Text: '{t1[:60]}' -> '{t2[:60]}'")
    
    # Compare children
    c1 = list(r1)
    c2 = list(r2)
    max_c = max(len(c1), len(c2))
    for i in range(max_c):
        child1 = c1[i] if i < len(c1) else None
        child2 = c2[i] if i < len(c2) else None
        if child1 is None and child2 is not None:
            tag = child2.tag.split('}')[-1]
            print(f"{sp}+ New element: {tag}")
        elif child2 is None and child1 is not None:
            tag = child1.tag.split('}')[-1]
            print(f"{sp}- Missing element: {tag}")
        elif ET.tostring(child1) != ET.tostring(child2):
            tag = child1.tag.split('}')[-1]
            print(f"{sp}~ Different: [{i}] {tag}")
            show_xml_diff(child1, child2, f"{path}/{tag}", indent + 2)
