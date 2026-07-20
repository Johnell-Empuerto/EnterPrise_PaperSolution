"""
Phase 7 — Forensic Workbook Comparison
========================================
Compares two workbooks at every level:
  1. ZIP entry inventory (name, size, hash)
  2. Workbook.xml element-by-element
  3. Styles.xml comparison
  4. SharedStrings.xml entry-by-entry
  5. Worksheet XMLs (rows, cells, values, merges)
  6. Comments (if present)
  7. Relationships
  8. Content Types
  9. Defined names

Usage:
    python forensic_compare_workbooks.py <original.xlsx> <edited.xlsx> [--verbose]
"""

import sys
import os
import hashlib
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from collections import defaultdict

S_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
RELS_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def format_xml(xml: str) -> str:
    """Pretty-print an XML snippet."""
    if not xml or xml.strip() == "(null)" or xml.strip() == "(missing)":
        return xml or "(empty)"
    try:
        root = ET.fromstring(xml)
        ET.indent(root)
        return ET.tostring(root, encoding="unicode")
    except Exception:
        return xml


def compare_workbooks(orig_path: str, edit_path: str, verbose: bool = False):
    print("=" * 70)
    print(f"  FORENSIC WORKBOOK COMPARISON")
    print(f"  Original: {os.path.basename(orig_path)} ({os.path.getsize(orig_path):,} bytes)")
    print(f"  Edited:   {os.path.basename(edit_path)} ({os.path.getsize(edit_path):,} bytes)")
    print("=" * 70)

    with ZipFile(orig_path) as oz, ZipFile(edit_path) as ez:
        orig_entries = {e.filename: e for e in oz.infolist()}
        edit_entries = {e.filename: e for e in ez.infolist()}

        # ── 1. ZIP Entry Inventory ──
        print("\n" + "─" * 70)
        print("  SECTION 1: ZIP ENTRY INVENTORY")
        print("─" * 70)

        all_names = sorted(set(list(orig_entries.keys()) + list(edit_entries.keys())))

        print(f"\n  {'Entry':<45} {'Orig Size':<12} {'Edit Size':<12} {'Orig Hash':<44} {'Edit Hash':<44} {'Match'}")
        print(f"  {'-'*45:<45} {'-'*12:<12} {'-'*12:<12} {'-'*44:<44} {'-'*44:<44} {'-'*5}")

        diffs = []
        for name in all_names:
            oe = orig_entries.get(name)
            ee = edit_entries.get(name)

            orig_sz = oe.file_size if oe else 0
            edit_sz = ee.file_size if ee else 0

            orig_bytes = oz.read(name) if oe else b""
            edit_bytes = ez.read(name) if ee else b""

            orig_hash = sha256(orig_bytes) if oe else "(MISSING)"
            edit_hash = sha256(edit_bytes) if ee else "(MISSING)"

            match = "✅" if oe and ee and orig_hash == edit_hash else "❌" if oe or ee else "?"
            status = match

            if verbose or status == "❌":
                print(f"  {name:<45} {orig_sz:<12,} {edit_sz:<12,} {orig_hash:<44} {edit_hash:<44} {status}")

            if oe and not ee:
                diffs.append(f"  MISSING IN EDITED: {name}")
            elif ee and not oe:
                diffs.append(f"  NEW IN EDITED: {name}")
            elif oe and ee and orig_hash != edit_hash:
                diffs.append(f"  CHANGED: {name} ({orig_sz:,} → {edit_sz:,} bytes)")

        if diffs:
            print(f"\n  Differences found: {len(diffs)}")
            for d in diffs:
                print(f"    {d}")
        else:
            print(f"\n  All {len(all_names)} entries MATCH — workbooks are byte-identical.")
            return

        # ── 2. Compare XML parts deeply ──
        xml_parts = [
            "xl/workbook.xml",
            "xl/styles.xml",
            "xl/theme/theme1.xml",
            "xl/sharedStrings.xml",
            "[Content_Types].xml",
        ]

        for part_name in xml_parts:
            oe = orig_entries.get(part_name)
            ee = edit_entries.get(part_name)
            if not oe or not ee:
                print(f"\n  [{part_name}] SKIPPED — missing in one workbook")
                continue

            orig_xml = oz.read(part_name)
            edit_xml = ez.read(part_name)
            orig_hash = sha256(orig_xml)
            edit_hash = sha256(edit_xml)

            print(f"\n  ── {part_name} ──")
            if orig_hash == edit_hash:
                print(f"  ✅ IDENTICAL ({len(orig_xml):,} bytes)")
                continue

            print(f"  ❌ DIFFERENT ({len(orig_xml):,} → {len(edit_xml):,} bytes)")
            print(f"     Original hash: {orig_hash}")
            print(f"     Edited hash:   {edit_hash}")

            try:
                o_root = ET.fromstring(orig_xml)
                e_root = ET.fromstring(edit_xml)
                compare_xml_elements(o_root, e_root, part_name)
            except Exception as ex:
                print(f"     XML parse error: {ex}")

        # ── 3. Worksheet comparison ──
        print("\n" + "─" * 70)
        print("  SECTION 3: WORKSHEET COMPARISON")
        print("─" * 70)

        orig_sheets = [n for n in all_names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
        edit_sheets = [n for n in all_names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]

        print(f"\n  Original worksheets ({len(orig_sheets)}): {orig_sheets}")
        print(f"  Edited worksheets ({len(edit_sheets)}): {edit_sheets}")

        for name in sorted(set(orig_sheets) | set(edit_sheets)):
            oe = orig_entries.get(name)
            ee = edit_entries.get(name)
            if not oe or not ee:
                print(f"\n  [{name}] {'MISSING IN EDITED' if oe else 'NEW IN EDITED' if ee else ''}")
                continue

            orig_xml = oz.read(name)
            edit_xml = ez.read(name)
            orig_hash = sha256(orig_xml)
            edit_hash = sha256(edit_xml)

            print(f"\n  ── {name} ──")
            if orig_hash == edit_hash:
                print(f"  ✅ IDENTICAL")
                continue

            print(f"  ❌ DIFFERENT ({len(orig_xml):,} → {len(edit_xml):,} bytes)")

            # Compare cell values
            try:
                o_root = ET.fromstring(orig_xml)
                e_root = ET.fromstring(edit_xml)
                compare_cells(o_root, e_root, name)
                compare_sheet_structure(o_root, e_root, name)
            except Exception as ex:
                print(f"     XML parse error: {ex}")

        # ── 4. Worksheet Relationships ──
        print("\n" + "─" * 70)
        print("  SECTION 4: WORKSHEET RELATIONSHIPS")
        print("─" * 70)

        for name in sorted(set(orig_entries.keys()) | set(edit_entries.keys())):
            if not name.startswith("xl/worksheets/_rels/") or not name.endswith(".rels"):
                continue
            oe = orig_entries.get(name)
            ee = edit_entries.get(name)
            if not oe or not ee:
                print(f"\n  [{name}] {'MISSING IN EDITED' if oe else 'NEW IN EDITED' if ee else ''}")
                continue

            orig_xml = oz.read(name)
            edit_xml = ez.read(name)
            if orig_xml == edit_xml:
                print(f"  ✅ {name}: IDENTICAL")
            else:
                print(f"  ❌ {name}: DIFFERENT")
                try:
                    o_root = ET.fromstring(orig_xml)
                    e_root = ET.fromstring(edit_xml)
                    compare_relationships(o_root, e_root, name)
                except Exception as ex:
                    print(f"     Parse error: {ex}")

        # ── 5. Defined Names ──
        print("\n" + "─" * 70)
        print("  SECTION 5: DEFINED NAMES")
        print("─" * 70)

        try:
            wb_xml_orig = oz.read("xl/workbook.xml")
            wb_xml_edit = ez.read("xl/workbook.xml")
            o_wb = ET.fromstring(wb_xml_orig)
            e_wb = ET.fromstring(wb_xml_edit)
            compare_defined_names(o_wb, e_wb)
        except Exception as ex:
            print(f"  Error: {ex}")

        # ── 6. VML/Drawings comparison ──
        print("\n" + "─" * 70)
        print("  SECTION 6: VML DRAWINGS")
        print("─" * 70)

        for name in sorted(set(orig_entries.keys()) | set(edit_entries.keys())):
            if not name.lower().endswith(".vml"):
                continue
            oe = orig_entries.get(name)
            ee = edit_entries.get(name)
            if not oe or not ee:
                print(f"  [{name}] {'MISSING IN EDITED' if oe else 'NEW IN EDITED' if ee else ''}")
                continue

            orig_bytes = oz.read(name)
            edit_bytes = ez.read(name)
            if orig_bytes == edit_bytes:
                print(f"  ✅ {name}: IDENTICAL")
            else:
                orig_hash = sha256(orig_bytes)
                edit_hash = sha256(edit_bytes)
                print(f"  ❌ {name}: DIFFERENT ({len(orig_bytes):,} → {len(edit_bytes):,} bytes)")
                print(f"     Orig hash: {orig_hash}")
                print(f"     Edit hash: {edit_hash}")
                try:
                    o_text = orig_bytes.decode("utf-8", errors="replace")
                    e_text = edit_bytes.decode("utf-8", errors="replace")
                    if "uid" in o_text or "uid" in e_text:
                        # Check for GUID changes
                        import re
                        o_guids = set(re.findall(r'uid="{[^}]+}"', o_text))
                        e_guids = set(re.findall(r'uid="{[^}]+}"', e_text))
                        if o_guids != e_guids:
                            print(f"     GUIDs changed: {len(o_guids)} → {len(e_guids)}")
                            print(f"     Removed GUIDs: {o_guids - e_guids}" if o_guids - e_guids else "")
                            print(f"     New GUIDs: {e_guids - o_guids}" if e_guids - o_guids else "")
                        if len(o_text) > 200 and len(e_text) > 200:
                            print(f"     Orig first 200: {o_text[:200]}")
                            print(f"     Edit first 200: {e_text[:200]}")
                except Exception as ex:
                    print(f"     Parse error: {ex}")

        # ── 7. docProps comparison ──
        print("\n" + "─" * 70)
        print("  SECTION 7: DOCUMENT PROPERTIES (docProps)")
        print("─" * 70)

        for name in sorted(set(orig_entries.keys()) | set(edit_entries.keys())):
            if not name.startswith("docProps/"):
                continue
            oe = orig_entries.get(name)
            ee = edit_entries.get(name)
            if not oe or not ee:
                print(f"  [{name}] {'MISSING IN EDITED' if oe else 'NEW IN EDITED' if ee else ''}")
                continue

            orig_bytes = oz.read(name)
            edit_bytes = ez.read(name)
            if orig_bytes == edit_bytes:
                print(f"  ✅ {name}: IDENTICAL")
            else:
                orig_hash = sha256(orig_bytes)
                edit_hash = sha256(edit_bytes)
                print(f"  ❌ {name}: DIFFERENT ({len(orig_bytes):,} → {len(edit_bytes):,} bytes)")
                print(f"     Orig hash: {orig_hash}")
                print(f"     Edit hash: {edit_hash}")
                try:
                    o_text = orig_bytes.decode("utf-8", errors="replace")
                    e_text = edit_bytes.decode("utf-8", errors="replace")
                    if len(o_text) > 400:
                        print(f"     Orig first 400: {o_text[:400]}")
                        print(f"     Edit first 400: {e_text[:400]}")
                except Exception:
                    pass

        # ── 8. Shared Strings ──
        print("\n" + "─" * 70)
        print("  SECTION 6: SHARED STRINGS")
        print("─" * 70)

        try:
            if "xl/sharedStrings.xml" in orig_entries and "xl/sharedStrings.xml" in edit_entries:
                o_ss = ET.fromstring(oz.read("xl/sharedStrings.xml"))
                e_ss = ET.fromstring(ez.read("xl/sharedStrings.xml"))
                compare_shared_strings(o_ss, e_ss)
        except Exception as ex:
            print(f"  Error: {ex}")

        # ── 7. Comments ──
        print("\n" + "─" * 70)
        print("  SECTION 7: COMMENTS")
        print("─" * 70)

        for name in sorted(set(orig_entries.keys()) | set(edit_entries.keys())):
            if not name.startswith("xl/comments") or not name.endswith(".xml"):
                continue
            oe = orig_entries.get(name)
            ee = edit_entries.get(name)
            if not oe or not ee:
                print(f"\n  [{name}] {'MISSING IN EDITED' if oe else 'NEW IN EDITED' if ee else ''}")
                continue

            orig_xml = oz.read(name)
            edit_xml = ez.read(name)
            if orig_xml == edit_xml:
                print(f"  ✅ {name}: IDENTICAL")
            else:
                print(f"  ❌ {name}: DIFFERENT")
                try:
                    o_root = ET.fromstring(orig_xml)
                    e_root = ET.fromstring(edit_xml)
                    compare_comments(o_root, e_root, name)
                except Exception as ex:
                    print(f"     Parse error: {ex}")

    print("\n" + "=" * 70)
    print("  COMPARISON COMPLETE")
    print("=" * 70)


def compare_xml_elements(o_root, e_root, part_name, path=""):
    """Compare two XML trees element-by-element."""
    o_children = list(o_root)
    e_children = list(e_root)

    # Compare root attributes
    o_attrs = set(o_root.attrib.items())
    e_attrs = set(e_root.attrib.items())
    if o_attrs != e_attrs:
        print(f"     Root attributes differ at {path}:")
        for k, v in sorted(o_attrs - e_attrs):
            print(f"       Orig has: {k}={v}")
        for k, v in sorted(e_attrs - o_attrs):
            print(f"       Edit has: {k}={v}")

    # Compare child elements
    max_children = max(len(o_children), len(e_children))
    for i in range(max_children):
        o_child = o_children[i] if i < len(o_children) else None
        e_child = e_children[i] if i < len(e_children) else None
        o_tag = o_child.tag if o_child is not None else "(missing)"
        e_tag = e_child.tag if e_child is not None else "(missing)"
        child_path = f"{path}/{o_tag}" if o_child else f"{path}/{e_tag}"

        if o_child is None:
            print(f"     + [NEW] {e_tag}: first 200 chars = {ET.tostring(e_child, encoding='unicode')[:200]}")
            continue
        if e_child is None:
            print(f"     - [MISSING] {o_tag}: first 200 chars = {ET.tostring(o_child, encoding='unicode')[:200]}")
            continue

        o_text = (o_child.text or "").strip()
        e_text = (e_child.text or "").strip()
        o_attrs = sorted(o_child.attrib.items())
        e_attrs = sorted(e_child.attrib.items())

        o_xml = ET.tostring(o_child, encoding="unicode")
        e_xml = ET.tostring(e_child, encoding="unicode")

        if o_xml != e_xml:
            print(f"     ❌ [{i}] {child_path}:")
            if o_attrs != e_attrs:
                print(f"        Orig attributes: {o_attrs}")
                print(f"        Edit attributes: {e_attrs}")
            if o_text != e_text:
                print(f"        Orig text: '{o_text[:100]}'")
                print(f"        Edit text: '{e_text[:100]}'")
            print(f"        Orig XML: {o_xml[:300]}")
            print(f"        Edit XML: {e_xml[:300]}")

        # Recurse into children
        compare_xml_elements(o_child, e_child, part_name, child_path)


def compare_cells(o_root, e_root, sheet_name):
    """Compare cell values between two worksheet XMLs."""
    o_cells = {}
    e_cells = {}

    for c in o_root.iter(f"{S_NS}c"):
        ref = c.get("r", "")
        v = c.find(f"{S_NS}v")
        f = c.find(f"{S_NS}f")
        formula = f.text if f is not None else ""
        value = v.text if v is not None else ""
        style = c.get("s", "0")
        o_cells[ref] = (value, style, formula)

    for c in e_root.iter(f"{S_NS}c"):
        ref = c.get("r", "")
        v = c.find(f"{S_NS}v")
        f = c.find(f"{S_NS}f")
        formula = f.text if f is not None else ""
        value = v.text if v is not None else ""
        style = c.get("s", "0")
        e_cells[ref] = (value, style, formula)

    # Find differences
    all_refs = sorted(set(list(o_cells.keys()) + list(e_cells.keys())))
    value_changes = []
    style_changes = []
    formula_changes = []

    for ref in all_refs:
        o = o_cells.get(ref, ("", "0", ""))
        e = e_cells.get(ref, ("", "0", ""))

        if o[0] != e[0]:
            value_changes.append(f"        {ref}: '{o[0]}' → '{e[0]}'")
        if o[1] != e[1]:
            style_changes.append(f"        {ref}: style {o[1]} → {e[1]}")
        if o[2] != e[2]:
            formula_changes.append(f"        {ref}: formula '{o[2]}' → '{e[2]}'")

    missing_in_edit = [r for r in all_refs if r not in e_cells]
    new_in_edit = [r for r in all_refs if r not in o_cells]

    if value_changes:
        print(f"     Cell VALUE changes ({len(value_changes)}):")
        for c in value_changes[:30]:
            print(c)
        if len(value_changes) > 30:
            print(f"       ... and {len(value_changes) - 30} more")
    if style_changes:
        print(f"     Cell STYLE changes ({len(style_changes)}):")
        for c in style_changes[:20]:
            print(c)
    if formula_changes:
        print(f"     Cell FORMULA changes ({len(formula_changes)}):")
        for c in formula_changes:
            print(c)
    if missing_in_edit:
        print(f"     Cells MISSING in edited ({len(missing_in_edit)}): {missing_in_edit[:20]}")
    if new_in_edit:
        print(f"     Cells NEW in edited ({len(new_in_edit)}): {new_in_edit[:20]}")


def compare_sheet_structure(o_root, e_root, sheet_name):
    """Compare sheet structure: merges, page setup, margins."""
    # Merges
    o_merges = o_root.find(f"{S_NS}mergeCells")
    e_merges = e_root.find(f"{S_NS}mergeCells")
    if o_merges is not None or e_merges is not None:
        o_refs = sorted([m.get("ref", "") for m in (o_merges or [])])
        e_refs = sorted([m.get("ref", "") for m in (e_merges or [])])
        if o_refs != e_refs:
            print(f"     MERGE CELLS DIFFER:")
            print(f"       Orig: {o_refs}")
            print(f"       Edit: {e_refs}")

    # PageSetup
    o_ps = o_root.find(f"{S_NS}pageSetup")
    e_ps = e_root.find(f"{S_NS}pageSetup")
    if o_ps is not None or e_ps is not None:
        o_attrs = dict(sorted(o_ps.attrib.items())) if o_ps is not None else {}
        e_attrs = dict(sorted(e_ps.attrib.items())) if e_ps is not None else {}
        if o_attrs != e_attrs:
            print(f"     PAGESETUP DIFFERS:")
            for k in set(list(o_attrs.keys()) + list(e_attrs.keys())):
                ov = o_attrs.get(k, "(missing)")
                ev = e_attrs.get(k, "(missing)")
                if ov != ev:
                    print(f"       {k}: '{ov}' → '{ev}'")

    # PrintOptions
    o_po = o_root.find(f"{S_NS}printOptions")
    e_po = e_root.find(f"{S_NS}printOptions")
    if o_po is not None or e_po is not None:
        o_attrs = dict(sorted(o_po.attrib.items())) if o_po is not None else {}
        e_attrs = dict(sorted(e_po.attrib.items())) if e_po is not None else {}
        if o_attrs != e_attrs:
            print(f"     PRINTOPTIONS DIFFERS:")
            for k in set(list(o_attrs.keys()) + list(e_attrs.keys())):
                ov = o_attrs.get(k, "(missing)")
                ev = e_attrs.get(k, "(missing)")
                if ov != ev:
                    print(f"       {k}: '{ov}' → '{ev}'")

    # PageMargins
    o_pm = o_root.find(f"{S_NS}pageMargins")
    e_pm = e_root.find(f"{S_NS}pageMargins")
    if o_pm is not None or e_pm is not None:
        o_attrs = dict(sorted(o_pm.attrib.items())) if o_pm is not None else {}
        e_attrs = dict(sorted(e_pm.attrib.items())) if e_pm is not None else {}
        if o_attrs != e_attrs:
            print(f"     PAGEMARGINS DIFFERS:")
            for k in set(list(o_attrs.keys()) + list(e_attrs.keys())):
                ov = o_attrs.get(k, "(missing)")
                ev = e_attrs.get(k, "(missing)")
                if ov != ev:
                    print(f"       {k}: '{ov}' → '{ev}'")


def compare_relationships(o_root, e_root, name):
    """Compare relationship XMLs."""
    o_rels = {}
    e_rels = {}
    rel_ns = o_root.tag if o_root.tag.startswith("{") else RELS_NS

    for rel in o_root:
        rid = rel.get("Id", "")
        target = rel.get("Target", "")
        rel_type = rel.get("Type", "")
        o_rels[rid] = (target, rel_type)

    for rel in e_root:
        rid = rel.get("Id", "")
        target = rel.get("Target", "")
        rel_type = rel.get("Type", "")
        e_rels[rid] = (target, rel_type)

    missing = set(o_rels.keys()) - set(e_rels.keys())
    new = set(e_rels.keys()) - set(o_rels.keys())
    common = set(o_rels.keys()) & set(e_rels.keys())

    if missing:
        for rid in sorted(missing):
            t, r = o_rels[rid]
            print(f"     - [REMOVED] {rid}: {t} ({r})")
    if new:
        for rid in sorted(new):
            t, r = e_rels[rid]
            print(f"     + [ADDED] {rid}: {t} ({r})")
    for rid in sorted(common):
        ot, _ = o_rels[rid]
        et, _ = e_rels[rid]
        if ot != et:
            print(f"     ~ [CHANGED] {rid}: '{ot}' → '{et}'")

    if not missing and not new:
        print(f"     ✅ Relationships match: {len(common)} relationships, all unchanged.")


def compare_defined_names(o_wb, e_wb):
    """Compare defined names between two workbook XMLs."""
    o_names = {}
    e_names = {}
    for dn in o_wb.iter(f"{S_NS}definedName"):
        name = dn.get("name", "")
        local = dn.get("localSheetId")
        text = (dn.text or "").strip()
        o_names[name] = (text, local)

    for dn in e_wb.iter(f"{S_NS}definedName"):
        name = dn.get("name", "")
        local = dn.get("localSheetId")
        text = (dn.text or "").strip()
        e_names[name] = (text, local)

    missing = sorted(set(o_names.keys()) - set(e_names.keys()))
    new = sorted(set(e_names.keys()) - set(o_names.keys()))
    common = sorted(set(o_names.keys()) & set(e_names.keys()))

    if missing:
        print(f"  DefinedNames MISSING in edited ({len(missing)}):")
        for n in missing:
            t, l = o_names[n]
            print(f"    - '{n}': {t}")
    if new:
        print(f"  DefinedNames NEW in edited ({len(new)}):")
        for n in new:
            t, l = e_names[n]
            print(f"    + '{n}': {t}")
    if not missing and not new:
        changed = [n for n in common if o_names[n] != e_names[n]]
        if changed:
            print(f"  DefinedNames changed ({len(changed)}):")
            for n in changed:
                ot, ol = o_names[n]
                et, el = e_names[n]
                print(f"    ~ '{n}': '{ot}' → '{et}'")
        else:
            print(f"  ✅ DefinedNames: {len(common)} names, all unchanged.")


def compare_shared_strings(o_ss, e_ss):
    """Compare shared strings tables."""
    o_items = [si for si in o_ss.iter(f"{S_NS}si")]
    e_items = [si for si in e_ss.iter(f"{S_NS}si")]

    print(f"  Original count: {len(o_items)}")
    print(f"  Edited count:   {len(e_items)}")

    # Compare common entries
    min_count = min(len(o_items), len(e_items))
    diffs = 0
    for i in range(min_count):
        o_xml = ET.tostring(o_items[i], encoding="unicode")
        e_xml = ET.tostring(e_items[i], encoding="unicode")
        if o_xml != e_xml:
            diffs += 1
            if diffs <= 5:
                o_txt = "".join(o_items[i].itertext())[:80]
                e_txt = "".join(e_items[i].itertext())[:80]
                print(f"    ❌ Entry #{i}:")
                print(f"       Orig: '{o_txt}'")
                print(f"       Edit: '{e_txt}'")

    if diffs == 0 and len(o_items) == len(e_items):
        print(f"  ✅ All {min_count} entries match.")
    elif diffs == 0:
        print(f"  ✅ First {min_count} entries match. {abs(len(e_items) - len(o_items))} entries added/removed.")
    else:
        print(f"  ❌ {diffs} differences in first {min_count} entries.")

    # Show appended entries
    if len(e_items) > len(o_items):
        print(f"  New entries appended ({len(e_items) - len(o_items)}):")
        for i in range(len(o_items), len(e_items)):
            txt = "".join(e_items[i].itertext())[:120]
            print(f"    [+{i}] '{txt}'")
    elif len(o_items) > len(e_items):
        print(f"  Entries removed ({len(o_items) - len(e_items)}): last {len(o_items) - len(e_items)} missing")


def compare_comments(o_root, e_root, name):
    """Compare comment XMLs — full OuterXml comparison, not just text."""
    o_comments = {}
    e_comments = {}
    for c in o_root.iter(f"{S_NS}comment"):
        ref = c.get("ref", "")
        full_xml = ET.tostring(c, encoding="unicode")
        text = "".join(c.itertext())[:100]
        o_comments[ref] = (full_xml, text)
    for c in e_root.iter(f"{S_NS}comment"):
        ref = c.get("ref", "")
        full_xml = ET.tostring(c, encoding="unicode")
        text = "".join(c.itertext())[:100]
        e_comments[ref] = (full_xml, text)

    missing = sorted(set(o_comments.keys()) - set(e_comments.keys()))
    new = sorted(set(e_comments.keys()) - set(o_comments.keys()))
    common = sorted(set(o_comments.keys()) & set(e_comments.keys()))

    if missing:
        print(f"    Comments MISSING in edited ({len(missing)}):")
        for r in missing:
            _, txt = o_comments[r]
            print(f"      - {r}: '{txt}'")
    if new:
        print(f"    Comments NEW in edited ({len(new)}):")
        for r in new:
            _, txt = e_comments[r]
            print(f"      + {r}: '{txt}'")
    if common:
        changed = [r for r in common if o_comments[r][0] != e_comments[r][0]]
        text_changed = [r for r in common if o_comments[r][1] != e_comments[r][1]]
        if changed:
            print(f"    Comments XML CHANGED ({len(changed)}):")
            for r in changed[:5]:
                o_xml = o_comments[r][0]
                e_xml = e_comments[r][0]
                print(f"      ~ {r}:")
                print(f"        Orig XML: {o_xml[:300]}")
                print(f"        Edit XML: {e_xml[:300]}")
                if r in text_changed:
                    print(f"        Text also changed: '{o_comments[r][1]}' → '{e_comments[r][1]}'")
        else:
            print(f"    ✅ {len(common)} comments, full XML match. All unchanged.")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python forensic_compare_workbooks.py <original.xlsx> <edited.xlsx> [--verbose]")
        sys.exit(1)

    orig_path = sys.argv[1]
    edit_path = sys.argv[2]
    verbose = "--verbose" in sys.argv

    if not os.path.exists(orig_path):
        print(f"Error: Original file not found: {orig_path}")
        sys.exit(1)
    if not os.path.exists(edit_path):
        print(f"Error: Edited file not found: {edit_path}")
        sys.exit(1)

    compare_workbooks(orig_path, edit_path, verbose)
