#!/usr/bin/env python3
"""
Phase 5.3 — Forensic OOXML Validation
Inspects a generated workbook ZIP and identifies every structural issue
that would cause Microsoft Excel to enter repair mode.
"""

import sys
import os
from zipfile import ZipFile, BadZipFile
from xml.etree import ElementTree as ET
from collections import defaultdict
import hashlib


NCX = {
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
}


def ns_attr(tag, ns_prefix="s"):
    """Return a namespace-prefixed attribute name, e.g. {ns}s:sheet."""
    uri = NCX.get(ns_prefix, ns_prefix)
    return f"{{{uri}}}{tag}"


def ns_tag(tag, ns_prefix="s"):
    """Return a namespace-prefixed tag name."""
    uri = NCX.get(ns_prefix, ns_prefix)
    return f"{{{uri}}}{tag}"


def find(xml, tag):
    """Find first matching child with namespace."""
    return xml.find(ns_tag(tag))


def findall(xml, tag):
    """Find all matching children with namespace."""
    return xml.findall(ns_tag(tag))


def parse_xml(zipfile, path):
    """Parse an XML file from the ZIP, return ElementTree root or None."""
    try:
        with zipfile.open(path) as f:
            return ET.parse(f).getroot()
    except (KeyError, ET.ParseError, BadZipFile) as e:
        return None


def get_rels(zipfile, xml_path):
    """Get relationships for a given XML part. Returns dict of {Id: {Target, Type}}."""
    rels_path = f"{os.path.dirname(xml_path)}/_rels/{os.path.basename(xml_path)}.rels"
    try:
        with zipfile.open(rels_path) as f:
            root = ET.parse(f).getroot()
        rels = {}
        for rel in root:
            rid = rel.get("Id")
            target = rel.get("Target")
            rtype = rel.get("Type")
            rels[rid] = {"Target": target, "Type": rtype}
        return rels
    except (KeyError, ET.ParseError, BadZipFile) as e:
        return {}


def validate_content_types(zipfile, report):
    """Check [Content_Types].xml has all required Override entries."""
    ct = parse_xml(zipfile, "[Content_Types].xml")
    if ct is None:
        report.append("[CRITICAL] [Content_Types].xml missing or unparseable")
        return set()

    overrides = {}
    for ov in findall(ct, "Override"):
        pn = ov.get("PartName", "")
        ct_val = ov.get("ContentType", "")
        overrides[pn] = ct_val

    defaults = {}
    for d in findall(ct, "Default"):
        ext = d.get("Extension", "")
        ct_val = d.get("ContentType", "")
        defaults[ext] = ct_val

    # Check essential defaults
    for ext in ["xml", "rels"]:
        if ext not in defaults:
            report.append(f"[WARNING] Missing Default ContentType for .{ext}")

    # Check all .xml entries in the zip have Overrides
    all_parts_in_zip = [n for n in zipfile.namelist()
                        if n.endswith(".xml") and not n.startswith("_rels") and "/_rels/" not in n]

    for part in all_parts_in_zip:
        pn = f"/{part}" if not part.startswith("/") else part
        if pn not in overrides:
            report.append(f"[MISSING] No Override in [Content_Types].xml for: {part}")

    # Validate worksheet override names match actual sheets
    ws_override_parts = {k for k in overrides
                         if "worksheet" in overrides[k]}

    return overrides


def validate_workbook_rels(zipfile, report):
    """Check xl/_rels/workbook.xml.rels has all required relationships."""
    rels = get_rels(zipfile, "xl/workbook.xml")
    if not rels:
        report.append("[CRITICAL] xl/_rels/workbook.xml.rels missing or empty")

    rel_targets = {}
    for rid, info in rels.items():
        rel_targets[rid] = info
        # Check target exists
        tgt = info["Target"]
        if not tgt.startswith("/"):
            tgt = f"xl/{tgt}"
        if tgt not in zipfile.namelist():
            report.append(f"[MISSING] Relationship {rid} targets {tgt} but file does not exist in ZIP")

    return rel_targets


def validate_workbook_xml(zipfile, report):
    """Check xl/workbook.xml for consistency."""
    wb = parse_xml(zipfile, "xl/workbook.xml")
    if wb is None:
        report.append("[CRITICAL] xl/workbook.xml missing or unparseable")
        return [], {}, {}

    sheets_elem = find(wb, "sheets")
    sheets = []
    if sheets_elem is not None:
        for s in findall(sheets_elem, "sheet"):
            name = s.get("name", "?")
            sid = s.get("sheetId", "?")
            rid = s.get(ns_attr("id", "r"), "")
            state = s.get("state", "visible")
            sheets.append({"name": name, "sheetId": sid, "rId": rid, "state": state})

    # Check each sheet has a corresponding relationship
    rels = get_rels(zipfile, "xl/workbook.xml")

    defined_names_elem = find(wb, "definedNames")
    defined_names = []
    if defined_names_elem is not None:
        for dn in findall(defined_names_elem, "definedName"):
            dn_name = dn.get("name", "?")
            dn_text = dn.text or ""
            defined_names.append({"name": dn_name, "text": dn_text.strip()})

    return sheets, rels, defined_names


def validate_worksheet_xml(zipfile, sheet_name, sheet_path, report):
    """Validate a worksheet XML file."""
    ws = parse_xml(zipfile, sheet_path)
    if ws is None:
        report.append(f"[CRITICAL] Worksheet at {sheet_path} is missing or unparseable")
        return

    # Check for legacyDrawing (VML comments)
    legacy_drawing = find(ws, "legacyDrawing")
    if legacy_drawing is not None:
        rid = legacy_drawing.get(ns_attr("id", "r"), "")
        rels = get_rels(zipfile, sheet_path)
        if rid and rid not in rels:
            report.append(f"[BROKEN] Sheet '{sheet_name}' ({sheet_path}) has legacyDrawing rId={rid} but no matching relationship in rels")

    # Check for drawing
    drawing = find(ws, "drawing")
    if drawing is not None:
        rid = drawing.get(ns_attr("id", "r"), "")
        rels = get_rels(zipfile, sheet_path)
        if rid and rid not in rels:
            report.append(f"[BROKEN] Sheet '{sheet_name}' ({sheet_path}) has drawing rId={rid} but no matching relationship")


def check_sheet_rels(zipfile, rels, report):
    """For each sheet rel, check that the target worksheet has matching sheet-level rels."""
    workbook_rels_path = "xl/_rels/workbook.xml.rels"
    wb_rels = get_rels(zipfile, "xl/workbook.xml")

    sheet_rel_ids = {}
    for rid, info in wb_rels.items():
        tgt = info["Target"]
        if tgt.startswith("worksheets/"):
            sheet_rel_ids[rid] = tgt

    # For each worksheet relationship, check its own .rels file
    for rid, tgt in sheet_rel_ids.items():
        ws_rels_path = f"xl/worksheets/_rels/{os.path.basename(tgt)}.rels"
        ws_rels = get_rels(zipfile, f"xl/{tgt}")
        if ws_rels is None:
            continue  # no rels file needed


def main():
    if len(sys.argv) < 2:
        print("Usage: python validate_forensic.py <generated.xlsx> [original.xlsx] [conmas.xlsx]")
        sys.exit(1)

    generated_path = sys.argv[1]
    original_path = sys.argv[2] if len(sys.argv) > 2 else None
    conmas_path = sys.argv[3] if len(sys.argv) > 3 else None

    report = []
    report.append("=" * 70)
    report.append("PAPERLESS FORENSIC OOXML VALIDATION REPORT")
    report.append("=" * 70)
    report.append("")

    # 1. Validate generated workbook
    report.append("─" * 50)
    report.append("FILE: " + generated_path)
    report.append("─" * 50)

    try:
        with ZipFile(generated_path, "r") as gen_zf:
            entries = gen_zf.namelist()
            report.append(f"ZIP entries: {len(entries)}")
            report.append("")

            # --- [Content_Types].xml ---
            report.append("─── [1] [Content_Types].xml Validation ───")
            cts = validate_content_types(gen_zf, report)

            # --- workbook.xml ---
            report.append("")
            report.append("─── [2] workbook.xml Validation ───")
            sheets, wb_rels, defined_names = validate_workbook_xml(gen_zf, report)

            report.append(f"  Sheets defined: {len(sheets)}")
            for s in sheets:
                report.append(f"    name='{s['name']}' sheetId={s['sheetId']} rId={s['rId']} state={s['state']}")

            if defined_names:
                report.append(f"  Defined names: {len(defined_names)}")
                for dn in defined_names:
                    report.append(f"    {dn['name']} = {dn['text'][:80]}")

            # Check each sheet against ZIP entries
            report.append("")
            report.append("─── [3] Sheet Relationship Validation ───")
            for s in sheets:
                rid = s["rId"]
                if rid in wb_rels:
                    target = wb_rels[rid]["Target"]
                    full_path = f"xl/{target}" if not target.startswith("/") else target
                    exists = full_path in entries
                    report.append(f"  Sheet '{s['name']}' ({rid}): rel->{target} -> {'EXISTS' if exists else 'MISSING!'}")
                    if exists:
                        validate_worksheet_xml(gen_zf, s["name"], full_path, report)
                else:
                    report.append(f"  [BROKEN] Sheet '{s['name']}' ({rid}): NO MATCHING RELATIONSHIP IN workbook.xml.rels!")

            # Check for worksheet entries in ZIP not referenced in workbook.xml
            sheet_xmls = [e for e in entries if e.startswith("xl/worksheets/sheet") and e.endswith(".xml")]
            referenced_sheets = set()
            for rid, info in wb_rels.items():
                tgt = info["Target"]
                if tgt.startswith("worksheets/"):
                    referenced_sheets.add(f"xl/{tgt}")
            for sxml in sheet_xmls:
                if sxml not in referenced_sheets:
                    report.append(f"  [ORPHAN] Worksheet {sxml} exists in ZIP but is NOT referenced in workbook.xml.rels!")

            # --- workbook.xml.rels ---
            report.append("")
            report.append("─── [4] xl/_rels/workbook.xml.rels Validation ───")
            workbook_rel_path = "xl/_rels/workbook.xml.rels"
            workbook_rels = get_rels(gen_zf, "xl/workbook.xml")
            if workbook_rels:
                for rid, info in sorted(workbook_rels.items()):
                    tgt = info["Target"]
                    full_path = f"xl/{tgt}" if not tgt.startswith("/") else tgt
                    exists = full_path in entries
                    icon = "EXISTS" if exists else "MISSING!"
                    report.append(f"    {rid}: Type='...{info['Type'][:50]}' Target='{tgt}' -> {icon}")
            else:
                report.append("  [CRITICAL] workbook.xml.rels not found or empty")

            # --- Shared Strings ---
            report.append("")
            report.append("─── [5] SharedStrings Validation ───")
            ss_entry = [e for e in entries if "sharedstrings" in e.lower()]
            if ss_entry:
                ss_path = ss_entry[0]
                ss_xml = parse_xml(gen_zf, ss_path)
                if ss_xml is not None:
                    si_count = len(findall(ss_xml, "si"))
                    report.append(f"  SharedStringTable: {si_count} items ({ss_path})")
                else:
                    report.append(f"  [CRITICAL] {ss_path} unparseable!")
            else:
                report.append("  [MISSING] No sharedStrings.xml found in ZIP!")

            # --- Comments ---
            report.append("")
            report.append("─── [6] Comments & VML Validation ───")
            comment_entries = [e for e in entries if "comments" in e.lower()]
            vml_entries = [e for e in entries if "vml" in e.lower() or "drawing" in e.lower()]
            if comment_entries:
                report.append(f"  Comments parts found: {comment_entries}")
                for ce in comment_entries:
                    # Check if matching VML exists
                    sheet_idx = None
                    for s in sheets:
                        rid = s["rId"]
                        if rid in wb_rels:
                            target = wb_rels[rid]["Target"]
                            sheet_num = target.replace("worksheets/sheet", "").replace(".xml", "")
                            comment_num = ce.replace("xl/comments", "").replace(".xml", "")
                            if sheet_num == comment_num:
                                sheet_idx = sheet_num
                                break
                    if sheet_idx:
                        # Check for VML drawing
                        ws_path = f"xl/worksheets/sheet{sheet_idx}.xml"
                        ws_xml = parse_xml(gen_zf, ws_path)
                        if ws_xml is not None:
                            leg_drawing = find(ws_xml, "legacyDrawing")
                            if leg_drawing is None:
                                report.append(f"  [CORRUPTION] Comments part {ce} exists but sheet{sheet_idx}.xml has NO legacyDrawing reference!")
                            else:
                                # Check that VML file exists
                                ws_rels = get_rels(gen_zf, ws_path)
                                rid = leg_drawing.get(ns_attr("id", "r"), "")
                                if rid and rid in ws_rels:
                                    vml_target = ws_rels[rid]["Target"]
                                    vml_path = f"xl/worksheets/_rels/{vml_target}"
                                    vml_full = f"xl/worksheets/{vml_target}"
                                    if vml_full in entries:
                                        report.append(f"  VML drawing {vml_full} EXISTS for comment {ce} ✅")
                                    else:
                                        report.append(f"  [CORRUPTION] Comments part {ce} references VML via legacyDrawing rId={rid} target='{vml_target}' but VML file NOT FOUND in ZIP!")
                                else:
                                    report.append(f"  [CORRUPTION] Comments part {ce} has legacyDrawing rId={rid} but no matching relationship in worksheet rels!")
            else:
                report.append("  No comments parts found")
            if vml_entries:
                report.append(f"  VML/drawing parts found: {vml_entries}")
            else:
                report.append("  No VML/drawing parts found")

            # --- Styles ---
            report.append("")
            report.append("─── [7] Styles Validation ───")
            style_entries = [e for e in entries if e.endswith("styles.xml")]
            if style_entries:
                for se in style_entries:
                    sx = parse_xml(gen_zf, se)
                    if sx is not None:
                        fonts = len(findall(sx, "fonts"))
                        fills = len(findall(sx, "fills"))
                        borders = len(findall(sx, "borders"))
                        cellxf = len(findall(sx, "cellXfs"))
                        report.append(f"  {se}: {fonts} fonts, {fills} fills, {borders} borders, {cellxf} cellXfs")
                    else:
                        report.append(f"  [CRITICAL] {se} unparseable!")
            else:
                report.append("  [MISSING] No styles.xml found!")

    except BadZipFile:
        report.append("[CRITICAL] File is not a valid ZIP/OOXML file!")
    except FileNotFoundError:
        report.append(f"[CRITICAL] File not found: {generated_path}")

    # 2. Compare with original if provided
    if original_path:
        report.append("")
        report.append("=" * 70)
        report.append("COMPARISON WITH ORIGINAL")
        report.append("=" * 70)
        try:
            with ZipFile(original_path, "r") as orig_zf:
                orig_entries = set(orig_zf.namelist())
                report.append(f"Original ZIP entries: {len(orig_entries)}")

                # Compare workbook.xml sheets
                orig_sheets, _, _ = validate_workbook_xml(orig_zf, [])
                report.append(f"Original sheets: {len(orig_sheets)}")
                for s in orig_sheets:
                    report.append(f"  name='{s['name']}' sheetId={s['sheetId']} state={s['state']}")

                # Check for comments in original
                orig_comments = [e for e in orig_entries if "comments" in e.lower()]
                orig_vml = [e for e in orig_entries if "vml" in e.lower()]
                report.append(f"Original comments: {orig_comments}")
                report.append(f"Original VML: {orig_vml}")

                # Compare all entries
                gen_entries = set(ZipFile(generated_path, "r").namelist())
                added = gen_entries - orig_entries
                removed = orig_entries - gen_entries
                if added:
                    report.append(f"Entries ADDED in generated: {sorted(added)}")
                if removed:
                    report.append(f"Entries REMOVED from generated: {sorted(removed)}")

        except (BadZipFile, FileNotFoundError) as e:
            report.append(f"Cannot compare original: {e}")

    report.append("")
    report.append("=" * 70)
    report.append("VALIDATION COMPLETE")
    report.append("=" * 70)

    # Print report — write to file to avoid encoding issues
    report_path = generated_path + ".validation.log"
    with open(report_path, "w", encoding="utf-8") as f:
        for line in report:
            f.write(line + "\n")
    print(f"Report written to: {report_path}")
    # Also print to stdout with safe encoding
    for line in report:
        try:
            print(line)
        except UnicodeEncodeError:
            print(line.encode('ascii', errors='replace').decode('ascii'))

    # Return non-zero if critical issues found
    critical_issues = [l for l in report if "[CRITICAL]" in l]
    corruption_issues = [l for l in report if "[CORRUPTION]" in l]
    broken_issues = [l for l in report if "[BROKEN]" in l]
    missing_issues = [l for l in report if "[MISSING]" in l]

    if critical_issues:
        print(f"\n❌ {len(critical_issues)} CRITICAL issues, {len(corruption_issues)} CORRUPTION issues, {len(broken_issues)} BROKEN issues, {len(missing_issues)} MISSING issues")
        return 1
    if corruption_issues or broken_issues:
        print(f"\n⚠️  {len(critical_issues)} CRITICAL, {len(corruption_issues)} CORRUPTION, {len(broken_issues)} BROKEN, {len(missing_issues)} MISSING issues")
        return 1
    print(f"\n✅ No critical issues found ({len(missing_issues)} warnings)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
