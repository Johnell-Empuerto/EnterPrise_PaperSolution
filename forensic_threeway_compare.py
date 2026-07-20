#!/usr/bin/env python3
"""
Phase 5.5.3 — Three-Way Forensic OOXML Comparison

Compares:
  1. Original template workbook
  2. Legacy ConMas output (known good)
  3. Current PaperLess output (corrupted)

Identifies the EXACT first structural difference that causes Excel repair.
"""

import sys
import os
import hashlib
from zipfile import ZipFile, BadZipFile
from xml.etree import ElementTree as ET
from collections import defaultdict
from difflib import HtmlDiff
import xml.parsers.expat

NCX = {
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    "dc": "http://purl.org/dc/elements/1.1/",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dcterms": "http://purl.org/dc/terms/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

def ns_attr(tag, ns_prefix="s"):
    uri = NCX.get(ns_prefix, ns_prefix)
    return f"{{{uri}}}{tag}"

def ns_tag(tag, ns_prefix="s"):
    uri = NCX.get(ns_prefix, ns_prefix)
    return f"{{{uri}}}{tag}"

def find(xml, tag):
    return xml.find(ns_tag(tag))

def findall(xml, tag):
    return xml.findall(ns_tag(tag))

def sha256_file(data):
    return hashlib.sha256(data).hexdigest()

def parse_xml_safe(data):
    """Parse XML bytes, return root or None with error."""
    try:
        return ET.fromstring(data)
    except ET.ParseError as e:
        return None

def format_xml(node, indent=0):
    """Simple XML pretty-printer for comparison output."""
    if node is None:
        return "(null)"
    if len(node) == 0:
        # Leaf node
        tag = node.tag.split('}')[-1] if '}' in node.tag else node.tag
        attrs = " ".join(f'{k.split("}")[-1]}="{v}"' for k, v in sorted(node.attrib.items()))
        text = (node.text or "").strip()[:40]
        if attrs:
            return f"<{tag} {attrs}>{text}</{tag}>"
        else:
            return f"<{tag}>{text}</{tag}>"
    return f"<{node.tag.split('}')[-1]}>...({len(node)} children)"

class WorkbookForensicAnalyzer:
    def __init__(self, path, label):
        self.path = path
        self.label = label
        self.entries = {}
        self.content_types = None
        self.workbook_xml = None
        self.workbook_rels = {}
        self.sheets = []
        self.defined_names = []
        self.comments = {}
        self.vml_drawings = {}
        self.worksheet_data = {}
        self.worksheet_rels = {}
        self.styles = None
        self.shared_strings = None
        self.relationships_graph = {}  # part -> {rels}
        self.orphan_parts = set()
        self.unreferenced_parts = set()
        
    def load(self):
        try:
            with ZipFile(self.path, 'r') as zf:
                self.entries = {name: zf.read(name) for name in zf.namelist()}
        except (BadZipFile, FileNotFoundError) as e:
            print(f"  ERROR loading {self.path}: {e}")
            return False
        
        # Parse content types
        ct_data = self.entries.get("[Content_Types].xml")
        if ct_data:
            self.content_types = parse_xml_safe(ct_data)
        
        # Parse workbook.xml
        wb_data = self.entries.get("xl/workbook.xml")
        if wb_data:
            self.workbook_xml = parse_xml_safe(wb_data)
            if self.workbook_xml is not None:
                self._extract_sheets()
                self._extract_defined_names()
        
        # Parse workbook.xml.rels
        wb_rels_data = self.entries.get("xl/_rels/workbook.xml.rels")
        if wb_rels_data:
            self.workbook_rels = self._parse_rels(wb_rels_data)
        
        # Parse all worksheets
        for name, data in self.entries.items():
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                ws = parse_xml_safe(data)
                self.worksheet_data[name] = ws
            
            # Parse worksheet rels
            if name.startswith("xl/worksheets/_rels/sheet") and name.endswith(".xml.rels"):
                self.worksheet_rels[name] = self._parse_rels(data)
            
            # Parse comments
            if name.startswith("xl/comments") and name.endswith(".xml"):
                self.comments[name] = parse_xml_safe(data)
            
            # Parse VML
            if "vml" in name.lower() and name.endswith(".vml"):
                try:
                    self.vml_drawings[name] = data.decode('utf-8', errors='replace')
                except:
                    self.vml_drawings[name] = "(binary or unparseable)"
        
        # Parse styles
        styles_data = self.entries.get("xl/styles.xml")
        if styles_data:
            self.styles = parse_xml_safe(styles_data)
        
        # Parse shared strings
        ss_data = self.entries.get("xl/sharedStrings.xml")
        if ss_data:
            self.shared_strings = parse_xml_safe(ss_data)
        
        # Build relationship graph
        self._build_relationship_graph()
        
        return True
    
    def _parse_rels(self, data):
        """Parse .rels file into dict of {Id: {Target, Type}}."""
        root = parse_xml_safe(data)
        if root is None:
            return {}
        rels = {}
        rel_ns = NCX["rel"]
        for rel in root:
            rid = rel.get("Id")
            target = rel.get("Target")
            rtype = rel.get("Type")
            rels[rid] = {"Target": target, "Type": rtype}
        return rels
    
    def _extract_sheets(self):
        wb = self.workbook_xml
        if wb is None:
            return
        sheets_elem = find(wb, "sheets")
        if sheets_elem is None:
            return
        for s in findall(sheets_elem, "sheet"):
            name = s.get("name", "?")
            sid = s.get("sheetId", "?")
            rid = s.get(ns_attr("id", "r"), "")
            state = s.get("state", "visible")
            self.sheets.append({"name": name, "sheetId": sid, "rId": rid, "state": state})
    
    def _extract_defined_names(self):
        wb = self.workbook_xml
        if wb is None:
            return
        dn_elem = find(wb, "definedNames")
        if dn_elem is None:
            return
        for dn in findall(dn_elem, "definedName"):
            dn_name = dn.get("name", "?")
            dn_text = (dn.text or "").strip()
            self.defined_names.append({"name": dn_name, "text": dn_text})
    
    def _build_relationship_graph(self):
        """Build complete relationship graph for all parts."""
        # Start with _rels/.rels
        dot_rels = self.entries.get("_rels/.rels")
        if dot_rels:
            rels = self._parse_rels(dot_rels)
            self.relationships_graph["_rels/.rels"] = rels
            for rid, info in rels.items():
                tgt = info["Target"]
                full = f"/{tgt}" if not tgt.startswith("/") else tgt
                if full.startswith("/"):
                    full = full[1:]
                if full not in self.entries:
                    self.orphan_parts.add(f"  _rels/.rels -> {rid} -> {full} (MISSING)")
                else:
                    self.unreferenced_parts.discard(full)
        
        # Mark all known parts as potentially unreferenced initially
        all_parts = set(self.entries.keys())
        self.unreferenced_parts = set(all_parts)
        # _rels files don't need to be referenced by other parts
        self.unreferenced_parts.discard("_rels/.rels")
        self.unreferenced_parts.discard("[Content_Types].xml")
        
        # Process workbook.xml.rels
        wb_rels = self.workbook_rels
        self.relationships_graph["xl/_rels/workbook.xml.rels"] = wb_rels
        for rid, info in wb_rels.items():
            tgt = info["Target"]
            full = f"xl/{tgt}" if not tgt.startswith("/") else tgt[1:]
            if full in self.entries:
                self.unreferenced_parts.discard(full)
            else:
                self.orphan_parts.add(f"  xl/_rels/workbook.xml.rels -> {rid} -> {full} (MISSING)")
        
        # Process each worksheet's rels
        for rels_name, rels in self.worksheet_rels.items():
            self.relationships_graph[rels_name] = rels
            for rid, info in rels.items():
                tgt = info["Target"]
                ws_dir = os.path.dirname(rels_name.replace("_rels/", ""))
                full = os.path.normpath(f"{ws_dir}/{tgt}").replace("\\", "/")
                if full in self.entries:
                    self.unreferenced_parts.discard(full)
                else:
                    self.orphan_parts.add(f"  {rels_name} -> {rid} -> {full} (MISSING)")
        
        # Remove known metadata files from unreferenced check
        for skip in ["_rels/.rels", "[Content_Types].xml", "docProps/core.xml", "docProps/app.xml"]:
            self.unreferenced_parts.discard(skip)
    
    def dump_structure(self):
        """Dump complete workbook structure for comparison."""
        lines = []
        lines.append(f"=== {self.label} ===")
        lines.append(f"Path: {self.path}")
        lines.append(f"ZIP entries: {len(self.entries)}")
        lines.append("")
        
        # All ZIP entries
        lines.append("--- ALL ZIP ENTRIES ---")
        for name in sorted(self.entries.keys()):
            data = self.entries[name]
            size = len(data)
            sha = sha256_file(data)
            lines.append(f"  [{size:8d}B] {sha[:12]} {name}")
        lines.append("")
        
        # Content Types
        lines.append("--- [Content_Types].xml ---")
        if self.content_types is not None:
            for ov in findall(self.content_types, "Override"):
                pn = ov.get("PartName", "")
                ct = ov.get("ContentType", "")
                lines.append(f"  Override: {pn}  -> {ct[:60]}")
            for d in findall(self.content_types, "Default"):
                ext = d.get("Extension", "")
                ct = d.get("ContentType", "")
                lines.append(f"  Default: .{ext}  -> {ct}")
        else:
            lines.append("  (unparseable)")
        lines.append("")
        
        # Workbook sheets
        lines.append("--- workbook.xml: SHEETS ---")
        for s in self.sheets:
            lines.append(f"  name='{s['name']}' sheetId={s['sheetId']} rId={s['rId']} state={s['state']}")
        if self.defined_names:
            lines.append("--- workbook.xml: DEFINED NAMES ---")
            for dn in self.defined_names:
                lines.append(f"  {dn['name']} = {dn['text'][:80]}")
        lines.append("")
        
        # Workbook rels
        lines.append("--- xl/_rels/workbook.xml.rels ---")
        for rid, info in sorted(self.workbook_rels.items()):
            typ = info["Type"].split("/")[-1]
            lines.append(f"  {rid}: Type={typ} Target={info['Target']}")
        lines.append("")
        
        # Worksheets
        lines.append("--- WORKSHEETS ---")
        for ws_name, ws in sorted(self.worksheet_data.items()):
            sheet_num = ws_name.split("sheet")[1].split(".")[0]
            # Get sheet name from workbook
            sheet_name = next((s['name'] for s in self.sheets if s['sheetId'] == sheet_num), "?")
            lines.append(f"  {ws_name} (sheet {sheet_num}, '{sheet_name}'):")
            if ws is not None:
                # Check for legacyDrawing
                leg = find(ws, "legacyDrawing")
                if leg is not None:
                    rid = leg.get(ns_attr("id", "r"), "")
                    lines.append(f"    legacyDrawing: rId={rid}")
                else:
                    lines.append(f"    legacyDrawing: NONE")
                # Check for drawing
                dr = find(ws, "drawing")
                if dr is not None:
                    rid = dr.get(ns_attr("id", "r"), "")
                    lines.append(f"    drawing: rId={rid}")
                # Check sheet data
                sd = find(ws, "sheetData")
                if sd is not None:
                    row_count = len(findall(sd, "row"))
                    cell_count = sum(len(findall(r, "c")) for r in findall(sd, "row"))
                    lines.append(f"    sheetData: {row_count} rows, {cell_count} cells")
                # Count merge cells
                mc = find(ws, "mergeCells")
                if mc is not None:
                    merge_count = len(findall(mc, "mergeCell"))
                    lines.append(f"    mergeCells: {merge_count}")
                # Page setup
                ps = find(ws, "pageSetup")
                if ps is not None:
                    lines.append(f"    pageSetup: paperSize={ps.get('paperSize')} orientation={ps.get('orientation')}")
        lines.append("")
        
        # Worksheet rels
        lines.append("--- WORKSHEET RELATIONSHIPS ---")
        for rels_name, rels in sorted(self.worksheet_rels.items()):
            lines.append(f"  {rels_name}:")
            for rid, info in sorted(rels.items()):
                typ = info["Type"].split("/")[-1]
                lines.append(f"    {rid}: Type={typ} Target={info['Target']}")
        lines.append("")
        
        # Comments
        lines.append("--- COMMENTS ---")
        if self.comments:
            for name, cmt in sorted(self.comments.items()):
                if cmt is not None:
                    cl = find(cmt, "commentList")
                    count = len(findall(cl, "comment")) if cl is not None else 0
                    lines.append(f"  {name}: {count} comments")
                    if cl is not None:
                        for c in findall(cl, "comment"):
                            ref = c.get("ref", "?")
                            lines.append(f"    comment ref={ref}")
                else:
                    lines.append(f"  {name}: (unparseable)")
        else:
            lines.append("  (none)")
        lines.append("")
        
        # VML
        lines.append("--- VML DRAWINGS ---")
        if self.vml_drawings:
            for name, vml in sorted(self.vml_drawings.items()):
                lines.append(f"  {name}: {len(vml)} chars")
        else:
            lines.append("  (none)")
        lines.append("")
        
        # Styles
        lines.append("--- STYLES ---")
        if self.styles is not None:
            f = len(findall(self.styles, "fonts"))
            fl = len(findall(self.styles, "fills"))
            b = len(findall(self.styles, "borders"))
            c = len(findall(self.styles, "cellXfs"))
            lines.append(f"  {f} fonts, {fl} fills, {b} borders, {c} cellXfs")
            # Show first few fonts
            for i, ft in enumerate(findall(self.styles, "fonts")):
                for font in findall(ft, "font"):
                    name_elem = find(font, "name")
                    sz_elem = find(font, "sz")
                    fn = name_elem.get("val") if name_elem is not None else "?"
                    fs = sz_elem.get("val") if sz_elem is not None else "?"
                    lines.append(f"    font: name={fn} size={fs}")
        else:
            lines.append("  (unparseable)")
        lines.append("")
        
        # Shared strings
        lines.append("--- SHARED STRINGS ---")
        if self.shared_strings is not None:
            si_count = len(findall(self.shared_strings, "si"))
            lines.append(f"  {si_count} items")
        else:
            lines.append("  (none)")
        lines.append("")
        
        # Relationship graph
        lines.append("--- RELATIONSHIP GRAPH ---")
        lines.append("  Orphan targets (referenced but missing):")
        if self.orphan_parts:
            for o in sorted(self.orphan_parts):
                lines.append(f"    {o}")
        else:
            lines.append("    (none)")
        lines.append("  Unreferenced parts (exist but not referenced):")
        unreferenced = [p for p in self.unreferenced_parts if p not in ["_rels/.rels", "[Content_Types].xml"]]
        if unreferenced:
            for u in sorted(unreferenced):
                lines.append(f"    {u}")
        else:
            lines.append("    (none)")
        
        return "\n".join(lines)
    
    def get_part_bytes(self, name):
        return self.entries.get(name)
    
    def get_part_sha256(self, name):
        data = self.entries.get(name)
        if data:
            return sha256_file(data)
        return None


def deep_compare_xml_tree(node_a, node_b, path="/", depth=0):
    """Compare two XML trees element-by-element, return list of differences."""
    diffs = []
    indent = "  " * depth
    
    if node_a is None and node_b is None:
        return diffs
    if node_a is None:
        diffs.append(f"{indent}DIFF at {path}: LEFT missing, RIGHT exists")
        return diffs
    if node_b is None:
        diffs.append(f"{indent}DIFF at {path}: LEFT exists, RIGHT missing")
        return diffs
    
    # Compare tags
    tag_a = node_a.tag.split('}')[-1] if '}' in node_a.tag else node_a.tag
    tag_b = node_b.tag.split('}')[-1] if '}' in node_b.tag else node_b.tag
    if tag_a != tag_b:
        diffs.append(f"{indent}DIFF at {path}: tag='{tag_a}' vs '{tag_b}'")
        return diffs
    
    # Compare attributes
    attrs_a = {k.split('}')[-1]: v for k, v in sorted(node_a.attrib.items())}
    attrs_b = {k.split('}')[-1]: v for k, v in sorted(node_b.attrib.items())}
    if attrs_a != attrs_b:
        added = set(attrs_b.keys()) - set(attrs_a.keys())
        removed = set(attrs_a.keys()) - set(attrs_b.keys())
        common = set(attrs_a.keys()) & set(attrs_b.keys())
        diff_attrs = []
        for k in sorted(removed):
            diff_attrs.append(f"{k}={attrs_a[k]} (removed)")
        for k in sorted(added):
            diff_attrs.append(f"{k}={attrs_b[k]} (added)")
        for k in sorted(common):
            if attrs_a[k] != attrs_b[k]:
                diff_attrs.append(f"{k}: '{attrs_a[k]}' -> '{attrs_b[k]}'")
        if diff_attrs:
            diffs.append(f"{indent}DIFF at {path}/{tag_a}: [{' '.join(diff_attrs)}]")
            return diffs  # Don't recurse further into children if attrs differ
    
    # Compare text
    text_a = (node_a.text or "").strip()
    text_b = (node_b.text or "").strip()
    if text_a != text_b and (text_a or text_b):
        diffs.append(f"{indent}DIFF at {path}/{tag_a}: text='{text_a[:60]}' vs '{text_b[:60]}'")
        return diffs
    
    # Compare children
    children_a = list(node_a)
    children_b = list(node_b)
    for i in range(max(len(children_a), len(children_b))):
        child_path = f"{path}/{tag_a}[{i}]"
        child_a = children_a[i] if i < len(children_a) else None
        child_b = children_b[i] if i < len(children_b) else None
        diffs.extend(deep_compare_xml_tree(child_a, child_b, child_path, depth + 1))
    
    return diffs


def compare_workbook_parts(analyzer_ref, analyzer_target, label_ref, label_target):
    """Compare two workbooks byte-by-byte and structurally."""
    diffs = {}
    
    # Compare ZIP entries
    entries_ref = set(analyzer_ref.entries.keys())
    entries_target = set(analyzer_target.entries.keys())
    
    added = entries_target - entries_ref
    removed = entries_ref - entries_target
    common = entries_ref & entries_target
    
    diffs["zip_added"] = sorted(added)
    diffs["zip_removed"] = sorted(removed)
    
    # Byte-level comparison of common entries
    byte_diffs = {}
    for name in sorted(common):
        data_ref = analyzer_ref.entries[name]
        data_target = analyzer_target.entries[name]
        if data_ref != data_target:
            sha_ref = sha256_file(data_ref)
            sha_target = sha256_file(data_target)
            byte_diffs[name] = {"ref_size": len(data_ref), "target_size": len(data_target),
                                 "ref_sha": sha_ref[:16], "target_sha": sha_target[:16]}
            
            # For XML files, do structural comparison
            if name.endswith(".xml") or name.endswith(".vml"):
                xml_ref = parse_xml_safe(data_ref)
                xml_target = parse_xml_safe(data_target)
                if xml_ref is not None and xml_target is not None:
                    xml_diffs = deep_compare_xml_tree(xml_ref, xml_target, f"/{name}")
                    if xml_diffs:
                        byte_diffs[f"{name}_xml"] = xml_diffs
    
    diffs["byte_diffs"] = byte_diffs
    
    # Compare sheets
    sheets_ref = {s['name']: s for s in analyzer_ref.sheets}
    sheets_target = {s['name']: s for s in analyzer_target.sheets}
    sheet_names_ref = set(sheets_ref.keys())
    sheet_names_target = set(sheets_target.keys())
    diffs["sheets_added"] = sorted(sheet_names_target - sheet_names_ref)
    diffs["sheets_removed"] = sorted(sheet_names_ref - sheet_names_target)
    
    # Compare defined names
    dn_ref = {(d['name'], d['text']): d for d in analyzer_ref.defined_names}
    dn_target = {(d['name'], d['text']): d for d in analyzer_target.defined_names}
    dn_names_ref = set((d['name'], d['text'][:40]) for d in analyzer_ref.defined_names)
    dn_names_target = set((d['name'], d['text'][:40]) for d in analyzer_target.defined_names)
    diffs["dn_added"] = sorted(dn_names_target - dn_names_ref)
    diffs["dn_removed"] = sorted(dn_names_ref - dn_names_target)
    
    # Compare workbook rels
    rels_ref = set((rid, info['Target']) for rid, info in analyzer_ref.workbook_rels.items())
    rels_target = set((rid, info['Target']) for rid, info in analyzer_target.workbook_rels.items())
    diffs["rels_added"] = sorted(rels_target - rels_ref)
    diffs["rels_removed"] = sorted(rels_ref - rels_target)
    
    # Compare content types
    diffs["ct_diffs"] = []
    if analyzer_ref.content_types is not None and analyzer_target.content_types is not None:
        ct_diffs = deep_compare_xml_tree(analyzer_ref.content_types, analyzer_target.content_types, "/[Content_Types].xml")
        if ct_diffs:
            diffs["ct_diffs"] = ct_diffs
    
    # Compare relationship graph
    orphan_ref = set(analyzer_ref.orphan_parts)
    orphan_target = set(analyzer_target.orphan_parts)
    diffs["orphan_added"] = sorted(orphan_target - orphan_ref)
    diffs["orphan_removed"] = sorted(orphan_ref - orphan_target)
    
    unreferenced_ref = set(analyzer_ref.unreferenced_parts)
    unreferenced_target = set(analyzer_target.unreferenced_parts)
    diffs["unreferenced_added"] = sorted(unreferenced_target - unreferenced_ref)
    diffs["unreferenced_removed"] = sorted(unreferenced_ref - unreferenced_target)
    
    return diffs


def format_diffs(diffs, label1, label2):
    """Format comparison results for display."""
    lines = []
    lines.append("=" * 70)
    lines.append(f"COMPARISON: {label1} vs {label2}")
    lines.append("=" * 70)
    lines.append("")
    
    # ZIP entry changes
    if diffs["zip_added"]:
        lines.append("--- ENTRIES ADDED ---")
        for e in diffs["zip_added"]:
            lines.append(f"  + {e}")
    if diffs["zip_removed"]:
        lines.append("--- ENTRIES REMOVED ---")
        for e in diffs["zip_removed"]:
            lines.append(f"  - {e}")
    
    # Byte diffs
    if diffs.get("byte_diffs"):
        lines.append("")
        lines.append("--- BYTE DIFFERENCES (common entries with different content) ---")
        for name, info in sorted(diffs["byte_diffs"].items()):
            if "_xml" not in name:
                lines.append(f"  {name}:")
                lines.append(f"    {label1}: {info['ref_size']}B [{info['ref_sha']}]")
                lines.append(f"    {label2}: {info['target_size']}B [{info['target_sha']}]")
            else:
                # XML structural diffs
                lines.append(f"  {name.replace('_xml', '')} — XML STRUCTURAL DIFFERENCES:")
                for d in info:
                    lines.append(f"    {d}")
    
    # Sheet changes
    if diffs["sheets_added"]:
        lines.append("")
        lines.append("--- SHEETS ADDED ---")
        for s in diffs["sheets_added"]:
            lines.append(f"  + {s}")
    if diffs["sheets_removed"]:
        lines.append("--- SHEETS REMOVED ---")
        for s in diffs["sheets_removed"]:
            lines.append(f"  - {s}")
    
    # Defined name changes
    if diffs["dn_added"]:
        lines.append("")
        lines.append("--- DEFINED NAMES ADDED ---")
        for d in diffs["dn_added"]:
            lines.append(f"  + {d[0]} = {d[1][:60]}")
    if diffs["dn_removed"]:
        lines.append("--- DEFINED NAMES REMOVED ---")
        for d in diffs["dn_removed"]:
            lines.append(f"  - {d[0]} = {d[1][:60]}")
    
    # Relationship changes
    if diffs["rels_added"]:
        lines.append("")
        lines.append("--- WORKBOOK RELATIONSHIPS ADDED ---")
        for r in diffs["rels_added"]:
            lines.append(f"  + {r[0]} -> {r[1]}")
    if diffs["rels_removed"]:
        lines.append("--- WORKBOOK RELATIONSHIPS REMOVED ---")
        for r in diffs["rels_removed"]:
            lines.append(f"  - {r[0]} -> {r[1]}")
    
    # Content type changes
    if diffs.get("ct_diffs"):
        lines.append("")
        lines.append("--- [Content_Types].xml DIFFERENCES ---")
        for d in diffs["ct_diffs"]:
            lines.append(f"  {d}")
    
    # Relationship graph changes
    if diffs.get("orphan_added"):
        lines.append("")
        lines.append("--- NEW ORPHAN REFERENCES (broken) ---")
        for o in diffs["orphan_added"]:
            lines.append(f"  {o}")
    if diffs.get("unreferenced_added"):
        lines.append("")
        lines.append("--- NEW UNREFERENCED PARTS (potential orphans) ---")
        for u in diffs["unreferenced_added"]:
            lines.append(f"  {u}")
    
    if not any([diffs["zip_added"], diffs["zip_removed"], diffs["byte_diffs"],
                diffs["sheets_added"], diffs["sheets_removed"],
                diffs["dn_added"], diffs["dn_removed"],
                diffs["rels_added"], diffs["rels_removed"],
                diffs.get("ct_diffs"), diffs.get("orphan_added"),
                diffs.get("unreferenced_added")]):
        lines.append("  (IDENTICAL — no differences found)")
    
    return "\n".join(lines)


def main():
    if len(sys.argv) < 4:
        print("Usage: python forensic_threeway_compare.py <original.xlsx> <conmas_output.xlsx> <generated.xlsx>")
        sys.exit(1)
    
    paths = {
        "original": sys.argv[1],
        "conmas": sys.argv[2],
        "generated": sys.argv[3],
    }
    
    # Load all three workbooks
    analyzers = {}
    for label, path in paths.items():
        print(f"Loading {label}... ({path})")
        ana = WorkbookForensicAnalyzer(path, label)
        if ana.load():
            analyzers[label] = ana
            print(f"  OK — {len(ana.entries)} entries, {len(ana.sheets)} sheets")
        else:
            print(f"  FAILED to load {path}")
            return 1
    
    # Dump full structure of each workbook
    for label, ana in analyzers.items():
        output = ana.dump_structure()
        output_path = f"{path.rsplit('.', 1)[0]}_{label}_structure.txt"
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(output)
        print(f"  Structure dumped to: {output_path}")
    
    # Three-way comparisons
    comparisons = [
        ("original", "conmas"),
        ("original", "generated"),
        ("conmas", "generated"),
    ]
    
    report_lines = []
    report_lines.append("=" * 70)
    report_lines.append("PHASE 5.5.3 — THREE-WAY FORENSIC OOXML COMPARISON REPORT")
    report_lines.append("=" * 70)
    report_lines.append("")
    
    for ref_label, target_label in comparisons:
        if ref_label in analyzers and target_label in analyzers:
            diffs = compare_workbook_parts(
                analyzers[ref_label], analyzers[target_label],
                ref_label, target_label
            )
            formatted = format_diffs(diffs, ref_label, target_label)
            report_lines.append(formatted)
            report_lines.append("")
    
    # Summary: which comparison matters most
    report_lines.append("=" * 70)
    report_lines.append("SUMMARY")
    report_lines.append("=" * 70)
    report_lines.append("")
    
    report_lines.append("The key comparison is: legacy ConMas output vs current PaperLess output.")
    report_lines.append("If the generated workbook matches the ConMas structure exactly,")
    report_lines.append("there should be no Excel repair.")
    report_lines.append("")
    
    if "conmas" in analyzers and "generated" in analyzers:
        diffs = compare_workbook_parts(
            analyzers["conmas"], analyzers["generated"],
            "conmas", "generated"
        )
        first_issues = []
        
        # Order issues by severity
        if diffs.get("byte_diffs"):
            for name in sorted(diffs["byte_diffs"].keys()):
                if not name.endswith("_xml"):
                    first_issues.append(f"BYTE DIFFERENCE: {name}")
                else:
                    xml_name = name.replace("_xml", "")
                    first_issues.append(f"XML STRUCTURAL DIFFERENCE: {xml_name}")
                    # Show first XML diff
                    xml_diffs = diffs["byte_diffs"][name]
                    if xml_diffs:
                        first_issues.append(f"  First XML diff: {xml_diffs[0]}")
        
        if diffs.get("orphan_added"):
            first_issues.append(f"BROKEN REFERENCES (orphan): {diffs['orphan_added'][0]}")
        
        if diffs.get("unreferenced_added"):
            first_issues.append(f"NEW UNREFERENCED PARTS: {diffs['unreferenced_added'][0]}")
        
        if diffs.get("rels_added"):
            first_issues.append(f"RELATIONSHIP ADDED: {diffs['rels_added'][0]}")
        
        if diffs.get("rels_removed"):
            first_issues.append(f"RELATIONSHIP REMOVED: {diffs['rels_removed'][0]}")
        
        if diffs.get("ct_diffs"):
            first_issues.append(f"CONTENT TYPE DIFF: {diffs['ct_diffs'][0]}")
        
        report_lines.append("Issues between ConMas and Generated (most likely cause of repair):")
        for issue in first_issues[:15]:  # Top 15
            report_lines.append(f"  ❌ {issue}")
        
        if not first_issues:
            report_lines.append("  ✅ No differences found between ConMas and Generated!")
    
    # Save full report
    report_path = paths["generated"] + ".threeway_report.txt"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))
    print(f"\nFull three-way report written to: {report_path}")
    
    # Print summary to stdout
    for line in report_lines[-20:]:
        try:
            print(line)
        except UnicodeEncodeError:
            print(line.encode('ascii', errors='replace').decode('ascii'))


if __name__ == "__main__":
    main()
