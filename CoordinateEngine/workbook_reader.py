import zipfile
import xml.etree.ElementTree as ET
import os
import re
from typing import Optional, List, Tuple, Dict
from CoordinateEngine.models.field_def import (MergeRange,
    
    WorkbookInfo, SheetInfo, CellFormat, CellFont, CellFill, CellBorder, CellBorders,
    CellAlignment, DataValidation, NamedRange, ProtectionInfo, PrintTitles, DrawingInfo, ControlInfo
)
from CoordinateEngine.utils.file_utils import col_letter, char_to_pt


def _find(el, tag):
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    return el.find(f'{{{ns}}}{tag}')


def _findall(el, tag):
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    return el.findall(f'{{{ns}}}{tag}')


def _attr(el, name, default=None):
    return el.get(name, default)


def _int_attr(el, name, default=0):
    v = el.get(name)
    try:
        return int(v) if v is not None else default
    except (ValueError, TypeError):
        return default


def _float_attr(el, name, default=0.0):
    v = el.get(name)
    try:
        return float(v) if v is not None else default
    except (ValueError, TypeError):
        return default


def _parse_cell_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)', ref.upper())
    if not m:
        return (1, 1)
    col = 0
    for ch in m.group(1):
        col = col * 26 + (ord(ch) - 64)
    return (col, int(m.group(2)))


def _find_relationships(zf, rels_path):
    rels = {}
    try:
        tree = ET.parse(zf.open(rels_path))
        root = tree.getroot()
        for rel in root:
            rid = rel.get('Id', '')
            target = rel.get('Target', '')
            rels[rid] = target
    except (KeyError, ET.ParseError):
        pass
    return rels


def _parse_style_xml(zf):
    try:
        tree = ET.parse(zf.open('xl/styles.xml'))
    except (KeyError, ET.ParseError):
        return {}
    root = tree.getroot()
    fonts = []
    fills = []
    borders = []
    num_fmts = {}
    cell_xfs = _find(root, 'cellXfs')
    fonts_el = _find(root, 'fonts')
    fills_el = _find(root, 'fills')
    borders_el = _find(root, 'borders')
    num_fmts_el = _find(root, 'numFmts')
    if num_fmts_el is not None:
        for nf in num_fmts_el:
            nf_id = _int_attr(nf, 'numFmtId')
            nf_code = _attr(nf, 'formatCode', 'General')
            num_fmts[nf_id] = nf_code
    if fonts_el is not None:
        for f_el in fonts_el:
            font = CellFont()
            name_el = _find(f_el, 'fontName') or _find(f_el, 'rFont')
            if name_el is not None:
                font.name = _attr(name_el, 'val')
            sz_el = _find(f_el, 'sz')
            if sz_el is not None:
                font.size = _float_attr(sz_el, 'val')
            font.bold = _find(f_el, 'b') is not None
            font.italic = _find(f_el, 'i') is not None
            u_el = _find(f_el, 'u')
            if u_el is not None:
                font.underline = _attr(u_el, 'val', 'single')
            font.strike = _find(f_el, 'strike') is not None
            color_el = _find(f_el, 'color')
            if color_el is not None:
                font.color_argb = _attr(color_el, 'rgb') or _attr(color_el, 'indexed') or _attr(color_el, 'theme')
            fonts.append(font)
    if fills_el is not None:
        for fl_el in fills_el:
            fill = CellFill()
            pat_el = _find(fl_el, 'patternFill')
            if pat_el is not None:
                fill.pattern_type = _attr(pat_el, 'patternType')
                fg = _find(pat_el, 'fgColor')
                if fg is not None:
                    fill.fg_color_argb = _attr(fg, 'rgb') or _attr(fg, 'indexed') or _attr(fg, 'theme')
                bg = _find(pat_el, 'bgColor')
                if bg is not None:
                    fill.bg_color_argb = _attr(bg, 'rgb') or _attr(bg, 'indexed') or _attr(bg, 'theme')
            fills.append(fill)
    if borders_el is not None:
        for b_el in borders_el:
            cb_o = CellBorders()
            for side in ['left', 'right', 'top', 'bottom', 'diagonal']:
                side_el = _find(b_el, side)
                if side_el is not None:
                    cb = CellBorder(style=_attr(side_el, 'style'))
                    color = _find(side_el, 'color')
                    if color is not None:
                        cb.color_argb = _attr(color, 'rgb') or _attr(color, 'indexed') or _attr(color, 'theme')
                    setattr(cb_o, side, cb)
            borders.append(cb_o)
    formats = {}
    if cell_xfs is not None:
        for idx, xf in enumerate(cell_xfs):
            fmt = CellFormat()
            font_id = _int_attr(xf, 'fontId', -1)
            fill_id = _int_attr(xf, 'fillId', -1)
            border_id = _int_attr(xf, 'borderId', -1)
            nf_id = _int_attr(xf, 'numFmtId', 0)
            if font_id >= 0 and font_id < len(fonts):
                fmt.font = fonts[font_id]
            if fill_id >= 0 and fill_id < len(fills):
                fmt.fill = fills[fill_id]
            if border_id >= 0 and border_id < len(borders):
                fmt.borders = borders[border_id]
            if nf_id in num_fmts:
                fmt.number_format = num_fmts[nf_id]
            fmt.number_format_id = nf_id
            align_el = _find(xf, 'alignment')
            if align_el is not None:
                al = CellAlignment()
                al.horizontal = _attr(align_el, 'horizontal')
                al.vertical = _attr(align_el, 'vertical')
                al.wrap_text = _attr(align_el, 'wrapText', '') in ('1', 'true', 'True')
                al.indent = _int_attr(align_el, 'indent')
                al.rotation = _int_attr(align_el, 'textRotation')
                al.shrink_to_fit = _attr(align_el, 'shrinkToFit', '') in ('1', 'true', 'True')
                fmt.alignment = al
            formats[idx] = fmt
    return formats


def _parse_shared_strings(zf):
    strings = []
    try:
        tree = ET.parse(zf.open('xl/sharedStrings.xml'))
        root = tree.getroot()
        for si in _findall(root, 'si'):
            t_el = _find(si, 't')
            if t_el is not None and t_el.text:
                strings.append(t_el.text)
            else:
                parts = []
                for r in _findall(si, 'r'):
                    tr = _find(r, 't')
                    if tr is not None and tr.text:
                        parts.append(tr.text)
                strings.append(''.join(parts) if parts else '')
    except (KeyError, ET.ParseError):
        pass
    return strings


def _parse_comments(zf, comment_path):
    comments = {}
    try:
        tree = ET.parse(zf.open(comment_path))
        root = tree.getroot()
        cl = _find(root, 'commentList')
        if cl is None:
            for el in root:
                if 'commentList' in el.tag:
                    cl = el
                    break
        if cl is not None:
            for cm in cl:
                ref = _attr(cm, 'ref', '')
                text_el = _find(cm, 'text')
                if text_el is None:
                    for child in cm:
                        if 'text' in child.tag:
                            text_el = child
                            break
                t_el = _find(text_el, 't') if text_el is not None else None
                if t_el is not None and t_el.text:
                    comments[ref] = t_el.text
                elif text_el is not None:
                    parts = []
                    for r in _findall(text_el, 'r'):
                        tr = _find(r, 't')
                        if tr is not None and tr.text:
                            parts.append(tr.text)
                    if parts:
                        comments[ref] = ''.join(parts)
    except (KeyError, ET.ParseError):
        pass
    return comments


def _parse_drawings(zf, drawing_path, sheet_rels):
    drawings = []
    try:
        tree = ET.parse(zf.open(drawing_path))
        root = tree.getroot()
    except (KeyError, ET.ParseError):
        return drawings
    drawing_rels_path = drawing_path.rsplit('.', 1)[0] + '.xml.rels'
    drawing_rels = _find_relationships(zf, drawing_rels_path)
    for ws_drawing in root:
        if 'twoCellAnchor' not in ws_drawing.tag and 'oneCellAnchor' not in ws_drawing.tag and 'absoluteAnchor' not in ws_drawing.tag:
            continue
        di = DrawingInfo()
        row_el = ws_drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}row')
        col_el = ws_drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}col')
        if row_el is not None:
            di.anchor_row = int(row_el.text)
        if col_el is not None:
            di.anchor_col = int(col_el.text)
        pic = ws_drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}pic')
        if pic is not None:
            di.drawing_type = 'image'
            di.is_image = True
            nv_pr = pic.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr')
            if nv_pr is not None:
                di.name = _attr(nv_pr, 'name', '')
                di.description = _attr(nv_pr, 'descr', '')
            blip = pic.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
            if blip is not None:
                embed = _attr(blip, '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', '')
                if embed and embed in drawing_rels:
                    di.image_type = drawing_rels[embed].split('.')[-1]
            sp_pr = pic.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}spPr')
            if sp_pr is not None:
                xfrm = sp_pr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
                if xfrm is not None:
                    ext = xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ext')
                    if ext is not None:
                        di.width_emu = _int_attr(ext, 'cx')
                        di.height_emu = _int_attr(ext, 'cy')
                        di.width_pt = _int_attr(ext, 'cx', 0) / 12700.0
                        di.height_pt = _int_attr(ext, 'cy', 0) / 12700.0
        else:
            sp = ws_drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}sp')
            if sp is not None:
                di.drawing_type = 'shape'
                nv_pr = sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr')
                if nv_pr is not None:
                    di.name = _attr(nv_pr, 'name', '')
                    di.description = _attr(nv_pr, 'descr', '')
            else:
                gf = ws_drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}graphicFrame')
                if gf is not None:
                    di.drawing_type = 'chart'
        drawings.append(di)
    return drawings


def _parse_controls(zf, ctrl_path):
    controls = []
    try:
        tree = ET.parse(zf.open(ctrl_path))
        root = tree.getroot()
    except (KeyError, ET.ParseError):
        return controls
    for ctrl in root:
        ci = ControlInfo()
        ci.cell_ref = _attr(ctrl, 'cellRef', '')
        ci.name = _attr(ctrl, 'name', '')
        ci.alt_text = _attr(ctrl, 'altText', '')
        ci.linked_cell = _attr(ctrl, 'linkedCell', '')
        ci.list_fill_range = _attr(ctrl, 'listFillRange', '')
        name_lower = (ci.name or '').lower()
        if 'checkbox' in name_lower:
            ci.ctrl_type = 'checkbox'
        elif 'option' in name_lower or 'radio' in name_lower:
            ci.ctrl_type = 'radio'
        elif 'button' in name_lower:
            ci.ctrl_type = 'button'
        elif 'dropdown' in name_lower or 'combo' in name_lower:
            ci.ctrl_type = 'dropdown'
        elif 'list' in name_lower and 'list' != (ci.list_fill_range or ''):
            ci.ctrl_type = 'listbox'
        else:
            ci.ctrl_type = 'unknown'
        ci.is_activex = 'activeX' in ctrl_path
        controls.append(ci)
    return controls


def read_workbook(filepath, verbose=False):
    name = os.path.splitext(os.path.basename(filepath))[0]
    info = WorkbookInfo(filepath=filepath, name=name)

    if verbose:
        print(f"\n{'='*70}")
        print(f"WORKBOOK READER TRACE: {filepath}")
        print(f"{'='*70}")

    try:
        zf = zipfile.ZipFile(filepath, 'r')
    except (zipfile.BadZipFile, FileNotFoundError):
        return info

    # Stage 1: ZIP File Verification
    if verbose:
        print(f"\n{'-'*70}")
        print("STAGE 1: ZIP File Verification")
        print(f"{'-'*70}")
        print(f"  Total entries: {len(zf.namelist())}")
        total_size = sum(zi.file_size for zi in zf.infolist())
        print(f"  Total uncompressed size: {total_size} bytes")
        for zi in zf.infolist():
            crc_hex = format(zi.CRC, '08x') if zi.CRC else 'N/A'
            ct = {0: 'STORED', 8: 'DEFLATED', 12: 'BZIP2'}.get(zi.compress_type, f't={zi.compress_type}')
            print(f"  [{ct:7s}] {zi.filename:55s} {zi.file_size:8d} B  CRC:{crc_hex}")
        print(f"{'-'*70}")

    info.formats = _parse_style_xml(zf)
    if verbose:
        print(f"\nSTAGE 1b: Style Formats ({len(info.formats)} xf entries)")
        for idx, fmt in sorted(info.formats.items())[:15]:
            fn = fmt.font.name if fmt.font else 'default'
            fs = fmt.font.size if fmt.font else 11
            bd = 'BOLD ' if fmt.font and fmt.font.bold else ''
            fi = fmt.fill.pattern_type if fmt.fill else 'none'
            bi = 'border' if fmt.borders else 'no-border'
            al = ''
            if fmt.alignment and fmt.alignment.horizontal:
                al = f' align={fmt.alignment.horizontal}'
            print(f"    xf[{idx:3d}]: font={fn}/{fs}pt {bd}fill={fi} {bi}{al}")
        if len(info.formats) > 15:
            print(f"    ... ({len(info.formats) - 15} more)")

    shared_strings = _parse_shared_strings(zf)
    if verbose:
        print(f"\nSTAGE 1c: Shared Strings ({len(shared_strings)} entries)")
        for i, s in enumerate(shared_strings[:30]):
            print(f"    [{i:4d}]: {repr(s[:120])}")
        if len(shared_strings) > 30:
            print(f"    ... ({len(shared_strings) - 30} more)")

    try:
        wb_tree = ET.parse(zf.open('xl/workbook.xml'))
        wb_root = wb_tree.getroot()
        if verbose:
            print(f"\nSTAGE 1d: workbook.xml root tag={wb_root.tag} attrib={wb_root.attrib}")
            ns_counts = {}
            for el in wb_root.iter():
                tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
                ns_counts[tag] = ns_counts.get(tag, 0) + 1
            print(f"  Element count: {sum(ns_counts.values())}")
            for tag, cnt in sorted(ns_counts.items()):
                print(f"    <{tag}>: {cnt}")
    except (KeyError, ET.ParseError):
        zf.close()
        return info

    wb_ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    sheets_xml = wb_root.find(f'{{{wb_ns}}}sheets')
    if sheets_xml is not None:
        for idx, s_el in enumerate(sheets_xml):
            sheet_name = _attr(s_el, 'name', f'Sheet{idx + 1}')
            sheet_id = _attr(s_el, 'sheetId', str(idx + 1))
            state = _attr(s_el, 'state', 'visible')
            si = SheetInfo(name=sheet_name, index=idx, visible=state in ('visible', None), sheet_id=sheet_id)
            info.sheets.append(si)
    if verbose:
        print(f"\n{'-'*70}")
        print("STAGE 2: Sheet Enumeration")
        print(f"{'-'*70}")
        for s in info.sheets:
            print(f"  Sheet[{s.index}]: name={s.name!r} id={s.sheet_id} visible={s.visible}")

    defined_names = wb_root.find(f'{{{wb_ns}}}definedNames')
    if defined_names is not None:
        if verbose:
            print(f"\nSTAGE 2b: Defined Names / Named Ranges")
            print(f"  {'Name':30s} {'RefersTo':40s} {'Scope':15s} {'CellRef':15s}")
            print(f"  {'-'*100}")
        for dn in defined_names:
            dn_name = _attr(dn, 'name', '')
            dn_ref = (dn.text or '').strip()
            local_sheet = _attr(dn, 'localSheetId', None)
            nr = NamedRange(name=dn_name, refers_to=dn_ref, scope='sheet' if local_sheet is not None else 'workbook')
            if '!' in dn_ref:
                nr.sheet_name = dn_ref.split('!')[0].strip("'")
                ref_part = dn_ref.split('!')[1].strip()
                ref_part = ref_part.replace('$', '')
                if ':' in ref_part:
                    nr.cell_ref = ref_part.split(':')[0]
                else:
                    nr.cell_ref = ref_part
            info.named_ranges.append(nr)
            # Extract print area from defined names (_xlnm.Print_Area)
            if dn_name == '_xlnm.Print_Area':
                if '!' in dn_ref:
                    ref_part = dn_ref.split('!')[1].strip().replace('$', '')
                    info.print_area_addr = ref_part
            if verbose:
                scope_s = 'workbook' if local_sheet is None else f'sheet[{local_sheet}]'
                print(f"  {dn_name:30s} {dn_ref:40s} {scope_s:15s} {nr.cell_ref or '-':15s}")

    wb_rels = _find_relationships(zf, 'xl/_rels/workbook.xml.rels')
    sheet_paths = {}
    for s in info.sheets:
        for k, v in wb_rels.items():
            if v == f'worksheets/sheet{s.index + 1}.xml' or v.endswith(f'sheet{s.index + 1}.xml'):
                target = v
                if not target.startswith('xl/'):
                    target = 'xl/' + target
                sheet_paths[s.index] = target
                break
    if not sheet_paths and info.sheets:
        for si in info.sheets:
            test_path = f'xl/worksheets/sheet{si.index + 1}.xml'
            try:
                zf.getinfo(test_path)
                sheet_paths[si.index] = test_path
            except KeyError:
                pass

    if verbose:
        print(f"\n{'-'*70}")
        print("STAGE 3: Sheet Path Resolution")
        print(f"{'-'*70}")
        for s in info.sheets:
            p = sheet_paths.get(s.index, 'MISSING')
            print(f"  Sheet[{s.index}] ({s.name}): {p}")

    info.has_sheet2 = len(info.sheets) > 1 and info.sheets[1].visible

    # Initialize cell_values storage on info
    info.cell_values = {}
    info.cell_styles = {}

    for s_idx, sheet in enumerate(info.sheets):
        sheet_path = sheet_paths.get(s_idx, f'xl/worksheets/sheet{s_idx + 1}.xml')
        if verbose:
            print(f"\n{'-'*70}")
            print(f"STAGE 4: Parsing Sheet[{s_idx}] ({sheet.name}) — {sheet_path}")
            print(f"{'-'*70}")
        try:
            st = ET.parse(zf.open(sheet_path))
        except (KeyError, ET.ParseError):
            if verbose:
                print(f"  ERROR: Cannot open {sheet_path}")
            continue
        sr = st.getroot()

        # Namespace / element analysis
        if verbose:
            print(f"  Root tag: {sr.tag}")
            ns_counts = {}
            total_els = 0
            for el in sr.iter():
                total_els += 1
                tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
                ns_counts[tag] = ns_counts.get(tag, 0) + 1
            print(f"  Total XML elements: {total_els}")
            for tag, cnt in sorted(ns_counts.items()):
                print(f"    <{tag}>: {cnt}")

        cols_el = _find(sr, 'cols')
        col_count = 0
        if cols_el is not None:
            for col_el in cols_el:
                col_min = _int_attr(col_el, 'min', 1)
                col_max = _int_attr(col_el, 'max', 1)
                width_chars = _float_attr(col_el, 'width', info.default_col_width)
                width_pt = char_to_pt(width_chars)
                hidden = _attr(col_el, 'hidden', '') in ('1', 'true', True)
                for c in range(col_min, col_max + 1):
                    col_count += 1
                    info.col_widths_pt[(c, s_idx)] = width_pt
                    if s_idx == 0:
                        info.col_widths_pt[c] = width_pt
                    if hidden:
                        info.hidden_cols.add((c, s_idx))
                        if s_idx == 0:
                            info.hidden_cols.add(c)
        if verbose:
            print(f"\nSTAGE 4a: Column Widths")
            print(f"  Default col width: {info.default_col_width} chars -> {char_to_pt(info.default_col_width):.2f}pt")
            print(f"  Total columns found: {col_count}")
            cols_s1 = {c: info.col_widths_pt[(c, s_idx)] for c in range(1, 15) if (c, s_idx) in info.col_widths_pt}
            if cols_s1:
                print(f"  Columns 1-14:")
                for c, w in cols_s1.items():
                    h = ' HIDDEN' if (c, s_idx) in {(x, s) for x, s in info.hidden_cols if s == s_idx} else ''
                    print(f"    Col {c:2d} ({chr(64+c)}): {w:8.2f}pt{h}")

        sheet_data = _find(sr, 'sheetData')
        row_count = 0
        cell_count = 0
        if sheet_data is not None:
            for row_el in sheet_data:
                r = _int_attr(row_el, 'r', 1)
                ht = _float_attr(row_el, 'ht', info.default_row_height)
                hidden = _attr(row_el, 'hidden', '') in ('1', 'true', True)
                row_count += 1
                info.row_heights_pt[(r, s_idx)] = ht
                if s_idx == 0:
                    info.row_heights_pt[r] = ht
                if hidden:
                    info.hidden_rows.add((r, s_idx))
                    if s_idx == 0:
                        info.hidden_rows.add(r)
                for cell_el in row_el:
                    ref = _attr(cell_el, 'r', '')
                    t = _attr(cell_el, 't', '')
                    style_idx = _int_attr(cell_el, 's', -1)
                    cell_count += 1
                    v_el = _find(cell_el, 'v')
                    cell_value = None
                    if t == 's' and v_el is not None and v_el.text and shared_strings:
                        try:
                            idx_s = int(v_el.text)
                            if 0 <= idx_s < len(shared_strings):
                                cell_value = shared_strings[idx_s]
                        except (ValueError, IndexError):
                            pass
                    elif v_el is not None and v_el.text:
                        cell_value = v_el.text
                    f_el = _find(cell_el, 'f')
                    if f_el is not None and f_el.text:
                        if cell_value is None:
                            cell_value = f'={f_el.text}'

                    # Store cell value and style in info
                    info.cell_values[ref] = {
                        'value': cell_value,
                        'type': t,
                        'style': style_idx,
                        'row': r,
                    }
                    info.cell_styles[ref] = style_idx

                    if verbose:
                        v_display = repr(cell_value[:60]) if cell_value else 'None'
                        f_display = f' formula={f_el.text!r}' if f_el is not None and f_el.text else ''
                        h_display = ' HIDDEN' if hidden else ''
                        t_display = f' t={t!r}' if t else ''
                        print(f"    Cell {ref:8s}: row={r:3d} style={style_idx:3d}{t_display} value={v_display:30s}{f_display}{h_display}")

        if verbose:
            print(f"\nSTAGE 4b: Row / Cell Summary")
            print(f"  Rows parsed: {row_count}")
            print(f"  Cells parsed: {cell_count}")
            print(f"  Cell values stored: {len(info.cell_values)}")
            print(f"  Hidden rows: {len([x for x in info.hidden_rows if isinstance(x, tuple) and x[1] == s_idx])}")
            print(f"  Hidden cols: {len([x for x in info.hidden_cols if isinstance(x, tuple) and x[1] == s_idx])}")
            # Show a sample of row heights
            sample_rows = [r for r in range(1, min(15, row_count + 1)) if (r, s_idx) in info.row_heights_pt]
            if sample_rows:
                print(f"\n  Row height sample (1-{max(sample_rows)}):")
                for r in sample_rows:
                    rh = info.row_heights_pt[(r, s_idx)]
                    hh = ' HIDDEN' if (r, s_idx) in {(x, s) for x, s in info.hidden_rows if s == s_idx} else ''
                    print(f"    Row {r:3d}: {rh:8.2f}pt{hh}")

        merges_el = _find(sr, 'mergeCells')
        merge_count = 0
        if merges_el is not None:
            for mc in merges_el:
                merge_ref = _attr(mc, 'ref', '')
                if merge_ref:
                    merge_count += 1
                    parts = merge_ref.split(':')
                    p1 = parts[0]
                    c, r = _parse_cell_ref(p1)
                    if len(parts) > 1:
                        from CoordinateEngine.utils.file_utils import parse_range
                        pc = parse_range(parts[1])
                        c2, r2 = pc[2], pc[3] if pc else (c, r)
                    else:
                        c2, r2 = c, r
                    info.merges[(c, r, s_idx)] = MergeRange(first_col=c, first_row=r, last_col=c2, last_row=r2, ref=merge_ref)
                    if s_idx == 0:
                        info.merges[(c, r)] = MergeRange(first_col=c, first_row=r, last_col=c2, last_row=r2, ref=merge_ref)
        if verbose:
            print(f"\nSTAGE 5: Merge Cells — {merge_count} merges found")
            for key, mr in sorted(info.merges.items()):
                if isinstance(key, tuple) and len(key) == 3 and key[2] == s_idx:
                    print(f"    {mr.ref:15s} -> cols {mr.first_col}-{mr.last_col} rows {mr.first_row}-{mr.last_row}")

        pm = _find(sr, 'pageMargins')
        if pm is not None:
            ml = _float_attr(pm, 'left', info.margin_left_pt)
            mr = _float_attr(pm, 'right', info.margin_right_pt)
            mt = _float_attr(pm, 'top', info.margin_top_pt)
            mb = _float_attr(pm, 'bottom', info.margin_bottom_pt)
            if s_idx == 0:
                info.margin_left_pt = ml
                info.margin_right_pt = mr
                info.margin_top_pt = mt
                info.margin_bottom_pt = mb
            if s_idx == 1:
                info.sheet2_margin_left_pt = ml
                info.sheet2_margin_top_pt = mt
            sheet.margin_left_pt = ml
            sheet.margin_right_pt = mr
            sheet.margin_top_pt = mt
            sheet.margin_bottom_pt = mb
            if verbose:
                print(f"\nSTAGE 5b: Page Margins — L={ml:.2f} R={mr:.2f} T={mt:.2f} B={mb:.2f} pt")

        ps = _find(sr, 'pageSetup')
        if ps is not None:
            paper_size = _int_attr(ps, 'paperSize', 0)
            orientation = _attr(ps, 'orientation', 'portrait')
            zoom = _int_attr(ps, 'zoom', 0)
            fit_w = _int_attr(ps, 'fitToWidth', 0)
            fit_h = _int_attr(ps, 'fitToHeight', 0)
            paper_sizes = {1: 'Letter', 2: 'Letter Small', 3: 'Tabloid', 4: 'Ledger',
                           5: 'Legal', 8: 'A3', 9: 'A4', 10: 'A4 Small', 11: 'A5',
                           13: 'B5', 20: 'A3 Extra', 27: 'A4 Extra'}
            sheet.paper_size = paper_sizes.get(paper_size, f'Paper{paper_size}')
            sheet.orientation = orientation
            sheet.zoom = zoom
            sheet.fit_to_pages_wide = fit_w
            sheet.fit_to_pages_tall = fit_h
            if paper_size:
                paper_dims = {1: (612, 792), 2: (612, 792), 3: (792, 1224), 4: (1224, 792),
                              5: (612, 1008), 8: (841.89, 1190.55), 9: (595.28, 841.89),
                              10: (595.28, 841.89), 11: (419.53, 595.28), 13: (498.9, 708.66)}
                if paper_size in paper_dims:
                    w, h = paper_dims[paper_size]
                    info.page_width_pt = w if orientation == 'portrait' else h
                    info.page_height_pt = h if orientation == 'portrait' else w
                    sheet.page_width_pt = info.page_width_pt
                    sheet.page_height_pt = info.page_height_pt
            if verbose:
                print(f"\nSTAGE 5c: Page Setup")
                print(f"  Paper size: {paper_size} ({sheet.paper_size})")
                print(f"  Orientation: {orientation}")
                print(f"  Zoom: {zoom}")
                print(f"  Fit to: {fit_w} x {fit_h} pages")
                print(f"  Page dims: {info.page_width_pt:.0f} x {info.page_height_pt:.0f} pt")

        print_options = _find(sr, 'printOptions')
        if print_options is not None:
            sheet.center_horizontally = _attr(print_options, 'horizontalCentered', '') in ('1', 'true')
            sheet.center_vertically = _attr(print_options, 'verticalCentered', '') in ('1', 'true')
            if verbose:
                print(f"\nSTAGE 5d: Print Options — center H={sheet.center_horizontally} V={sheet.center_vertically}")

        dv = _find(sr, 'dataValidations')
        dv_count = 0
        if dv is not None:
            for dv_item in dv:
                dv_type = _attr(dv_item, 'type', '')
                dv_op = _attr(dv_item, 'operator', '')
                f1_el = _find(dv_item, 'formula1')
                f2_el = _find(dv_item, 'formula2')
                dv_obj = DataValidation(
                    type=dv_type, operator=dv_op,
                    formula1=f1_el.text if f1_el is not None else None,
                    formula2=f2_el.text if f2_el is not None else None,
                    allow_blank=_attr(dv_item, 'allowBlank', '') in ('1', 'true', ''),
                    show_input_message=_attr(dv_item, 'showInputMessage', '') in ('1', 'true', ''),
                    show_error_message=_attr(dv_item, 'showErrorMessage', '') in ('1', 'true', ''),
                    error_title=_attr(dv_item, 'errorTitle', None),
                    error_text=_attr(dv_item, 'errorText', None),
                    input_title=_attr(dv_item, 'promptTitle', None),
                    input_text=_attr(dv_item, 'prompt', None),
                )
                dv_count += 1
                info.validation_rules.append(dv_obj)
                if verbose:
                    dv_ref = _attr(dv_item, 'sqref', '?')
                    print(f"    Validation[{dv_count}]: type={dv_type} op={dv_op} ref={dv_ref} allowBlank={dv_obj.allow_blank}")

        hyperlinks_el = _find(sr, 'hyperlinks')
        hl_count = 0
        if hyperlinks_el is not None:
            if not hasattr(info, '_hyperlinks'):
                info._hyperlinks = {}
            for hl in hyperlinks_el:
                hl_ref = _attr(hl, 'ref', '')
                hl_display = _attr(hl, 'display', '') or _attr(hl, 'location', '') or ''
                info._hyperlinks[hl_ref] = hl_display
                hl_count += 1
        if verbose and hl_count:
            print(f"\nSTAGE 5e: Hyperlinks — {hl_count} found")
            for ref, target in info._hyperlinks.items():
                print(f"    {ref} -> {target}")

        sp = _find(sr, 'sheetProtection')
        if sp is not None:
            info.protection = ProtectionInfo(
                sheet_protected=True, sheet_password=_attr(sp, 'password', None),
                objects_locked=_attr(sp, 'objects', '') not in ('1', 'true'),
                scenarios_locked=_attr(sp, 'scenarios', '') not in ('1', 'true'),
                format_cells=_attr(sp, 'formatCells', '') in ('1', 'true'),
                format_columns=_attr(sp, 'formatColumns', '') in ('1', 'true'),
                format_rows=_attr(sp, 'formatRows', '') in ('1', 'true'),
                insert_columns=_attr(sp, 'insertColumns', '') in ('1', 'true'),
                insert_rows=_attr(sp, 'insertRows', '') in ('1', 'true'),
                insert_hyperlinks=_attr(sp, 'insertHyperlinks', '') in ('1', 'true'),
                delete_columns=_attr(sp, 'deleteColumns', '') in ('1', 'true'),
                delete_rows=_attr(sp, 'deleteRows', '') in ('1', 'true'),
                select_locked_cells=_attr(sp, 'selectLockedCells', '') not in ('0', 'false'),
                select_unlocked_cells=_attr(sp, 'selectUnlockedCells', '') not in ('0', 'false'),
                sort=_attr(sp, 'sort', '') in ('1', 'true'),
                auto_filter=_attr(sp, 'autoFilter', '') in ('1', 'true'),
                pivot_tables=_attr(sp, 'pivotTables', '') in ('1', 'true'),
            )
            if verbose:
                print(f"\nSTAGE 5f: Sheet Protection — ENABLED (password={'SET' if info.protection.sheet_password else 'none'})")

        drawings_rel_el = _find(sr, 'drawing')
        if drawings_rel_el is not None:
            drawing_rid = _attr(drawings_rel_el, '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            sheet_rels_dir = sheet_path.rsplit('/', 1)[0] + '/_rels/' + sheet_path.rsplit('/', 1)[1] + '.rels'
            sheet_rels = _find_relationships(zf, sheet_rels_dir)
            if drawing_rid and drawing_rid in sheet_rels:
                drawing_target = sheet_rels[drawing_rid]
                drawing_path = 'xl/' + drawing_target.lstrip('/') if not drawing_target.startswith('xl/') else drawing_target
                info.drawings = _parse_drawings(zf, drawing_path, sheet_rels)
                if verbose:
                    print(f"\nSTAGE 5g: Drawings — {len(info.drawings)} found")
                    for d in info.drawings:
                        print(f"    type={d.drawing_type} name={d.name} anchor={d.anchor_col},{d.anchor_row} size={d.width_pt:.1f}x{d.height_pt:.1f}pt")

    if verbose:
        print(f"\n{'-'*70}")
        print("STAGE 6: Comments Processing (post-sheet loop)")
        print(f"{'-'*70}")

    for s_idx, sheet in enumerate(info.sheets):
        sheet_path = sheet_paths.get(s_idx, f'xl/worksheets/sheet{s_idx + 1}.xml')
        if not sheet_path:
            continue
        sheet_rels_dir = sheet_path.rsplit('/', 1)[0] + '/_rels/' + sheet_path.rsplit('/', 1)[1] + '.rels'
        sheet_rels = _find_relationships(zf, sheet_rels_dir)
        for rid, target in sheet_rels.items():
            if 'comments' in target.lower():
                comment_path = os.path.normpath(os.path.join(os.path.dirname(sheet_path), target)).replace('\\', '/')
                if not hasattr(info, '_comments'):
                    info._comments = {}
                info._comments.update(_parse_comments(zf, comment_path))

    comments = getattr(info, '_comments', {})
    if verbose:
        print(f"  Total comments: {len(comments)}")
        for ref, text in comments.items():
            print(f"    {ref}: {text[:80]!r}")

    try:
        ctrl_count = 0
        for name in zf.namelist():
            if 'ctrlProps' in name.lower() and name.endswith('.xml'):
                new_ctrls = _parse_controls(zf, name)
                info.controls.extend(new_ctrls)
                ctrl_count += len(new_ctrls)
        if verbose and ctrl_count:
            print(f"\nSTAGE 6b: Controls — {ctrl_count} found")
            for c in info.controls:
                print(f"    type={c.ctrl_type} cell={c.cell_ref} name={c.name} linked={c.linked_cell}")
    except Exception:
        pass

    try:
        for name in zf.namelist():
            if 'printerSettings' in name.lower():
                info.printer_settings_path = name
                if verbose:
                    print(f"\nSTAGE 6c: Printer Settings — {name}")
                break
    except Exception:
        pass

    zf.close()

    # Stage 11: Final Coordinate Summary
    if verbose:
        print(f"\n{'='*70}")
        print("STAGE 11: Workbook Reader — Final Summary")
        print(f"{'='*70}")
        print(f"  Sheets: {len(info.sheets)}")
        print(f"  Page: {info.page_width_pt:.0f} x {info.page_height_pt:.0f} pt")
        print(f"  Margins: L={info.margin_left_pt:.2f} R={info.margin_right_pt:.2f} T={info.margin_top_pt:.2f} B={info.margin_bottom_pt:.2f}")
        print(f"  Print Area: {info.print_area_addr or 'None'}")
        print(f"  Columns: {len(info.col_widths_pt)} widths stored")
        print(f"  Rows: {len(info.row_heights_pt)} heights stored")
        print(f"  Merges: {len(info.merges)}")
        print(f"  Cell values: {len(info.cell_values)}")
        print(f"  Hidden cols: {len(info.hidden_cols)}  Hidden rows: {len(info.hidden_rows)}")
        print(f"  Named ranges: {len(info.named_ranges)}")
        print(f"  Validations: {len(info.validation_rules)}")
        print(f"  Drawings: {len(info.drawings)}")
        print(f"  Controls: {len(info.controls)}")
        print(f"  Formats (xf): {len(info.formats)}")
        print(f"  Shared strings: {len(shared_strings)}")
        print(f"{'='*70}\n")

    return info


def xml_dump_to_debug(filepath, debug_dir):
    """Dump raw XML contents from the XLSX ZIP to debug/xml/ folder."""
    import shutil
    xml_dir = os.path.join(debug_dir, 'xml')
    os.makedirs(xml_dir, exist_ok=True)
    try:
        zf = zipfile.ZipFile(filepath, 'r')
        xml_files = [n for n in zf.namelist() if n.endswith('.xml') or n.endswith('.rels')]
        for rel_path in xml_files:
            safe_name = rel_path.replace('/', '_').replace('\\', '_')
            out_path = os.path.join(xml_dir, safe_name)
            try:
                content = zf.read(rel_path)
                with open(out_path, 'wb') as f:
                    f.write(content)
            except Exception:
                pass
        # Also dump a JSON index
        index = {}
        for rel_path in xml_files:
            safe_name = rel_path.replace('/', '_').replace('\\', '_')
            zinfo = zf.getinfo(rel_path)
            index[rel_path] = {
                'dump_file': safe_name,
                'size': zinfo.file_size,
                'crc': format(zinfo.CRC, '08x') if zinfo.CRC else None,
            }
        zf.close()
        import json
        idx_path = os.path.join(xml_dir, '_index.json')
        with open(idx_path, 'w', encoding='utf-8') as f:
            json.dump(index, f, indent=2)
        print(f"  XML dump: {len(xml_files)} files → {xml_dir}")
        return xml_dir
    except Exception as e:
        print(f"  XML dump error: {e}")
        return None


def generate_workbook_reader_report(info, debug_dir):
    """Generate workbook_reader_report.md documenting the full parsed workbook state."""
    import datetime
    lines = []
    lines.append("# Workbook Reader Report")
    lines.append("")
    lines.append(f"**Generated:** {datetime.datetime.now().isoformat()}")
    lines.append(f"**File:** {info.filepath}")
    lines.append(f"**Name:** {info.name}")
    lines.append("")

    # 1. ZIP / Workbook Structure
    lines.append("## 1. Workbook Structure")
    lines.append("")
    lines.append(f"| Property | Value |")
    lines.append(f"|----------|-------|")
    lines.append(f"| Sheets | {len(info.sheets)} |")
    lines.append(f"| Page width | {info.page_width_pt:.0f} pt |")
    lines.append(f"| Page height | {info.page_height_pt:.0f} pt |")
    lines.append(f"| Margin L | {info.margin_left_pt:.2f} pt |")
    lines.append(f"| Margin R | {info.margin_right_pt:.2f} pt |")
    lines.append(f"| Margin T | {info.margin_top_pt:.2f} pt |")
    lines.append(f"| Margin B | {info.margin_bottom_pt:.2f} pt |")
    lines.append(f"| Print area | {info.print_area_addr or 'None'} |")
    lines.append(f"| Default col width | {info.default_col_width} chars ({char_to_pt(info.default_col_width):.2f} pt) |")
    lines.append(f"| Default row height | {info.default_row_height:.2f} pt |")
    lines.append(f"| Has Sheet2 | {info.has_sheet2} |")
    lines.append("")

    # 2. Sheets
    lines.append("## 2. Sheets")
    lines.append("")
    lines.append("| # | Name | Visible | Page Size (pt) | Orientation | Center H | Center V | Margins (LRTB) |")
    lines.append("|---|------|---------|----------------|-------------|----------|----------|-----------------|")
    for s in info.sheets:
        vis = "Yes" if s.visible else "No"
        sz = f"{s.page_width_pt:.0f}x{s.page_height_pt:.0f}"
        mg = f"{s.margin_left_pt:.1f} {s.margin_right_pt:.1f} {s.margin_top_pt:.1f} {s.margin_bottom_pt:.1f}"
        lines.append(f"| {s.index + 1} | {s.name} | {vis} | {sz} | {s.orientation} | {s.center_horizontally} | {s.center_vertically} | {mg} |")
    lines.append("")

    # 3. Column Widths
    lines.append("## 3. Column Widths")
    lines.append("")
    seen_cols = set()
    for key, w in sorted(info.col_widths_pt.items(), key=lambda x: (str(type(x[0])), str(x[0]))):
        if isinstance(key, int):
            h = " (hidden)" if key in info.hidden_cols else ""
            lines.append(f"- **{chr(64 + key) if key <= 26 else f'Col{key}'}:** {w:.2f} pt{h}")
            seen_cols.add(key)
        elif isinstance(key, tuple) and len(key) == 2 and key[1] == 0 and key[0] not in seen_cols:
            h = " (hidden)" if key in info.hidden_cols else ""
            lines.append(f"- **{chr(64 + key[0]) if key[0] <= 26 else f'Col{key[0]}'}:** {w:.2f} pt{h}")
            seen_cols.add(key[0])
    lines.append("")

    # 4. Row Heights
    lines.append("## 4. Row Heights")
    lines.append("")
    seen_rows = set()
    for key, h in sorted(info.row_heights_pt.items(), key=lambda x: (str(type(x[0])), str(x[0]))):
        if isinstance(key, int):
            hh = " (hidden)" if key in info.hidden_rows else ""
            lines.append(f"- **Row {key}:** {h:.2f} pt{hh}")
            seen_rows.add(key)
        elif isinstance(key, tuple) and len(key) == 2 and key[1] == 0 and key[0] not in seen_rows:
            hh = " (hidden)" if key in info.hidden_rows else ""
            lines.append(f"- **Row {key[0]}:** {h:.2f} pt{hh}")
            seen_rows.add(key[0])
    lines.append("")

    # 5. Named Ranges
    lines.append("## 5. Named Ranges")
    lines.append("")
    if info.named_ranges:
        lines.append("| Name | Refers To | Scope | Cell Ref |")
        lines.append("|------|-----------|-------|----------|")
        for nr in info.named_ranges:
            lines.append(f"| {nr.name} | {nr.refers_to} | {nr.scope} | {nr.cell_ref or '-'} |")
    else:
        lines.append("None.")
    lines.append("")

    # 6. Merges
    lines.append("## 6. Merged Cells")
    lines.append("")
    if info.merges:
        lines.append("| Ref | First Col | First Row | Last Col | Last Row |")
        lines.append("|-----|-----------|-----------|----------|----------|")
        for key, mr in sorted(info.merges.items()):
            if isinstance(key, tuple) and len(key) == 2:
                lines.append(f"| {mr.ref} | {mr.first_col} | {mr.first_row} | {mr.last_col} | {mr.last_row} |")
    else:
        lines.append("None.")
    lines.append("")

    # 7. Cell Values
    lines.append("## 7. Cell Values")
    lines.append("")
    cv = getattr(info, 'cell_values', {})
    if cv:
        lines.append(f"Total: {len(cv)} cells")
        lines.append("")
        lines.append("| Ref | Row | Value | Style |")
        lines.append("|-----|-----|-------|-------|")
        for ref, data in sorted(cv.items(), key=lambda x: (x[1].get('row', 0), x[0])):
            v = data.get('value', '')
            if v is None:
                v = ''
            v_str = str(v)[:80].replace('\n', '\\n')
            style = data.get('style', -1)
            lines.append(f"| {ref} | {data.get('row', '?')} | {v_str} | {style} |")
    else:
        lines.append("No cell values stored.")
    lines.append("")

    # 8. Hidden Rows/Cols
    lines.append("## 8. Hidden Rows & Columns")
    lines.append("")
    hc = [c for c, s in info.hidden_cols]
    hr = [r for r, s in info.hidden_rows]
    lines.append(f"- **Hidden columns:** {len(hc)} — {sorted(set(hc))}")
    lines.append(f"- **Hidden rows:** {len(hr)} — {sorted(set(hr))}")
    lines.append("")

    # 9. Data Validations
    lines.append("## 9. Data Validations")
    lines.append("")
    if info.validation_rules:
        lines.append("| # | Type | Operator | Formula1 | Formula2 | Allow Blank |")
        lines.append("|---|------|----------|----------|----------|-------------|")
        for i, dv in enumerate(info.validation_rules):
            lines.append(f"| {i + 1} | {dv.type} | {dv.operator or '-'} | {dv.formula1 or ''} | {dv.formula2 or ''} | {dv.allow_blank} |")
    else:
        lines.append("None.")
    lines.append("")

    # 10. Protection
    lines.append("## 10. Sheet Protection")
    lines.append("")
    if info.protection and info.protection.sheet_protected:
        p = info.protection
        lines.append(f"- **Protected:** Yes (password: {'SET' if p.sheet_password else 'none'})")
        lines.append(f"- **Select locked cells:** {p.select_locked_cells}")
        lines.append(f"- **Select unlocked cells:** {p.select_unlocked_cells}")
        lines.append(f"- **Format cells:** {p.format_cells}")
    else:
        lines.append("Not protected.")
    lines.append("")

    # 11. Drawings
    lines.append("## 11. Drawings / Images")
    lines.append("")
    if info.drawings:
        for d in info.drawings:
            lines.append(f"- **{d.name or 'Unnamed'}:** type={d.drawing_type} anchor={d.anchor_col},{d.anchor_row} size={d.width_pt:.1f}x{d.height_pt:.1f}pt")
    else:
        lines.append("None.")
    lines.append("")

    # 12. Controls
    lines.append("## 12. Controls (Form / ActiveX)")
    lines.append("")
    if info.controls:
        for c in info.controls:
            lines.append(f"- **{c.name or 'Unnamed'}:** type={c.ctrl_type} cell={c.cell_ref} activeX={c.is_activex}")
    else:
        lines.append("None.")
    lines.append("")

    # 13. Formats Summary
    lines.append("## 13. Cell Formats (xf entries)")
    lines.append("")
    lines.append(f"Total: {len(info.formats)} xf entries")
    if info.formats:
        lines.append("")
        lines.append("| Index | Font | Size | Bold | Fill | Border | Alignment |")
        lines.append("|-------|------|------|------|------|--------|-----------|")
        for idx, fmt in sorted(info.formats.items()):
            fn = fmt.font.name if fmt.font else 'default'
            fs = fmt.font.size if fmt.font else 11
            bd = 'Yes' if fmt.font and fmt.font.bold else ''
            fi = fmt.fill.pattern_type or ''
            bi = 'Yes' if fmt.borders else ''
            al = fmt.alignment.horizontal or '' if fmt.alignment else ''
            lines.append(f"| {idx} | {fn} | {fs} | {bd} | {fi} | {bi} | {al} |")
    lines.append("")

    # 14. Shared Strings
    lines.append("## 14. Shared Strings")
    lines.append("")
    ss_length = 0
    try:
        with zipfile.ZipFile(info.filepath, 'r') as zf:
            ss = _parse_shared_strings(zf)
            ss_length = len(ss)
    except Exception:
        pass
    lines.append(f"Total: {ss_length}")
    lines.append("")

    text = '\n'.join(lines)
    if debug_dir:
        os.makedirs(debug_dir, exist_ok=True)
        path = os.path.join(debug_dir, 'workbook_reader_report.md')
        with open(path, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"  Workbook Reader Report: {path}")
        return path
    return text


# Backward compat alias
read = read_workbook