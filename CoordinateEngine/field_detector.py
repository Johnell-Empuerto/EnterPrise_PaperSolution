from typing import List
from CoordinateEngine.models.field_def import FieldDef, WorkbookInfo, MergeRange, DataValidation, CellFormat, ControlInfo
from CoordinateEngine.utils.file_utils import col_num


def detect(workbook_info: WorkbookInfo, verbose=False) -> List[FieldDef]:
    """Detect fields from workbook data, applying all filtering and enrichment."""
    merges = workbook_info.merges
    _validate_merges(merges)
    hidden_cols = {(c, 0) for (c, s) in workbook_info.hidden_cols if s == 0}
    hidden_rows = {(r, 0) for (r, s) in workbook_info.hidden_rows if s == 0}
    named_ranges = _build_named_range_map(workbook_info)
    hyperlinks = getattr(workbook_info, '_hyperlinks', {})
    comments = getattr(workbook_info, '_comments', {})
    controls_map = _build_control_map(workbook_info.controls)
    validation_map = _build_validation_map(workbook_info.validation_rules)

    if verbose:
        print(f"\n{'='*70}")
        print("FIELD DETECTOR TRACE")
        print(f"{'='*70}")
        print(f"  Merges: {len(merges)}")
        print(f"  Hidden cols: {len(hidden_cols)}")
        print(f"  Hidden rows: {len(hidden_rows)}")
        print(f"  Named ranges: {len(named_ranges)}")
        print(f"  Hyperlinks: {len(hyperlinks)}")
        print(f"  Comments: {len(comments)}")
        print(f"  Controls: {len(controls_map)}")
        print(f"  Validations: {len(validation_map)}")
        print(f"{'─'*70}")

    cells = []
    seen = set()

    sheet_path = 'xl/worksheets/sheet1.xml'
    try:
        import zipfile
        import xml.etree.ElementTree as ET
        ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        zf = zipfile.ZipFile(workbook_info.filepath, 'r')
        tree = ET.parse(zf.open(sheet_path))
        root = tree.getroot()
        shared_strings = []
        try:
            sst = ET.parse(zf.open('xl/sharedStrings.xml'))
            ss_root = sst.getroot()
            for si in ss_root.findall('s:si', ns):
                t_el = si.find('s:t', ns)
                if t_el is not None and t_el.text:
                    shared_strings.append(t_el.text)
                else:
                    parts = []
                    for r_el in si.findall('s:r', ns):
                        rt = r_el.find('s:t', ns)
                        if rt is not None and rt.text:
                            parts.append(rt.text)
                    shared_strings.append(''.join(parts))
        except (KeyError, ET.ParseError):
            pass
        zf.close()
    except Exception as e:
        if verbose:
            print(f"  ERROR re-parsing sheet1.xml: {e}")
        return []

    for row_el in root.findall('.//s:sheetData/s:row', ns):
        rn = int(row_el.get('r', '0'))
        # Skip hidden rows
        if (rn, 0) in hidden_rows:
            if verbose:
                print(f"  SKIP row {rn}: hidden row")
            continue
        for c_el in row_el.findall('s:c', ns):
            ref = c_el.get('r', '')
            col_str = ''.join(ch for ch in ref if ch.isalpha())
            col = col_num(col_str)
            row = int(''.join(ch for ch in ref if ch.isdigit()))
            reasons = []
            if ref in seen:
                reasons.append('duplicate ref')
                if verbose:
                    print(f"  SKIP {ref:8s}: duplicate (already seen)")
                continue
            seen.add(ref)

            if col > 50:
                reasons.append('col > 50')
                if verbose:
                    print(f"  SKIP {ref:8s}: col={col} > 50 limit")
                continue

            # Skip hidden columns
            if (col, 0) in hidden_cols:
                reasons.append('hidden col')
                if verbose:
                    print(f"  SKIP {ref:8s}: col={col} is hidden")
                continue

            # Formula check
            f_el = c_el.find('s:f', ns)
            formula = f_el.text if f_el is not None else None

            v_el = c_el.find('s:v', ns)
            raw_value = v_el.text.strip() if v_el is not None and v_el.text else ''
            cell_type = c_el.get('t', '')
            if cell_type == 's' and raw_value and shared_strings:
                try:
                    ss_idx = int(raw_value)
                    if 0 <= ss_idx < len(shared_strings):
                        value = shared_strings[ss_idx]
                    else:
                        value = raw_value
                except ValueError:
                    value = raw_value
            elif cell_type == 'b':
                value = 'true' if raw_value == '1' else 'false'
            else:
                value = raw_value

            # Skip formula-only cells (constants only; no explicit field)
            if formula and not value:
                reasons.append('formula-only (no value)')
                if verbose:
                    print(f"  SKIP {ref:8s}: formula={formula!r} has no cached value")
                continue
            if formula and not _is_editable_by_validation(ref, validation_map, workbook_info):
                reasons.append('readonly formula')
                if verbose:
                    print(f"  SKIP {ref:8s}: formula={formula!r} not editable")
                continue

            # Determine data type
            data_type = _detect_data_type(c_el, value, ns)

            # Format info
            style_idx_attr = c_el.get('s', '-1')
            style_idx = int(style_idx_attr) if style_idx_attr and style_idx_attr != '-1' else -1
            fmt = workbook_info.formats.get(style_idx) if style_idx >= 0 else None

            # Locked check
            is_locked = _is_cell_locked(workbook_info, fmt)

            # Named range
            named_range = named_ranges.get(ref)

            # Control association
            control = controls_map.get(ref)

            # Validation
            validation = validation_map.get(ref)

            # Comment
            comment_text = comments.get(ref)

            # Hyperlink
            hyperlink = hyperlinks.get(ref)

            # Required from validation
            required = validation is not None and not validation.allow_blank

            mr = _find_merge_for(col, row, merges)
            if mr is not None:
                if col != mr.first_col or row != mr.first_row:
                    if verbose:
                        print(f"  SKIP {ref:8s}: merge interior cell (merge={mr.ref}, first={mr.first_col}{mr.first_row})")
                    continue
                cells.append(FieldDef(
                    addr=ref, col=col, row=row,
                    col_end=mr.last_col, row_end=mr.last_row,
                    merge_ref=mr.ref, value=value, is_merge=True,
                    formula=formula, is_hidden=False, is_locked=is_locked,
                    format=fmt, validation=validation, comment=comment_text,
                    hyperlink=hyperlink, named_range=named_range,
                    control=control, data_type=data_type, required=required,
                ))
                if verbose:
                    mr_info = f" merge={mr.ref}"
                    nr_info = f" named={named_range}" if named_range else ""
                    cv_info = f" ctrl={control.ctrl_type}" if control else ""
                    print(f"  KEEP  {ref:8s}: value={value!r:30s} type={data_type:10s}{mr_info}{nr_info}{cv_info}")
            else:
                cells.append(FieldDef(
                    addr=ref, col=col, row=row,
                    col_end=col, row_end=row, value=value,
                    formula=formula, is_hidden=False, is_locked=is_locked,
                    format=fmt, validation=validation, comment=comment_text,
                    hyperlink=hyperlink, named_range=named_range,
                    control=control, data_type=data_type, required=required,
                ))
                if verbose:
                    nr_info = f" named={named_range}" if named_range else ""
                    cv_info = f" ctrl={control.ctrl_type}" if control else ""
                    hl_info = f" hl={hyperlink}" if hyperlink else ""
                    print(f"  KEEP  {ref:8s}: value={value!r:30s} type={data_type:10s}{nr_info}{cv_info}{hl_info}")

    cells.sort(key=lambda c: (c.row, c.col))
    if verbose:
        print(f"{'─'*70}")
        print(f"  TOTAL FIELDS DETECTED: {len(cells)}")
        print(f"{'─'*70}")
        for c in cells:
            m = f" MERGE={c.merge_ref}" if c.is_merge else ""
            print(f"    {c.addr:8s} row={c.row:3d} col={c.col:3d} value={str(c.value)[:50]!r:52s} type={c.data_type:10s}{m}")
        print(f"{'='*70}\n")
    return cells


def _find_merge_for(col, row, merges):
    for key, mr in merges.items():
        if not isinstance(key, tuple) or len(key) < 2:
            continue
        fc, fr = key[0], key[1]
        if mr.first_col <= col <= mr.last_col and mr.first_row <= row <= mr.last_row:
            return mr
    return None


def _validate_merges(merges):
    if merges is None:
        return
    if not isinstance(merges, dict):
        raise TypeError(f"Expected dict for merges but received {type(merges).__name__}")
    for key, val in merges.items():
        if not isinstance(val, MergeRange):
            raise TypeError(
                f"Expected MergeRange for merge value but received {type(val).__name__} "
                f"at key {key}. Ensure workbook_reader and field_detector agree on merge format."
            )


def _build_named_range_map(workbook_info):
    """Build {cell_ref -> named_range_name} lookup."""
    result = {}
    for nr in workbook_info.named_ranges:
        if nr.cell_ref:
            result[nr.cell_ref] = nr.name
    return result


def _build_control_map(controls):
    """Build {cell_ref -> ControlInfo} lookup."""
    result = {}
    for c in controls:
        if c.cell_ref:
            result[c.cell_ref] = c
    return result


def _build_validation_map(validations):
    """Build {cell_ref -> DataValidation} from validation rules.
    This is approximate since validation rules apply to ranges, not individual refs in our current parsing."""
    result = {}
    for dv in validations:
        pass  # Full range matching would need the sqref attribute
    return result


def _is_editable_by_validation(ref, validation_map, workbook_info):
    """Check if a formula cell should still be a field due to validation."""
    # If it has data validation, it's an input cell even with formula
    return bool(validation_map.get(ref))


def _is_cell_locked(workbook_info, fmt):
    """Determine if a cell is locked (default True unless format says unlocked)."""
    if workbook_info.protection and workbook_info.protection.sheet_protected:
        pass  # locked state from xml; default is locked
    return True


def _detect_data_type(c_el, value, ns):
    """Detect data type from cell XML."""
    t = c_el.get('t', '')
    if t == 's':
        return 'text'
    if t == 'b':
        return 'boolean'
    if t == 'e':
        return 'error'
    if value:
        try:
            float(value.replace(',', ''))
            return 'number'
        except ValueError:
            return 'text'
    return 'text'