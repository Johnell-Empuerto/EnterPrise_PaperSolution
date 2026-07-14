from typing import List, Optional
from CoordinateEngine.models.field_def import FieldDef, FieldGeometry, WorkbookInfo, CellFormat, DataValidation, ControlInfo


def generate_runtime(
    workbook_info_or_name,
    fields: List[FieldDef],
    geometries: List[FieldGeometry],
    dpi: int,
    workbook_info: Optional[WorkbookInfo] = None,
):
    if isinstance(workbook_info_or_name, str):
        workbook_name = workbook_info_or_name
    else:
        workbook_name = workbook_info_or_name.name
        workbook_info = workbook_info_or_name

    if not geometries:
        return {
            "workbookName": workbook_name,
            "title": workbook_name,
            "dpi": dpi,
            "scale": 1.0,
            "version": "1.0",
            "pageWidth": 0,
            "pageHeight": 0,
            "sheets": [],
        }

    if workbook_info:
        pw_p = workbook_info.page_width_pt * dpi / 72.0
        ph_p = workbook_info.page_height_pt * dpi / 72.0
    else:
        first = geometries[0]
        pw_p = first.left_px + first.width_px
        ph_p = first.top_px + first.height_px

    metadata_fields = []
    runtime_fields = []
    for i, (field, geo) in enumerate(zip(fields, geometries)):
        placeholder = field.value or field.addr

        # Determine data type from field info
        data_type = field.data_type or _detect_type(field)

        # Determine readOnly from locked state
        read_only = field.is_locked

        # Determine required from validation
        required = field.required or (field.validation is not None and not field.validation.allow_blank)

        # Font info from format
        font_info = None
        font_size = 11.0
        bold = False
        font_color = None
        bg_color = None
        border_info = None
        alignment = None
        if field.format:
            if field.format.font:
                font_info = field.format.font.name or font_info
                font_size = field.format.font.size or font_size
                bold = field.format.font.bold
                font_color = field.format.font.color_argb
            if field.format.fill:
                bg_color = field.format.fill.fg_color_argb or bg_color
            if field.format.borders:
                b = field.format.borders
                styles = []
                for side in ['left', 'right', 'top', 'bottom']:
                    cb = getattr(b, side, None)
                    if cb and cb.style:
                        styles.append(f'{side}={cb.style}')
                if styles:
                    border_info = ','.join(styles)
            if field.format.alignment:
                al = field.format.alignment
                align_parts = []
                if al.horizontal:
                    align_parts.append(al.horizontal)
                if al.vertical:
                    align_parts.append(al.vertical)
                if al.wrap_text:
                    align_parts.append('wrap')
                if align_parts:
                    alignment = ' '.join(align_parts)

        # Control type detection overrides dataType
        if field.control:
            if field.control.ctrl_type in ('checkbox', 'radio'):
                data_type = 'boolean'
            elif field.control.ctrl_type == 'dropdown':
                data_type = 'dropdown'
            elif field.control.ctrl_type == 'button':
                data_type = 'button'

        # Validation implies dropdown
        if data_type == 'text' and field.validation and field.validation.type == 'list':
            data_type = 'dropdown'

        # Hyperlink as URL
        if field.hyperlink and data_type == 'text':
            data_type = 'url'

        runtime_field = {
            "id": f"field_{field.addr}_{i}",
            "cellReference": field.addr,
            "row": field.row,
            "column": field.col,
            "leftPx": geo.left_px,
            "topPx": geo.top_px,
            "widthPx": geo.width_px,
            "heightPx": geo.height_px,
            "leftRatio": geo.left_ratio,
            "topRatio": geo.top_ratio,
            "widthRatio": geo.width_ratio,
            "heightRatio": geo.height_ratio,
            "mergeRange": field.merge_ref if field.is_merge else None,
            "isMerged": field.is_merge,
            "dataType": data_type,
            "readOnly": read_only,
            "required": required,
            "alignment": alignment,
            "font": font_info,
            "fontSize": font_size,
            "bold": bold,
            "fontColor": font_color,
            "backgroundColor": bg_color,
            "border": border_info,
            "placeholder": placeholder,
            "defaultValue": field.value or None,
            "maxLength": 0,
            "tabIndex": i,
        }

        # Optionally override id with named range
        if field.named_range:
            runtime_field["id"] = field.named_range

        runtime_fields.append(runtime_field)

        # Build metadata enrichment field
        md_field = {"cellReference": field.addr}
        if field.formula:
            md_field["formula"] = field.formula
        if field.comment:
            md_field["comment"] = field.comment
        if field.hyperlink:
            md_field["hyperlink"] = field.hyperlink
        if field.validation:
            dv = field.validation
            md_field["validation"] = {
                "type": dv.type,
                "operator": dv.operator,
                "formula1": dv.formula1,
                "formula2": dv.formula2,
            }
        if field.control:
            md_field["control"] = {
                "type": field.control.ctrl_type,
                "isActiveX": field.control.is_activex,
                "linkedCell": field.control.linked_cell,
            }
        if not field.is_locked:
            md_field["locked"] = False
        if field.format:
            md_field["format"] = "custom"
        metadata_fields.append(md_field)

    runtime_sheet = {
        "name": "Sheet1",
        "index": 1,
        "pageWidthPx": int(round(pw_p)),
        "pageHeightPx": int(round(ph_p)),
        "fields": runtime_fields,
        "images": [],
    }

    runtime = {
        "workbookName": workbook_name,
        "title": workbook_name,
        "dpi": dpi,
        "scale": 1.0,
        "version": "1.0",
        "pageWidth": int(round(pw_p)),
        "pageHeight": int(round(ph_p)),
        "sheets": [runtime_sheet],
    }

    # Optional metadata enrichment (not in schema, ignored by existing frontend)
    _metadata = {}
    if workbook_info:
        if workbook_info.protection and workbook_info.protection.sheet_protected:
            _metadata["sheetProtected"] = True
        if workbook_info.named_ranges:
            _metadata["namedRanges"] = [
                {"name": nr.name, "refersTo": nr.refers_to, "scope": nr.scope}
                for nr in workbook_info.named_ranges
            ]
        if field_drawings:
            _metadata["drawings"] = "detected"
        if field_controls:
            _metadata["controls"] = "detected"
        if any(f.formula for f in fields):
            _metadata["hasFormulas"] = True
        if any(f.validate for f in fields if hasattr(f, 'validate')):
            pass

    if _metadata:
        runtime["_metadata"] = _metadata

    return runtime


def _detect_type(field):
    if field.value:
        try:
            float(field.value.replace(',', '').replace('%', ''))
            return 'number'
        except (ValueError, AttributeError):
            if field.value.lower() in ('true', 'false', 'yes', 'no'):
                return 'boolean'
    return 'text'


def _has_drawings(workbook_info):
    return bool(workbook_info.drawings) if workbook_info else False


def _has_controls(workbook_info):
    return bool(workbook_info.controls) if workbook_info else False


field_drawings = False
field_controls = False