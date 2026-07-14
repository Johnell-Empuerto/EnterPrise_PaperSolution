from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any


@dataclass
class CellAlignment:
    horizontal: Optional[str] = None
    vertical: Optional[str] = None
    wrap_text: bool = False
    indent: int = 0
    rotation: int = 0
    shrink_to_fit: bool = False
    reading_order: int = 0


@dataclass
class CellFont:
    name: Optional[str] = None
    size: Optional[float] = None
    bold: bool = False
    italic: bool = False
    underline: Optional[str] = None
    strike: bool = False
    color_argb: Optional[str] = None
    font_family: int = 0


@dataclass
class CellFill:
    pattern_type: Optional[str] = None
    fg_color_argb: Optional[str] = None
    bg_color_argb: Optional[str] = None


@dataclass
class CellBorder:
    style: Optional[str] = None
    color_argb: Optional[str] = None


@dataclass
class CellBorders:
    left: Optional[CellBorder] = None
    right: Optional[CellBorder] = None
    top: Optional[CellBorder] = None
    bottom: Optional[CellBorder] = None
    diagonal: Optional[CellBorder] = None


@dataclass
class CellFormat:
    font: Optional[CellFont] = None
    fill: Optional[CellFill] = None
    borders: Optional[CellBorders] = None
    alignment: Optional[CellAlignment] = None
    number_format_id: int = 0
    number_format: str = "General"


@dataclass
class DataValidation:
    type: str = ""  # list, whole, decimal, date, time, textLength, custom
    operator: Optional[str] = None
    formula1: Optional[str] = None
    formula2: Optional[str] = None
    allow_blank: bool = True
    show_input_message: bool = True
    show_error_message: bool = True
    error_title: Optional[str] = None
    error_text: Optional[str] = None
    input_title: Optional[str] = None
    input_text: Optional[str] = None


@dataclass
class ControlInfo:
    ctrl_type: str = ""  # checkbox, radio, button, dropdown, listbox, scrollbar, spinner
    cell_ref: Optional[str] = None
    name: Optional[str] = None
    alt_text: Optional[str] = None
    linked_cell: Optional[str] = None
    list_fill_range: Optional[str] = None
    value: Optional[int] = None
    is_activex: bool = False


@dataclass
class DrawingInfo:
    drawing_type: str = ""  # image, shape, chart, smartart, textbox
    name: Optional[str] = None
    description: Optional[str] = None
    anchor_col: Optional[int] = None
    anchor_row: Optional[int] = None
    col_offset_pt: float = 0.0
    row_offset_pt: float = 0.0
    width_emu: int = 0
    height_emu: int = 0
    width_pt: float = 0.0
    height_pt: float = 0.0
    is_image: bool = False
    image_type: Optional[str] = None


@dataclass
class NamedRange:
    name: str = ""
    refers_to: str = ""
    sheet_name: Optional[str] = None
    cell_ref: Optional[str] = None
    scope: str = "workbook"


@dataclass
class ProtectionInfo:
    sheet_protected: bool = False
    sheet_password: Optional[str] = None
    objects_locked: bool = True
    scenarios_locked: bool = True
    format_cells: bool = False
    format_columns: bool = False
    format_rows: bool = False
    insert_columns: bool = False
    insert_rows: bool = False
    insert_hyperlinks: bool = False
    delete_columns: bool = False
    delete_rows: bool = False
    select_locked_cells: bool = True
    select_unlocked_cells: bool = True
    sort: bool = False
    auto_filter: bool = False
    pivot_tables: bool = False


@dataclass
class PrintTitles:
    rows_to_repeat_at_top: Optional[str] = None
    cols_to_repeat_at_left: Optional[str] = None


@dataclass
class SheetInfo:
    name: str = ""
    index: int = 0
    visible: bool = True
    sheet_id: Optional[str] = None
    xml_path: Optional[str] = None
    page_width_pt: float = 612.0
    page_height_pt: float = 792.0
    margin_left_pt: float = 51.024
    margin_right_pt: float = 51.024
    margin_top_pt: float = 53.858
    margin_bottom_pt: float = 53.858
    print_area_addr: Optional[str] = None
    paper_size: Optional[str] = None
    orientation: str = "portrait"
    zoom: int = 0
    fit_to_pages_wide: int = 0
    fit_to_pages_tall: int = 0
    center_horizontally: bool = False
    center_vertically: bool = False


@dataclass
class FieldDef:
    addr: str
    col: int
    row: int
    col_end: int
    row_end: int
    sheet_index: int = 0
    merge_ref: Optional[str] = None
    value: Optional[str] = None
    is_merge: bool = False
    formula: Optional[str] = None
    is_hidden: bool = False
    is_locked: bool = True
    format: Optional[CellFormat] = None
    validation: Optional[DataValidation] = None
    comment: Optional[str] = None
    hyperlink: Optional[str] = None
    named_range: Optional[str] = None
    control: Optional[ControlInfo] = None
    data_type: str = "text"
    required: bool = False
    ratio_left: Optional[float] = None
    ratio_top: Optional[float] = None
    ratio_right: Optional[float] = None
    ratio_bottom: Optional[float] = None

    def __post_init__(self):
        if self.col_end is None:
            self.col_end = self.col
        if self.row_end is None:
            self.row_end = self.row
        if self.col != self.col_end or self.row != self.row_end:
            self.is_merge = True


@dataclass
class WorkbookInfo:
    filepath: str
    name: str
    col_widths_pt: dict = field(default_factory=dict)
    row_heights_pt: dict = field(default_factory=dict)
    hidden_cols: set = field(default_factory=set)
    hidden_rows: set = field(default_factory=set)
    default_col_width: float = 8.43
    default_row_height: float = 15.0
    page_width_pt: float = 612.0
    page_height_pt: float = 792.0
    margin_left_pt: float = 51.024
    margin_right_pt: float = 51.024
    margin_top_pt: float = 53.858
    margin_bottom_pt: float = 53.858
    print_area_addr: Optional[str] = None
    merges: dict = field(default_factory=dict)
    printer_settings_path: Optional[str] = None
    has_sheet2: bool = False
    sheet2_margin_left_pt: Optional[float] = None
    sheet2_margin_top_pt: Optional[float] = None
    sheets: List[SheetInfo] = field(default_factory=list)
    active_sheet_index: int = 0
    named_ranges: List[NamedRange] = field(default_factory=list)
    protection: Optional[ProtectionInfo] = None
    print_titles: Optional[PrintTitles] = None
    formats: Dict[str, CellFormat] = field(default_factory=dict)
    validation_rules: List[DataValidation] = field(default_factory=list)
    drawings: List[DrawingInfo] = field(default_factory=list)
    controls: List[ControlInfo] = field(default_factory=list)


@dataclass
class ContentBounds:
    x0: float
    y0: float
    x1: float
    y1: float
    width_pt: float
    height_pt: float
    page_width_pt: float
    page_height_pt: float
    method: str = "pdf"


@dataclass
class LayoutInfo:
    origin_x_pt: float
    origin_y_pt: float
    scale_w: float
    scale_h: float
    eff_w_pt: float
    eff_h_pt: float
    paw_pt: float
    pah_pt: float
    printable_w_pt: float
    printable_h_pt: float
    centered_h: bool
    centered_v: bool


@dataclass
class FieldGeometry:
    field: FieldDef
    range_left_pt: float
    range_top_pt: float
    range_width_pt: float
    range_height_pt: float
    page_left_pt: float
    page_top_pt: float
    page_width_pt: float
    page_height_pt: float
    left_px: float
    top_px: float
    width_px: float
    height_px: float
    left_ratio: float
    top_ratio: float
    width_ratio: float
    height_ratio: float

@dataclass
class MergeRange:
    first_col: int
    first_row: int
    last_col: int
    last_row: int
    ref: str
