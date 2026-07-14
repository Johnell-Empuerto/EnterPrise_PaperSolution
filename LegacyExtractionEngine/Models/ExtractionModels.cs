using System.Text.Json.Serialization;

namespace LegacyExtractionEngine.Models;

public class SheetMetadataItem
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";
    [JsonPropertyName("id")]
    public int Id { get; set; }
    [JsonPropertyName("state")]
    public string State { get; set; } = "visible";
    [JsonPropertyName("rel_id")]
    public string RelId { get; set; } = "";
}

public class WorkbookData
{
    [JsonPropertyName("properties")]
    public WorkbookPropertiesData Properties { get; set; } = new();
    [JsonPropertyName("sheets")]
    public List<SheetData> Sheets { get; set; } = new();
    [JsonPropertyName("sheet_metadata")]
    public List<SheetMetadataItem> SheetMetadata { get; set; } = new();
    [JsonPropertyName("defined_names")]
    public List<DefinedNameData> DefinedNames { get; set; } = new();
    [JsonPropertyName("custom_properties")]
    public List<CustomPropertyData> CustomProperties { get; set; } = new();
    [JsonPropertyName("workbook_protection")]
    public ProtectionData? WorkbookProtection { get; set; }
}

public class WorkbookPropertiesData
{
    [JsonPropertyName("title")]
    public string? Title { get; set; }
    [JsonPropertyName("subject")]
    public string? Subject { get; set; }
    [JsonPropertyName("creator")]
    public string? Creator { get; set; }
    [JsonPropertyName("description")]
    public string? Description { get; set; }
    [JsonPropertyName("last_modified_by")]
    public string? LastModifiedBy { get; set; }
    [JsonPropertyName("revision")]
    public string? Revision { get; set; }
    [JsonPropertyName("created")]
    public DateTime? Created { get; set; }
    [JsonPropertyName("modified")]
    public DateTime? Modified { get; set; }
    [JsonPropertyName("category")]
    public string? Category { get; set; }
    [JsonPropertyName("keywords")]
    public string? Keywords { get; set; }
    [JsonPropertyName("content_type")]
    public string? ContentType { get; set; }
    [JsonPropertyName("status")]
    public string? Status { get; set; }
    [JsonPropertyName("application")]
    public string? Application { get; set; }
    [JsonPropertyName("app_version")]
    public string? AppVersion { get; set; }
    [JsonPropertyName("hyperlink_base")]
    public string? HyperlinkBase { get; set; }
    [JsonPropertyName("manager")]
    public string? Manager { get; set; }
    [JsonPropertyName("company")]
    public string? Company { get; set; }
    [JsonPropertyName("all_properties")]
    public Dictionary<string, object?> AllProperties { get; set; } = new();
}

public class DefinedNameData
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";
    [JsonPropertyName("comment")]
    public string? Comment { get; set; }
    [JsonPropertyName("refers_to")]
    public string? RefersTo { get; set; }
    [JsonPropertyName("local_sheet_id")]
    public string? LocalSheetId { get; set; }
    [JsonPropertyName("hidden")]
    public bool Hidden { get; set; }
}

public class CustomPropertyData
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";
    [JsonPropertyName("value")]
    public string? Value { get; set; }
}

public class ProtectionData
{
    [JsonPropertyName("protected")]
    public bool Protected { get; set; }
    [JsonPropertyName("lock_objects")]
    public bool? LockObjects { get; set; }
    [JsonPropertyName("lock_scenarios")]
    public bool? LockScenarios { get; set; }
    [JsonPropertyName("algorithm_name")]
    public string? AlgorithmName { get; set; }
    [JsonPropertyName("hash_value")]
    public string? HashValue { get; set; }
    [JsonPropertyName("spin_count")]
    public int? SpinCount { get; set; }
    [JsonPropertyName("salt_value")]
    public string? SaltValue { get; set; }
    [JsonPropertyName("all_attributes")]
    public Dictionary<string, object?> AllAttributes { get; set; } = new();
}

public class SheetData
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";
    [JsonPropertyName("sheet_id")]
    public int SheetId { get; set; }
    [JsonPropertyName("visible")]
    public string Visible { get; set; } = "visible";
    [JsonPropertyName("page_setup")]
    public PageSetupData PageSetup { get; set; } = new();
    [JsonPropertyName("page_margins")]
    public PageMarginsData PageMargins { get; set; } = new();
    [JsonPropertyName("columns")]
    public List<ColumnData> Columns { get; set; } = new();
    [JsonPropertyName("rows")]
    public List<RowData> Rows { get; set; } = new();
    [JsonPropertyName("cells")]
    public List<CellData> Cells { get; set; } = new();
    [JsonPropertyName("merged_cells")]
    public List<MergedCellData> MergedCells { get; set; } = new();
    [JsonPropertyName("conditional_formats")]
    public List<ConditionalFormatData> ConditionalFormats { get; set; } = new();
    [JsonPropertyName("data_validations")]
    public List<DataValidationData> DataValidations { get; set; } = new();
    [JsonPropertyName("hyperlinks")]
    public List<HyperlinkData> Hyperlinks { get; set; } = new();
    [JsonPropertyName("comments")]
    public List<CommentData> Comments { get; set; } = new();
    [JsonPropertyName("images")]
    public List<ImageData> Images { get; set; } = new();
    [JsonPropertyName("shapes")]
    public List<ShapeData> Shapes { get; set; } = new();
    [JsonPropertyName("objects")]
    public List<ObjectData> Objects { get; set; } = new();
    [JsonPropertyName("protection")]
    public ProtectionData? Protection { get; set; }
    [JsonPropertyName("auto_filter")]
    public AutoFilterData? AutoFilter { get; set; }
    [JsonPropertyName("freeze_panes")]
    public FreezePanesData? FreezePanes { get; set; }
    [JsonPropertyName("print_titles")]
    public PrintTitlesData? PrintTitles { get; set; }
    [JsonPropertyName("print_area")]
    public string? PrintArea { get; set; }
    [JsonPropertyName("page_breaks")]
    public List<PageBreakData> PageBreaks { get; set; } = new();
    [JsonPropertyName("sheet_format_properties")]
    public SheetFormatPropertiesData? SheetFormatProperties { get; set; }
}

public class PageSetupData
{
    [JsonPropertyName("paper_size")]
    public int? PaperSize { get; set; }
    [JsonPropertyName("orientation")]
    public string? Orientation { get; set; }
    [JsonPropertyName("scale")]
    public int? Scale { get; set; }
    [JsonPropertyName("fit_to_width")]
    public int? FitToWidth { get; set; }
    [JsonPropertyName("fit_to_height")]
    public int? FitToHeight { get; set; }
    [JsonPropertyName("page_height")]
    public double? PageHeight { get; set; }
    [JsonPropertyName("page_width")]
    public double? PageWidth { get; set; }
    [JsonPropertyName("black_and_white")]
    public bool? BlackAndWhite { get; set; }
    [JsonPropertyName("draft")]
    public bool? Draft { get; set; }
    [JsonPropertyName("cell_comments")]
    public string? CellComments { get; set; }
    [JsonPropertyName("errors")]
    public string? Errors { get; set; }
    [JsonPropertyName("horizontal_dpi")]
    public int? HorizontalDpi { get; set; }
    [JsonPropertyName("vertical_dpi")]
    public int? VerticalDpi { get; set; }
    [JsonPropertyName("copies")]
    public int? Copies { get; set; }
    [JsonPropertyName("first_page_number")]
    public int? FirstPageNumber { get; set; }
    [JsonPropertyName("use_first_page_number")]
    public bool? UseFirstPageNumber { get; set; }
    [JsonPropertyName("page_order")]
    public string? PageOrder { get; set; }
    [JsonPropertyName("all_attributes")]
    public Dictionary<string, object?> AllAttributes { get; set; } = new();
}

public class PageMarginsData
{
    [JsonPropertyName("top")]
    public double Top { get; set; }
    [JsonPropertyName("bottom")]
    public double Bottom { get; set; }
    [JsonPropertyName("left")]
    public double Left { get; set; }
    [JsonPropertyName("right")]
    public double Right { get; set; }
    [JsonPropertyName("header")]
    public double Header { get; set; }
    [JsonPropertyName("footer")]
    public double Footer { get; set; }
}

public class ColumnData
{
    [JsonPropertyName("min")]
    public int Min { get; set; }
    [JsonPropertyName("max")]
    public int Max { get; set; }
    [JsonPropertyName("width")]
    public double Width { get; set; }
    [JsonPropertyName("custom_width")]
    public bool? CustomWidth { get; set; }
    [JsonPropertyName("hidden")]
    public bool Hidden { get; set; }
    [JsonPropertyName("best_fit")]
    public bool? BestFit { get; set; }
    [JsonPropertyName("style")]
    public int? Style { get; set; }
}

public class RowData
{
    [JsonPropertyName("row_index")]
    public int RowIndex { get; set; }
    [JsonPropertyName("height")]
    public double? Height { get; set; }
    [JsonPropertyName("custom_height")]
    public bool? CustomHeight { get; set; }
    [JsonPropertyName("hidden")]
    public bool Hidden { get; set; }
    [JsonPropertyName("collapsed")]
    public bool? Collapsed { get; set; }
    [JsonPropertyName("outline_level")]
    public int? OutlineLevel { get; set; }
    [JsonPropertyName("style")]
    public int? Style { get; set; }
}

public class CellData
{
    [JsonPropertyName("reference")]
    public string Reference { get; set; } = "";
    [JsonPropertyName("row")]
    public int Row { get; set; }
    [JsonPropertyName("column")]
    public int Column { get; set; }
    [JsonPropertyName("value")]
    public object? Value { get; set; }
    [JsonPropertyName("formula")]
    public string? Formula { get; set; }
    [JsonPropertyName("type")]
    public string? Type { get; set; }
    [JsonPropertyName("style_index")]
    public int? StyleIndex { get; set; }
    [JsonPropertyName("data_type")]
    public string? DataType { get; set; }
    [JsonPropertyName("show_formula")]
    public bool? ShowFormula { get; set; }
    [JsonPropertyName("number_format")]
    public string? NumberFormat { get; set; }
}

public class MergedCellData
{
    [JsonPropertyName("reference")]
    public string Reference { get; set; } = "";
    [JsonPropertyName("start_row")]
    public int StartRow { get; set; }
    [JsonPropertyName("start_column")]
    public int StartColumn { get; set; }
    [JsonPropertyName("end_row")]
    public int EndRow { get; set; }
    [JsonPropertyName("end_column")]
    public int EndColumn { get; set; }
}

public class ConditionalFormatData
{
    [JsonPropertyName("rule_type")]
    public string? RuleType { get; set; }
    [JsonPropertyName("priority")]
    public int Priority { get; set; }
    [JsonPropertyName("formula")]
    public string? Formula { get; set; }
    [JsonPropertyName("text")]
    public string? Text { get; set; }
    [JsonPropertyName("operator")]
    public string? Operator { get; set; }
    [JsonPropertyName("dxf_id")]
    public int? DxfId { get; set; }
    [JsonPropertyName("all_attributes")]
    public Dictionary<string, object?> AllAttributes { get; set; } = new();
}

public class DataValidationData
{
    [JsonPropertyName("type")]
    public string? Type { get; set; }
    [JsonPropertyName("error_style")]
    public string? ErrorStyle { get; set; }
    [JsonPropertyName("operator")]
    public string? Operator { get; set; }
    [JsonPropertyName("formula1")]
    public string? Formula1 { get; set; }
    [JsonPropertyName("formula2")]
    public string? Formula2 { get; set; }
    [JsonPropertyName("allow_blank")]
    public bool? AllowBlank { get; set; }
    [JsonPropertyName("show_input_message")]
    public bool? ShowInputMessage { get; set; }
    [JsonPropertyName("show_error_message")]
    public bool? ShowErrorMessage { get; set; }
    [JsonPropertyName("error_title")]
    public string? ErrorTitle { get; set; }
    [JsonPropertyName("error")]
    public string? Error { get; set; }
    [JsonPropertyName("prompt_title")]
    public string? PromptTitle { get; set; }
    [JsonPropertyName("prompt")]
    public string? Prompt { get; set; }
    [JsonPropertyName("sqref")]
    public string? SqRef { get; set; }
}

public class HyperlinkData
{
    [JsonPropertyName("reference")]
    public string Reference { get; set; } = "";
    [JsonPropertyName("id")]
    public string? Id { get; set; }
    [JsonPropertyName("location")]
    public string? Location { get; set; }
    [JsonPropertyName("tooltip")]
    public string? Tooltip { get; set; }
    [JsonPropertyName("display")]
    public string? Display { get; set; }
    [JsonPropertyName("target")]
    public string? Target { get; set; }
}

public class CommentData
{
    [JsonPropertyName("reference")]
    public string Reference { get; set; } = "";
    [JsonPropertyName("author")]
    public string? Author { get; set; }
    [JsonPropertyName("text")]
    public string? Text { get; set; }
    [JsonPropertyName("visible")]
    public bool? Visible { get; set; }
    [JsonPropertyName("width")]
    public int? Width { get; set; }
    [JsonPropertyName("height")]
    public int? Height { get; set; }
}

public class ImageData
{
    [JsonPropertyName("name")]
    public string? Name { get; set; }
    [JsonPropertyName("content_type")]
    public string? ContentType { get; set; }
    [JsonPropertyName("format")]
    public string? Format { get; set; }
    [JsonPropertyName("width")]
    public long? Width { get; set; }
    [JsonPropertyName("height")]
    public long? Height { get; set; }
    [JsonPropertyName("horizontal_resolution")]
    public long? HorizontalResolution { get; set; }
    [JsonPropertyName("vertical_resolution")]
    public long? VerticalResolution { get; set; }
    [JsonPropertyName("embed_id")]
    public string? EmbedId { get; set; }
    [JsonPropertyName("data")]
    public string? DataBase64 { get; set; }
}

public class ShapeData
{
    [JsonPropertyName("name")]
    public string? Name { get; set; }
    [JsonPropertyName("type")]
    public string? Type { get; set; }
    [JsonPropertyName("anchor")]
    public string? Anchor { get; set; }
    [JsonPropertyName("placement")]
    public string? Placement { get; set; }
    [JsonPropertyName("left")]
    public double? Left { get; set; }
    [JsonPropertyName("top")]
    public double? Top { get; set; }
    [JsonPropertyName("width")]
    public double? Width { get; set; }
    [JsonPropertyName("height")]
    public double? Height { get; set; }
    [JsonPropertyName("rotation")]
    public double? Rotation { get; set; }
    [JsonPropertyName("fill_color")]
    public string? FillColor { get; set; }
    [JsonPropertyName("line_color")]
    public string? LineColor { get; set; }
    [JsonPropertyName("line_width")]
    public double? LineWidth { get; set; }
    [JsonPropertyName("text")]
    public string? Text { get; set; }
    [JsonPropertyName("all_attributes")]
    public Dictionary<string, object?> AllAttributes { get; set; } = new();
}

public class ObjectData
{
    [JsonPropertyName("name")]
    public string? Name { get; set; }
    [JsonPropertyName("type")]
    public string? Type { get; set; }
    [JsonPropertyName("all_attributes")]
    public Dictionary<string, object?> AllAttributes { get; set; } = new();
}

public class AutoFilterData
{
    [JsonPropertyName("reference")]
    public string? Reference { get; set; }
    [JsonPropertyName("all_attributes")]
    public Dictionary<string, object?> AllAttributes { get; set; } = new();
}

public class FreezePanesData
{
    [JsonPropertyName("row")]
    public int? Row { get; set; }
    [JsonPropertyName("column")]
    public int? Column { get; set; }
}

public class PrintTitlesData
{
    [JsonPropertyName("columns_to_repeat")]
    public string? ColumnsToRepeat { get; set; }
    [JsonPropertyName("rows_to_repeat")]
    public string? RowsToRepeat { get; set; }
}

public class PageBreakData
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "";
    [JsonPropertyName("row")]
    public int? Row { get; set; }
    [JsonPropertyName("column")]
    public int? Column { get; set; }
}

public class SheetFormatPropertiesData
{
    [JsonPropertyName("default_row_height")]
    public double? DefaultRowHeight { get; set; }
    [JsonPropertyName("default_column_width")]
    public double? DefaultColumnWidth { get; set; }
    [JsonPropertyName("outline_level_row")]
    public int? OutlineLevelRow { get; set; }
    [JsonPropertyName("outline_level_column")]
    public int? OutlineLevelColumn { get; set; }
}

public class StyleData
{
    [JsonPropertyName("style_index")]
    public int StyleIndex { get; set; }
    [JsonPropertyName("font")]
    public FontData? Font { get; set; }
    [JsonPropertyName("fill")]
    public FillData? Fill { get; set; }
    [JsonPropertyName("border")]
    public BorderData? Border { get; set; }
    [JsonPropertyName("alignment")]
    public AlignmentData? Alignment { get; set; }
    [JsonPropertyName("number_format")]
    public NumberFormatData? NumberFormat { get; set; }
    [JsonPropertyName("builtin_id")]
    public int? BuiltinId { get; set; }
    [JsonPropertyName("custom_builtin")]
    public bool? CustomBuiltin { get; set; }
    [JsonPropertyName("hidden")]
    public bool? Hidden { get; set; }
    [JsonPropertyName("outline_style")]
    public bool? OutlineStyle { get; set; }
    [JsonPropertyName("all_attributes")]
    public Dictionary<string, object?> AllAttributes { get; set; } = new();
}

public class FontData
{
    [JsonPropertyName("bold")]
    public bool? Bold { get; set; }
    [JsonPropertyName("italic")]
    public bool? Italic { get; set; }
    [JsonPropertyName("underline")]
    public string? Underline { get; set; }
    [JsonPropertyName("strikethrough")]
    public bool? Strikethrough { get; set; }
    [JsonPropertyName("size")]
    public double? Size { get; set; }
    [JsonPropertyName("color")]
    public string? Color { get; set; }
    [JsonPropertyName("color_index")]
    public int? ColorIndex { get; set; }
    [JsonPropertyName("color_theme")]
    public int? ColorTheme { get; set; }
    [JsonPropertyName("color_tint")]
    public double? ColorTint { get; set; }
    [JsonPropertyName("font_name")]
    public string? FontName { get; set; }
    [JsonPropertyName("font_family")]
    public int? FontFamily { get; set; }
    [JsonPropertyName("font_scheme")]
    public string? FontScheme { get; set; }
    [JsonPropertyName("charset")]
    public int? Charset { get; set; }
    [JsonPropertyName("condense")]
    public bool? Condense { get; set; }
    [JsonPropertyName("extend")]
    public bool? Extend { get; set; }
    [JsonPropertyName("vertical_align")]
    public string? VerticalAlign { get; set; }
}

public class FillData
{
    [JsonPropertyName("pattern_type")]
    public string? PatternType { get; set; }
    [JsonPropertyName("foreground_color")]
    public string? ForegroundColor { get; set; }
    [JsonPropertyName("foreground_color_index")]
    public int? ForegroundColorIndex { get; set; }
    [JsonPropertyName("foreground_color_theme")]
    public int? ForegroundColorTheme { get; set; }
    [JsonPropertyName("foreground_color_tint")]
    public double? ForegroundColorTint { get; set; }
    [JsonPropertyName("background_color")]
    public string? BackgroundColor { get; set; }
    [JsonPropertyName("background_color_index")]
    public int? BackgroundColorIndex { get; set; }
    [JsonPropertyName("background_color_theme")]
    public int? BackgroundColorTheme { get; set; }
    [JsonPropertyName("background_color_tint")]
    public double? BackgroundColorTint { get; set; }
    [JsonPropertyName("gradient")]
    public GradientData? Gradient { get; set; }
}

public class GradientData
{
    [JsonPropertyName("type")]
    public string? Type { get; set; }
    [JsonPropertyName("degree")]
    public double? Degree { get; set; }
    [JsonPropertyName("stops")]
    public List<GradientStopData> Stops { get; set; } = new();
}

public class GradientStopData
{
    [JsonPropertyName("position")]
    public double Position { get; set; }
    [JsonPropertyName("color")]
    public string? Color { get; set; }
    [JsonPropertyName("color_index")]
    public int? ColorIndex { get; set; }
    [JsonPropertyName("color_theme")]
    public int? ColorTheme { get; set; }
    [JsonPropertyName("color_tint")]
    public double? ColorTint { get; set; }
}

public class BorderData
{
    [JsonPropertyName("left")]
    public BorderEdgeData? Left { get; set; }
    [JsonPropertyName("right")]
    public BorderEdgeData? Right { get; set; }
    [JsonPropertyName("top")]
    public BorderEdgeData? Top { get; set; }
    [JsonPropertyName("bottom")]
    public BorderEdgeData? Bottom { get; set; }
    [JsonPropertyName("diagonal")]
    public BorderEdgeData? Diagonal { get; set; }
    [JsonPropertyName("diagonal_up")]
    public bool? DiagonalUp { get; set; }
    [JsonPropertyName("diagonal_down")]
    public bool? DiagonalDown { get; set; }
}

public class BorderEdgeData
{
    [JsonPropertyName("style")]
    public string? Style { get; set; }
    [JsonPropertyName("color")]
    public string? Color { get; set; }
    [JsonPropertyName("color_index")]
    public int? ColorIndex { get; set; }
    [JsonPropertyName("color_theme")]
    public int? ColorTheme { get; set; }
    [JsonPropertyName("color_tint")]
    public double? ColorTint { get; set; }
}

public class AlignmentData
{
    [JsonPropertyName("horizontal")]
    public string? Horizontal { get; set; }
    [JsonPropertyName("vertical")]
    public string? Vertical { get; set; }
    [JsonPropertyName("text_rotation")]
    public int? TextRotation { get; set; }
    [JsonPropertyName("wrap_text")]
    public bool? WrapText { get; set; }
    [JsonPropertyName("indent")]
    public int? Indent { get; set; }
    [JsonPropertyName("justify_last_line")]
    public bool? JustifyLastLine { get; set; }
    [JsonPropertyName("shrink_to_fit")]
    public bool? ShrinkToFit { get; set; }
    [JsonPropertyName("reading_order")]
    public int? ReadingOrder { get; set; }
}

public class NumberFormatData
{
    [JsonPropertyName("format_code")]
    public string? FormatCode { get; set; }
}

public class ExtractionResult
{
    [JsonPropertyName("workbook")]
    public WorkbookData Workbook { get; set; } = new();
    [JsonPropertyName("styles")]
    public List<StyleData> Styles { get; set; } = new();
    [JsonPropertyName("extracted_at")]
    public DateTime ExtractedAt { get; set; } = DateTime.UtcNow;
    [JsonPropertyName("source_file")]
    public string? SourceFile { get; set; }
    [JsonPropertyName("source")]
    public string Source { get; set; } = "generated";
}
