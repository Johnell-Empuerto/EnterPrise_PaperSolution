// ────────────────────────────────────────────────────────────────────────────
// DesignerModel — Comprehensive Designer State for Round-Trip Editing
//
// Phase 20: The single source of truth for the entire editing session.
// The workbook is the serialized form of this model.
// Serialization and deserialization must be deterministic.
// ────────────────────────────────────────────────────────────────────────────

using ExcelAPI.Models.WorkbookDefinition;

namespace ExcelAPI.Models
{
    /// <summary>
    /// Root model for the complete designer state.
    /// Represents everything needed to reconstruct the editing session.
    ///
    /// Architecture:
    ///   DesignerModel
    ///   ├── Workbook meta (Info, SessionId)
    ///   ├── Pages (one per printable worksheet)
    ///   │   ├── Page metadata (dimensions, print area, page setup)
    ///   │   └── Fields (interactive form fields with formatting)
    ///   ├── Config sheets (_Fields, ExcelOutputSetting)
    ///   └── Comments (all cell comments across all sheets)
    ///
    /// This is the model sent to the frontend for editor reconstruction.
    /// </summary>
    public class DesignerModel
    {
        /// <summary>Workbook metadata (title, author, dates).</summary>
        public DesignerWorkbookInfo Info { get; set; } = new();

        /// <summary>Server-side session identifier.</summary>
        public string SessionId { get; set; } = string.Empty;

        /// <summary>Printable pages in the workbook (one per non-config worksheet).</summary>
        public List<DesignerPage> Pages { get; set; } = new();

        /// <summary>Configuration sheet metadata.</summary>
        public DesignerConfiguration Configuration { get; set; } = new();

        /// <summary>All cell comments across all sheets.</summary>
        public List<DesignerComment> Comments { get; set; } = new();
    }

    /// <summary>Workbook-level metadata for the designer.</summary>
    public class DesignerWorkbookInfo
    {
        public string Title { get; set; } = "Untitled";
        public string Author { get; set; } = "";
        public string Description { get; set; } = "";
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        public string Version { get; set; } = "1.0";
    }

    /// <summary>A single printable page in the designer.</summary>
    public class DesignerPage
    {
        /// <summary>Unique page identifier.</summary>
        public string Id { get; set; } = "";

        /// <summary>Worksheet name in Excel.</summary>
        public string Name { get; set; } = "";

        /// <summary>0-based index.</summary>
        public int Index { get; set; }

        /// <summary>Page layout and print settings.</summary>
        public DesignerPageLayout Layout { get; set; } = new();

        /// <summary>Interactive fields on this page.</summary>
        public List<DesignerField> Fields { get; set; } = new();

        /// <summary>Background image URL or path.</summary>
        public string? BackgroundImage { get; set; }

        /// <summary>Background image dimensions (pixels).</summary>
        public int BackgroundWidth { get; set; }

        /// <summary>Background image dimensions (pixels).</summary>
        public int BackgroundHeight { get; set; }
    }

    /// <summary>Page layout and print settings.</summary>
    public class DesignerPageLayout
    {
        /// <summary>Print area address (e.g., "A1:H40").</summary>
        public string PrintArea { get; set; } = "";

        /// <summary>Paper size name (e.g., "Letter", "A4").</summary>
        public string PaperSize { get; set; } = "Letter";

        /// <summary>Page orientation.</summary>
        public string Orientation { get; set; } = "portrait";

        /// <summary>Page width in points.</summary>
        public double WidthPt { get; set; }

        /// <summary>Page height in points.</summary>
        public double HeightPt { get; set; }

        /// <summary>Margins in points.</summary>
        public DesignerMargins Margins { get; set; } = new();

        /// <summary>Zoom percentage (100 = 100%).</summary>
        public int Zoom { get; set; } = 100;

        /// <summary>Fit to page settings.</summary>
        public int FitToPagesWide { get; set; }

        /// <summary>Fit to page settings.</summary>
        public int FitToPagesTall { get; set; }

        /// <summary>Center horizontally on page.</summary>
        public bool CenterHorizontally { get; set; }

        /// <summary>Center vertically on page.</summary>
        public bool CenterVertically { get; set; }

        /// <summary>Row definitions (height, visibility).</summary>
        public List<DesignerRowInfo> Rows { get; set; } = new();

        /// <summary>Column definitions (width, visibility).</summary>
        public List<DesignerColumnInfo> Columns { get; set; } = new();

        /// <summary>Merged cell ranges.</summary>
        public List<string> MergedCells { get; set; } = new();
    }

    /// <summary>Page margins in points.</summary>
    public class DesignerMargins
    {
        public double Top { get; set; } = 54;
        public double Bottom { get; set; } = 54;
        public double Left { get; set; } = 50.4;
        public double Right { get; set; } = 50.4;
        public double Header { get; set; }
        public double Footer { get; set; }
    }

    /// <summary>Row information.</summary>
    public class DesignerRowInfo
    {
        public int Index { get; set; }
        public double HeightPt { get; set; }
        public bool Hidden { get; set; }
    }

    /// <summary>Column information.</summary>
    public class DesignerColumnInfo
    {
        public int Index { get; set; }
        public double WidthChars { get; set; }
        public bool Hidden { get; set; }
    }

    /// <summary>An interactive form field on a page.</summary>
    public class DesignerField
    {
        /// <summary>Unique field identifier (e.g., "p1f1").</summary>
        public string Id { get; set; } = "";

        /// <summary>User-visible field name.</summary>
        public string Name { get; set; } = "";

        /// <summary>Cell address this field is anchored to (e.g., "B5").</summary>
        public string CellAddress { get; set; } = "";

        /// <summary>Field type (text, number, date, checkbox, signature, dropdown).</summary>
        public string Type { get; set; } = "text";

        /// <summary>Position in points (worksheet-relative).</summary>
        public DesignerFieldBounds Bounds { get; set; } = new();

        /// <summary>Visual formatting of the field's cell.</summary>
        public DesignerFieldStyle Style { get; set; } = new();

        /// <summary>Field behavior flags.</summary>
        public DesignerFieldBehavior Behavior { get; set; } = new();

        /// <summary>Data validation rules.</summary>
        public DesignerFieldValidation? Validation { get; set; }

        /// <summary>Current value (user-entered).</summary>
        public string? Value { get; set; }

        /// <summary>Placeholder text.</summary>
        public string? Placeholder { get; set; }

        /// <summary>Default value.</summary>
        public string? DefaultValue { get; set; }

        /// <summary>Maximum input length.</summary>
        public int MaxLength { get; set; }

        /// <summary>Dropdown options (if type is dropdown).</summary>
        public List<string> Options { get; set; } = new();

        /// <summary>Whether the field is visible on the form.</summary>
        public bool Visible { get; set; } = true;

        /// <summary>Field label/description text.</summary>
        public string? Label { get; set; }

        /// <summary>Description / tooltip text.</summary>
        public string? Description { get; set; }
    }

    /// <summary>Field bounds in points.</summary>
    public class DesignerFieldBounds
    {
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
    }

    /// <summary>Visual formatting of a field's cell.</summary>
    public class DesignerFieldStyle
    {
        // Font
        public string? FontFamily { get; set; }
        public double? FontSize { get; set; }
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public bool? Underline { get; set; }
        public string? FontColor { get; set; }
        public string? FillColor { get; set; }

        // Alignment
        public string? HorizontalAlignment { get; set; }
        public string? VerticalAlignment { get; set; }
        public bool? WrapText { get; set; }
        public int? Rotation { get; set; }

        // Number format
        public string? NumberFormat { get; set; }

        // Borders (CSS-style strings like "1px solid #000000")
        public string? BorderTop { get; set; }
        public string? BorderBottom { get; set; }
        public string? BorderLeft { get; set; }
        public string? BorderRight { get; set; }
    }

    /// <summary>Field behavior flags.</summary>
    public class DesignerFieldBehavior
    {
        public bool Required { get; set; }
        public bool ReadOnly { get; set; }
        public string? KeyboardType { get; set; }
    }

    /// <summary>Data validation rules.</summary>
    public class DesignerFieldValidation
    {
        public string Type { get; set; } = "any";
        public string? Operator { get; set; }
        public string? Formula1 { get; set; }
        public string? Formula2 { get; set; }
        public bool AllowBlank { get; set; } = true;
        public string? ErrorMessage { get; set; }
    }

    /// <summary>Configuration sheet metadata.</summary>
    public class DesignerConfiguration
    {
        /// <summary>Whether _Fields sheet exists.</summary>
        public bool HasFieldsSheet { get; set; }

        /// <summary>Whether ExcelOutputSetting sheet exists.</summary>
        public bool HasExcelOutputSetting { get; set; }

        /// <summary>Number of rows in _Fields sheet.</summary>
        public int FieldsSheetRowCount { get; set; }

        /// <summary>Field IDs extracted from _Fields sheet.</summary>
        public List<string> FieldsSheetFieldIds { get; set; } = new();

        /// <summary>Whether configuration is duplicated.</summary>
        public bool IsDuplicated { get; set; }
    }

    /// <summary>A cell comment.</summary>
    public class DesignerComment
    {
        public string CellAddress { get; set; } = "";
        public string Worksheet { get; set; } = "";
        public string Text { get; set; } = "";
        public string Author { get; set; } = "";
    }
}
