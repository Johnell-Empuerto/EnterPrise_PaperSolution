// ────────────────────────────────────────────────────────────────────────────
// SheetDefinition — Worksheet in the Canonical Model
//
// Describes a single worksheet: its layout, rows, columns, named ranges,
// comments, merged ranges, and all interactive fields.
//
// Ownership: Shared (Designer → Runtime + Rendering)
// ────────────────────────────────────────────────────────────────────────────

namespace ExcelAPI.Models.WorkbookDefinition
{
    /// <summary>
    /// A single worksheet in the workbook definition.
    /// Contains all structural elements needed to understand the template.
    /// </summary>
    public class SheetDefinition
    {
        /// <summary>Unique sheet identifier within the workbook (e.g., "sheet_0").</summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>Sheet name as it appears in Excel (e.g., "Sheet1", "Invoice").</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>0-based index in the workbook's worksheet list.</summary>
        public int Index { get; set; }

        /// <summary>Print layout configuration (paper, margins, scaling, print area).</summary>
        public PrintLayout PrintLayout { get; set; } = new();

        /// <summary>Interactive form fields detected on this sheet.</summary>
        public List<FieldDefinition> Fields { get; set; } = new();

        /// <summary>Embedded images (pictures) on this sheet.</summary>
        public List<ImageDefinition> Images { get; set; } = new();

        /// <summary>Drawing shapes (rectangles, text boxes, arrows, etc.) on this sheet.</summary>
        public List<ShapeDefinition> Shapes { get; set; } = new();

        /// <summary>Defined names scoped to this sheet (or workbook).</summary>
        public List<NamedRangeDefinition> NamedRanges { get; set; } = new();

        /// <summary>Cell comments/notes on this sheet.</summary>
        public List<CommentDefinition> Comments { get; set; } = new();

        /// <summary>Merged cell ranges on this sheet.</summary>
        public List<MergedRangeDefinition> MergedRanges { get; set; } = new();

        /// <summary>Row definitions (height, visibility, outline level).</summary>
        public List<RowDefinition> Rows { get; set; } = new();

        /// <summary>Column definitions (width, visibility, outline level).</summary>
        public List<ColumnDefinition> Columns { get; set; } = new();

        /// <summary>
        /// Freeze pane cell (e.g., "B2" means rows above and columns left are frozen).
        /// Null if no freeze panes.
        /// </summary>
        public string? FreezePane { get; set; }

        /// <summary>
        /// Default column width in character units for this sheet.
        /// </summary>
        public double DefaultColumnWidthChars { get; set; } = 8.43;

        /// <summary>
        /// Default row height in points for this sheet.
        /// </summary>
        public double DefaultRowHeightPt { get; set; } = 15.0;

        /// <summary>
        /// Whether the sheet is visible in the Excel UI.
        /// </summary>
        public bool IsVisible { get; set; } = true;
    }

    /// <summary>
    /// Row dimension on a worksheet.
    /// </summary>
    public class RowDefinition
    {
        /// <summary>1-based row index.</summary>
        public uint Index { get; set; }

        /// <summary>Row height in points (null = default height).</summary>
        public double? HeightPt { get; set; }

        /// <summary>Whether the row is hidden.</summary>
        public bool Hidden { get; set; }

        /// <summary>Whether a custom height has been set.</summary>
        public bool CustomHeight { get; set; }

        /// <summary>Outline/group level (0 = not grouped).</summary>
        public uint OutlineLevel { get; set; }
    }

    /// <summary>
    /// Column dimension on a worksheet.
    /// </summary>
    public class ColumnDefinition
    {
        /// <summary>1-based column index.</summary>
        public uint Index { get; set; }

        /// <summary>Column width in character units (Excel's standard unit).</summary>
        public double WidthChars { get; set; }

        /// <summary>Column width in points (computed from character width at max digit width).</summary>
        public double WidthPt { get; set; }

        /// <summary>Whether the column is hidden.</summary>
        public bool Hidden { get; set; }

        /// <summary>Whether a custom width has been set.</summary>
        public bool CustomWidth { get; set; }

        /// <summary>Outline/group level (0 = not grouped).</summary>
        public uint OutlineLevel { get; set; }
    }

    /// <summary>
    /// A named range (defined name) in the workbook.
    /// </summary>
    public class NamedRangeDefinition
    {
        /// <summary>The defined name (e.g., "Print_Area", "MyRange").</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>The range address the name refers to (e.g., "Sheet1!$A$1:$H$40").</summary>
        public string Address { get; set; } = string.Empty;

        /// <summary>
        /// Scope of the name: "Workbook" for workbook-level names,
        /// or the sheet name for sheet-scoped names.
        /// </summary>
        public string Scope { get; set; } = "Workbook";

        /// <summary>Optional comment/description.</summary>
        public string? Comment { get; set; }
    }

    /// <summary>
    /// A cell comment/note.
    /// </summary>
    public class CommentDefinition
    {
        /// <summary>A1-style address of the commented cell (e.g., "B5").</summary>
        public string Address { get; set; } = string.Empty;

        /// <summary>Comment author (from Excel).</summary>
        public string Author { get; set; } = string.Empty;

        /// <summary>Comment text content.</summary>
        public string Text { get; set; } = string.Empty;

        /// <summary>When the comment was added (if available).</summary>
        public DateTime? DateTime { get; set; }

        /// <summary>Whether the comment is visible.</summary>
        public bool IsVisible { get; set; } = true;
    }

    /// <summary>
    /// A merged cell range.
    /// </summary>
    public class MergedRangeDefinition
    {
        /// <summary>Merge range address (e.g., "A1:C3").</summary>
        public string Address { get; set; } = string.Empty;

        /// <summary>The cell range of the merge.</summary>
        public CellRange? Range { get; set; }

        /// <summary>Merge bounds in points.</summary>
        public Rectangle? BoundsPt { get; set; }
    }
}
