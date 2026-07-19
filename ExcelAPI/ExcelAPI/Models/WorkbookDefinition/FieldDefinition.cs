// ────────────────────────────────────────────────────────────────────────────
// FieldDefinition — Interactive Form Field
//
// Describes an interactive/form field on the worksheet. Each field corresponds
// to a cell (or merged cell range) that the user can fill in.
//
// Ownership: Shared (Designer → Runtime)
// ────────────────────────────────────────────────────────────────────────────

namespace ExcelAPI.Models.WorkbookDefinition
{
    /// <summary>
    /// An interactive form field on a worksheet.
    /// Produced by the Designer from cell comments or _Fields sheet metadata.
    /// Consumed by the Runtime to produce RuntimeField.
    /// </summary>
    public class FieldDefinition
    {
        /// <summary>Unique field identifier within the workbook (e.g., "p1f1").</summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>User-visible field name (from comment first line or _Fields sheet).</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>The Excel cell this field is anchored to.</summary>
        public CellReference Cell { get; set; } = new();

        /// <summary>Field bounds in points (worksheet-relative, from Range.Left/Top/Width/Height).</summary>
        public Rectangle BoundsPt { get; set; } = new();

        /// <summary>Field bounds as ratio of page dimensions (legacy ConMas compatibility).</summary>
        public RatioRectangle? BoundsRatio { get; set; }

        /// <summary>Field type (determines UI widget).</summary>
        public FieldType Type { get; set; } = FieldType.Text;

        /// <summary>Cell style at this field's position (font, border, fill, alignment).</summary>
        public CellStyle Style { get; set; } = new();

        /// <summary>Whether the field is visible on the form.</summary>
        public bool Visible { get; set; } = true;

        /// <summary>Whether the field is locked (read-only).</summary>
        public bool Locked { get; set; }

        /// <summary>Cell formula expression (if calculated field).</summary>
        public string? Formula { get; set; }

        /// <summary>Data validation rules for this field.</summary>
        public DataValidationDefinition? DataValidation { get; set; }

        /// <summary>Placeholder text (from cell value or default).</summary>
        public string? Placeholder { get; set; }

        /// <summary>Default value for the field.</summary>
        public string? DefaultValue { get; set; }

        /// <summary>Maximum input length.</summary>
        public int MaxLength { get; set; }

        /// <summary>Tab index for keyboard navigation.</summary>
        public int TabIndex { get; set; }

        /// <summary>
        /// Whether the field is required (must be filled before submit).
        /// </summary>
        public bool Required { get; set; }

        /// <summary>
        /// User-entered value for this field (set by the frontend editor).
        /// Written back to the original XLSX by WorkbookValueWriter.
        /// </summary>
        public string? Value { get; set; }

        /// <summary>
        /// Custom metadata key-value pairs for this field.
        /// </summary>
        public Dictionary<string, string> Metadata { get; set; } = new();

        /// <summary>
        /// Merge info if this field's cell is part of a merged range.
        /// </summary>
        public MergedFieldInfo? MergeInfo { get; set; }
    }

    /// <summary>
    /// The type of an interactive field.
    /// Maps to the frontend's overlay type.
    /// </summary>
    public enum FieldType
    {
        /// <summary>Free-form text input.</summary>
        Text,
        /// <summary>Numeric input.</summary>
        Number,
        /// <summary>Date picker.</summary>
        Date,
        /// <summary>Checkbox / boolean toggle.</summary>
        Checkbox,
        /// <summary>Signature pad.</summary>
        Signature,
        /// <summary>Dropdown / combo box selection.</summary>
        Dropdown,
        /// <summary>Auto-calculated field (from a formula).</summary>
        Calculated
    }

    /// <summary>
    /// Information about a merged cell that this field belongs to.
    /// </summary>
    public class MergedFieldInfo
    {
        /// <summary>Whether this field's cell is part of a merged range.</summary>
        public bool IsMerged { get; set; }

        /// <summary>Merge range address (e.g., "A1:C3").</summary>
        public string? MergeAddress { get; set; }

        /// <summary>Bounds of the entire merged range in points.</summary>
        public Rectangle? MergeBoundsPt { get; set; }
    }

    /// <summary>
    /// Data validation rules applied to a cell/field.
    /// Maps to Excel's Data Validation feature.
    /// </summary>
    public class DataValidationDefinition
    {
        /// <summary>
        /// Validation type. One of:
        /// "any", "whole", "decimal", "list", "date", "time", "textLength", "custom".
        /// </summary>
        public string Type { get; set; } = "any";

        /// <summary>
        /// Comparison operator. One of:
        /// "between", "notBetween", "equal", "notEqual", "greaterThan",
        /// "lessThan", "greaterOrEqual", "lessOrEqual".
        /// </summary>
        public string? Operator { get; set; }

        /// <summary>First formula/value for validation.</summary>
        public string? Formula1 { get; set; }

        /// <summary>Second formula/value (used for "between" and "notBetween").</summary>
        public string? Formula2 { get; set; }

        /// <summary>Whether blank values are allowed.</summary>
        public bool AllowBlank { get; set; } = true;

        /// <summary>Whether to show the input prompt message.</summary>
        public bool ShowInputMessage { get; set; }

        /// <summary>Input prompt title.</summary>
        public string? PromptTitle { get; set; }

        /// <summary>Input prompt text.</summary>
        public string? PromptText { get; set; }

        /// <summary>Whether to show the error alert.</summary>
        public bool ShowErrorMessage { get; set; }

        /// <summary>Error alert title.</summary>
        public string? ErrorTitle { get; set; }

        /// <summary>Error alert text.</summary>
        public string? ErrorText { get; set; }

        /// <summary>
        /// Error alert style: "stop", "warning", "information".
        /// </summary>
        public string? ErrorStyle { get; set; }
    }
}
