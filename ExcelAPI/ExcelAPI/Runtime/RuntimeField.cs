using System.Text.Json.Serialization;

namespace ExcelAPI.Runtime
{
    /// <summary>
    /// Represents a single editable field in the runtime form.
    /// Produced by the FormRuntimeBuilder and consumed by the Next.js frontend.
    /// Contains all metadata needed for rendering an editable overlay field.
    /// </summary>
    public class RuntimeField
    {
        /// <summary>Unique internal field identifier (e.g., "page1field1"). Never shown to users.</summary>
        public string Id { get; set; } = "";

        /// <summary>User-visible field name (from comment first line or default). Never duplicate IDs.</summary>
        public string Name { get; set; } = "";

        /// <summary>Excel cell reference (e.g., "A1", "B2").</summary>
        public string CellReference { get; set; } = "";

        /// <summary>Row index (1-based).</summary>
        public uint Row { get; set; }

        /// <summary>Column index (1-based).</summary>
        public uint Column { get; set; }

        // ── Pixel Coordinates (matches rendering engine) ───────────────

        /// <summary>Left edge in pixels (page-relative).</summary>
        public double LeftPx { get; set; }

        /// <summary>Top edge in pixels (page-relative).</summary>
        public double TopPx { get; set; }

        /// <summary>Width in pixels.</summary>
        public double WidthPx { get; set; }

        /// <summary>Height in pixels.</summary>
        public double HeightPx { get; set; }

        // ── Ratio-Based Coordinates (legacy ConMas compatibility) ──────

        /// <summary>Left edge as ratio of page width (0-1).</summary>
        public double LeftRatio { get; set; }

        /// <summary>Top edge as ratio of page height (0-1).</summary>
        public double TopRatio { get; set; }

        /// <summary>Width as ratio of page width (0-1).</summary>
        public double WidthRatio { get; set; }

        /// <summary>Height as ratio of page height (0-1).</summary>
        public double HeightRatio { get; set; }

        // ── Merged Cell Info ──────────────────────────────────────────

        /// <summary>Merge range reference (e.g., "A1:C3"), empty if not merged.</summary>
        public string? MergeRange { get; set; }

        /// <summary>Whether this cell is part of a merged range.</summary>
        public bool IsMerged { get; set; }

        // ── Data Type ─────────────────────────────────────────────────

        /// <summary>Data type: "text", "number", "date", "checkbox", "signature", "dropdown", "calculated".</summary>
        public string DataType { get; set; } = "text";

        /// <summary>Whether the field is read-only.</summary>
        public bool ReadOnly { get; set; }

        /// <summary>Whether the field is required.</summary>
        public bool Required { get; set; }

        // ── Style / Display ───────────────────────────────────────────

        /// <summary>Horizontal alignment: "left", "center", "right", "general".</summary>
        public string? Alignment { get; set; }

        /// <summary>Font name.</summary>
        public string? Font { get; set; }

        /// <summary>Font size in points.</summary>
        public double FontSize { get; set; } = 11;

        /// <summary>Whether the font is bold.</summary>
        public bool Bold { get; set; }

        /// <summary>Font color in #AARRGGBB format.</summary>
        public string? FontColor { get; set; }

        /// <summary>Background color in #AARRGGBB format.</summary>
        public string? BackgroundColor { get; set; }

        /// <summary>Border style: "none", "thin", "medium", "thick", "double".</summary>
        public string? Border { get; set; }

        // ── Input Constraints ─────────────────────────────────────────

        /// <summary>Placeholder text (from cell value or comment).</summary>
        public string? Placeholder { get; set; }

        /// <summary>Default value (from cell value).</summary>
        public string? DefaultValue { get; set; }

        /// <summary>Maximum input length.</summary>
        public int MaxLength { get; set; }

        /// <summary>Tab index for keyboard navigation.</summary>
        public int TabIndex { get; set; }

        /// <summary>Validation pattern (regex) for input validation.</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? ValidationPattern { get; set; }

        /// <summary>Validation message shown when pattern fails.</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? ValidationMessage { get; set; }
    }
}
