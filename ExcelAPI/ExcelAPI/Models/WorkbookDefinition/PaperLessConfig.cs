using System.Text.Json.Serialization;

namespace ExcelAPI.Models.WorkbookDefinition
{
    /// <summary>
    /// Root configuration object serialized as JSON and embedded in the PaperLessConfig worksheet.
    /// Preserves PaperLess-specific field metadata across export → re-upload round trips.
    /// This is NOT a runtime value store — only field identity, style, and configuration.
    /// </summary>
    public class PaperLessConfig
    {
        /// <summary>Schema version for forward/backward compatibility. Must be 1 for this version.</summary>
        public int SchemaVersion { get; set; } = 1;

        /// <summary>PaperLess application metadata.</summary>
        public PaperLessInfo PaperLess { get; set; } = new();

        /// <summary>Per-sheet field configuration.</summary>
        public List<PaperLessSheet> Sheets { get; set; } = new();
    }

    /// <summary>
    /// PaperLess application metadata embedded in the config.
    /// </summary>
    public class PaperLessInfo
    {
        /// <summary>PaperLess runtime version that wrote this config.</summary>
        public string Version { get; set; } = "1.0";
    }

    /// <summary>
    /// Configuration for a single worksheet.
    /// </summary>
    public class PaperLessSheet
    {
        /// <summary>Sheet name matching the Excel worksheet name.</summary>
        public string Name { get; set; } = "";

        /// <summary>Fields on this sheet.</summary>
        public List<PaperLessField> Fields { get; set; } = new();
    }

    /// <summary>
    /// PaperLess-specific field configuration.
    /// Preserves the stable field identity, type, style, and config that
    /// Excel natively cannot represent or cannot preserve across round trips.
    /// </summary>
    public class PaperLessField
    {
        /// <summary>Stable PaperLess field ID (e.g., "p1f1"). Never "samples" or a generated ID.</summary>
        public string Id { get; set; } = "";

        /// <summary>Cell address or range (e.g., "$A$1:$B$2"). Used for matching during re-upload.</summary>
        public string Cell { get; set; } = "";

        /// <summary>Field type name (e.g., "KeyboardText", "KeyboardNumber", "Calendar", "Check", etc.).</summary>
        public string Type { get; set; } = "KeyboardText";

        /// <summary>PaperLess-style configuration (font, fill, alignment). Reuses existing CellStyle model.</summary>
        public CellStyle? Style { get; set; }

        /// <summary>PaperLess input configuration (required, max length, restrictions).</summary>
        public PaperLessFieldConfig? Config { get; set; }
    }

    /// <summary>
    /// PaperLess-specific input configuration for a field.
    /// These are PaperLess editorial settings, not native Excel properties.
    /// </summary>
    public class PaperLessFieldConfig
    {
        /// <summary>Whether the field is required before submit.</summary>
        public bool Required { get; set; }

        /// <summary>Minimum input length.</summary>
        public int MinLength { get; set; }

        /// <summary>Maximum input length (0 = unlimited).</summary>
        public int MaxLength { get; set; }

        /// <summary>Input restriction type: "None", "Numeric", "Date", etc.</summary>
        public string InputRestriction { get; set; } = "None";

        /// <summary>Number of visible lines for multi-line text fields.</summary>
        public int Lines { get; set; } = 1;

        /// <summary>Whether the field requires validation during editing (not just on submit).</summary>
        public bool ValidateOnEditing { get; set; }

        /// <summary>Whether the field is read-only.</summary>
        public bool ReadOnly { get; set; }

        /// <summary>Whether the field is hidden.</summary>
        public bool Hidden { get; set; }

        /// <summary>Placeholder text displayed when the field is empty.</summary>
        public string? Placeholder { get; set; }

        /// <summary>Default value for the field (configuration, not user-entered).</summary>
        public string? DefaultValue { get; set; }
    }
}
