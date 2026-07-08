namespace ExcelAPI.Models
{
    /// <summary>
    /// Represents a form field extracted from an Excel cell comment/note.
    /// Coordinates are converted from Excel points to PNG pixels using the
    /// actual export scale (PNG dimensions / Print Area dimensions).
    /// Include debug metadata for verifying coordinate alignment.
    /// </summary>
    public class ExcelField
    {
        /// <summary>A unique identifier for this field, derived from the cell address.</summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>The Excel cell address (e.g., "B5").</summary>
        public string Cell { get; set; } = string.Empty;

        /// <summary>The field type extracted from the comment (e.g., "Text", "Date", "Checkbox").</summary>
        public string Type { get; set; } = string.Empty;

        /// <summary>Left edge in PNG pixels (relative to Print Area origin).</summary>
        public double Left { get; set; }

        /// <summary>Top edge in PNG pixels (relative to Print Area origin).</summary>
        public double Top { get; set; }

        /// <summary>Width in PNG pixels.</summary>
        public double Width { get; set; }

        /// <summary>Height in PNG pixels.</summary>
        public double Height { get; set; }

        /// <summary>The full raw comment text from the cell.</summary>
        public string Comment { get; set; } = string.Empty;

        // --- Debug Metadata (temporary — for verifying alignment) ---

        /// <summary>Raw cell Left in Excel points (for debug verification).</summary>
        public double ExcelLeft { get; set; }

        /// <summary>Raw cell Top in Excel points (for debug verification).</summary>
        public double ExcelTop { get; set; }

        /// <summary>Print Area origin Left in Excel points (for debug verification).</summary>
        public double PrintAreaLeft { get; set; }

        /// <summary>Print Area origin Top in Excel points (for debug verification).</summary>
        public double PrintAreaTop { get; set; }
    }
}
