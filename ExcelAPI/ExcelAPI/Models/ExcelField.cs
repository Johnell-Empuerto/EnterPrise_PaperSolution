namespace ExcelAPI.Models
{
    /// <summary>
    /// Represents a form field extracted from an Excel cell comment/note.
    /// Coordinates are converted from Excel points to PNG pixels at 300 DPI.
    /// </summary>
    public class ExcelField
    {
        /// <summary>A unique identifier for this field, derived from the cell address.</summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>The Excel cell address (e.g., "B5").</summary>
        public string Cell { get; set; } = string.Empty;

        /// <summary>The field type extracted from the comment (e.g., "Text", "Date", "Checkbox").</summary>
        public string Type { get; set; } = string.Empty;

        /// <summary>Left edge in PNG pixels (300 DPI).</summary>
        public double Left { get; set; }

        /// <summary>Top edge in PNG pixels (300 DPI).</summary>
        public double Top { get; set; }

        /// <summary>Width in PNG pixels (300 DPI).</summary>
        public double Width { get; set; }

        /// <summary>Height in PNG pixels (300 DPI).</summary>
        public double Height { get; set; }

        /// <summary>The full raw comment text from the cell.</summary>
        public string Comment { get; set; } = string.Empty;
    }
}
