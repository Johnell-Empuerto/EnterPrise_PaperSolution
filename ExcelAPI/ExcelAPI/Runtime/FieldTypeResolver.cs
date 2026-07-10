using ExcelAPI.Rendering;

namespace ExcelAPI.Runtime
{
    /// <summary>
    /// Resolves the runtime data type of a cell based on its value, style, and format.
    /// Uses the parsed RenderCell properties and ResolvedCellStyle to determine
    /// the appropriate field type for the frontend overlay.
    /// </summary>
    public class FieldTypeResolver
    {
        /// <summary>
        /// Resolve the data type for a cell.
        /// </summary>
        /// <param name="cell">The parsed render cell.</param>
        /// <returns>One of: "text", "number", "date", "checkbox", "dropdown", "calculated", "signature".</returns>
        public string ResolveType(RenderCell cell)
        {
            // Check for checkbox patterns:
            // Cells with ☐, ☑, ✓, ✗ characters or specific checkbox-like values
            if (cell.Value != null)
            {
                string val = cell.Value.Trim();
                if (val is "☐" or "☑" or "✓" or "✗" or "□" or "■")
                    return "checkbox";

                if (string.Equals(val, "true", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(val, "false", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(val, "yes", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(val, "no", StringComparison.OrdinalIgnoreCase))
                    return "checkbox";
            }

            // Check for date patterns in the value
            if (cell.Value != null && IsDateValue(cell.Value))
                return "date";

            // Check for number patterns
            if (cell.Value != null && IsNumericValue(cell.Value))
                return "number";

            // Check if the cell has a formula (dataType = "str" or cached value with formula)
            if (cell.DataType == "str" && cell.Value != null)
                return "calculated";

            // Check style for horizontal alignment
            if (cell.HorizontalAlignment != null)
            {
                string align = cell.HorizontalAlignment.ToLowerInvariant();
                if (align is "right" or "fill")
                    return "number";
            }

            // Default to text
            return "text";
        }

        /// <summary>
        /// Determine whether a cell should be treated as a date field.
        /// Checks for date-like patterns and numeric date serial numbers.
        /// </summary>
        private static bool IsDateValue(string value)
        {
            // Check for common date formats
            if (DateTime.TryParse(value, out _))
                return true;

            // Check for date serial numbers (Excel dates are integers 1-2958465)
            if (int.TryParse(value, out int num) && num >= 1 && num <= 2958465)
                return true;

            return false;
        }

        /// <summary>
        /// Determine whether a value represents a number.
        /// </summary>
        private static bool IsNumericValue(string value)
        {
            return double.TryParse(value,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out _);
        }

        /// <summary>
        /// Determine whether a cell is read-only based on its style and content.
        /// Cells with values but no border are typically labels, not editable fields.
        /// </summary>
        public bool IsReadOnly(RenderCell cell)
        {
            // Cells with no border are typically read-only labels
            if (cell.Border == null)
                return true;

            // Cells with text but no border decoration are labels
            if (!string.IsNullOrEmpty(cell.Value) && cell.Border != null)
            {
                bool hasBorder = cell.Border.Left != null || cell.Border.Right != null
                    || cell.Border.Top != null || cell.Border.Bottom != null;
                if (!hasBorder)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Determine whether a cell is required (marked with specific indicators).
        /// </summary>
        public bool IsRequired(RenderCell cell, string? placeholder)
        {
            // Cells with placeholder text containing "required" markers
            if (placeholder != null)
            {
                string lower = placeholder.ToLowerInvariant();
                if (lower.Contains("required") || lower.Contains("*") || lower.Contains("mandatory"))
                    return true;
            }

            // Empty bordered cells are typically required
            if (string.IsNullOrEmpty(cell.Value) && cell.Border != null)
                return true;

            return false;
        }

        /// <summary>
        /// Resolve the border style for a cell: "none", "thin", "medium", "thick", "double".
        /// Returns the strongest border from any side.
        /// </summary>
        public static string ResolveBorder(RenderCell cell)
        {
            if (cell.Border == null) return "none";

            var items = new[] { cell.Border.Left, cell.Border.Right, cell.Border.Top, cell.Border.Bottom };
            string strongest = "none";

            foreach (var item in items)
            {
                if (item == null) continue;
                string style = item.Style.ToLowerInvariant();

                // Prefer the strongest border style
                if (style == "double") return "double";
                if (style == "thick" && strongest != "double") strongest = "thick";
                if (style == "medium" && strongest is not ("double" or "thick")) strongest = "medium";
                if (style == "thin" && strongest is not ("double" or "thick" or "medium")) strongest = "thin";
            }

            return strongest;
        }
    }
}
