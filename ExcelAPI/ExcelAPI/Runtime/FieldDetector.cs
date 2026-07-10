using ExcelAPI.Rendering;

namespace ExcelAPI.Runtime
{
    /// <summary>
    /// Detects editable fields in a parsed worksheet.
    /// Rules:
    ///   - Empty cells with borders → editable field
    ///   - Empty merged cells → editable field
    ///   - Cells with placeholder text → editable field
    ///   - Named ranges → editable field
    ///   - Cells with specific styles → editable field
    ///
    /// Uses CoordinateEngine for pixel coordinates — never recomputes geometry.
    /// </summary>
    public class FieldDetector
    {
        private readonly CoordinateEngine _coords;
        private readonly FieldTypeResolver _typeResolver;

        public FieldDetector(CoordinateEngine coords, FieldTypeResolver typeResolver)
        {
            _coords = coords;
            _typeResolver = typeResolver;
        }

        /// <summary>
        /// Detect editable fields in a sheet.
        /// </summary>
        /// <param name="sheet">The parsed render sheet with computed geometry.</param>
        /// <param name="originXPt">Page origin X in points (margin + centering).</param>
        /// <param name="originYPt">Page origin Y in points (margin + centering).</param>
        /// <param name="dpi">Rendering DPI for pixel conversion.</param>
        /// <returns>List of runtime fields.</returns>
        public List<RuntimeField> DetectFields(
            RenderSheet sheet, double originXPt, double originYPt, int dpi)
        {
            var fields = new List<RuntimeField>();
            var renderedMerges = new HashSet<int>();
            int tabIndex = 0;

            foreach (var cell in sheet.Cells)
            {
                // Skip non-anchor cells in merged ranges
                if (cell.MergeIndex.HasValue)
                {
                    var merge = sheet.Merges[cell.MergeIndex.Value];
                    if (cell.ColumnIndex != merge.FirstCol || cell.RowIndex != merge.FirstRow)
                        continue;

                    if (!renderedMerges.Add(cell.MergeIndex.Value))
                        continue;
                }

                // Determine if this cell should be an editable field
                if (!IsEditable(cell))
                    continue;

                // Compute pixel bounds using CoordinateEngine
                var bounds = _coords.GetCellOrMergePixelBounds(sheet, cell, originXPt, originYPt);

                // Build the runtime field
                var field = BuildField(cell, sheet, bounds, tabIndex, dpi);

                fields.Add(field);
                tabIndex++;
            }

            return fields;
        }

        /// <summary>
        /// Determine whether a cell should be treated as an editable field.
        /// </summary>
        private bool IsEditable(RenderCell cell)
        {
            // Cells without borders are typically not editable (labels)
            bool hasBorder = cell.Border != null &&
                (cell.Border.Left != null || cell.Border.Right != null ||
                 cell.Border.Top != null || cell.Border.Bottom != null);

            // Empty cells with borders are always editable
            if (hasBorder && string.IsNullOrEmpty(cell.Value))
                return true;

            // Cells with values AND borders are editable input fields
            if (hasBorder && !string.IsNullOrEmpty(cell.Value))
                return true;

            // Merged cells with borders are editable
            if (hasBorder && cell.MergeIndex.HasValue)
                return true;

            // Cells with certain alignment patterns suggest editable fields
            if (!string.IsNullOrEmpty(cell.Value) && cell.HorizontalAlignment != null)
            {
                string align = cell.HorizontalAlignment.ToLowerInvariant();
                if (align is "center" or "right")
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Build a RuntimeField from a RenderCell.
        /// </summary>
        private RuntimeField BuildField(
            RenderCell cell, RenderSheet sheet,
            SkiaSharp.SKRect bounds, int tabIndex, int dpi)
        {
            double ptsToPx = dpi / 72.0;

            string? mergeRange = null;
            if (cell.MergeIndex.HasValue && cell.MergeIndex.Value >= 0
                && cell.MergeIndex.Value < sheet.Merges.Count)
            {
                mergeRange = sheet.Merges[cell.MergeIndex.Value].Reference;
            }

            string? placeholder = cell.Value;
            string? defaultValue = cell.Value;

            // For non-empty cells with values, the value is the default (pre-filled)
            if (!string.IsNullOrEmpty(cell.Value))
            {
                placeholder = null;
                defaultValue = cell.Value;
            }
            else
            {
                // Empty cell — no default, placeholder from cell reference
                placeholder = cell.Reference;
            }

            string dataType = _typeResolver.ResolveType(cell);
            bool readOnly = _typeResolver.IsReadOnly(cell);
            bool required = _typeResolver.IsRequired(cell, placeholder);
            string border = FieldTypeResolver.ResolveBorder(cell);

            return new RuntimeField
            {
                Id = $"field_{cell.Reference ?? $"R{cell.RowIndex}C{cell.ColumnIndex}"}_{tabIndex}",
                CellReference = cell.Reference ?? $"R{cell.RowIndex}C{cell.ColumnIndex}",
                Row = cell.RowIndex,
                Column = cell.ColumnIndex,

                LeftPx = bounds.Left,
                TopPx = bounds.Top,
                WidthPx = bounds.Width,
                HeightPx = bounds.Height,

                MergeRange = mergeRange,
                IsMerged = cell.MergeIndex.HasValue,

                DataType = dataType,
                ReadOnly = readOnly,
                Required = required,

                Alignment = cell.HorizontalAlignment,
                Font = cell.FontName,
                FontSize = cell.FontSize,
                Bold = cell.Bold,
                FontColor = cell.FontColorArgb,
                BackgroundColor = cell.FillColorArgb,
                Border = border,

                Placeholder = placeholder,
                DefaultValue = defaultValue,
                MaxLength = 0,
                TabIndex = tabIndex
            };
        }
    }
}
