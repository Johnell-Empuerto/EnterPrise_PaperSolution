using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// A single text draw command produced by TextLayoutEngine.
    /// Contains the pixel-level position, bounds, clipping, font, and text content
    /// ready for SkiaSharp rendering. No data here is approximate — all positions
    /// come from GeometryBuilder's computed cell rectangles.
    /// </summary>
    public class TextDrawCommand
    {
        /// <summary>Text content to render.</summary>
        public string Text { get; set; } = "";

        /// <summary>Pixel X of the text layout area (not yet accounting for alignment).</summary>
        public float X { get; set; }

        /// <summary>Pixel Y baseline of the first line of text.</summary>
        public float Y { get; set; }

        /// <summary>Width of the text layout area in pixels.</summary>
        public float Width { get; set; }

        /// <summary>Height of the text layout area in pixels.</summary>
        public float Height { get; set; }

        /// <summary>Horizontal alignment within the cell rect.</summary>
        public string HorizontalAlignment { get; set; } = "left";

        /// <summary>Vertical alignment within the cell rect.</summary>
        public string VerticalAlignment { get; set; } = "bottom";

        /// <summary>Whether text should wrap to fit Width.</summary>
        public bool WrapText { get; set; }

        /// <summary>Text rotation in degrees (0, 90, -90, 180, etc.).</summary>
        public float RotationDegrees { get; set; }

        /// <summary>Font typeface to use for rendering.</summary>
        public SKTypeface? Typeface { get; set; }

        /// <summary>Font size in points.</summary>
        public float FontSizePt { get; set; } = 11;

        /// <summary>Font color as SKColor.</summary>
        public SKColor FontColor { get; set; } = SKColors.Black;

        /// <summary>Whether to render underline.</summary>
        public bool Underline { get; set; }

        /// <summary>Whether to render strikethrough.</summary>
        public bool Strikeout { get; set; }

        /// <summary>Clipping rectangle in pixels (usually the cell/merge bounds).</summary>
        public SKRect ClipRect { get; set; }

        /// <summary>Indent level (number of characters).</summary>
        public int Indent { get; set; }

        /// <summary>Indent width in pixels.</summary>
        public float IndentPixels => Indent * FontSizePt * 0.5f;
    }

    /// <summary>
    /// Computes text layout from cell data.
    /// Produces TextDrawCommand objects with pixel-accurate positioning.
    /// No rendering here — only layout calculation.
    /// </summary>
    public class TextLayoutEngine
    {
        private readonly CellGeometryEngine _geometry;
        private readonly FontResolver _fontResolver;

        public TextLayoutEngine(CellGeometryEngine geometry, FontResolver fontResolver)
        {
            _geometry = geometry;
            _fontResolver = fontResolver;
        }

        /// <summary>
        /// Compute the text layout for a given cell.
        /// Returns null if the cell has no value to render.
        /// </summary>
        public TextDrawCommand? LayoutCell(RenderSheet sheet, RenderCell cell,
            double originXPt, double originYPt, double ptsToPx)
        {
            // Skip cells without values
            string? text = cell.Value;
            if (string.IsNullOrEmpty(text)) return null;

            // Get pixel bounds for the cell or merge
            var bounds = _geometry.GetCellOrMergePixelBounds(sheet, cell, originXPt, originYPt);

            // Resolve typeface
            var typeface = _fontResolver.ResolveTypeface(cell.FontName, cell.Bold, cell.Italic);

            // Parse font color
            var fontColor = ParseColor(cell.FontColorArgb);

            // Determine rotation
            float rotationDeg = (float)cell.TextRotation;

            // Create the draw command
            var cmd = new TextDrawCommand
            {
                Text = text,
                X = bounds.Left,
                Y = bounds.Top,
                Width = bounds.Width,
                Height = bounds.Height,
                HorizontalAlignment = ResolveHorizontalAlignment(cell.HorizontalAlignment, cell.DataType),
                VerticalAlignment = cell.VerticalAlignment ?? "bottom",
                WrapText = cell.WrapText,
                RotationDegrees = rotationDeg,
                Typeface = typeface,
                FontSizePt = (float)cell.FontSize,
                FontColor = fontColor,
                Underline = cell.Underline,
                Strikeout = cell.Strikeout,
                ClipRect = bounds,
                Indent = (int)cell.Indent
            };

            return cmd;
        }

        /// <summary>
        /// Resolve the effective horizontal alignment from cell settings and data type.
        /// Excel's "general" alignment = text left, numbers right, booleans center.
        /// </summary>
        private static string ResolveHorizontalAlignment(string? alignment, string? dataType)
        {
            if (!string.IsNullOrEmpty(alignment) && !alignment.Equals("general", StringComparison.OrdinalIgnoreCase))
                return alignment;

            // General alignment depends on data type
            if (dataType == "number" || dataType == "Number")
                return "right";

            return "left"; // text, sharedString, booleans, dates
        }

        /// <summary>
        /// Parse a #RRGGBB or #AARRGGBB string to SKColor. Returns black if null.
        /// Delegates to FillEngine.ParseColor to avoid duplication.
        /// </summary>
        private static SKColor ParseColor(string? argb)
        {
            if (string.IsNullOrEmpty(argb)) return SKColors.Black;
            return FillEngine.ParseColor(argb);
        }
    }
}
