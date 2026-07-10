using ExcelAPI.Rendering;
using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Renders cell background fills onto a SkiaSharp canvas.
    /// Supports solid fills, pattern fills (gray125, gray0625), and no-fill.
    /// Implements IRenderLayer for pluggable pipeline use.
    /// Rendering order: Layer 1 (after page background).
    /// </summary>
    public class FillEngine : IRenderLayer
    {
        private readonly CellGeometryEngine _geometry;

        public FillEngine(CellGeometryEngine geometry)
        {
            _geometry = geometry;
        }

        public string Name => "FillLayer";

        void IRenderLayer.Render(SKCanvas canvas, RenderingContext context)
        {
            RenderFills(canvas, context.Sheet, context.OriginXPt, context.OriginYPt);
        }

        /// <summary>
        /// Render all cell background fills for the given sheet.
        /// </summary>
        public void RenderFills(SKCanvas canvas, RenderSheet sheet,
            double originXPt, double originYPt)
        {
            foreach (var cell in sheet.Cells)
            {
                RenderCellFill(canvas, sheet, cell, originXPt, originYPt);
            }
        }

        /// <summary>
        /// Render fill for a single cell (respects merged cells — only renders
        /// at the anchor cell position for merged ranges).
        /// </summary>
        public void RenderCellFill(SKCanvas canvas, RenderSheet sheet, RenderCell cell,
            double originXPt, double originYPt)
        {
            // Skip cells without fill
            if (string.IsNullOrEmpty(cell.FillColorArgb)) return;
            if (cell.PatternType == "none") return;

            // For merged cells: only render fill at the top-left anchor cell
            if (cell.MergeIndex.HasValue)
            {
                var merge = sheet.Merges[cell.MergeIndex.Value];
                // Only the first cell in the merge renders the fill
                if (cell.ColumnIndex != merge.FirstCol || cell.RowIndex != merge.FirstRow)
                    return;
            }

            var bounds = _geometry.GetCellOrMergePixelBounds(sheet, cell, originXPt, originYPt);

            using var paint = new SKPaint
            {
                Style = SKPaintStyle.Fill,
                IsAntialias = false
            };

            // Parse color
            var color = ParseColor(cell.FillColorArgb);
            paint.Color = color;

            // Handle pattern fills
            if (cell.PatternType == "gray125")
            {
                // 12.5% gray — draw with low opacity
                paint.Color = color.WithAlpha(32);
                canvas.DrawRect(bounds, paint);
            }
            else if (cell.PatternType == "gray0625")
            {
                // 6.25% gray — even lighter
                paint.Color = color.WithAlpha(16);
                canvas.DrawRect(bounds, paint);
            }
            else if (cell.PatternType == "solid" || cell.PatternType == null)
            {
                // Solid fill (most common)
                canvas.DrawRect(bounds, paint);
            }
            else
            {
                // Unknown pattern type — fall through to solid fill
                canvas.DrawRect(bounds, paint);
            }
        }

        /// <summary>
        /// Parse a #RRGGBB or #AARRGGBB string to SKColor.
        /// </summary>
        public static SKColor ParseColor(string? argb)
        {
            if (string.IsNullOrEmpty(argb)) return SKColors.White;

            string hex = argb.TrimStart('#');
            if (hex.Length == 6)
            {
                byte r = Convert.ToByte(hex[..2], 16);
                byte g = Convert.ToByte(hex[2..4], 16);
                byte b = Convert.ToByte(hex[4..6], 16);
                return new SKColor(r, g, b, 255);
            }
            else if (hex.Length == 8)
            {
                byte a = Convert.ToByte(hex[..2], 16);
                byte r = Convert.ToByte(hex[2..4], 16);
                byte g = Convert.ToByte(hex[4..6], 16);
                byte b = Convert.ToByte(hex[6..8], 16);
                return new SKColor(r, g, b, a);
            }
            return SKColors.White;
        }

        /// <summary>
        /// Check if a color is considered a "yellow separator" (common in Japanese forms).
        /// </summary>
        public static bool IsYellowSeparator(SKColor color)
        {
            return color.Red > 200 && color.Green > 180 && color.Blue < 100;
        }
    }
}
