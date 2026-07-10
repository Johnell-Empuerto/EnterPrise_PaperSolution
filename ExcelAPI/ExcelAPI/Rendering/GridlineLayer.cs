using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Renders default light-gray gridlines (Excel default color).
    /// Layer 2 (between fills and borders).
    /// </summary>
    public class GridlineLayer : IRenderLayer
    {
        public string Name => "GridlineLayer";

        public void Render(SKCanvas canvas, RenderingContext context)
        {
            var sheet = context.Sheet;
            double originXPt = context.OriginXPt;
            double originYPt = context.OriginYPt;
            double ptsToPx = context.PointsToPixels;

            using var gridPaint = new SKPaint
            {
                Color = new SKColor(208, 215, 222, 100),  // Light gray (Excel default)
                Style = SKPaintStyle.Stroke,
                StrokeWidth = 0.5f,
                IsAntialias = true
            };

            float leftPx = (float)(originXPt * ptsToPx);
            float topPx = (float)(originYPt * ptsToPx);
            float totalWPx = (float)(sheet.TotalWidthPt * ptsToPx);
            float totalHPx = (float)(sheet.TotalHeightPt * ptsToPx);

            // Vertical gridlines (right edge of each column)
            foreach (var kv in sheet.ComputedColCumLeftPt)
            {
                float x = leftPx + (float)(kv.Value * ptsToPx);
                canvas.DrawLine(x, topPx, x, topPx + totalHPx, gridPaint);
            }

            // Horizontal gridlines (bottom edge of each row)
            foreach (var kv in sheet.ComputedRowCumTopPt)
            {
                float y = topPx + (float)(kv.Value * ptsToPx);
                canvas.DrawLine(leftPx, y, leftPx + totalWPx, y, gridPaint);
            }
        }
    }
}
