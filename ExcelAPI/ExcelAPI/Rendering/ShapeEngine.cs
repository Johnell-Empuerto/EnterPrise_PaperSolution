using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Renders DrawingML shapes (rectangles, ellipses, lines, arrows, text boxes, polygons)
    /// onto the SkiaSharp canvas.
    ///
    /// Implements IRenderLayer — automatically executed by RendererCoordinator
    /// after ImageEngine.
    ///
    /// Rendering order within this layer:
    ///   1. For each drawing object that IsShape, resolve to RenderShape via ShapeResolver
    ///   2. Compute pixel bounds via CoordinateEngine
    ///   3. Draw fill (solid color)
    ///   4. Draw border (with dash style and width)
    ///   5. Draw text (for text boxes and shapes with text)
    ///   6. Apply rotation transform
    ///
    /// Uses CoordinateEngine for position — never computes coordinates independently.
    /// </summary>
    public class ShapeEngine : IRenderLayer
    {
        private readonly ShapeResolver _shapeResolver;

        public ShapeEngine(ShapeResolver shapeResolver)
        {
            _shapeResolver = shapeResolver;
        }

        public string Name => "ShapeLayer";

        public void Render(SKCanvas canvas, RenderingContext context)
        {
            var sheet = context.Sheet;
            if (sheet.DrawingObjects == null || sheet.DrawingObjects.Count == 0)
                return;

            double ptsToPx = context.PointsToPixels;
            double originXPt = context.OriginXPt;
            double originYPt = context.OriginYPt;

            foreach (var drawObj in sheet.DrawingObjects)
            {
                if (!drawObj.IsShape) continue;

                var shape = _shapeResolver.Resolve(drawObj);
                if (shape == null) continue;

                // Compute pixel bounds — use CoordinateEngine conversions
                var bounds = ComputeShapePixelBounds(shape, originXPt, originYPt, ptsToPx);

                if (bounds.Width <= 0 || bounds.Height <= 0) continue;

                RenderShape(canvas, shape, bounds);
            }
        }

        /// <summary>
        /// Compute pixel bounds using context's point-to-pixel conversion.
        /// </summary>
        private static SKRect ComputeShapePixelBounds(
            RenderShape shape,
            double originXPt,
            double originYPt,
            double ptsToPx)
        {
            double leftPt = shape.LeftPt + originXPt;
            double topPt = shape.TopPt + originYPt;

            return new SKRect(
                (float)(leftPt * ptsToPx),
                (float)(topPt * ptsToPx),
                (float)((leftPt + shape.WidthPt) * ptsToPx),
                (float)((topPt + shape.HeightPt) * ptsToPx));
        }

        private static void RenderShape(SKCanvas canvas, RenderShape shape, SKRect bounds)
        {
            canvas.Save();

            try
            {
                if (shape.RotationDegrees != 0)
                {
                    float cx = bounds.MidX;
                    float cy = bounds.MidY;
                    canvas.RotateDegrees((float)shape.RotationDegrees, cx, cy);
                }

                // ── Fill ──
                if (!string.IsNullOrEmpty(shape.FillColorArgb))
                {
                    using var fillPaint = new SKPaint
                    {
                        Style = SKPaintStyle.Fill,
                        IsAntialias = true
                    };

                    fillPaint.Color = FillEngine.ParseColor(shape.FillColorArgb);
                    if (shape.FillOpacity < 1.0)
                    {
                        fillPaint.Color = fillPaint.Color.WithAlpha(
                            (byte)(fillPaint.Color.Alpha * shape.FillOpacity));
                    }

                    DrawShapeFill(canvas, shape, bounds, fillPaint);
                }

                // ── Border / Outline ──
                if (!string.IsNullOrEmpty(shape.BorderColorArgb) || shape.BorderWidthPt > 0)
                {
                    using var borderPaint = new SKPaint
                    {
                        Style = SKPaintStyle.Stroke,
                        IsAntialias = true,
                        StrokeWidth = (float)(shape.BorderWidthPt * (300.0 / 72.0)),
                        Color = FillEngine.ParseColor(shape.BorderColorArgb ?? "#FF000000")
                    };

                    shape.BorderDashStyle = shape.BorderDashStyle.ToLowerInvariant();
                    switch (shape.BorderDashStyle)
                    {
                        case "dash":
                            borderPaint.PathEffect = SKPathEffect.CreateDash(new[] { 4f, 3f }, 0);
                            break;
                        case "dot":
                            borderPaint.PathEffect = SKPathEffect.CreateDash(new[] { 1f, 3f }, 0);
                            break;
                        case "dashdot":
                            borderPaint.PathEffect = SKPathEffect.CreateDash(new[] { 4f, 3f, 1f, 3f }, 0);
                            break;
                        case "lgndash":
                            borderPaint.PathEffect = SKPathEffect.CreateDash(new[] { 8f, 3f }, 0);
                            break;
                        case "lgndashdot":
                            borderPaint.PathEffect = SKPathEffect.CreateDash(new[] { 8f, 3f, 1f, 3f }, 0);
                            break;
                        case "lgndashdotdot":
                            borderPaint.PathEffect = SKPathEffect.CreateDash(new[] { 8f, 3f, 1f, 3f, 1f, 3f }, 0);
                            break;
                    }

                    DrawShapeBorder(canvas, shape, bounds, borderPaint);
                }

                // ── Text ──
                if (!string.IsNullOrEmpty(shape.Text))
                {
                    DrawShapeText(canvas, shape, bounds);
                }
            }
            finally
            {
                canvas.Restore();
            }
        }

        private static void DrawShapeFill(SKCanvas canvas, RenderShape shape, SKRect bounds, SKPaint paint)
        {
            switch (shape.ShapeType)
            {
                case "ellipse":
                    canvas.DrawOval(bounds.MidX, bounds.MidY,
                        bounds.Width / 2f, bounds.Height / 2f, paint);
                    break;
                case "roundedRect":
                    float radius = Math.Min(bounds.Width, bounds.Height) * 0.15f;
                    canvas.DrawRoundRect(bounds, radius, radius, paint);
                    break;
                case "line":
                    break;
                case "polygon":
                    canvas.DrawRect(bounds, paint);
                    break;
                default:
                    canvas.DrawRect(bounds, paint);
                    break;
            }
        }

        private static void DrawShapeBorder(SKCanvas canvas, RenderShape shape, SKRect bounds, SKPaint paint)
        {
            switch (shape.ShapeType)
            {
                case "ellipse":
                    canvas.DrawOval(bounds.MidX, bounds.MidY,
                        bounds.Width / 2f, bounds.Height / 2f, paint);
                    break;
                case "roundedRect":
                    float radius = Math.Min(bounds.Width, bounds.Height) * 0.15f;
                    canvas.DrawRoundRect(bounds, radius, radius, paint);
                    break;
                case "line":
                    canvas.DrawLine(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom, paint);
                    break;
                default:
                    canvas.DrawRect(bounds, paint);
                    break;
            }
        }

        private static void DrawShapeText(SKCanvas canvas, RenderShape shape, SKRect bounds)
        {
            float fontSizePx = (float)(shape.FontSizePt * (300.0 / 72.0));

            using var typeface = SKTypeface.FromFamilyName(
                shape.FontName,
                shape.Bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal,
                SKFontStyleWidth.Normal,
                shape.Italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright);

            using var font = new SKFont(typeface, fontSizePx);
            using var paint = new SKPaint
            {
                Color = FillEngine.ParseColor(shape.FontColorArgb ?? "#FF000000"),
                IsAntialias = true
            };

            float textWidth = font.MeasureText(shape.Text);
            float textHeight = font.Spacing;

            float textX = shape.HorizontalAlignment.ToLowerInvariant() switch
            {
                "center" => bounds.Left + (bounds.Width - textWidth) / 2f,
                "right" => bounds.Right - textWidth,
                _ => bounds.Left + 4f
            };

            float textY = shape.VerticalAlignment.ToLowerInvariant() switch
            {
                "center" => bounds.Top + (bounds.Height - textHeight) / 2f + textHeight - font.Metrics.Descent,
                "bottom" => bounds.Bottom - 4f,
                _ => bounds.Top + textHeight - font.Metrics.Descent
            };

            canvas.Save();
            canvas.ClipRect(bounds);
            canvas.DrawText(shape.Text, textX, textY, font, paint);
            canvas.Restore();
        }
    }
}
