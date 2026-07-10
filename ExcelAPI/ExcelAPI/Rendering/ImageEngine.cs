using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Renders worksheet images (PNG, JPEG, GIF, BMP, TIFF) onto the canvas.
    /// Implements IRenderLayer — automatically executed by RendererCoordinator
    /// after text rendering.
    ///
    /// Rendering order: Layer 5 (after text, before shapes).
    ///
    /// Image positions use CoordinateEngine for cell-relative anchors.
    /// Images are pre-resolved during parsing by ImageResolver.
    /// </summary>
    public class ImageEngine : IRenderLayer
    {
        public ImageEngine()
        {
        }

        public string Name => "ImageLayer";

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
                if (!drawObj.IsImage) continue;
                if (drawObj.ImageData == null) continue;

                var bounds = ComputeImagePixelBounds(
                    sheet, drawObj, originXPt, originYPt, ptsToPx);

                if (bounds.Width <= 0 || bounds.Height <= 0) continue;

                DrawImage(canvas, drawObj.ImageData.Bitmap, bounds);
            }
        }

        /// <summary>
        /// Compute pixel bounds using CoordinateEngine point-to-pixel conversion.
        /// </summary>
        private static SKRect ComputeImagePixelBounds(
            RenderSheet sheet,
            DrawingParser.DrawingObject drawObj,
            double originXPt,
            double originYPt,
            double ptsToPx)
        {
            double leftPt, topPt, widthPt, heightPt;

            switch (drawObj.AnchorType)
            {
                case "oneCellAnchor":
                {
                    double colLeftPt = sheet.ComputedColCumLeftPt.GetValueOrDefault(drawObj.AnchorCol, 0);
                    double colWidthPt = sheet.ComputedColWidthsPt.GetValueOrDefault(drawObj.AnchorCol, 50);

                    // Offset is in EMU — convert to points
                    double colOffPt = DrawingParser.EmuToPt(drawObj.ColOffsetEmu);
                    double rowOffPt = DrawingParser.EmuToPt(drawObj.RowOffsetEmu);

                    leftPt = originXPt + colLeftPt + colOffPt;
                    topPt = originYPt +
                            sheet.ComputedRowCumTopPt.GetValueOrDefault(drawObj.AnchorRow, 0) + rowOffPt;
                    widthPt = DrawingParser.EmuToPt(drawObj.RightEmu);
                    heightPt = DrawingParser.EmuToPt(drawObj.BottomEmu);
                    break;
                }
                case "twoCellAnchor":
                {
                    double colLeftPt = sheet.ComputedColCumLeftPt.GetValueOrDefault(drawObj.AnchorCol, 0);
                    double colOffPt = DrawingParser.EmuToPt(drawObj.ColOffsetEmu);
                    double rowOffPt = DrawingParser.EmuToPt(drawObj.RowOffsetEmu);

                    leftPt = originXPt + colLeftPt + colOffPt;
                    topPt = originYPt +
                            sheet.ComputedRowCumTopPt.GetValueOrDefault(drawObj.AnchorRow, 0) + rowOffPt;

                    double rightPt = originXPt +
                                     sheet.ComputedColCumLeftPt.GetValueOrDefault(drawObj.EndCol + 1, 0);
                    double bottomPt = originYPt +
                                      sheet.ComputedRowCumTopPt.GetValueOrDefault(drawObj.EndRow + 1, 0);
                    widthPt = rightPt - leftPt;
                    heightPt = bottomPt - topPt;
                    break;
                }
                case "absoluteAnchor":
                {
                    leftPt = DrawingParser.EmuToPt(drawObj.LeftEmu) + originXPt;
                    topPt = DrawingParser.EmuToPt(drawObj.TopEmu) + originYPt;
                    widthPt = DrawingParser.EmuToPt(drawObj.RightEmu - drawObj.LeftEmu);
                    heightPt = DrawingParser.EmuToPt(drawObj.BottomEmu - drawObj.TopEmu);
                    break;
                }
                default:
                    return SKRect.Empty;
            }

            return new SKRect(
                (float)(leftPt * ptsToPx),
                (float)(topPt * ptsToPx),
                (float)((leftPt + widthPt) * ptsToPx),
                (float)((topPt + heightPt) * ptsToPx));
        }

        /// <summary>
        /// Draw an SKBitmap with aspect-ratio preservation and transparency.
        /// </summary>
        private static void DrawImage(SKCanvas canvas, SKBitmap bitmap, SKRect bounds)
        {
            if (bitmap == null) return;

            using var paint = new SKPaint
            {
                IsAntialias = true,
                FilterQuality = SKFilterQuality.High
            };

            float sourceAspect = (float)bitmap.Width / bitmap.Height;
            float destAspect = bounds.Width / bounds.Height;

            float drawW, drawH, drawX, drawY;

            if (sourceAspect > destAspect)
            {
                drawW = bounds.Width;
                drawH = bounds.Width / sourceAspect;
                drawX = bounds.Left;
                drawY = bounds.Top + (bounds.Height - drawH) / 2f;
            }
            else
            {
                drawH = bounds.Height;
                drawW = bounds.Height * sourceAspect;
                drawX = bounds.Left + (bounds.Width - drawW) / 2f;
                drawY = bounds.Top;
            }

            canvas.DrawBitmap(bitmap, new SKRect(drawX, drawY, drawX + drawW, drawY + drawH), paint);
        }
    }
}
