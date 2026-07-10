using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Renders cell text onto the SkiaSharp canvas.
    /// Implements IRenderLayer — automatically executed by RendererCoordinator
    /// after fills, gridlines, and borders.
    ///
    /// Rendering order within this layer:
    ///   1. Layout: compute TextDrawCommand for each cell
    ///   2. Clip: set clipping to cell/merge bounds
    ///   3. Align: compute text position based on horizontal/vertical alignment
    ///   4. Wrap: break text into lines if WrapText is enabled
    ///   5. Rotate: apply text rotation transform
    ///   6. Draw: render each line with SkiaSharp
    ///   7. Decor: draw underline and strikethrough
    ///
    /// Merged cells: only the anchor cell renders text. Non-anchor cells are skipped.
    /// This matches Excel's behavior where merged content is stored in the top-left cell.
    /// </summary>
    public class TextEngine : IRenderLayer
    {
        private readonly TextLayoutEngine _layoutEngine;

        public TextEngine(TextLayoutEngine layoutEngine)
        {
            _layoutEngine = layoutEngine;
        }

        public string Name => "TextLayer";

        public void Render(SKCanvas canvas, RenderingContext context)
        {
            var sheet = context.Sheet;
            double ptsToPx = context.PointsToPixels;

            // Track which merged cell anchors we've rendered (first cell in merge)
            var renderedMerges = new HashSet<int>();

            foreach (var cell in sheet.Cells)
            {
                // Skip non-anchor cells in merged ranges
                if (cell.MergeIndex.HasValue)
                {
                    var merge = sheet.Merges[cell.MergeIndex.Value];
                    if (cell.ColumnIndex != merge.FirstCol || cell.RowIndex != merge.FirstRow)
                        continue;

                    // Only render each merge once
                    if (!renderedMerges.Add(cell.MergeIndex.Value))
                        continue;
                }

                // Compute layout
                var cmd = _layoutEngine.LayoutCell(sheet, cell,
                    context.OriginXPt, context.OriginYPt, ptsToPx);

                if (cmd == null) continue;

                // Render the text command
                RenderDrawCommand(canvas, cmd);
            }
        }

        /// <summary>
        /// Render a single TextDrawCommand onto the canvas.
        /// </summary>
        private void RenderDrawCommand(SKCanvas canvas, TextDrawCommand cmd)
        {
            if (string.IsNullOrEmpty(cmd.Text)) return;

            // Save state before clip + transform
            canvas.Save();

            try
            {
                // Apply clipping to cell/merge bounds
                canvas.ClipRect(cmd.ClipRect);

                // Apply rotation around cell center
                float centerX = cmd.ClipRect.MidX;
                float centerY = cmd.ClipRect.MidY;

                if (cmd.RotationDegrees != 0)
                {
                    canvas.RotateDegrees(cmd.RotationDegrees, centerX, centerY);
                }

                // Compute font paint
                float fontSizePx = cmd.FontSizePt * (float)(300.0 / 72.0); // Convert pt to px at 300 DPI
                using var font = new SKFont(cmd.Typeface, fontSizePx);
                using var paint = new SKPaint
                {
                    Color = cmd.FontColor,
                    IsAntialias = true
                };

                // Compute text metrics
                float textHeight = font.Spacing; // Line height in pixels
                float capHeight = font.Metrics.CapHeight; // Height of capital letters

                // Calculate content area (with indent)
                float contentLeft = cmd.X + cmd.IndentPixels;
                float contentWidth = cmd.Width - cmd.IndentPixels;
                float contentTop = cmd.Y;
                float contentHeight = cmd.Height;

                // Handle WrapText — break into lines
                string[] lines;
                float[] lineWidths;

                if (cmd.WrapText && contentWidth > 0)
                {
                    lines = WrapTextLines(cmd.Text, cmd.Typeface, fontSizePx, contentWidth);
                }
                else
                {
                    // Single line, possibly truncated
                    lines = new[] { cmd.Text };
                }

                // Measure line widths
                lineWidths = new float[lines.Length];
                for (int i = 0; i < lines.Length; i++)
                {
                    lineWidths[i] = font.MeasureText(lines[i]);
                }

                // Compute vertical position based on alignment
                float totalTextHeight = textHeight * lines.Length;
                float baselineY;

                switch (cmd.VerticalAlignment.ToLowerInvariant())
                {
                    case "top":
                        baselineY = contentTop + textHeight - (textHeight - capHeight) / 2f;
                        break;
                    case "center":
                        baselineY = contentTop + (contentHeight - totalTextHeight) / 2f
                                     + textHeight - (textHeight - capHeight) / 2f;
                        break;
                    case "justify":
                    case "distributed":
                        // Justify/Distributed behaves like top when single line
                        baselineY = contentTop + textHeight - (textHeight - capHeight) / 2f;
                        break;
                    default: // "bottom" (Excel default)
                        baselineY = contentTop + contentHeight - totalTextHeight
                                     + textHeight - (textHeight - capHeight) / 2f;
                        break;
                }

                // Draw each line
                for (int i = 0; i < lines.Length; i++)
                {
                    float lineWidth = lineWidths[i];
                    float lineX;

                    // Compute horizontal position based on alignment
                    switch (cmd.HorizontalAlignment.ToLowerInvariant())
                    {
                        case "center":
                        case "centerContinuous":
                            lineX = contentLeft + (contentWidth - lineWidth) / 2f;
                            break;
                        case "right":
                        case "fill":
                            lineX = contentLeft + contentWidth - lineWidth;
                            break;
                        default: // "left", "general"
                            lineX = contentLeft;
                            break;
                    }

                    // Calculate Y for this line
                    float lineY = baselineY + i * textHeight;

                    // Draw the text
                    canvas.DrawText(lines[i], lineX, lineY, font, paint);

                    // Draw underline
                    if (cmd.Underline)
                    {
                        float underlineY = lineY + 1.5f;
                        using var underlinePaint = new SKPaint
                        {
                            Color = cmd.FontColor,
                            Style = SKPaintStyle.Stroke,
                            StrokeWidth = Math.Max(1f, fontSizePx * 0.04f),
                            IsAntialias = true
                        };
                        canvas.DrawLine(lineX, underlineY, lineX + lineWidth, underlineY, underlinePaint);
                    }

                    // Draw strikethrough
                    if (cmd.Strikeout)
                    {
                        float strikeY = lineY - capHeight * 0.3f;
                        using var strikePaint = new SKPaint
                        {
                            Color = cmd.FontColor,
                            Style = SKPaintStyle.Stroke,
                            StrokeWidth = Math.Max(1f, fontSizePx * 0.04f),
                            IsAntialias = true
                        };
                        canvas.DrawLine(lineX, strikeY, lineX + lineWidth, strikeY, strikePaint);
                    }
                }
            }
            finally
            {
                canvas.Restore();
            }
        }

        /// <summary>
        /// Wrap text into lines that fit within the given width.
        /// Splits on word boundaries (spaces) or at any character if necessary.
        /// </summary>
        private static string[] WrapTextLines(string text, SKTypeface? typeface,
            float fontSizePx, float maxWidthPx)
        {
            if (string.IsNullOrEmpty(text)) return Array.Empty<string>();

            var lines = new List<string>();
            using var font = new SKFont(typeface, fontSizePx);

            string remaining = text;

            while (remaining.Length > 0)
            {
                if (font.MeasureText(remaining) <= maxWidthPx)
                {
                    // Entire remaining text fits
                    // Check for newlines
                    int nlIndex = remaining.IndexOf('\n');
                    if (nlIndex >= 0)
                    {
                        lines.Add(remaining[..nlIndex]);
                        remaining = remaining[(nlIndex + 1)..];
                    }
                    else
                    {
                        lines.Add(remaining);
                        break;
                    }
                }
                else
                {
                    // Find the longest segment that fits
                    int breakIndex = FindBreakIndex(remaining, font, maxWidthPx);

                    // Try to break at a word boundary
                    int wordBreak = remaining.LastIndexOf(' ', breakIndex - 1, breakIndex);
                    if (wordBreak > 0)
                        breakIndex = wordBreak;

                    // Check for newline before the break
                    int nlIndex = remaining.IndexOf('\n');
                    if (nlIndex >= 0 && nlIndex < breakIndex)
                    {
                        lines.Add(remaining[..nlIndex]);
                        remaining = remaining[(nlIndex + 1)..];
                        continue;
                    }

                    lines.Add(remaining[..breakIndex].TrimEnd());
                    remaining = remaining[breakIndex..].TrimStart();
                }
            }

            return lines.ToArray();
        }

        /// <summary>
        /// Binary search to find the character index where text fits within maxWidth.
        /// </summary>
        private static int FindBreakIndex(string text, SKFont font, float maxWidthPx)
        {
            int lo = 1;
            int hi = text.Length;
            int best = 1;

            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                float w = font.MeasureText(text[..mid]);
                if (w <= maxWidthPx)
                {
                    best = mid;
                    lo = mid + 1;
                }
                else
                {
                    hi = mid - 1;
                }
            }

            return Math.Max(1, best);
        }
    }
}
