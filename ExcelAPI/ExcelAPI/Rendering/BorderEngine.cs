using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Renders cell borders onto a SkiaSharp canvas.
    /// Implements Excel's border priority rules (right/bottom cell wins),
    /// border collapse (shared borders), and all 14 border styles.
    /// Implements IRenderLayer for pluggable pipeline use.
    /// Rendering order: Layer 3 (after gridlines, before text).
    /// </summary>
    public class BorderEngine : IRenderLayer
    {
        private readonly CellGeometryEngine _geometry;

        public BorderEngine(CellGeometryEngine geometry)
        {
            _geometry = geometry;
        }

        public string Name => "BorderLayer";

        void IRenderLayer.Render(SKCanvas canvas, RenderingContext context)
        {
            RenderBorders(canvas, context.Sheet, context.OriginXPt, context.OriginYPt);
        }

        /// <summary>
        /// Render all cell borders for the given sheet.
        /// Excel draws borders with right/bottom priority:
        ///   - For each cell, draw RIGHT border
        ///   - For each cell, draw BOTTOM border
        ///   - This ensures adjacent cells share one border line
        ///   - Left/top borders are only drawn for edge cells
        /// </summary>
        public void RenderBorders(SKCanvas canvas, RenderSheet sheet,
            double originXPt, double originYPt)
        {
            // Build a complete border grid
            var borderGrid = BuildBorderGrid(sheet);

            // Render right and bottom borders for all cells
            foreach (var cell in sheet.Cells)
            {
                if (cell.Border == null) continue;

                // Skip non-anchor cells in merged ranges
                if (cell.MergeIndex.HasValue)
                {
                    var merge = sheet.Merges[cell.MergeIndex.Value];
                    if (cell.ColumnIndex != merge.FirstCol || cell.RowIndex != merge.FirstRow)
                        continue;
                }

                var bounds = _geometry.GetCellOrMergePixelBounds(sheet, cell, originXPt, originYPt);

                // Draw Right border
                if (cell.Border.Right != null)
                {
                    DrawBorderLine(canvas, bounds.Right, bounds.Top, bounds.Right, bounds.Bottom,
                        cell.Border.Right, 1.0f);
                }

                // Draw Bottom border
                if (cell.Border.Bottom != null)
                {
                    DrawBorderLine(canvas, bounds.Left, bounds.Bottom, bounds.Right, bounds.Bottom,
                        cell.Border.Bottom, 1.0f);
                }

                // Draw Left border (only if first column or explicitly set)
                if (cell.Border.Left != null)
                {
                    var leftCol = cell.MergeIndex.HasValue
                        ? sheet.Merges[cell.MergeIndex.Value].FirstCol
                        : cell.ColumnIndex;
                    if (leftCol == 1 || !HasLeftBorderFromLeftCell(sheet, cell))
                    {
                        DrawBorderLine(canvas, bounds.Left, bounds.Top, bounds.Left, bounds.Bottom,
                            cell.Border.Left, 1.0f);
                    }
                }

                // Draw Top border (only if first row or explicitly set)
                if (cell.Border.Top != null)
                {
                    var topRow = cell.MergeIndex.HasValue
                        ? sheet.Merges[cell.MergeIndex.Value].FirstRow
                        : cell.RowIndex;
                    if (topRow == 1 || !HasTopBorderFromAboveCell(sheet, cell))
                    {
                        DrawBorderLine(canvas, bounds.Left, bounds.Top, bounds.Right, bounds.Top,
                            cell.Border.Top, 1.0f);
                    }
                }

                // Draw Diagonal borders
                if (cell.Border.DiagonalUp != null)
                {
                    DrawBorderLine(canvas, bounds.Left, bounds.Bottom, bounds.Right, bounds.Top,
                        cell.Border.DiagonalUp, 1.0f);
                }
                if (cell.Border.DiagonalDown != null)
                {
                    DrawBorderLine(canvas, bounds.Left, bounds.Top, bounds.Right, bounds.Bottom,
                        cell.Border.DiagonalDown, 1.0f);
                }
            }
        }

        /// <summary>
        /// Check if the cell to the left has a right border that would overlap.
        /// </summary>
        private bool HasLeftBorderFromLeftCell(RenderSheet sheet, RenderCell cell)
        {
            // In Excel, cell B1's left border is drawn by cell A1's right border
            // If A1 has no right border, B1 draws its own left border
            if (cell.ColumnIndex <= 1) return false;

            // Find cell to the left
            var leftCell = sheet.Cells.FirstOrDefault(c =>
                c.RowIndex == cell.RowIndex && c.ColumnIndex == cell.ColumnIndex - 1);
            if (leftCell?.Border?.Right != null)
                return !IsNoBorder(leftCell.Border.Right);

            return false;
        }

        /// <summary>
        /// Check if the cell above has a bottom border that would overlap.
        /// </summary>
        private bool HasTopBorderFromAboveCell(RenderSheet sheet, RenderCell cell)
        {
            if (cell.RowIndex <= 1) return false;

            var aboveCell = sheet.Cells.FirstOrDefault(c =>
                c.ColumnIndex == cell.ColumnIndex && c.RowIndex == cell.RowIndex - 1);
            if (aboveCell?.Border?.Bottom != null)
                return !IsNoBorder(aboveCell.Border.Bottom);

            return false;
        }

        private bool IsNoBorder(RenderBorderItem item)
        {
            return string.Equals(item.Style, "none", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Draw a single border line segment.
        /// </summary>
        private void DrawBorderLine(SKCanvas canvas, float x1, float y1, float x2, float y2,
            RenderBorderItem border, float pixelOffset)
        {
            using var paint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                IsAntialias = true,
                StrokeWidth = (float)(border.WeightPt * (300.0 / 72.0))
            };

            // Color
            if (border.ColorArgb != null)
            {
                paint.Color = FillEngine.ParseColor(border.ColorArgb);
            }
            else
            {
                paint.Color = SKColors.Black;
            }

            // Line style
            // Line style (string-based to avoid OpenXml type conflicts)
            string style = border.Style.ToLowerInvariant();
            switch (style)
            {
                case "thin":
                case "hair":
                case "medium":
                case "thick":
                    paint.PathEffect = null;
                    break;

                case "dotted":
                    paint.PathEffect = SKPathEffect.CreateDash(new[] { 1f, 3f }, 0);
                    break;

                case "dashed":
                    paint.PathEffect = SKPathEffect.CreateDash(new[] { 4f, 3f }, 0);
                    break;

                case "mediumdashed":
                    paint.PathEffect = SKPathEffect.CreateDash(new[] { 6f, 4f }, 0);
                    paint.StrokeWidth = (float)(1.0 * 300.0 / 72.0);
                    break;

                case "dashdot":
                    paint.PathEffect = SKPathEffect.CreateDash(new[] { 4f, 3f, 1f, 3f }, 0);
                    break;

                case "mediumdashdot":
                    paint.PathEffect = SKPathEffect.CreateDash(new[] { 6f, 4f, 2f, 4f }, 0);
                    paint.StrokeWidth = (float)(1.0 * 300.0 / 72.0);
                    break;

                case "dashdotdot":
                    paint.PathEffect = SKPathEffect.CreateDash(new[] { 4f, 3f, 1f, 3f, 1f, 3f }, 0);
                    break;

                case "mediumdashdotdot":
                    paint.PathEffect = SKPathEffect.CreateDash(new[] { 6f, 4f, 2f, 4f, 2f, 4f }, 0);
                    paint.StrokeWidth = (float)(1.0 * 300.0 / 72.0);
                    break;

                case "double":
                    // Double border: draw two parallel lines
                    DrawDoubleBorderLine(canvas, x1, y1, x2, y2, border, pixelOffset);
                    return; // Already drawn

                default:
                    paint.PathEffect = null;
                    break;
            }

            // For horizontal lines, slightly offset to prevent anti-aliasing blur
            if (y1 == y2)  // Horizontal
            {
                canvas.DrawLine(x1, y1 + pixelOffset, x2, y2 + pixelOffset, paint);
            }
            else
            {
                canvas.DrawLine(x1 + pixelOffset, y1, x2 + pixelOffset, y2, paint);
            }
        }

        /// <summary>
        /// Draw double border (two parallel thin lines with a gap).
        /// </summary>
        private void DrawDoubleBorderLine(SKCanvas canvas, float x1, float y1, float x2, float y2,
            RenderBorderItem border, float pixelOffset)
        {
            using var paint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                IsAntialias = true,
                Color = border.ColorArgb != null ? FillEngine.ParseColor(border.ColorArgb) : SKColors.Black,
                StrokeWidth = (float)(0.5 * 300.0 / 72.0)  // ~2px at 300 DPI
            };

            float gap = (float)(0.5 * 300.0 / 72.0);  // ~2px gap

            if (y1 == y2)  // Horizontal
            {
                // Top line
                canvas.DrawLine(x1, y1 + pixelOffset - gap, x2, y2 + pixelOffset - gap, paint);
                // Bottom line
                canvas.DrawLine(x1, y1 + pixelOffset + gap, x2, y2 + pixelOffset + gap, paint);
            }
            else if (x1 == x2)  // Vertical
            {
                // Left line
                canvas.DrawLine(x1 - gap + pixelOffset, y1, x2 - gap + pixelOffset, y2, paint);
                // Right line
                canvas.DrawLine(x1 + gap + pixelOffset, y1, x2 + gap + pixelOffset, y2, paint);
            }
        }

        /// <summary>
        /// Build a grid of border presence for collapse detection.
        /// </summary>
        private Dictionary<(uint row, uint col), bool> BuildBorderGrid(RenderSheet sheet)
        {
            var grid = new Dictionary<(uint, uint), bool>();

            foreach (var cell in sheet.Cells)
            {
                if (cell.Border != null)
                {
                    grid[(cell.RowIndex, cell.ColumnIndex)] = true;
                }
            }

            return grid;
        }
    }
}
