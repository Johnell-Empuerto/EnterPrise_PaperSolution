namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Computes cell geometry from parsed sheet data.
    /// Populates ComputedColWidthsPt, ComputedColCumLeftPt, ComputedRowHeightsPt,
    /// ComputedRowCumTopPt, TotalWidthPt, TotalHeightPt, and merge geometry.
    ///
    /// This is the ONLY place geometry is computed — not in the parser, not in layers.
    /// </summary>
    public class GeometryBuilder
    {
        /// <summary>Convert character width to points using ECMA-376 formula.</summary>
        /// <remarks>Delegates to CellGeometryEngine as the canonical implementation (§4).</remarks>
        public static double CharWidthToPoints(double charWidth, double maxDigitWidth = 7.33)
        {
            return CellGeometryEngine.CharWidthToPoints(charWidth, maxDigitWidth);
        }

        /// <summary>
        /// Compute all geometry for a sheet: column positions, row positions, merge bounds.
        /// Must be called after parser populates columns, rows, cells, and merges.
        /// </summary>
        public void ComputeGeometry(RenderSheet sheet)
        {
            // Column cumulative positions (points)
            double cumX = 0;
            uint maxCol = 0;
            foreach (var c in sheet.Cells)
                if (c.ColumnIndex > maxCol) maxCol = c.ColumnIndex;
            if (sheet.Columns.Count > 0)
            {
                uint lastCol = sheet.Columns.Max(c => c.Max);
                if (lastCol > maxCol) maxCol = lastCol;
            }

            for (uint col = 1; col <= maxCol; col++)
            {
                double w = sheet.DefaultColumnWidth;
                foreach (var c in sheet.Columns)
                {
                    if (c.Min <= col && col <= c.Max && !c.Hidden)
                    {
                        w = c.Width;
                        break;
                    }
                }
                double pt = CharWidthToPoints(w, sheet.MaxDigitWidth);
                sheet.ComputedColWidthsPt[col] = pt;
                sheet.ComputedColCumLeftPt[col] = cumX;
                cumX += pt;
            }
            sheet.TotalWidthPt = cumX;

            // Row cumulative positions (points)
            double cumY = 0;
            uint maxRow = 0;
            foreach (var c in sheet.Cells)
                if (c.RowIndex > maxRow) maxRow = c.RowIndex;
            foreach (var r in sheet.Rows)
                if (r.RowIndex > maxRow) maxRow = r.RowIndex;

            for (uint row = 1; row <= maxRow; row++)
            {
                double h = sheet.DefaultRowHeight;
                foreach (var r in sheet.Rows)
                {
                    if (r.RowIndex == row && r.Height.HasValue)
                    {
                        h = r.Height.Value;
                        break;
                    }
                }
                sheet.ComputedRowHeightsPt[row] = h;
                sheet.ComputedRowCumTopPt[row] = cumY;
                cumY += h;
            }
            sheet.TotalHeightPt = cumY;

            // Compute merge geometry
            foreach (var m in sheet.Merges)
            {
                m.LeftPt = sheet.ComputedColCumLeftPt.GetValueOrDefault(m.FirstCol, 0);
                m.TopPt = sheet.ComputedRowCumTopPt.GetValueOrDefault(m.FirstRow, 0);

                double right = m.LastCol < uint.MaxValue
                    ? sheet.ComputedColCumLeftPt.GetValueOrDefault(m.LastCol + 1, cumX)
                    : cumX;
                double bottom = m.LastRow < uint.MaxValue
                    ? sheet.ComputedRowCumTopPt.GetValueOrDefault(m.LastRow + 1, cumY)
                    : cumY;

                m.WidthPt = right - m.LeftPt;
                m.HeightPt = bottom - m.TopPt;
            }
        }
    }
}
