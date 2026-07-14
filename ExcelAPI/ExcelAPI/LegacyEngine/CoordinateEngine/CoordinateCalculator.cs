using ExcelAPI.LegacyEngine.LayoutEngine;
using ExcelAPI.LegacyEngine.Models;

namespace ExcelAPI.LegacyEngine.CoordinateEngine;

public class CoordinateCalculator : ICoordinateCalculator
{
    private readonly IColumnEngine _columnEngine;
    private readonly IRowEngine _rowEngine;

    public CoordinateCalculator(IColumnEngine columnEngine, IRowEngine rowEngine)
    {
        _columnEngine = columnEngine;
        _rowEngine = rowEngine;
    }

    public CoordinateRect CalculateRaw(ClusterModel cluster, SheetModel sheet)
    {
        double leftPt = _columnEngine.GetLeftPoints(sheet, cluster.StartColumn);
        double topPt = _rowEngine.GetTopPoints(sheet, cluster.StartRow);
        double widthPt = _columnEngine.GetWidthPoints(sheet, cluster.StartColumn, cluster.EndColumn);
        double heightPt = _rowEngine.GetHeightPoints(sheet, cluster.StartRow, cluster.EndRow);

        return new CoordinateRect
        {
            Left = leftPt,
            Top = topPt,
            Right = leftPt + widthPt,
            Bottom = topPt + heightPt
        };
    }

    public CoordinateRect CalculateNormalized(ClusterModel cluster, SheetModel sheet, double originX, double originY)
    {
        double pageW = sheet.PageSetup.PageWidthPoints;
        double pageH = sheet.PageSetup.PageHeightPoints;

        CoordinateRect raw = CalculateRaw(cluster, sheet);

        return new CoordinateRect
        {
            Left = Math.Round((raw.Left + originX) / pageW, 7),
            Top = Math.Round((raw.Top + originY) / pageH, 7),
            Right = Math.Round((raw.Right + originX) / pageW, 7),
            Bottom = Math.Round((raw.Bottom + originY) / pageH, 7)
        };
    }
}
