using ExcelAPI.LegacyEngine.Models;

namespace ExcelAPI.LegacyEngine.LayoutEngine;

public class OriginEngine : IOriginEngine
{
    private readonly IColumnEngine _columnEngine;
    private readonly IRowEngine _rowEngine;

    public OriginEngine(IColumnEngine columnEngine, IRowEngine rowEngine)
    {
        _columnEngine = columnEngine;
        _rowEngine = rowEngine;
    }

    public (double originX, double originY) CalculateOrigin(SheetModel sheet)
    {
        var ps = sheet.PageSetup;
        double pageW = ps.PageWidthPoints;
        double pageH = ps.PageHeightPoints;

        double marginLeftPt = ps.MarginLeft * 72.0;
        double marginRightPt = ps.MarginRight * 72.0;
        double marginTopPt = ps.MarginTop * 72.0;
        double marginBottomPt = ps.MarginBottom * 72.0;

        double contentW = _columnEngine.GetTotalWidthPoints(sheet);
        double contentH = _rowEngine.GetTotalHeightPoints(sheet);

        double originX, originY;

        if (ps.CenterHorizontally)
        {
            originX = (pageW - contentW) / 2.0;
        }
        else
        {
            originX = marginLeftPt;
        }

        if (ps.CenterVertically)
        {
            originY = (pageH - contentH) / 2.0;
        }
        else
        {
            originY = marginTopPt;
        }

        return (originX, originY);
    }
}
