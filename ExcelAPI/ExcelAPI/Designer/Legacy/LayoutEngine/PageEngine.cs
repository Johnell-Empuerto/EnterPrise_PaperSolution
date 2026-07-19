using ExcelAPI.Designer.Legacy.Models;

namespace ExcelAPI.Designer.Legacy.LayoutEngine;

public class PageEngine : IPageEngine
{
    public double PageWidthPoints { get; }
    public double PageHeightPoints { get; }

    public PageEngine()
    {
        PageWidthPoints = 612;
        PageHeightPoints = 792;
    }

    public static (double width, double height) GetPageSize(string orientation, double paperWidthInch, double paperHeightInch)
    {
        double w = paperWidthInch * 72.0;
        double h = paperHeightInch * 72.0;

        if (orientation == "Landscape" && w < h)
            return (h, w);

        return (w, h);
    }
}
