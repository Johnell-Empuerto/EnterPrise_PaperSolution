using ExcelAPI.LegacyEngine.Models;

namespace ExcelAPI.LegacyEngine.LayoutEngine;

public class RowEngine : IRowEngine
{
    public double GetTopPoints(SheetModel sheet, int rowIndex)
    {
        double top = 0;
        foreach (var row in sheet.Rows)
        {
            if (row.Index < rowIndex)
                top += row.HeightPoints;
        }
        return top;
    }

    public double GetHeightPoints(SheetModel sheet, int startRow, int endRow)
    {
        double total = 0;
        foreach (var row in sheet.Rows)
        {
            if (row.Index >= startRow && row.Index <= endRow)
                total += row.HeightPoints;
        }
        return total;
    }

    public double GetTotalHeightPoints(SheetModel sheet)
    {
        return sheet.Rows.Sum(r => r.HeightPoints);
    }
}
