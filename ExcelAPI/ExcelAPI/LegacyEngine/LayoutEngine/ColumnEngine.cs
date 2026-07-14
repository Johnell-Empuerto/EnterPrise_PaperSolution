using ExcelAPI.LegacyEngine.Models;

namespace ExcelAPI.LegacyEngine.LayoutEngine;

public class ColumnEngine : IColumnEngine
{
    public double GetLeftPoints(SheetModel sheet, int columnIndex)
    {
        double left = 0;
        foreach (var col in sheet.Columns)
        {
            if (col.Index < columnIndex)
                left += col.WidthPoints;
        }
        return left;
    }

    public double GetWidthPoints(SheetModel sheet, int startColumn, int endColumn)
    {
        double total = 0;
        foreach (var col in sheet.Columns)
        {
            if (col.Index >= startColumn && col.Index <= endColumn)
                total += col.WidthPoints;
        }
        return total;
    }

    public double GetTotalWidthPoints(SheetModel sheet)
    {
        return sheet.Columns.Sum(c => c.WidthPoints);
    }
}
