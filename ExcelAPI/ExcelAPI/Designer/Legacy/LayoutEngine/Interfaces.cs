using ExcelAPI.Designer.Legacy.Models;

namespace ExcelAPI.Designer.Legacy.LayoutEngine;

public interface IColumnEngine
{
    double GetLeftPoints(SheetModel sheet, int columnIndex);
    double GetWidthPoints(SheetModel sheet, int startColumn, int endColumn);
    double GetTotalWidthPoints(SheetModel sheet);
}

public interface IRowEngine
{
    double GetTopPoints(SheetModel sheet, int rowIndex);
    double GetHeightPoints(SheetModel sheet, int startRow, int endRow);
    double GetTotalHeightPoints(SheetModel sheet);
}

public interface IOriginEngine
{
    (double originX, double originY) CalculateOrigin(SheetModel sheet);
}

public interface IPageEngine
{
    double PageWidthPoints { get; }
    double PageHeightPoints { get; }
}
