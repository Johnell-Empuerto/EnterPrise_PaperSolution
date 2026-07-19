namespace ExcelAPI.Designer.Legacy.Models;

public class SheetModel
{
    public string Name { get; set; } = "";
    public int SheetIndex { get; set; }
    public PrintAreaInfo PrintArea { get; set; } = new();
    public List<ColumnModel> Columns { get; set; } = new();
    public List<RowModel> Rows { get; set; } = new();
    public List<CellModel> Cells { get; set; } = new();
    public List<CellModel> MergedCells { get; set; } = new();
    public List<CommentModel> Comments { get; set; } = new();
    public PageSetupModel PageSetup { get; set; } = new();
    public double ContentWidthPoints => Columns.Sum(c => c.WidthPoints);
    public double ContentHeightPoints => Rows.Sum(r => r.HeightPoints);
}

public class ColumnModel
{
    public int Index { get; set; }
    public double WidthPoints { get; set; }
}

public class RowModel
{
    public int Index { get; set; }
    public double HeightPoints { get; set; }
}

public class CellModel
{
    public int Row { get; set; }
    public int Column { get; set; }
    public int EndRow { get; set; }
    public int EndColumn { get; set; }
    public string? Value { get; set; }
    public string? Reference =>
        IsMergeCell
            ? $"${GetColumnLetter(Column)}${Row}:${GetColumnLetter(EndColumn)}${EndRow}"
            : $"${GetColumnLetter(Column)}${Row}";
    public bool IsMergeCell => Row != EndRow || Column != EndColumn;

    public static string GetColumnLetter(int col)
    {
        if (col < 1) return "";
        col--;
        string result = "";
        while (col >= 0)
        {
            result = (char)('A' + col % 26) + result;
            col = col / 26 - 1;
        }
        return result;
    }
}

public class PrintAreaInfo
{
    public string Address { get; set; } = "";
    public int StartRow { get; set; } = 1;
    public int EndRow { get; set; } = 1;
    public int StartColumn { get; set; } = 1;
    public int EndColumn { get; set; } = 1;
    public bool IsValid => !string.IsNullOrEmpty(Address);
}

public class CommentModel
{
    public int Row { get; set; }
    public int Column { get; set; }
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public string Text { get; set; } = "";

    public int EndRow => Row + RowCount - 1;
    public int EndColumn => Column + ColumnCount - 1;

    public string CellAddress
    {
        get
        {
            string start = "$" + CellModel.GetColumnLetter(Column) + "$" + Row;
            if (RowCount == 1 && ColumnCount == 1)
                return start;
            string end = "$" + CellModel.GetColumnLetter(Column + ColumnCount - 1) + "$" + (Row + RowCount - 1);
            return start + ":" + end;
        }
    }
}

public class PageSetupModel
{
    public double PageWidthPoints { get; set; } = 612;
    public double PageHeightPoints { get; set; } = 792;
    public double MarginLeft { get; set; } = 0.75;
    public double MarginRight { get; set; } = 0.75;
    public double MarginTop { get; set; } = 0.75;
    public double MarginBottom { get; set; } = 0.75;
    public double MarginHeader { get; set; } = 0.5;
    public double MarginFooter { get; set; } = 0.5;
    public bool CenterHorizontally { get; set; }
    public bool CenterVertically { get; set; }
    public int? Zoom { get; set; }
    public int? FitToPagesWide { get; set; }
    public int? FitToPagesTall { get; set; }
    public string Orientation { get; set; } = "Portrait";
    public double PaperWidthPoints { get; set; } = 612;
    public double PaperHeightPoints { get; set; } = 792;
}
