namespace ExcelAPI.Designer.Legacy.OpenXml;

public sealed class OpenXmlWorkbookModel
{
    public string FilePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public IReadOnlyList<OpenXmlSheetModel> Sheets { get; set; } = Array.Empty<OpenXmlSheetModel>();
}

public sealed class OpenXmlSheetModel
{
    public string Name { get; set; } = "";
    public int SheetIndex { get; set; }
    public int SheetId { get; set; }
    public string RelationshipId { get; set; } = "";
    public OpenXmlPrintAreaModel PrintArea { get; set; } = new();
    public OpenXmlPageSetupModel PageSetup { get; set; } = new();
    public IReadOnlyList<double> ColumnWidths { get; set; } = Array.Empty<double>();
    public IReadOnlyList<double> RowHeights { get; set; } = Array.Empty<double>();
    public IReadOnlyList<OpenXmlMergeAreaModel> MergeAreas { get; set; } = Array.Empty<OpenXmlMergeAreaModel>();
    public IReadOnlyList<OpenXmlCommentModel> Comments { get; set; } = Array.Empty<OpenXmlCommentModel>();
    public IReadOnlyList<OpenXmlImageModel> Images { get; set; } = Array.Empty<OpenXmlImageModel>();
    public IReadOnlyList<OpenXmlCellModel> Cells { get; set; } = Array.Empty<OpenXmlCellModel>();
    public IReadOnlyDictionary<string, OpenXmlCellStyleModel> CellStyles { get; set; } = new Dictionary<string, OpenXmlCellStyleModel>();
}

public sealed class OpenXmlPrintAreaModel
{
    public bool IsValid => StartColumn > 0 && StartRow > 0;
    public int StartColumn { get; set; }
    public int StartRow { get; set; }
    public int EndColumn { get; set; }
    public int EndRow { get; set; }
}

public sealed class OpenXmlPageSetupModel
{
    public double PageWidthPoints { get; set; } = 612;
    public double PageHeightPoints { get; set; } = 792;
    public double PaperWidthPoints { get; set; } = 612;
    public double PaperHeightPoints { get; set; } = 792;
    public double MarginLeft { get; set; } = 0.75;
    public double MarginRight { get; set; } = 0.75;
    public double MarginTop { get; set; } = 0.75;
    public double MarginBottom { get; set; } = 0.75;
    public double MarginHeader { get; set; } = 0.5;
    public double MarginFooter { get; set; } = 0.5;
    public bool CenterHorizontally { get; set; }
    public bool CenterVertically { get; set; }
    public string Orientation { get; set; } = "Portrait";
}

public sealed class OpenXmlMergeAreaModel
{
    public int StartRow { get; set; }
    public int StartColumn { get; set; }
    public int EndRow { get; set; }
    public int EndColumn { get; set; }
    public string Reference => $"${GetColumnLetter(StartColumn)}${StartRow}:${GetColumnLetter(EndColumn)}${EndRow}";

    private static string GetColumnLetter(int col)
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

public sealed class OpenXmlCommentModel
{
    public int Row { get; set; }
    public int Column { get; set; }
    public int RowCount { get; set; } = 1;
    public int ColumnCount { get; set; } = 1;
    public string Text { get; set; } = "";
    public int EndRow => Row + RowCount - 1;
    public int EndColumn => Column + ColumnCount - 1;
}

public sealed class OpenXmlImageModel
{
    public int Row { get; set; }
    public int Column { get; set; }
    public int RowOffsetPx { get; set; }
    public int ColumnOffsetPx { get; set; }
    public long ImageBytesLength { get; set; }
    public string ContentType { get; set; } = "";
    public string FileName { get; set; } = "";
}

public sealed class OpenXmlStyleModel
{
    public double DefaultFontSize { get; set; } = 11;
    public string DefaultFontName { get; set; } = "Calibri";
    public double MaxDigitWidth { get; set; } = 7;
}

public sealed class OpenXmlCellModel
{
    public int Row { get; set; }
    public int Column { get; set; }
    public string Value { get; set; } = "";
    public int StyleIndex { get; set; } = -1;
}

public sealed class OpenXmlCellStyleModel
{
    public string FontName { get; set; } = "";
    public double FontSize { get; set; }
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public string FontColor { get; set; } = "";
    public string FillColor { get; set; } = "";
    public string BorderTop { get; set; } = "";
    public string BorderBottom { get; set; } = "";
    public string BorderLeft { get; set; } = "";
    public string BorderRight { get; set; } = "";
    public string HorizontalAlignment { get; set; } = "";
    public string VerticalAlignment { get; set; } = "";
    public bool WrapText { get; set; }
    public int Indent { get; set; }
}
