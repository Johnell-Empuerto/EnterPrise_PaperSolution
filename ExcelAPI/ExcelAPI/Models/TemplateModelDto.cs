namespace ExcelAPI.Models;

public sealed class TemplateModelDto
{
    public string SheetName { get; set; } = "";
    public TemplatePageSetupDto PageSetup { get; set; } = new();
    public TemplatePrintAreaDto PrintArea { get; set; } = new();
    public List<double> ColumnWidths { get; set; } = new();
    public List<double> RowHeights { get; set; } = new();
    public List<TemplateMergedCellDto> MergedCells { get; set; } = new();
    public Dictionary<string, string> CellValues { get; set; } = new();
    public Dictionary<string, TemplateCellStyleDto> CellStyles { get; set; } = new();
    public List<TemplateCommentDto> Comments { get; set; } = new();
    public List<TemplateImageDto> Images { get; set; } = new();
}

public sealed class TemplatePageSetupDto
{
    public double PaperWidthPt { get; set; }
    public double PaperHeightPt { get; set; }
    public double MarginLeftIn { get; set; }
    public double MarginRightIn { get; set; }
    public double MarginTopIn { get; set; }
    public double MarginBottomIn { get; set; }
    public double MarginHeaderIn { get; set; }
    public double MarginFooterIn { get; set; }
    public bool CenterHorizontally { get; set; }
    public bool CenterVertically { get; set; }
    public string Orientation { get; set; } = "Portrait";
    public int PaperSize { get; set; } = 1;
}

public sealed class TemplatePrintAreaDto
{
    public int StartCol { get; set; }
    public int StartRow { get; set; }
    public int EndCol { get; set; }
    public int EndRow { get; set; }
}

public sealed class TemplateMergedCellDto
{
    public int StartCol { get; set; }
    public int StartRow { get; set; }
    public int EndCol { get; set; }
    public int EndRow { get; set; }
}

public sealed class TemplateCellStyleDto
{
    public string? FontName { get; set; }
    public double? FontSize { get; set; }
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public bool? Underline { get; set; }
    public string? FontColor { get; set; }
    public string? FillColor { get; set; }
    public string? BorderTop { get; set; }
    public string? BorderBottom { get; set; }
    public string? BorderLeft { get; set; }
    public string? BorderRight { get; set; }
    public string? HorizontalAlignment { get; set; }
    public string? VerticalAlignment { get; set; }
    public bool? WrapText { get; set; }
    public int? Indent { get; set; }
}

public sealed class TemplateCommentDto
{
    public string Address { get; set; } = "";
    public string Text { get; set; } = "";
}

public sealed class TemplateImageDto
{
    public int AnchorCol { get; set; }
    public int AnchorRow { get; set; }
    public double WidthPt { get; set; }
    public double HeightPt { get; set; }
}
