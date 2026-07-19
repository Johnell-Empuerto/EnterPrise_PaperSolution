namespace ExcelAPI.Models
{
    public class FormDefinition
    {
        public WorkbookMetadata Workbook { get; set; } = new();
        public List<SheetDefinition> Sheets { get; set; } = new();
        public List<ClusterDefinition> Clusters { get; set; } = new();
        public PageSettings? FieldsPageSettings { get; set; }
        public List<ImageDefinition> Images { get; set; } = new();
        public Dictionary<string, string> Metadata { get; set; } = new();
    }

    public class WorkbookMetadata
    {
        public string Title { get; set; } = "Untitled Form";
        public string Author { get; set; } = "";
        public string Created { get; set; } = "";
        public string Modified { get; set; } = "";
        public string Version { get; set; } = "1.0";
        public string Description { get; set; } = "";
    }

    public class SheetDefinition
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public int Index { get; set; }
        public PageSettings PageSettings { get; set; } = new();
        public PrintAreaInfo? PrintArea { get; set; }
        public string? BackgroundImage { get; set; }
        public int BackgroundWidth { get; set; }
        public int BackgroundHeight { get; set; }
        public string? Thumbnail { get; set; }
        public Dictionary<int, double> RowHeights { get; set; } = new();
        public Dictionary<int, double> ColumnWidths { get; set; } = new();
        public List<MergedCellInfo> MergedCells { get; set; } = new();
        public string? FreezePane { get; set; }
        public Dictionary<string, CellStyleInfo> CellStyles { get; set; } = new();
        public Dictionary<string, string> CellValues { get; set; } = new();
        public Dictionary<string, string> Metadata { get; set; } = new();
    }

    public class PageSettings
    {
        public string PaperSize { get; set; } = "Letter";
        public string Orientation { get; set; } = "portrait";
        public double WidthPt { get; set; } = 612;
        public double HeightPt { get; set; } = 792;
        public double LeftMargin { get; set; } = 70;
        public double TopMargin { get; set; } = 70;
        public double RightMargin { get; set; } = 70;
        public double BottomMargin { get; set; } = 70;
        public bool CenterHorizontally { get; set; }
        public bool CenterVertically { get; set; }
        public int Zoom { get; set; } = 100;
        public int FitToPagesWide { get; set; }
        public int FitToPagesTall { get; set; }
    }

    public class PrintAreaInfo
    {
        public string Address { get; set; } = "";
        public double LeftPt { get; set; }
        public double TopPt { get; set; }
        public double WidthPt { get; set; }
        public double HeightPt { get; set; }
        public int Cols { get; set; }
        public int Rows { get; set; }
    }

    public class MergedCellInfo
    {
        public string Address { get; set; } = "";
        public string CellAddress { get; set; } = "";
        public double LeftPt { get; set; }
        public double TopPt { get; set; }
        public double WidthPt { get; set; }
        public double HeightPt { get; set; }
    }

    public class CellStyleInfo
    {
        public string? FontName { get; set; }
        public double? FontSize { get; set; }
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public bool? Underline { get; set; }
        public string? Color { get; set; }
        public string? FillColor { get; set; }
        public string? BorderTop { get; set; }
        public string? BorderBottom { get; set; }
        public string? BorderLeft { get; set; }
        public string? BorderRight { get; set; }
        public string? HorizontalAlignment { get; set; }
        public string? VerticalAlignment { get; set; }
        public bool? WrapText { get; set; }
    }

    public class ClusterDefinition
    {
        public string ClusterId { get; set; } = "";
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public string SheetId { get; set; } = "";
        public string CellAddress { get; set; } = "";
        public double Left { get; set; }
        public double Right { get; set; }
        public double Top { get; set; }
        public double Bottom { get; set; }
        public double LeftPt { get; set; }
        public double TopPt { get; set; }
        public double WidthPt { get; set; }
        public double HeightPt { get; set; }
        public Dictionary<string, string> InputParameters { get; set; } = new();
        public string Visibility { get; set; } = "visible";
        public bool Readonly { get; set; }
        public string Remarks { get; set; } = "";
        public List<string> Functions { get; set; } = new();
        public Dictionary<string, string> Metadata { get; set; } = new();
    }

    public class ImageDefinition
    {
        public string Id { get; set; } = "";
        public string SheetId { get; set; } = "";
        public string Name { get; set; } = "";
        public double LeftPt { get; set; }
        public double TopPt { get; set; }
        public double WidthPt { get; set; }
        public double HeightPt { get; set; }
        public byte[]? Data { get; set; }
        public string Format { get; set; } = "png";
    }
}
