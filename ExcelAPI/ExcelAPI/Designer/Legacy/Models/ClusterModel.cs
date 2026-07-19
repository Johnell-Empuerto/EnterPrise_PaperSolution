namespace ExcelAPI.Designer.Legacy.Models;

public class ClusterModel
{
    public int ClusterId { get; set; }
    public int SheetIndex { get; set; }

    public int StartRow { get; set; }
    public int EndRow { get; set; }
    public int StartColumn { get; set; }
    public int EndColumn { get; set; }

    public string CellAddress { get; set; } = "";
    public string? Comment { get; set; }
    public string? TypeKey { get; set; }
    public string? TypeName { get; set; }
    public bool IsSelected { get; set; }
    public bool IsUnknown { get; set; }
    public string CommentName { get; set; } = "";
    public string CommentType { get; set; } = "";

    public int RowSpan => EndRow - StartRow + 1;
    public int ColumnSpan => EndColumn - StartColumn + 1;

    public double WidthPoints { get; set; }
    public double HeightPoints { get; set; }
    public double LeftPoints { get; set; }
    public double TopPoints { get; set; }
    public double AreaPercent { get; set; }
    public double AspectRatio { get; set; }

    public CoordinateRect? CoordinatesBeforeTransform { get; set; }
    public CoordinateRect? CoordinatesAfterTransform { get; set; }
}

public class CoordinateRect
{
    public double Left { get; set; }
    public double Top { get; set; }
    public double Right { get; set; }
    public double Bottom { get; set; }
    public double Width => Right - Left;
    public double Height => Bottom - Top;
}
