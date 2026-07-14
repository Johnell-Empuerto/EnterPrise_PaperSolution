using System.Text;
using System.Xml.Linq;

namespace ExcelAPI.LegacyEngine.Models;

public class PublishResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public string? XmlData { get; set; }
    public string? BackgroundImagePath { get; set; }
    public string? ThumbnailPath { get; set; }
    public List<PublishCluster> Clusters { get; set; } = new();
    public List<PublishSheetInfo> Sheets { get; set; } = new();
}

public class PublishCluster
{
    public int ClusterId { get; set; }
    public int SheetNo { get; set; }
    public string CellAddress { get; set; } = "";
    public double Left { get; set; }
    public double Top { get; set; }
    public double Right { get; set; }
    public double Bottom { get; set; }
    public string? Value { get; set; }
    public string? Type { get; set; }
    public string? InputParameters { get; set; }
}

public class PublishSheetInfo
{
    public int SheetNo { get; set; }
    public string SheetName { get; set; } = "";
    public double WidthPoints { get; set; }
    public double HeightPoints { get; set; }
    public string? BackgroundImagePath { get; set; }
}
