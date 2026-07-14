namespace ExcelAPI.Models;

public sealed class LegacyRuntimeDocument
{
    public string BackgroundImage { get; set; } = "";
    public int PageWidth { get; set; }
    public int PageHeight { get; set; }
    public List<LegacyRuntimeField> Fields { get; set; } = new();
}

public sealed class LegacyRuntimeField
{
    public string Id { get; set; } = "";
    public string Label { get; set; } = "";
    public double LeftPx { get; set; }
    public double TopPx { get; set; }
    public double WidthPx { get; set; }
    public double HeightPx { get; set; }
    public string Type { get; set; } = "text";
    public bool Required { get; set; }
}
