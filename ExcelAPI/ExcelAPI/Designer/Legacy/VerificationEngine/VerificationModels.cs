using System.Text.Json.Serialization;

namespace ExcelAPI.Designer.Legacy.VerificationEngine;

public class VerificationReport
{
    [JsonPropertyName("defTopId")]
    public int DefTopId { get; set; }

    [JsonPropertyName("generatedAt")]
    public DateTime GeneratedAt { get; set; } = DateTime.UtcNow;

    [JsonPropertyName("overall")]
    public string Overall { get; set; } = "";

    [JsonPropertyName("xmlComparison")]
    public XmlComparisonResult? XmlComparison { get; set; }

    [JsonPropertyName("coordinateComparison")]
    public CoordinateComparisonResult? CoordinateComparison { get; set; }

    [JsonPropertyName("imageComparison")]
    public ImageComparisonResult? ImageComparison { get; set; }

    [JsonPropertyName("summary")]
    public VerificationSummary Summary { get; set; } = new();
}

public class VerificationSummary
{
    [JsonPropertyName("totalChecks")]
    public int TotalChecks { get; set; }

    [JsonPropertyName("passed")]
    public int Passed { get; set; }

    [JsonPropertyName("failed")]
    public int Failed { get; set; }

    [JsonPropertyName("overallMatchPercentage")]
    public double OverallMatchPercentage { get; set; }
}

public class XmlComparisonResult
{
    [JsonPropertyName("status")]
    public string Status { get; set; } = "";

    [JsonPropertyName("nodeCountGenerated")]
    public int NodeCountGenerated { get; set; }

    [JsonPropertyName("nodeCountDatabase")]
    public int NodeCountDatabase { get; set; }

    [JsonPropertyName("matchingNodes")]
    public int MatchingNodes { get; set; }

    [JsonPropertyName("differingNodes")]
    public int DifferingNodes { get; set; }

    [JsonPropertyName("missingNodes")]
    public int MissingNodes { get; set; }

    [JsonPropertyName("extraNodes")]
    public int ExtraNodes { get; set; }

    [JsonPropertyName("matchPercentage")]
    public double MatchPercentage { get; set; }

    [JsonPropertyName("differences")]
    public List<XmlNodeDiff> Differences { get; set; } = new();
}

public class XmlNodeDiff
{
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("status")]
    public string Status { get; set; } = "";

    [JsonPropertyName("attribute")]
    public string? Attribute { get; set; }

    [JsonPropertyName("databaseValue")]
    public string? DatabaseValue { get; set; }

    [JsonPropertyName("generatedValue")]
    public string? GeneratedValue { get; set; }
}

public class CoordinateComparisonResult
{
    [JsonPropertyName("status")]
    public string Status { get; set; } = "";

    [JsonPropertyName("totalClusters")]
    public int TotalClusters { get; set; }

    [JsonPropertyName("passed")]
    public int Passed { get; set; }

    [JsonPropertyName("failed")]
    public int Failed { get; set; }

    [JsonPropertyName("matchPercentage")]
    public double MatchPercentage { get; set; }

    [JsonPropertyName("differences")]
    public List<CoordinateDiff> Differences { get; set; } = new();
}

public class CoordinateDiff
{
    [JsonPropertyName("clusterId")]
    public int ClusterId { get; set; }

    [JsonPropertyName("sheetNo")]
    public int SheetNo { get; set; }

    [JsonPropertyName("cellAddress")]
    public string CellAddress { get; set; } = "";

    [JsonPropertyName("status")]
    public string Status { get; set; } = "";

    [JsonPropertyName("expectedLeft")]
    public double ExpectedLeft { get; set; }

    [JsonPropertyName("actualLeft")]
    public double ActualLeft { get; set; }

    [JsonPropertyName("diffLeftPoints")]
    public double DiffLeftPoints { get; set; }

    [JsonPropertyName("diffLeftRatio")]
    public double DiffLeftRatio { get; set; }

    [JsonPropertyName("expectedTop")]
    public double ExpectedTop { get; set; }

    [JsonPropertyName("actualTop")]
    public double ActualTop { get; set; }

    [JsonPropertyName("diffTopPoints")]
    public double DiffTopPoints { get; set; }

    [JsonPropertyName("diffTopRatio")]
    public double DiffTopRatio { get; set; }

    [JsonPropertyName("expectedRight")]
    public double ExpectedRight { get; set; }

    [JsonPropertyName("actualRight")]
    public double ActualRight { get; set; }

    [JsonPropertyName("diffRightPoints")]
    public double DiffRightPoints { get; set; }

    [JsonPropertyName("diffRightRatio")]
    public double DiffRightRatio { get; set; }

    [JsonPropertyName("expectedBottom")]
    public double ExpectedBottom { get; set; }

    [JsonPropertyName("actualBottom")]
    public double ActualBottom { get; set; }

    [JsonPropertyName("diffBottomPoints")]
    public double DiffBottomPoints { get; set; }

    [JsonPropertyName("diffBottomRatio")]
    public double DiffBottomRatio { get; set; }
}

public class ImageComparisonResult
{
    [JsonPropertyName("status")]
    public string Status { get; set; } = "";

    [JsonPropertyName("databaseImageExists")]
    public bool DatabaseImageExists { get; set; }

    [JsonPropertyName("generatedImageExists")]
    public bool GeneratedImageExists { get; set; }

    [JsonPropertyName("databaseWidth")]
    public int? DatabaseWidth { get; set; }

    [JsonPropertyName("databaseHeight")]
    public int? DatabaseHeight { get; set; }

    [JsonPropertyName("generatedWidth")]
    public int? GeneratedWidth { get; set; }

    [JsonPropertyName("generatedHeight")]
    public int? GeneratedHeight { get; set; }

    [JsonPropertyName("mse")]
    public double? Mse { get; set; }

    [JsonPropertyName("ssim")]
    public double? Ssim { get; set; }

    [JsonPropertyName("pixelDiffPercent")]
    public double? PixelDiffPercent { get; set; }

    [JsonPropertyName("similarityScore")]
    public double? SimilarityScore { get; set; }
}

public class DefTopData
{
    public int DefTopId { get; set; }
    public string? XmlData { get; set; }
    public byte[]? BackgroundImageData { get; set; }
}

public class DefClusterData
{
    public int DefSheetId { get; set; }
    public int ClusterId { get; set; }
    public string? ClusterName { get; set; }
    public string? ClusterType { get; set; }
    public string? LeftPosition { get; set; }
    public string? RightPosition { get; set; }
    public string? TopPosition { get; set; }
    public string? BottomPosition { get; set; }
    public string? CellAddr { get; set; }
    public string? InputParameter { get; set; }
}
