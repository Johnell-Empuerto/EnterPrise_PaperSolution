using System.Text.Json.Serialization;

namespace LegacyExtractionEngine.Models;

public class DefTop
{
    [JsonPropertyName("def_top_id")]
    public int DefTopId { get; set; }
    [JsonPropertyName("def_top_org")]
    public int? DefTopOrg { get; set; }
    [JsonPropertyName("rev_no")]
    public int? RevNo { get; set; }
    [JsonPropertyName("def_top_name")]
    public string? DefTopName { get; set; }
    [JsonPropertyName("report_type")]
    public decimal? ReportType { get; set; }
    [JsonPropertyName("def_sheet_count")]
    public int? DefSheetCount { get; set; }
    [JsonPropertyName("auto_gen")]
    public decimal? AutoGen { get; set; }
    [JsonPropertyName("public_status")]
    public decimal? PublicStatus { get; set; }
    [JsonPropertyName("thumbnail")]
    public string? Thumbnail { get; set; }
    [JsonPropertyName("xml_data")]
    public string? XmlData { get; set; }
    [JsonPropertyName("thumbnail_file")]
    public byte[]? ThumbnailFile { get; set; }
    [JsonPropertyName("background_image_file")]
    public byte[]? BackgroundImageFile { get; set; }
    [JsonPropertyName("designer_version")]
    public string? DesignerVersion { get; set; }
    [JsonPropertyName("designer_display_version")]
    public string? DesignerDisplayVersion { get; set; }
    [JsonPropertyName("image_size")]
    public int? ImageSize { get; set; }
    [JsonPropertyName("def_file")]
    public byte[]? DefFile { get; set; }
    [JsonPropertyName("pdf_auto_output_sheet_select")]
    public string? PdfAutoOutputSheetSelect { get; set; }
    [JsonPropertyName("server_version")]
    public string? ServerVersion { get; set; }
    [JsonPropertyName("updated_at")]
    public DateTime? UpdatedAt { get; set; }
    [JsonPropertyName("sys_regist_time")]
    public DateTime? SysRegistTime { get; set; }
    [JsonPropertyName("all_columns")]
    public Dictionary<string, object?> AllColumns { get; set; } = new();
}

public class DefSheet
{
    [JsonPropertyName("def_sheet_id")]
    public int DefSheetId { get; set; }
    [JsonPropertyName("def_top_id")]
    public int DefTopId { get; set; }
    [JsonPropertyName("def_sheet_no")]
    public int? DefSheetNo { get; set; }
    [JsonPropertyName("def_sheet_name")]
    public string? DefSheetName { get; set; }
    [JsonPropertyName("xml_data")]
    public string? XmlData { get; set; }
    [JsonPropertyName("background_image_file")]
    public byte[]? BackgroundImageFile { get; set; }
    [JsonPropertyName("thumbnail_file")]
    public byte[]? ThumbnailFile { get; set; }
    [JsonPropertyName("updated_at")]
    public DateTime? UpdatedAt { get; set; }
    [JsonPropertyName("all_columns")]
    public Dictionary<string, object?> AllColumns { get; set; } = new();
    [JsonPropertyName("clusters")]
    public List<DefCluster> Clusters { get; set; } = new();
}

public class DefCluster
{
    [JsonPropertyName("def_sheet_id")]
    public int DefSheetId { get; set; }
    [JsonPropertyName("cluster_id")]
    public int ClusterId { get; set; }
    [JsonPropertyName("cluster_name")]
    public string? ClusterName { get; set; }
    [JsonPropertyName("cluster_type")]
    public string? ClusterType { get; set; }
    [JsonPropertyName("left_position")]
    public string? LeftPosition { get; set; }
    [JsonPropertyName("right_position")]
    public string? RightPosition { get; set; }
    [JsonPropertyName("top_position")]
    public string? TopPosition { get; set; }
    [JsonPropertyName("bottom_position")]
    public string? BottomPosition { get; set; }
    [JsonPropertyName("cell_addr")]
    public string? CellAddr { get; set; }
    [JsonPropertyName("input_parameter")]
    public string? InputParameter { get; set; }
    [JsonPropertyName("updated_at")]
    public DateTime? UpdatedAt { get; set; }
    [JsonPropertyName("all_columns")]
    public Dictionary<string, object?> AllColumns { get; set; } = new();
}

public class DefTopSize
{
    [JsonPropertyName("def_top_id")]
    public int DefTopId { get; set; }
    [JsonPropertyName("excel_size")]
    public long? ExcelSize { get; set; }
    [JsonPropertyName("pdf_size")]
    public long? PdfSize { get; set; }
    [JsonPropertyName("sheet_count")]
    public long? SheetCount { get; set; }
    [JsonPropertyName("cluster_count")]
    public long? ClusterCount { get; set; }
    [JsonPropertyName("updated_at")]
    public DateTime? UpdatedAt { get; set; }
}

public class DefClusterRefer
{
    [JsonPropertyName("def_sheet_id")]
    public int DefSheetId { get; set; }
    [JsonPropertyName("cluster_id")]
    public int ClusterId { get; set; }
    [JsonPropertyName("all_columns")]
    public Dictionary<string, object?> AllColumns { get; set; } = new();
}

public class DefLabel
{
    [JsonPropertyName("def_label_id")]
    public int DefLabelId { get; set; }
    [JsonPropertyName("def_top_id")]
    public int DefTopId { get; set; }
    [JsonPropertyName("all_columns")]
    public Dictionary<string, object?> AllColumns { get; set; } = new();
}

public class DefCurrent
{
    [JsonPropertyName("all_columns")]
    public Dictionary<string, object?> AllColumns { get; set; } = new();
}

public class DefTopOption
{
    [JsonPropertyName("def_top_id")]
    public int DefTopId { get; set; }
    [JsonPropertyName("all_columns")]
    public Dictionary<string, object?> AllColumns { get; set; } = new();
}

public class DatabaseDump
{
    [JsonPropertyName("def_top")]
    public DefTop? DefTop { get; set; }
    [JsonPropertyName("def_sheets")]
    public List<DefSheet> DefSheets { get; set; } = new();
    [JsonPropertyName("def_top_size")]
    public DefTopSize? DefTopSize { get; set; }
    [JsonPropertyName("def_top_options")]
    public List<DefTopOption> DefTopOptions { get; set; } = new();
    [JsonPropertyName("def_cluster_refers")]
    public List<DefClusterRefer> DefClusterRefers { get; set; } = new();
    [JsonPropertyName("def_labels")]
    public List<DefLabel> DefLabels { get; set; } = new();
    [JsonPropertyName("def_current")]
    public DefCurrent? DefCurrent { get; set; }
    [JsonPropertyName("all_related_tables")]
    public Dictionary<string, List<Dictionary<string, object?>>> AllRelatedTables { get; set; } = new();
    [JsonPropertyName("extracted_at")]
    public DateTime ExtractedAt { get; set; } = DateTime.UtcNow;
    [JsonPropertyName("source")]
    public string Source { get; set; } = "database";
}
