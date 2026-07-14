using System.Text.Json.Serialization;

namespace LegacyExtractionEngine.Models;

public class ComparisonReport
{
    [JsonPropertyName("overall")]
    public string Overall { get; set; } = "";
    [JsonPropertyName("def_top_id")]
    public int DefTopId { get; set; }
    [JsonPropertyName("generated_at")]
    public DateTime GeneratedAt { get; set; } = DateTime.UtcNow;
    [JsonPropertyName("sections")]
    public Dictionary<string, ComparisonSection> Sections { get; set; } = new();
    [JsonPropertyName("summary")]
    public ComparisonSummary Summary { get; set; } = new();
}

public class ComparisonSection
{
    [JsonPropertyName("status")]
    public string Status { get; set; } = ""; // PASS, FAIL, DIFFERENT, MISSING, EXTRA
    [JsonPropertyName("matches")]
    public int Matches { get; set; }
    [JsonPropertyName("mismatches")]
    public int Mismatches { get; set; }
    [JsonPropertyName("missing")]
    public int Missing { get; set; }
    [JsonPropertyName("extra")]
    public int Extra { get; set; }
    [JsonPropertyName("match_percentage")]
    public double MatchPercentage { get; set; }
    [JsonPropertyName("details")]
    public List<ComparisonDetail> Details { get; set; } = new();
}

public class ComparisonDetail
{
    [JsonPropertyName("property")]
    public string Property { get; set; } = "";
    [JsonPropertyName("status")]
    public string Status { get; set; } = ""; // PASS, FAIL, DIFFERENT, MISSING, EXTRA
    [JsonPropertyName("database_value")]
    public object? DatabaseValue { get; set; }
    [JsonPropertyName("generated_value")]
    public object? GeneratedValue { get; set; }
    [JsonPropertyName("reason")]
    public string? Reason { get; set; }
    [JsonPropertyName("suggested_fix")]
    public string? SuggestedFix { get; set; }
}

public class ComparisonSummary
{
    [JsonPropertyName("total_sections")]
    public int TotalSections { get; set; }
    [JsonPropertyName("passed")]
    public int Passed { get; set; }
    [JsonPropertyName("failed")]
    public int Failed { get; set; }
    [JsonPropertyName("different")]
    public int Different { get; set; }
    [JsonPropertyName("missing")]
    public int Missing { get; set; }
    [JsonPropertyName("extra")]
    public int Extra { get; set; }
    [JsonPropertyName("overall_match_percentage")]
    public double OverallMatchPercentage { get; set; }
}

public class DeepDiffResult
{
    [JsonPropertyName("diff_type")]
    public string DiffType { get; set; } = ""; // EQUAL, MODIFIED, ADDED, REMOVED, TYPE_MISMATCH
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";
    [JsonPropertyName("database_value")]
    public object? DatabaseValue { get; set; }
    [JsonPropertyName("generated_value")]
    public object? GeneratedValue { get; set; }
    [JsonPropertyName("database_type")]
    public string? DatabaseType { get; set; }
    [JsonPropertyName("generated_type")]
    public string? GeneratedType { get; set; }
    [JsonPropertyName("reason")]
    public string? Reason { get; set; }
    [JsonPropertyName("children")]
    public List<DeepDiffResult> Children { get; set; } = new();
}

public class DifferenceReport
{
    [JsonPropertyName("def_top_id")]
    public int DefTopId { get; set; }
    [JsonPropertyName("generated_at")]
    public DateTime GeneratedAt { get; set; } = DateTime.UtcNow;
    [JsonPropertyName("total_differences")]
    public int TotalDifferences { get; set; }
    [JsonPropertyName("differences")]
    public List<DifferenceEntry> Differences { get; set; } = new();
}

public class DifferenceEntry
{
    [JsonPropertyName("section")]
    public string Section { get; set; } = "";
    [JsonPropertyName("property")]
    public string Property { get; set; } = "";
    [JsonPropertyName("database_value")]
    public string? DatabaseValue { get; set; }
    [JsonPropertyName("generated_value")]
    public string? GeneratedValue { get; set; }
    [JsonPropertyName("reason")]
    public string Reason { get; set; } = "";
    [JsonPropertyName("suggested_fix")]
    public string SuggestedFix { get; set; } = "";
}
