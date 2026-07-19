using System.Text.Json;

namespace ExcelAPI.Designer.Legacy.Models;

public class DebugSnapshot
{
    public string SessionId { get; set; } = Guid.NewGuid().ToString("N")[..12];
    public DateTime Timestamp { get; set; } = DateTime.UtcNow;

    public WorkbookModel? Workbook { get; set; }

    public string? SheetsPath { get; set; }
    public string? ColumnsPath { get; set; }
    public string? RowsPath { get; set; }

    public string? ClustersPath { get; set; }
    public string? CoordinatesBeforePath { get; set; }
    public string? CoordinatesAfterPath { get; set; }

    public string? XmlPath { get; set; }
    public string? BackgroundPath { get; set; }
    public string? VerificationPath { get; set; }
    public string? SummaryPath { get; set; }

    public void SaveTo(string directory)
    {
        Directory.CreateDirectory(directory);
        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };
        SummaryPath = Path.Combine(directory, "summary.json");
        File.WriteAllText(SummaryPath, JsonSerializer.Serialize(new
        {
            SessionId,
            Timestamp,
            Sheets = SheetsPath != null ? Path.GetFileName(SheetsPath) : null,
            Columns = ColumnsPath != null ? Path.GetFileName(ColumnsPath) : null,
            Rows = RowsPath != null ? Path.GetFileName(RowsPath) : null,
            Clusters = ClustersPath != null ? Path.GetFileName(ClustersPath) : null,
            CoordinatesBefore = CoordinatesBeforePath != null ? Path.GetFileName(CoordinatesBeforePath) : null,
            CoordinatesAfter = CoordinatesAfterPath != null ? Path.GetFileName(CoordinatesAfterPath) : null,
            Xml = XmlPath != null ? Path.GetFileName(XmlPath) : null,
            Background = BackgroundPath != null ? Path.GetFileName(BackgroundPath) : null,
            Verification = VerificationPath != null ? Path.GetFileName(VerificationPath) : null
        }, options));
    }
}
