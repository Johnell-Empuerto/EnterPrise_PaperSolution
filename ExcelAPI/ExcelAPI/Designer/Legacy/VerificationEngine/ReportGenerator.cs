using System.Text;
using System.Text.Json;

namespace ExcelAPI.Designer.Legacy.VerificationEngine;

public class VerificationReportGenerator : IVerificationReportGenerator
{
    public async Task<string> GenerateJsonAsync(VerificationReport report, string outputDir)
    {
        Directory.CreateDirectory(outputDir);
        string path = Path.Combine(outputDir, "report.json");
        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };
        string json = JsonSerializer.Serialize(report, options);
        await File.WriteAllTextAsync(path, json);
        return path;
    }

    public async Task<string> GenerateHtmlAsync(VerificationReport report, string outputDir)
    {
        Directory.CreateDirectory(outputDir);
        string path = Path.Combine(outputDir, "report.html");
        var sb = new StringBuilder();

        sb.AppendLine("<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'>");
        sb.AppendLine("<title>Verification Report - DefTop " + report.DefTopId + "</title>");
        sb.AppendLine("<style>");
        sb.AppendLine("body{font-family:-apple-system,sans-serif;margin:20px;color:#333}");
        sb.AppendLine("h1{color:#1a1a2e;border-bottom:2px solid #e94560;padding-bottom:8px}");
        sb.AppendLine("h2{color:#16213e;margin-top:24px}");
        sb.AppendLine(".pass{color:#2e7d32;font-weight:bold}");
        sb.AppendLine(".fail{color:#c62828;font-weight:bold}");
        sb.AppendLine(".partial{color:#f57f17;font-weight:bold}");
        sb.AppendLine(".skip{color:#666;font-weight:bold}");
        sb.AppendLine("table{border-collapse:collapse;width:100%;margin:12px 0;font-size:14px}");
        sb.AppendLine("th,td{border:1px solid #ddd;padding:8px;text-align:left}");
        sb.AppendLine("th{background:#1a1a2e;color:white}");
        sb.AppendLine("tr:nth-child(even){background:#f9f9f9}");
        sb.AppendLine(".summary{display:flex;gap:16px;margin:16px 0}");
        sb.AppendLine(".summary-card{flex:1;padding:16px;border-radius:8px;text-align:center;color:white}");
        sb.AppendLine(".summary-card.pass{background:#2e7d32}");
        sb.AppendLine(".summary-card.fail{background:#c62828}");
        sb.AppendLine(".summary-card.total{background:#1565c0}");
        sb.AppendLine("</style></head><body>");

        sb.AppendLine("<h1>PaperLess Verification Report</h1>");
        sb.AppendLine("<p><strong>DefTop ID:</strong> " + report.DefTopId + "</p>");
        sb.AppendLine("<p><strong>Generated:</strong> " + report.GeneratedAt.ToString("yyyy-MM-dd HH:mm:ss") + "</p>");
        sb.AppendLine("<p><strong>Overall:</strong> <span class='" + report.Overall.ToLower() + "'>" + report.Overall + "</span></p>");

        sb.AppendLine("<div class='summary'>");
        sb.AppendLine("<div class='summary-card total'>Total Checks<br/><strong>" + report.Summary.TotalChecks + "</strong></div>");
        sb.AppendLine("<div class='summary-card pass'>Passed<br/><strong>" + report.Summary.Passed + "</strong></div>");
        sb.AppendLine("<div class='summary-card fail'>Failed<br/><strong>" + report.Summary.Failed + "</strong></div>");
        sb.AppendLine("</div>");

        if (report.XmlComparison != null)
        {
            sb.AppendLine("<h2>XML Comparison</h2>");
            sb.AppendLine("<p>Status: <span class='" + report.XmlComparison.Status.ToLower() + "'>" + report.XmlComparison.Status + "</span>");
            sb.AppendLine(" | Match: " + report.XmlComparison.MatchPercentage.ToString("F1") + "%</p>");
            sb.AppendLine("<table><tr><th>Metric</th><th>Value</th></tr>");
            sb.AppendLine("<tr><td>Cluster Nodes (Generated)</td><td>" + report.XmlComparison.NodeCountGenerated + "</td></tr>");
            sb.AppendLine("<tr><td>Cluster Nodes (Database)</td><td>" + report.XmlComparison.NodeCountDatabase + "</td></tr>");
            sb.AppendLine("<tr><td>Matching Nodes</td><td>" + report.XmlComparison.MatchingNodes + "</td></tr>");
            sb.AppendLine("<tr><td>Differing Nodes</td><td>" + report.XmlComparison.DifferingNodes + "</td></tr>");
            sb.AppendLine("<tr><td>Missing Nodes</td><td>" + report.XmlComparison.MissingNodes + "</td></tr>");
            sb.AppendLine("<tr><td>Extra Nodes</td><td>" + report.XmlComparison.ExtraNodes + "</td></tr>");
            sb.AppendLine("</table>");

            if (report.XmlComparison.Differences.Count > 0)
            {
                sb.AppendLine("<h3>XML Differences</h3><table><tr><th>Path</th><th>Status</th><th>Attribute</th><th>Database</th><th>Generated</th></tr>");
                foreach (var d in report.XmlComparison.Differences)
                {
                    sb.AppendLine("<tr><td>" + d.Path + "</td><td class='" + d.Status.ToLower() + "'>" + d.Status + "</td>");
                    sb.AppendLine("<td>" + (d.Attribute ?? "") + "</td><td>" + (d.DatabaseValue ?? "") + "</td><td>" + (d.GeneratedValue ?? "") + "</td></tr>");
                }
                sb.AppendLine("</table>");
            }
        }

        if (report.CoordinateComparison != null)
        {
            sb.AppendLine("<h2>Coordinate Comparison</h2>");
            sb.AppendLine("<p>Status: <span class='" + report.CoordinateComparison.Status.ToLower() + "'>" + report.CoordinateComparison.Status + "</span>");
            sb.AppendLine(" | Match: " + report.CoordinateComparison.MatchPercentage.ToString("F1") + "%");
            sb.AppendLine(" | Total: " + report.CoordinateComparison.TotalClusters);
            sb.AppendLine(" | Passed: " + report.CoordinateComparison.Passed);
            sb.AppendLine(" | Failed: " + report.CoordinateComparison.Failed + "</p>");

            if (report.CoordinateComparison.Differences.Count > 0)
            {
                sb.AppendLine("<h3>Coordinate Differences</h3><table>");
                sb.AppendLine("<tr><th>Cluster</th><th>Cell</th><th>Status</th><th>Expected</th><th>Actual</th><th>Diff (Ratio)</th></tr>");
                foreach (var d in report.CoordinateComparison.Differences)
                {
                    sb.AppendLine("<tr><td>" + d.ClusterId + "</td><td>" + d.CellAddress + "</td>");
                    sb.AppendLine("<td class='" + d.Status.ToLower() + "'>" + d.Status + "</td>");
                    sb.AppendLine("<td>L:" + d.ExpectedLeft.ToString("F4") + " T:" + d.ExpectedTop.ToString("F4") + " R:" + d.ExpectedRight.ToString("F4") + " B:" + d.ExpectedBottom.ToString("F4") + "</td>");
                    sb.AppendLine("<td>L:" + d.ActualLeft.ToString("F4") + " T:" + d.ActualTop.ToString("F4") + " R:" + d.ActualRight.ToString("F4") + " B:" + d.ActualBottom.ToString("F4") + "</td>");
                    sb.AppendLine("<td>L:" + d.DiffLeftRatio.ToString("F6") + " T:" + d.DiffTopRatio.ToString("F6") + " R:" + d.DiffRightRatio.ToString("F6") + " B:" + d.DiffBottomRatio.ToString("F6") + "</td>");
                    sb.AppendLine("</tr>");
                }
                sb.AppendLine("</table>");
            }
        }

        if (report.ImageComparison != null)
        {
            sb.AppendLine("<h2>Background Image Comparison</h2>");
            sb.AppendLine("<p>Status: <span class='" + report.ImageComparison.Status.ToLower() + "'>" + report.ImageComparison.Status + "</span></p>");
            sb.AppendLine("<table><tr><th>Property</th><th>Database</th><th>Generated</th></tr>");
            sb.AppendLine("<tr><td>Exists</td><td>" + report.ImageComparison.DatabaseImageExists + "</td><td>" + report.ImageComparison.GeneratedImageExists + "</td></tr>");
            sb.AppendLine("<tr><td>Width</td><td>" + (report.ImageComparison.DatabaseWidth?.ToString() ?? "N/A") + "</td><td>" + (report.ImageComparison.GeneratedWidth?.ToString() ?? "N/A") + "</td></tr>");
            sb.AppendLine("<tr><td>Height</td><td>" + (report.ImageComparison.DatabaseHeight?.ToString() ?? "N/A") + "</td><td>" + (report.ImageComparison.GeneratedHeight?.ToString() ?? "N/A") + "</td></tr>");
            sb.AppendLine("<tr><td>MSE</td><td colspan='2'>" + (report.ImageComparison.Mse?.ToString("F4") ?? "N/A") + "</td></tr>");
            sb.AppendLine("<tr><td>SSIM</td><td colspan='2'>" + (report.ImageComparison.Ssim?.ToString("F4") ?? "N/A") + "</td></tr>");
            sb.AppendLine("<tr><td>Pixel Diff %</td><td colspan='2'>" + (report.ImageComparison.PixelDiffPercent?.ToString("F2") ?? "N/A") + "</td></tr>");
            sb.AppendLine("<tr><td>Similarity</td><td colspan='2'>" + (report.ImageComparison.SimilarityScore?.ToString("P1") ?? "N/A") + "</td></tr>");
            sb.AppendLine("</table>");
        }

        sb.AppendLine("</body></html>");

        await File.WriteAllTextAsync(path, sb.ToString());
        return path;
    }

    public async Task<string> GenerateCsvAsync(VerificationReport report, string outputDir)
    {
        Directory.CreateDirectory(outputDir);
        string path = Path.Combine(outputDir, "report.csv");
        var sb = new StringBuilder();

        sb.AppendLine("Section,Property,Status,Expected,Actual");

        sb.AppendLine($"Summary,Overall,{report.Overall},");
        sb.AppendLine($"Summary,TotalChecks,{report.Summary.TotalChecks},");
        sb.AppendLine($"Summary,Passed,{report.Summary.Passed},");
        sb.AppendLine($"Summary,Failed,{report.Summary.Failed},");
        sb.AppendLine($"Summary,MatchPercentage,{report.Summary.OverallMatchPercentage:F1}%,");

        if (report.XmlComparison != null)
        {
            sb.AppendLine($"Xml,Status,{report.XmlComparison.Status},");
            sb.AppendLine($"Xml,MatchPercentage,{report.XmlComparison.MatchPercentage:F1}%,");
            sb.AppendLine($"Xml,MatchingNodes,{report.XmlComparison.MatchingNodes},");
            sb.AppendLine($"Xml,DifferingNodes,{report.XmlComparison.DifferingNodes},");
            sb.AppendLine($"Xml,MissingNodes,{report.XmlComparison.MissingNodes},");
            sb.AppendLine($"Xml,ExtraNodes,{report.XmlComparison.ExtraNodes},");
            foreach (var d in report.XmlComparison.Differences)
            {
                sb.AppendLine($"Xml,{EscapeCsv(d.Path)},{d.Status},{EscapeCsv(d.DatabaseValue)},{EscapeCsv(d.GeneratedValue)}");
            }
        }

        if (report.CoordinateComparison != null)
        {
            sb.AppendLine($"Coordinate,Status,{report.CoordinateComparison.Status},");
            sb.AppendLine($"Coordinate,MatchPercentage,{report.CoordinateComparison.MatchPercentage:F1}%,");
            sb.AppendLine($"Coordinate,TotalClusters,{report.CoordinateComparison.TotalClusters},");
            sb.AppendLine($"Coordinate,Passed,{report.CoordinateComparison.Passed},");
            sb.AppendLine($"Coordinate,Failed,{report.CoordinateComparison.Failed},");
            foreach (var d in report.CoordinateComparison.Differences)
            {
                sb.AppendLine($"Coordinate,{d.ClusterId},{d.Status},L:{d.ExpectedLeft:F4} T:{d.ExpectedTop:F4},L:{d.ActualLeft:F4} T:{d.ActualTop:F4}");
            }
        }

        if (report.ImageComparison != null)
        {
            sb.AppendLine($"Image,Status,{report.ImageComparison.Status},");
            sb.AppendLine($"Image,Similarity,{report.ImageComparison.SimilarityScore?.ToString("P1") ?? "N/A"},");
            sb.AppendLine($"Image,MSE,{report.ImageComparison.Mse?.ToString("F4") ?? "N/A"},");
            sb.AppendLine($"Image,SSIM,{report.ImageComparison.Ssim?.ToString("F4") ?? "N/A"},");
        }

        await File.WriteAllTextAsync(path, sb.ToString());
        return path;
    }

    private static string EscapeCsv(string? value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        return value;
    }
}
