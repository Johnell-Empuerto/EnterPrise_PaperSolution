using System.Text.Json;

namespace ExcelAPI.Rendering.Validation
{
    /// <summary>
    /// Static helper for generating and saving validation reports.
    /// Provides convenience methods for creating JSON and HTML reports
    /// from a ValidationSummary.
    /// </summary>
    public static class ValidationReport
    {
        /// <summary>
        /// Save the validation summary as a JSON report file.
        /// </summary>
        /// <param name="summary">The validation summary to serialize.</param>
        /// <param name="outputPath">Path to the output JSON file.</param>
        public static void SaveJson(ValidationSummary summary, string outputPath)
        {
            var dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            var json = JsonSerializer.Serialize(summary, new JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
            });

            System.IO.File.WriteAllText(outputPath, json);
        }

        /// <summary>
        /// Save the validation summary as an HTML report file with embedded previews.
        /// </summary>
        /// <param name="summary">The validation summary to render.</param>
        /// <param name="outputPath">Path to the output HTML file.</param>
        public static void SaveHtml(ValidationSummary summary, string outputPath)
        {
            var dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            var html = GenerateHtml(summary);
            System.IO.File.WriteAllText(outputPath, html);
        }

        /// <summary>
        /// Generate an HTML report string from a validation summary.
        /// </summary>
        public static string GenerateHtml(ValidationSummary summary)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang=\"en\">");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta charset=\"UTF-8\">");
            sb.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            sb.AppendLine("<title>FormLess Validation Report</title>");
            sb.AppendLine("<style>");
            sb.AppendLine("  * { box-sizing: border-box; margin: 0; padding: 0; }");
            sb.AppendLine("  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 0; padding: 20px; background: #f0f2f5; color: #333; }");
            sb.AppendLine("  h1 { font-size: 24px; margin-bottom: 8px; color: #1a1a2e; }");
            sb.AppendLine("  h2 { font-size: 18px; margin: 20px 0 12px; color: #333; }");
            sb.AppendLine("  .subtitle { color: #666; font-size: 14px; margin-bottom: 24px; }");
            sb.AppendLine("  .pass { color: #16a34a; font-weight: 600; }");
            sb.AppendLine("  .fail { color: #dc2626; font-weight: 600; }");
            sb.AppendLine("  .summary-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; margin-bottom: 24px; }");
            sb.AppendLine("  .card { background: white; padding: 16px; border-radius: 10px; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }");
            sb.AppendLine("  .card .value { font-size: 28px; font-weight: 700; }");
            sb.AppendLine("  .card .label { font-size: 12px; color: #888; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 4px; }");
            sb.AppendLine("  .card.pass-bg { border-left: 4px solid #16a34a; }");
            sb.AppendLine("  .card.fail-bg { border-left: 4px solid #dc2626; }");
            sb.AppendLine("  .workbook { background: white; border-radius: 10px; padding: 16px 20px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }");
            sb.AppendLine("  .workbook-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }");
            sb.AppendLine("  .workbook-header h3 { font-size: 16px; }");
            sb.AppendLine("  .workbook-meta { font-size: 13px; color: #888; margin-bottom: 12px; }");
            sb.AppendLine("  table { width: 100%; border-collapse: collapse; font-size: 13px; }");
            sb.AppendLine("  th { background: #f8f9fa; padding: 8px 10px; text-align: left; border-bottom: 2px solid #dee2e6; font-weight: 600; }");
            sb.AppendLine("  td { padding: 8px 10px; border-bottom: 1px solid #eee; }");
            sb.AppendLine("  .previews { display: flex; gap: 12px; flex-wrap: wrap; margin: 12px 0 4px; }");
            sb.AppendLine("  .previews .frame { text-align: center; }");
            sb.AppendLine("  .previews img { max-width: 280px; max-height: 200px; border: 1px solid #ddd; border-radius: 6px; }");
            sb.AppendLine("  .previews .label { font-size: 11px; color: #888; margin-top: 4px; }");
            sb.AppendLine("  @media (max-width: 768px) { .previews img { max-width: 100%; } }");
            sb.AppendLine("</style>");
            sb.AppendLine("</head><body>");

            sb.AppendLine($"<h1>FormLess Rendering Engine — Validation Report</h1>");
            sb.AppendLine($"<div class=\"subtitle\">Run: {summary.Timestamp} | Engine v1.3 | {summary.TotalWorkbooks} workbook(s), {summary.TotalPages} page(s)</div>");

            // Summary cards
            sb.AppendLine("<div class=\"summary-cards\">");
            sb.AppendLine($"<div class=\"card pass-bg\"><div class=\"value pass\">{summary.Passed}</div><div class=\"label\">Passed</div></div>");
            sb.AppendLine($"<div class=\"card fail-bg\"><div class=\"value fail\">{summary.Failed}</div><div class=\"label\">Failed</div></div>");
            sb.AppendLine($"<div class=\"card\"><div class=\"value\">{summary.AveragePixelDifference:F2}%</div><div class=\"label\">Avg Pixel Diff</div></div>");
            sb.AppendLine($"<div class=\"card\"><div class=\"value\">{summary.AverageRenderingTimeMs:F0}ms</div><div class=\"label\">Avg Render Time</div></div>");
            sb.AppendLine($"<div class=\"card\"><div class=\"value\">{summary.TotalRenderingTimeMs}ms</div><div class=\"label\">Total Time</div></div>");
            sb.AppendLine("</div>");

            // Per-workbook results
            foreach (var report in summary.Reports
                .OrderBy(r => r.Passed ? 0 : 1)
                .ThenBy(r => r.WorkbookName))
            {
                string statusClass = report.Passed ? "pass" : "fail";
                string statusText = report.Passed ? "✓ PASS" : "✗ FAIL";

                sb.AppendLine("<div class=\"workbook\">");
                sb.AppendLine("<div class=\"workbook-header\">");
                sb.AppendLine($"<h3>{report.WorkbookName} <span class=\"{statusClass}\">({statusText})</span></h3>");
                sb.AppendLine("</div>");
                sb.AppendLine($"<div class=\"workbook-meta\">");
                sb.AppendLine($"Pixel Diff: {report.PixelDifference:F4}% | Time: {report.RenderingTimeMs}ms | Pages: {report.Pages}");
                sb.AppendLine("</div>");

                if (report.PageResults.Count > 0)
                {
                    sb.AppendLine("<table>");
                    sb.AppendLine("<tr><th>Page</th><th>Diff %</th><th>Diff Px</th><th>Max Err</th><th>Avg Err</th><th>Dim</th><th>Status</th></tr>");
                    foreach (var page in report.PageResults)
                    {
                        string pgStatus = page.Passed
                            ? "<span class=\"pass\">PASS</span>"
                            : "<span class=\"fail\">FAIL</span>";

                        sb.AppendLine($"<tr>" +
                            $"<td>{page.PageNumber}</td>" +
                            $"<td>{page.PixelDifference:F4}%</td>" +
                            $"<td>{page.DifferentPixels}/{page.TotalPixels}</td>" +
                            $"<td>{page.MaxError}</td>" +
                            $"<td>{page.AverageError}</td>" +
                            $"<td>{page.Dimensions}</td>" +
                            $"<td>{pgStatus}</td></tr>");

                        // Preview row
                        sb.AppendLine("<tr><td colspan=\"7\"><div class=\"previews\">");
                        if (!string.IsNullOrEmpty(page.ExpectedImagePath) && System.IO.File.Exists(page.ExpectedImagePath))
                            sb.AppendLine($"<div class=\"frame\"><img src=\"{MakeRelativePath(page.ExpectedImagePath)}\" /><div class=\"label\">Expected</div></div>");
                        if (!string.IsNullOrEmpty(page.ActualImagePath) && System.IO.File.Exists(page.ActualImagePath))
                            sb.AppendLine($"<div class=\"frame\"><img src=\"{MakeRelativePath(page.ActualImagePath)}\" /><div class=\"label\">Actual</div></div>");
                        if (!string.IsNullOrEmpty(page.DiffImagePath) && System.IO.File.Exists(page.DiffImagePath))
                            sb.AppendLine($"<div class=\"frame\"><img src=\"{MakeRelativePath(page.DiffImagePath)}\" /><div class=\"label\">Diff</div></div>");
                        sb.AppendLine("</div></td></tr>");
                    }
                    sb.AppendLine("</table>");
                }

                sb.AppendLine("</div>");
            }

            sb.AppendLine("</body></html>");
            return sb.ToString();
        }

        /// <summary>
        /// Convert an absolute path to a relative path for HTML img src.
        /// </summary>
        private static string MakeRelativePath(string absolutePath)
        {
            // For HTML reports saved in RegressionTests/Reports/, we need paths relative to that directory.
            // This is a best-effort conversion.
            try
            {
                string reportsDir = Path.GetFullPath("RegressionTests/Reports");
                string fullPath = Path.GetFullPath(absolutePath);
                if (fullPath.StartsWith(reportsDir, StringComparison.OrdinalIgnoreCase))
                    return fullPath[reportsDir.Length..].TrimStart('\\', '/');

                // Fall back to absolute path
                return fullPath;
            }
            catch
            {
                return absolutePath;
            }
        }
    }
}
