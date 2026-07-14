using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;
using Npgsql;
using LegacyExtractionEngine.Models;
using LegacyExtractionEngine.Services.Importer;

namespace LegacyExtractionEngine.Services;

/// <summary>
/// Phase 7 — Discover the Missing Coordinate Transform.
/// 
/// For every mismatched cluster across all templates, measure the per-cluster
/// error (DB vs COM-derived) and determine whether the error correlates with:
///   - Row number (vertical gap accumulation)
///   - Column number (font metric / column-width mismatch)
///   - Distance from origin (cumulative gap)
///   - Merged cells vs standalone cells
///   - Print scaling / FitToPages / Zoom
///   - Hidden rows / columns
/// 
/// Then build an error model: constant, linear, piecewise, proportional, or scaled.
/// </summary>
public class Phase7ErrorInvestigator
{
    private readonly string _connectionString;
    private readonly IProgress<string> _progress;
    private readonly string _outputDir;
    private readonly string _tempDir;

    public Phase7ErrorInvestigator(string connectionString, IProgress<string> progress)
    {
        _connectionString = connectionString;
        _progress = progress;
        _outputDir = Path.Combine(
            @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise",
            "Test Folder Final Test", "Phase7");
        _tempDir = Path.Combine(Path.GetTempPath(), "phase7_workbooks");
        Directory.CreateDirectory(_outputDir);
        Directory.CreateDirectory(_tempDir);
    }

    public async Task RunInvestigationAsync()
    {
        _progress.Report("============================================");
        _progress.Report(" Phase 7 — Missing Coordinate Transform");
        _progress.Report(" Per-Cluster Error Analysis & Regression");
        _progress.Report("============================================");
        _progress.Report("");

        // Step 1: Gather all templates with workbooks
        var templates = await GatherTemplatesAsync();

        // Step 2: Run per-cluster error analysis for each template
        var allResults = new List<TemplateErrorReport>();
        foreach (var (id, wbPath) in templates)
            if (File.Exists(wbPath))
                await AnalyzeTemplateAsync(id, wbPath, allResults);

        // Step 3: Generate comprehensive report
        await GenerateReportAsync(allResults);
    }

    private async Task<List<(int id, string path)>> GatherTemplatesAsync()
    {
        var result = new List<(int id, string path)>();

        // Template 546
        var path546 = @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx";
        if (File.Exists(path546)) result.Add((546, path546));

        // Template 547
        var discovery = new TemplateDiscoveryService(_connectionString, _progress);
        var wb547 = discovery.FindWorkbookForTemplate(547, null);
        if (wb547 != null) result.Add((547, wb547));

        // Extracted workbooks from Phase 6 temp dir
        var phase6Dir = Path.Combine(Path.GetTempPath(), "phase6_workbooks");
        if (Directory.Exists(phase6Dir))
            foreach (var f in Directory.GetFiles(phase6Dir, "*.xlsx"))
                if (int.TryParse(Path.GetFileNameWithoutExtension(f).Split('_')[0], out var id))
                    if (!result.Any(r => r.id == id))
                        result.Add((id, f));

        // Also try extracting more workbooks from DB if we don't have many
        if (result.Count < 10)
        {
            var more = await ExtractMoreWorkbooksAsync(result.Select(r => r.id).ToHashSet());
            result.AddRange(more);
        }

        _progress.Report($"Found {result.Count} templates with workbooks for Phase 7 analysis");
        return result.OrderBy(t => t.id).ToList();
    }

    private async Task<List<(int id, string path)>> ExtractMoreWorkbooksAsync(HashSet<int> exclude)
    {
        var result = new List<(int id, string path)>();
        await using var conn = new NpgsqlConnection(_connectionString);
        await conn.OpenAsync();

        var sql = @"
            SELECT def_top_id, def_file, def_top_name
            FROM def_top
            WHERE def_file IS NOT NULL
              AND def_top_id NOT IN (SELECT unnest(@exclude::int[]))
            ORDER BY def_top_id
            LIMIT 10";
        await using var cmd = new NpgsqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("exclude", exclude.ToArray());

        await using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var id = reader.GetInt32(0);
            var bytes = reader["def_file"] as byte[];
            var name = reader.IsDBNull(2) ? null : reader.GetString(2);
            if (bytes == null || bytes.Length == 0) continue;

            var safeName = (name ?? $"template_{id}")
                .Replace("<", "").Replace(">", "").Replace(":", "")
                .Replace("\"", "").Replace("/", "").Replace("\\", "")
                .Replace("|", "").Replace("?", "").Replace("*", "");
            var path = Path.Combine(_tempDir, $"{id}_{safeName}.xlsx");
            await File.WriteAllBytesAsync(path, bytes);
            result.Add((id, path));
        }
        return result;
    }

    private async Task AnalyzeTemplateAsync(int defTopId, string workbookPath,
        List<TemplateErrorReport> results)
    {
        _progress.Report($"--- Analyzing template {defTopId} ---");

        if (!File.Exists(workbookPath))
        {
            _progress.Report($"  SKIP: workbook not found at {workbookPath}");
            return;
        }

        var report = new TemplateErrorReport { DefTopId = defTopId };

        try
        {
            // Load DB data
            var dbReader = new DatabaseReader(_connectionString, _progress);
            var db = await dbReader.ReadAllAsync(defTopId);
            if (db?.DefSheets.Count == 0)
            {
                _progress.Report($"  SKIP: no DB data for template {defTopId}");
                return;
            }

            var dbClusters = db!.DefSheets.SelectMany(s => s.Clusters).ToList();
            report.ClusterCount = dbClusters.Count;
            report.DefTopName = db.DefTop?.DefTopName ?? $"Template {defTopId}";

            // Load workbook characteristics
            report.WorkbookChars = await ExtractWorkbookCharacteristicsAsync(workbookPath);

            // Extract COM data
            using var com = new ComCoordinateService(_progress);
            var comExt = com.Extract(workbookPath, dbClusters);

            // Record page setup
            report.PageWidth = comExt.PageWidth;
            report.PageHeight = comExt.PageHeight;
            report.PrintedOriginX = comExt.PrintedOriginX;
            report.PrintedOriginY = comExt.PrintedOriginY;
            report.ClusterGapX = comExt.ClusterGapX;
            report.ClusterGapY = comExt.ClusterGapY;
            report.CenterHorizontally = comExt.CenterHorizontally;
            report.CenterVertically = comExt.CenterVertically;
            report.PrintArea = comExt.PrintArea;
            report.Zoom = comExt.Zoom ?? 100;
            report.FitToPagesWide = comExt.FitToPagesWide ?? 0;
            report.FitToPagesTall = comExt.FitToPagesTall ?? 0;
            report.PaperSize = comExt.PaperSize ?? 0;
            report.Orientation = comExt.Orientation ?? 0;

            // Per-cluster error analysis
            foreach (var dbc in dbClusters.OrderBy(c => c.ClusterId))
            {
                var addr = dbc.CellAddr?.Trim() ?? "";
                if (string.IsNullOrEmpty(addr)) continue;

                double dbL = ParseDoubleSafe(dbc.LeftPosition);
                double dbT = ParseDoubleSafe(dbc.TopPosition);
                double dbR = ParseDoubleSafe(dbc.RightPosition);
                double dbB = ParseDoubleSafe(dbc.BottomPosition);

                if (comExt.ComPositions.TryGetValue(addr, out var comPos))
                {
                    // COM-derived ratios (our engine's generated values)
                    double genL = ComCoordinateService.ComputeRatio(comPos.Left,
                        comExt.PrintedOriginX, comExt.PageWidth);
                    double genT = ComCoordinateService.ComputeRatio(comPos.Top,
                        comExt.PrintedOriginY, comExt.PageHeight);
                    double genR = ComCoordinateService.ComputeRatio(
                        comPos.Left + comPos.Width + comExt.ClusterGapX,
                        comExt.PrintedOriginX, comExt.PageWidth);
                    double genB = ComCoordinateService.ComputeRatio(
                        comPos.Top + comPos.Height + comExt.ClusterGapY,
                        comExt.PrintedOriginY, comExt.PageHeight);

                    // Delas (DB - Generated)
                    double deltaL = dbL - genL;
                    double deltaT = dbT - genT;
                    double deltaR = dbR - genR;
                    double deltaB = dbB - genB;

                    // Delta in points
                    double deltaXPt = deltaL * comExt.PageWidth;
                    double deltaYPt = deltaT * comExt.PageHeight;

                    // Resolve row and column from address
                    int col = ParseCol(addr);
                    int row = ParseRow(addr);

                    report.ClusterErrors.Add(new ClusterError
                    {
                        CellAddress = addr,
                        Row = row,
                        Column = col,
                        DbLeft = dbL, DbTop = dbT, DbRight = dbR, DbBottom = dbB,
                        ComLeft = comPos.Left, ComTop = comPos.Top,
                        ComWidth = comPos.Width, ComHeight = comPos.Height,
                        GeneratedLeft = genL, GeneratedTop = genT,
                        GeneratedRight = genR, GeneratedBottom = genB,
                        DeltaLeft = deltaL, DeltaTop = deltaT,
                        DeltaRight = deltaR, DeltaBottom = deltaB,
                        DeltaXPt = deltaXPt, DeltaYPt = deltaYPt,
                        IsMerged = comPos.MergeAddress != null,
                        DistanceFromOrigin = Math.Sqrt(
                            Math.Pow(comPos.Left + comExt.PrintedOriginX, 2) +
                            Math.Pow(comPos.Top + comExt.PrintedOriginY, 2))
                    });
                }
                else
                {
                    // No COM position — mark as unmatched
                    report.UnmatchedClusters.Add(dbc);
                }
            }

            // Compute per-template statistics
            report.ComputeStatistics();

            _progress.Report($"  Clusters: {report.ClusterErrors.Count} matched, " +
                $"{report.UnmatchedClusters.Count} unmatched");
            _progress.Report($"  Mean Error X: {report.MeanErrorXPt:F4} pt, " +
                $"Mean Error Y: {report.MeanErrorYPt:F4} pt");
            _progress.Report($"  Std Dev X: {report.StdDevXPt:F4}, " +
                $"Std Dev Y: {report.StdDevYPt:F4}");
            _progress.Report($"  Row regression: slope={report.RowRegSlope:F7}, " +
                $"intercept={report.RowRegIntercept:F7}");
            _progress.Report($"  Col regression: slope={report.ColRegSlope:F7}, " +
                $"intercept={report.ColRegIntercept:F7}");
        }
        catch (Exception ex)
        {
            _progress.Report($"  ERROR: {ex.Message}");
            report.ErrorMessage = ex.Message;
        }

        results.Add(report);
    }

    /// <summary>
    /// Extract workbook characteristics from OpenXML.
    /// </summary>
    private async Task<WorkbookCharacteristics> ExtractWorkbookCharacteristicsAsync(string workbookPath)
    {
        var chars = new WorkbookCharacteristics();

        try
        {
            var importer = new ExcelImporter(_progress);
            var extraction = importer.Import(workbookPath);
            var sheet = extraction.Workbook.Sheets.FirstOrDefault();
            if (sheet == null) return chars;

            chars.ColumnCount = sheet.Columns.Count;
            chars.CellCount = sheet.Cells.Count;
            chars.MergeCount = sheet.MergedCells.Count;
            chars.HiddenColumns = sheet.Columns.Count(c => c.Hidden);
            chars.HiddenRows = sheet.Rows.Count(r => r.Hidden);

            // Default font from workbook styles
            var defaultFont = extraction.Styles
                .FirstOrDefault(s => s.Font != null)?.Font;
            if (defaultFont != null)
            {
                chars.DefaultFontName = defaultFont.FontName ?? "Calibri";
                chars.DefaultFontSize = defaultFont.Size ?? 11;
            }

            // Compute content width from columns
            double contentWidth = 0;
            int maxCol = Math.Max(sheet.Cells.Count > 0 ? sheet.Cells.Max(c => c.Column) : 0,
                sheet.MergedCells.Count > 0 ? sheet.MergedCells.Max(m => m.EndColumn) : 0);
            for (int i = 1; i <= maxCol; i++)
            {
                var col = sheet.Columns.FirstOrDefault(c => i >= c.Min && i <= c.Max);
                double charsW = col?.Width ?? 8.43;
                contentWidth += charsW * 50.04 / 8.43;
            }
            chars.ContentWidth = contentWidth;

            // Print area from defined names
            var printAreaDef = extraction.Workbook.DefinedNames
                .FirstOrDefault(d => d.Name.Contains("Print_Area"));
            if (printAreaDef != null)
                chars.HasPrintArea = true;

            // Sheet format properties
            if (sheet.SheetFormatProperties != null)
                chars.DefaultRowHeight = sheet.SheetFormatProperties.DefaultRowHeight;

            // Print page setup
            chars.PageSetupScale = sheet.PageSetup.Scale;
            chars.FitToWidth = sheet.PageSetup.FitToWidth;
            chars.FitToHeight = sheet.PageSetup.FitToHeight;
            chars.PageWidth = sheet.PageSetup.PageWidth;
            chars.PageHeight = sheet.PageSetup.PageHeight;
            chars.Orientation = sheet.PageSetup.Orientation;
        }
        catch (Exception ex)
        {
            _progress.Report($"  Warn: could not read workbook characteristics: {ex.Message}");
        }

        return chars;
    }

    /// <summary>
    /// Generate comprehensive Phase 7 report.
    /// </summary>
    private async Task GenerateReportAsync(List<TemplateErrorReport> results)
    {
        _progress.Report("Generating Phase 7 report...");

        var sb = new StringBuilder();
        sb.AppendLine("# Phase 7 — Missing Coordinate Transform Investigation");
        sb.AppendLine();
        sb.AppendLine($"**Generated:** {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC");
        sb.AppendLine($"**Templates Analyzed:** {results.Count}");
        sb.AppendLine();

        // =========================================================
        // Section 1: Executive Summary
        // =========================================================
        sb.AppendLine("## Executive Summary");
        sb.AppendLine();

        var pureCom = results.Where(r => Math.Abs(r.MeanErrorXPt) < 0.5 && Math.Abs(r.MeanErrorYPt) < 0.5 && r.ClusterErrors.Count > 0).ToList();
        var horizOnly = results.Where(r => Math.Abs(r.MeanErrorXPt) >= 0.5 && Math.Abs(r.MeanErrorYPt) < 1.0 && r.ClusterErrors.Count > 0).ToList();
        var vertOnly = results.Where(r => Math.Abs(r.MeanErrorXPt) < 0.5 && Math.Abs(r.MeanErrorYPt) >= 1.0 && r.ClusterErrors.Count > 0).ToList();
        var bothAxes = results.Where(r => Math.Abs(r.MeanErrorXPt) >= 0.5 && Math.Abs(r.MeanErrorYPt) >= 1.0 && r.ClusterErrors.Count > 0).ToList();

        sb.AppendLine("### Classification by Error Pattern");
        sb.AppendLine();
        sb.AppendLine("| Pattern | Templates | Description |");
        sb.AppendLine("|---------|-----------|-------------|");
        if (pureCom.Count > 0)
            sb.AppendLine($"| Pure COM match | {string.Join(", ", pureCom.Select(r => r.DefTopId))} | Error < 0.5pt on both axes — formula works perfectly |");
        if (horizOnly.Count > 0)
            sb.AppendLine($"| Horizontal error only | {string.Join(", ", horizOnly.Select(r => r.DefTopId))} | Column-width/font metric mismatch |");
        if (vertOnly.Count > 0)
            sb.AppendLine($"| Vertical error only | {string.Join(", ", vertOnly.Select(r => r.DefTopId))} | Row-gap accumulation or missing vertical transform |");
        if (bothAxes.Count > 0)
            sb.AppendLine($"| Both axes | {string.Join(", ", bothAxes.Select(r => r.DefTopId))} | Multiple missing transforms |");
        sb.AppendLine();

        // =========================================================
        // Section 2: Per-Template Summary
        // =========================================================
        sb.AppendLine("## Per-Template Summary");
        sb.AppendLine();
        sb.AppendLine("| ID | Name | Clusters | Matched | Unmatched | " +
            "Mean ΔX (pt) | Mean ΔY (pt) | StdDev ΔX | StdDev ΔY | " +
            "Max |ΔX| | Max |ΔY| | Row Slope | Col Slope | Pattern |");
        sb.AppendLine("|----|------|----------|---------|-----------|" +
            "-------------|-------------|-----------|-----------|" +
            "----------|----------|-----------|-----------|---------|");

        foreach (var r in results.OrderBy(r => r.DefTopId))
        {
            var pattern = ClassifyPattern(r);
            sb.AppendLine($"| {r.DefTopId} | {r.DefTopName} | {r.ClusterCount} | " +
                $"{r.ClusterErrors.Count} | {r.UnmatchedClusters.Count} | " +
                $"{r.MeanErrorXPt:F4} | {r.MeanErrorYPt:F4} | " +
                $"{r.StdDevXPt:F4} | {r.StdDevYPt:F4} | " +
                $"{r.MaxAbsDeltaXPt:F4} | {r.MaxAbsDeltaYPt:F4} | " +
                $"{r.RowRegSlope:F7} | {r.ColRegSlope:F7} | {pattern} |");
        }
        sb.AppendLine();

        // =========================================================
        // Section 3: Workload Characteristics Comparison
        // =========================================================
        sb.AppendLine("## Workbook Characteristics vs Error");
        sb.AppendLine();

        sb.AppendLine("| ID | Cols | Cells | Merges | HiddenCol | HiddenRow | " +
            "Font | FontSz | RowHt | ContentW | HasPA | Scale | FitW×H | " +
            "Paper | Orient | Mean ΔX | Mean ΔY |");
        sb.AppendLine("|----|------|-------|--------|-----------|-----------|" +
            "-----|--------|-------|----------|-------|--------|--------|" +
            "-------|--------|---------|---------|");

        foreach (var r in results.OrderBy(r => r.DefTopId))
        {
            var c = r.WorkbookChars;
            sb.AppendLine($"| {r.DefTopId} | {c.ColumnCount} | {c.CellCount} | {c.MergeCount} | " +
                $"{c.HiddenColumns} | {c.HiddenRows} | {c.DefaultFontName ?? "?"} | " +
                $"{(c.DefaultFontSize ?? 11):F0} | {(c.DefaultRowHeight ?? 14.4):F1} | " +
                $"{c.ContentWidth:F0} | " +
                $"{(c.HasPrintArea ? "✓" : "✗")} | {c.PageSetupScale ?? 100} | " +
                $"{c.FitToWidth ?? 0}×{c.FitToHeight ?? 0} | " +
                $"{r.PaperSize} | " +
                $"{c.Orientation ?? "portrait"} | " +
                $"{r.MeanErrorXPt:F4} | {r.MeanErrorYPt:F4} |");
        }
        sb.AppendLine();

        // =========================================================
        // Section 4: Per-Template Detailed Error Data
        // =========================================================
        foreach (var r in results.OrderBy(r => r.DefTopId))
        {
            sb.AppendLine($"## Template {r.DefTopId} — {r.DefTopName}");
            sb.AppendLine();

            if (r.ClusterErrors.Count == 0)
            {
                sb.AppendLine("No matched clusters to analyze.");
                sb.AppendLine();
                continue;
            }

            sb.AppendLine("### Page Setup");
            sb.AppendLine();
            sb.AppendLine($"| Property | Value |");
            sb.AppendLine($"|----------|-------|");
            sb.AppendLine($"| Page Width | {r.PageWidth:F1} pt |");
            sb.AppendLine($"| Page Height | {r.PageHeight:F1} pt |");
            sb.AppendLine($"| PrintedOriginX | {r.PrintedOriginX:F4} pt |");
            sb.AppendLine($"| PrintedOriginY | {r.PrintedOriginY:F4} pt |");
            sb.AppendLine($"| ClusterGapX | {r.ClusterGapX:F4} pt |");
            sb.AppendLine($"| ClusterGapY | {r.ClusterGapY:F4} pt |");
            sb.AppendLine($"| CenterHorizontally | {r.CenterHorizontally} |");
            sb.AppendLine($"| CenterVertically | {r.CenterVertically} |");
            sb.AppendLine($"| PrintArea | {r.PrintArea ?? "(none)"} |");
            sb.AppendLine($"| Zoom | {r.Zoom}% |");
            sb.AppendLine($"| FitToPagesWide | {r.FitToPagesWide} |");
            sb.AppendLine($"| FitToPagesTall | {r.FitToPagesTall} |");
            sb.AppendLine($"| Orientation | {(r.Orientation == 2 ? "Landscape" : "Portrait")} |");
            sb.AppendLine();

            sb.AppendLine("### Statistics");
            sb.AppendLine();
            sb.AppendLine($"| Metric | X (Left) | Y (Top) |");
            sb.AppendLine($"|--------|----------|---------|");
            sb.AppendLine($"| Mean Error (pt) | {r.MeanErrorXPt:F7} | {r.MeanErrorYPt:F7} |");
            sb.AppendLine($"| Mean Error (ratio) | {r.MeanErrorXRatio:F9} | {r.MeanErrorYRatio:F9} |");
            sb.AppendLine($"| Std Deviation (pt) | {r.StdDevXPt:F7} | {r.StdDevYPt:F7} |");
            sb.AppendLine($"| Max |Δ| (pt) | {r.MaxAbsDeltaXPt:F7} | {r.MaxAbsDeltaYPt:F7} |");
            sb.AppendLine($"| Min Δ (pt) | {r.MinDeltaXPt:F7} | {r.MinDeltaYPt:F7} |");
            sb.AppendLine($"| Max Δ (pt) | {r.MaxDeltaXPt:F7} | {r.MaxDeltaYPt:F7} |");
            sb.AppendLine();

            // Regression analysis
            sb.AppendLine("### Regression Analysis");
            sb.AppendLine();
            sb.AppendLine("#### X Error vs Row Number");
            sb.AppendLine();
            sb.AppendLine($"- **Slope:** {r.RowRegSlope:F9} (ΔX pt per row)");
            sb.AppendLine($"- **Intercept:** {r.RowRegIntercept:F9} pt");
            sb.AppendLine($"- **R²:** {r.RowRegRSquared:F6}");
            sb.AppendLine($"- **Interpretation:** {(Math.Abs(r.RowRegRSquared) > 0.7 ? "Strong correlation — X error grows with row count" : Math.Abs(r.RowRegRSquared) > 0.3 ? "Moderate correlation" : "Weak or no correlation")}");
            sb.AppendLine();

            sb.AppendLine("#### X Error vs Column Number");
            sb.AppendLine();
            sb.AppendLine($"- **Slope:** {r.ColRegSlope:F9} (ΔX pt per column)");
            sb.AppendLine($"- **Intercept:** {r.ColRegIntercept:F9} pt");
            sb.AppendLine($"- **R²:** {r.ColRegRSquared:F6}");
            sb.AppendLine($"- **Interpretation:** {(Math.Abs(r.ColRegRSquared) > 0.7 ? "Strong correlation — X error grows with column count" : Math.Abs(r.ColRegRSquared) > 0.3 ? "Moderate correlation" : "Weak or no correlation")}");
            sb.AppendLine();

            sb.AppendLine("#### Y Error vs Row Number");
            sb.AppendLine();
            sb.AppendLine($"- **Slope:** {r.YRowRegSlope:F9} (ΔY pt per row)");
            sb.AppendLine($"- **Intercept:** {r.YRowRegIntercept:F9} pt");
            sb.AppendLine($"- **R²:** {r.YRowRegRSquared:F6}");
            sb.AppendLine($"- **Interpretation:** {(Math.Abs(r.YRowRegRSquared) > 0.7 ? "Strong correlation — Y error grows with row count" : Math.Abs(r.YRowRegRSquared) > 0.3 ? "Moderate correlation" : "Weak or no correlation")}");
            sb.AppendLine();

            sb.AppendLine("#### Y Error vs Distance from Origin (pt)");
            sb.AppendLine();
            sb.AppendLine($"- **Slope:** {r.DistRegSlope:F9} (ΔY pt per pt of distance)");
            sb.AppendLine($"- **Intercept:** {r.DistRegIntercept:F9} pt");
            sb.AppendLine($"- **R²:** {r.DistRegRSquared:F6}");
            sb.AppendLine($"- **Interpretation:** {(Math.Abs(r.DistRegRSquared) > 0.7 ? "Strong correlation — error grows with distance from origin" : Math.Abs(r.DistRegRSquared) > 0.3 ? "Moderate correlation" : "Weak or no correlation")}");
            sb.AppendLine();

            // Per-cluster error table (top 50)
            sb.AppendLine("### Per-Cluster Error Data");
            sb.AppendLine();
            sb.AppendLine("| # | Cell | Row | Col | DB Left | DB Top | COM Left | COM Top | " +
                "Gen Left | Gen Top | ΔX (pt) | ΔY (pt) | ΔL | ΔT | Dist | Merged? |");
            sb.AppendLine("|---|------|-----|-----|---------|--------|----------|---------|" +
                "---------|---------|---------|---------|----|----|------|---------|");

            int displayCount = Math.Min(r.ClusterErrors.Count, 100);
            for (int i = 0; i < displayCount; i++)
            {
                var e = r.ClusterErrors[i];
                var dL = Math.Abs(e.DeltaLeft) < 0.000001 ? "✓" : $"Δ={e.DeltaLeft:F7}";
                var dT = Math.Abs(e.DeltaTop) < 0.000001 ? "✓" : $"Δ={e.DeltaTop:F7}";
                sb.AppendLine($"| {i + 1} | {e.CellAddress} | {e.Row} | {e.Column} | " +
                    $"{e.DbLeft:F7} | {e.DbTop:F7} | {e.ComLeft:F2} | {e.ComTop:F2} | " +
                    $"{e.GeneratedLeft:F7} | {e.GeneratedTop:F7} | " +
                    $"{e.DeltaXPt:+0.0000;-0.0000} | {e.DeltaYPt:+0.0000;-0.0000} | " +
                    $"{dL} | {dT} | {e.DistanceFromOrigin:F0} | " +
                    $"{(e.IsMerged ? "✓" : "✗")} |");
            }
            if (r.ClusterErrors.Count > displayCount)
                sb.AppendLine($"| ... | *{r.ClusterErrors.Count - displayCount} more clusters* |");
            sb.AppendLine();

            // Chart: error vs row number (simple text chart)
            if (r.ClusterErrors.Count >= 3)
            {
                sb.AppendLine("### Error vs Row (Visual)");
                sb.AppendLine();
                sb.AppendLine("```");
                sb.AppendLine($"Row   ΔX(pt)   ΔY(pt)");
                sb.AppendLine($"----  -------  -------");
                foreach (var e in r.ClusterErrors.OrderBy(e => e.Row).ThenBy(e => e.Column).Take(30))
                    sb.AppendLine($"{e.Row,4}  {e.DeltaXPt,7:0.0000}  {e.DeltaYPt,7:0.0000}");
                if (r.ClusterErrors.Count > 30)
                    sb.AppendLine($"... ({r.ClusterErrors.Count - 30} more)");
                sb.AppendLine("```");
                sb.AppendLine();
            }
        }

        // =========================================================
        // Section 5: Cross-Template Error Pattern Analysis
        // =========================================================
        sb.AppendLine("## Cross-Template Error Pattern Analysis");
        sb.AppendLine();

        // Row regression slopes across templates
        sb.AppendLine("### Vertical Error (ΔY) Consistency");
        sb.AppendLine();
        sb.AppendLine("If the vertical error is caused by a per-row gap accumulation, " +
            "the Y error vs row regression should show a consistent slope across templates " +
            "that is proportional to row height.");
        sb.AppendLine();

        sb.AppendLine("| ID | Y~Row Slope | Y~Row R² | Row Height | Predicted Gap | Notes |");
        sb.AppendLine("|----|-------------|----------|------------|---------------|-------|");
        foreach (var r in results.OrderBy(r => r.DefTopId))
        {
            var rowH = r.WorkbookChars.DefaultRowHeight ?? 14.4;
            double predictedGap = rowH * 0.36; // PerRowExtraPt
            sb.AppendLine($"| {r.DefTopId} | {r.YRowRegSlope:+0.0000000;-0.0000000} | " +
                $"{r.YRowRegRSquared:F6} | {rowH:F1} | {predictedGap:F4} | " +
                $"{(Math.Abs(r.YRowRegSlope - predictedGap) < 0.5 ? "Matches PerRowExtra" : Math.Abs(r.YRowRegSlope) < 0.01 ? "No vertical error" : "Different pattern")} |");
        }
        sb.AppendLine();

        // Column regression slopes across templates
        sb.AppendLine("### Horizontal Error (ΔX) Consistency");
        sb.AppendLine();
        sb.AppendLine("If the horizontal error is caused by font metric differences, " +
            "the X error vs column regression should show a consistent per-column offset.");
        sb.AppendLine();

        sb.AppendLine("| ID | X~Col Slope | X~Col R² | Default Font | ContentWidth | Notes |");
        sb.AppendLine("|----|-------------|----------|--------------|-------------|-------|");
        foreach (var r in results.OrderBy(r => r.DefTopId))
        {
            sb.AppendLine($"| {r.DefTopId} | {r.ColRegSlope:+0.0000000;-0.0000000} | " +
                $"{r.ColRegRSquared:F6} | {r.WorkbookChars.DefaultFontName ?? "?"} {r.WorkbookChars.DefaultFontSize ?? 11} | " +
                $"{r.WorkbookChars.ContentWidth:F0} | " +
                $"{(Math.Abs(r.ColRegSlope) < 0.01 ? "No horizontal error" : Math.Abs(r.ColRegSlope) < 1 ? "Small font-metric offset" : "Significant column-width mismatch")} |");
        }
        sb.AppendLine();

        // =========================================================
        // Section 6: Pattern Discovery
        // =========================================================
        sb.AppendLine("## Pattern Discovery");
        sb.AppendLine();

        sb.AppendLine("### Question 1: Is the error constant across the template?");
        sb.AppendLine();
        foreach (var r in results.OrderBy(r => r.DefTopId))
        {
            var isConstant = r.StdDevYPt < 1.0 && r.StdDevXPt < 0.5;
            sb.AppendLine($"- **Template {r.DefTopId}** ({r.DefTopName}): " +
                $"{(isConstant ? "YES — constant offset of " : "NO — error varies (")}" +
                $"ΔX={r.MeanErrorXPt:F2}±{r.StdDevXPt:F2}, " +
                $"ΔY={r.MeanErrorYPt:F2}±{r.StdDevYPt:F2}" +
                $"{(isConstant ? "" : ")")}");
        }
        sb.AppendLine();

        sb.AppendLine("### Question 2: Is the error linear with row count?");
        sb.AppendLine();
        foreach (var r in results.OrderBy(r => r.DefTopId))
        {
            sb.AppendLine($"- **Template {r.DefTopId}** ({r.DefTopName}): " +
                $"Y~Row R²={r.YRowRegRSquared:F4}, " +
                $"slope={r.YRowRegSlope:+0.0000;-0.0000} pt/row" +
                $"{(Math.Abs(r.YRowRegRSquared) > 0.7 ? " — **Strong linear correlation**" : Math.Abs(r.YRowRegRSquared) > 0.3 ? " — Moderate correlation" : " — Weak/no correlation")}");
        }
        sb.AppendLine();

        sb.AppendLine("### Question 3: Is the error proportional to content width/page width?");
        sb.AppendLine();
        foreach (var r in results.OrderBy(r => r.DefTopId))
        {
            bool hasPrintArea = !string.IsNullOrEmpty(r.PrintArea);
            bool hasFitToPages = r.FitToPagesWide > 0 || r.FitToPagesTall > 0;
            double contentRatio = r.PageWidth > 0 ? r.WorkbookChars.ContentWidth / r.PageWidth : 0;
            sb.AppendLine($"- **Template {r.DefTopId}** ({r.DefTopName}): " +
                $"contentWidth/pageWidth={contentRatio:F2}, " +
                $"PrintArea={(hasPrintArea ? "✓" : "✗")}, " +
                $"FitToPages={(hasFitToPages ? "✓" : "✗")}, " +
                $"Mean ΔX={r.MeanErrorXPt:+0.000;-0.000}");
        }
        sb.AppendLine();

        sb.AppendLine("### Question 4: Piecewise error (different behavior in different regions)?");
        sb.AppendLine();
        foreach (var r in results.OrderBy(r => r.DefTopId))
        {
            if (r.ClusterErrors.Count < 4) continue;

            // Split clusters into top half and bottom half
            var sorted = r.ClusterErrors.OrderBy(e => e.Row).ToList();
            int mid = sorted.Count / 2;
            var topHalf = sorted.Take(mid).ToList();
            var bottomHalf = sorted.Skip(mid).ToList();

            double topMeanY = topHalf.Average(e => e.DeltaYPt);
            double bottomMeanY = bottomHalf.Average(e => e.DeltaYPt);
            double diff = Math.Abs(topMeanY - bottomMeanY);

            sb.AppendLine($"- **Template {r.DefTopId}** ({r.DefTopName}): " +
                $"Top-half mean ΔY={topMeanY:+0.000;-0.000}, " +
                $"Bottom-half mean ΔY={bottomMeanY:+0.000;-0.000}, " +
                $"Difference={diff:F4}" +
                $"{(diff > 3 ? " — **Possible piecewise behavior**" : " — Uniform")}");
        }
        sb.AppendLine();

        // =========================================================
        // Section 7: Switching Rule Analysis
        // =========================================================
        sb.AppendLine("## Switching Rule Analysis");
        sb.AppendLine();
        sb.AppendLine("Based on the error patterns, what property determines " +
            "whether a template needs a different coordinate transform?");
        sb.AppendLine();

        // Compare templates that match vs don't match
        var matching = results.Where(r => r.ClusterErrors.Count > 0 &&
            Math.Abs(r.MeanErrorXPt) < 0.5 && Math.Abs(r.MeanErrorYPt) < 1.0).ToList();
        var nonMatching = results.Where(r => r.ClusterErrors.Count > 0 &&
            (Math.Abs(r.MeanErrorXPt) >= 0.5 || Math.Abs(r.MeanErrorYPt) >= 1.0)).ToList();

        sb.AppendLine($"Templates with Pure COM match: {string.Join(", ", matching.Select(r => r.DefTopId))}");
        sb.AppendLine($"Templates with residual error: {string.Join(", ", nonMatching.Select(r => r.DefTopId))}");
        sb.AppendLine();

        if (matching.Count >= 2 && nonMatching.Count >= 2)
        {
            sb.AppendLine("### Discriminant Analysis");
            sb.AppendLine();
            sb.AppendLine("| Property | Match Avg | Non-Match Avg | Discriminant? |");
            sb.AppendLine("|----------|-----------|---------------|:-------------:|");

            // Mean values for each property
            double matchContentW = matching.Average(m => m.WorkbookChars.ContentWidth);
            double nonMatchContentW = nonMatching.Average(n => n.WorkbookChars.ContentWidth);
            sb.AppendLine($"| Content Width | {matchContentW:F0} | {nonMatchContentW:F0} | " +
                $"{(Math.Abs(matchContentW - nonMatchContentW) > 100 ? "✓" : "✗")} |");

            double matchMergeRatio = matching.Average(m => (double)m.WorkbookChars.MergeCount / Math.Max(m.ClusterErrors.Count, 1));
            double nonMatchMergeRatio = nonMatching.Average(n => (double)n.WorkbookChars.MergeCount / Math.Max(n.ClusterErrors.Count, 1));
            sb.AppendLine($"| Merge Ratio | {matchMergeRatio:F3} | {nonMatchMergeRatio:F3} | " +
                $"{(Math.Abs(matchMergeRatio - nonMatchMergeRatio) > 0.3 ? "✓" : "✗")} |");

            double matchPrintArea = matching.Count(m => !string.IsNullOrEmpty(m.PrintArea));
            double nonMatchPrintArea = nonMatching.Count(n => !string.IsNullOrEmpty(n.PrintArea));
            sb.AppendLine($"| Has PrintArea | {matchPrintArea}/{matching.Count} | {nonMatchPrintArea}/{nonMatching.Count} | " +
                $"{(Math.Abs(matchPrintArea / matching.Count - nonMatchPrintArea / nonMatching.Count) > 0.3 ? "✓" : "✗")} |");

            double matchZoom = matching.Average(m => m.Zoom);
            double nonMatchZoom = nonMatching.Average(n => n.Zoom);
            sb.AppendLine($"| Zoom | {matchZoom:F0}% | {nonMatchZoom:F0}% | " +
                $"{(Math.Abs(matchZoom - nonMatchZoom) > 10 ? "✓" : "✗")} |");

            double matchFitPages = matching.Count(m => m.FitToPagesWide > 0 || m.FitToPagesTall > 0);
            double nonMatchFitPages = nonMatching.Count(n => n.FitToPagesWide > 0 || n.FitToPagesTall > 0);
            sb.AppendLine($"| Has FitToPages | {matchFitPages}/{matching.Count} | {nonMatchFitPages}/{nonMatching.Count} | " +
                $"{(Math.Abs(matchFitPages / matching.Count - nonMatchFitPages / nonMatching.Count) > 0.3 ? "✓" : "✗")} |");

            double matchCenterH = matching.Count(m => m.CenterHorizontally);
            double nonMatchCenterH = nonMatching.Count(n => n.CenterHorizontally);
            sb.AppendLine($"| CenterH | {matchCenterH}/{matching.Count} | {nonMatchCenterH}/{nonMatching.Count} | " +
                $"{(Math.Abs(matchCenterH / matching.Count - nonMatchCenterH / nonMatching.Count) > 0.3 ? "✓" : "✗")} |");

            double matchCols = matching.Average(m => m.WorkbookChars.ColumnCount);
            double nonMatchCols = nonMatching.Average(n => n.WorkbookChars.ColumnCount);
            sb.AppendLine($"| Column Count | {matchCols:F0} | {nonMatchCols:F0} | " +
                $"{(Math.Abs(matchCols - nonMatchCols) > 5 ? "✓" : "✗")} |");

            double matchRows = matching.Average(m => m.WorkbookChars.HiddenRows);
            double nonMatchRows = nonMatching.Average(n => n.WorkbookChars.HiddenRows);
            sb.AppendLine($"| Hidden Rows | {matchRows:F0} | {nonMatchRows:F0} | " +
                $"{(Math.Abs(matchRows - nonMatchRows) > 1 ? "✓" : "✗")} |");
        }
        sb.AppendLine();

        // =========================================================
        // Section 8: Final Deliverable
        // =========================================================
        sb.AppendLine("## Final Deliverable — Error Pattern Classification");
        sb.AppendLine();

        sb.AppendLine("| Pattern | Templates | Evidence | Confidence |");
        sb.AppendLine("|---------|-----------|----------|:----------:|");
        foreach (var r in results.OrderBy(r => r.DefTopId))
        {
            if (r.ClusterErrors.Count == 0) continue;

            bool isPureCom = Math.Abs(r.MeanErrorXPt) < 0.5 && Math.Abs(r.MeanErrorYPt) < 0.5;
            bool fontIssue = Math.Abs(r.MeanErrorXPt) >= 0.5 && Math.Abs(r.MeanErrorYPt) < 1.0;
            bool vertIssue = Math.Abs(r.MeanErrorYPt) >= 1.0;

            string pattern, evidence, confidence;
            if (isPureCom)
            {
                pattern = "Pure COM";
                evidence = "100% coordinate match";
                confidence = "High";
            }
            else if (fontIssue && vertIssue)
            {
                pattern = "COM + Font + Vertical Transform";
                evidence = $"ΔX={r.MeanErrorXPt:F2}pt (font), ΔY={r.MeanErrorYPt:F2}pt (vertical), " +
                    $"Y~Row R²={r.YRowRegRSquared:F4}";
                confidence = Math.Abs(r.YRowRegRSquared) > 0.7 ? "High" : "Medium";
            }
            else if (fontIssue)
            {
                pattern = "COM + Font Metric Adjustment";
                evidence = $"Column offset only: ΔX={r.MeanErrorXPt:F2}pt";
                confidence = "High";
            }
            else if (vertIssue)
            {
                pattern = "COM + Additional Vertical Transform";
                evidence = $"Y~Row slope={r.YRowRegSlope:+0.0000;-0.0000}pt/row, " +
                    $"R²={r.YRowRegRSquared:F4}";
                confidence = Math.Abs(r.YRowRegRSquared) > 0.7 ? "High" : "Medium";
            }
            else
            {
                pattern = "Unknown";
                evidence = "Needs more investigation";
                confidence = "Low";
            }

            sb.AppendLine($"| {pattern} | {r.DefTopId} | {evidence} | {confidence} |");
        }
        sb.AppendLine();

        // =========================================================
        // Section 9: Recommendations
        // =========================================================
        sb.AppendLine("## Recommendations");
        sb.AppendLine();

        sb.AppendLine("### 1. Implement Per-Row Gap Correction");
        sb.AppendLine();
        sb.AppendLine("For templates where Y~Row R² > 0.7, the missing vertical transform is:");
        sb.AppendLine();
        sb.AppendLine("```");
        sb.AppendLine("VerticalGap(row) = baseGap + perRowExtra × (row - firstRow)");
        sb.AppendLine("Corrected_Top = COM_Top + PrintedOriginY + VerticalGap(row)");
        sb.AppendLine("```");
        sb.AppendLine();

        sb.AppendLine("### 2. Implement Per-Column Font Metric Correction");
        sb.AppendLine();
        sb.AppendLine("For templates where the default font is not Calibri, adjust column widths:");
        sb.AppendLine();
        sb.AppendLine("```");
        sb.AppendLine("ActualFontWidth = COM.ColumnWidth(chars)  // from EntireColumn.Width");
        sb.AppendLine("FontMetricFactor = ActualFontWidth / OpenXML_Width");
        sb.AppendLine("Corrected_Left = sum(columnWidth × FontMetricFactor) for columns before cluster");
        sb.AppendLine("```");
        sb.AppendLine();

        sb.AppendLine("### 3. Implement Print-Scaled Strategy");
        sb.AppendLine();
        sb.AppendLine("For templates with FitToPages or large content/page ratio:");
        sb.AppendLine("```");
        sb.AppendLine("ScaleX = PageWidth / ContentWidth");
        sb.AppendLine("Scaled_Position = COM_Position × ScaleX");
        sb.AppendLine("```");
        sb.AppendLine();

        // Save report
        var reportPath = Path.Combine(_outputDir, "phase7_error_investigation_report.md");
        await File.WriteAllTextAsync(reportPath, sb.ToString());
        _progress.Report($"Wrote phase7_error_investigation_report.md");

        // Also save per-template CSV data for external analysis
        await SaveCsvDataAsync(results);

        _progress.Report("Phase 7 investigation complete.");
    }

    /// <summary>
    /// Save per-cluster error data as CSV for external analysis/plotting.
    /// </summary>
    private async Task SaveCsvDataAsync(List<TemplateErrorReport> results)
    {
        var csvPath = Path.Combine(_outputDir, "phase7_per_cluster_errors.csv");
        var sb = new StringBuilder();
        sb.AppendLine("TemplateID,TemplateName,CellAddress,Row,Column,DBLeft,DBTop,DBRight,DBBottom," +
            "COMLeft,COMTop,COMWidth,COMHeight," +
            "GenLeft,GenTop,GenRight,GenBottom," +
            "DeltaLeft,DeltaTop,DeltaRight,DeltaBottom," +
            "DeltaXPt,DeltaYPt,DistanceFromOrigin,IsMerged," +
            "PageWidth,PageHeight,PrintedOriginX,PrintedOriginY," +
            "CenterHorizontally,CenterVertically,HasPrintArea,Zoom,FitToPagesWide,FitToPagesTall");

        foreach (var r in results.OrderBy(r => r.DefTopId))
        {
            foreach (var e in r.ClusterErrors)
            {
                sb.AppendLine(
                    $"{r.DefTopId}," +
                    $"\"{r.DefTopName}\"," +
                    $"{e.CellAddress}," +
                    $"{e.Row},{e.Column}," +
                    $"{e.DbLeft:F9},{e.DbTop:F9},{e.DbRight:F9},{e.DbBottom:F9}," +
                    $"{e.ComLeft:F4},{e.ComTop:F4},{e.ComWidth:F4},{e.ComHeight:F4}," +
                    $"{e.GeneratedLeft:F9},{e.GeneratedTop:F9},{e.GeneratedRight:F9},{e.GeneratedBottom:F9}," +
                    $"{e.DeltaLeft:F9},{e.DeltaTop:F9},{e.DeltaRight:F9},{e.DeltaBottom:F9}," +
                    $"{e.DeltaXPt:F6},{e.DeltaYPt:F6},{e.DistanceFromOrigin:F2},{e.IsMerged}," +
                    $"{r.PageWidth:F1},{r.PageHeight:F1},{r.PrintedOriginX:F4},{r.PrintedOriginY:F4}," +
                    $"{r.CenterHorizontally},{r.CenterVertically}," +
                    $"{(string.IsNullOrEmpty(r.PrintArea) ? "false" : "true")}," +
                    $"{r.Zoom},{r.FitToPagesWide},{r.FitToPagesTall}");
            }
        }

        await File.WriteAllTextAsync(csvPath, sb.ToString());
        _progress.Report($"Wrote phase7_per_cluster_errors.csv");
    }

    private static string ClassifyPattern(TemplateErrorReport r)
    {
        if (r.ClusterErrors.Count == 0) return "NO_DATA";
        double mx = r.MeanErrorXPt, my = r.MeanErrorYPt;
        if (Math.Abs(mx) < 0.5 && Math.Abs(my) < 0.5) return "MATCH";
        if (Math.Abs(mx) >= 0.5 && Math.Abs(my) < 1.0) return "HORIZ-ONLY";
        if (Math.Abs(mx) < 0.5 && Math.Abs(my) >= 1.0) return "VERT-ONLY";
        return "BOTH-AXES";
    }

    private static int ParseCol(string addr)
    {
        // Handle formats: "$A$1", "$A$1:$B$2", "A1", "A1:B2"
        addr = addr.Replace("$", "").Split(':')[0];
        var letters = new string(addr.TakeWhile(char.IsLetter).ToArray());
        int col = 0;
        foreach (char c in letters)
            col = col * 26 + (c - 'A' + 1);
        return col;
    }

    private static int ParseRow(string addr)
    {
        // Handle formats: "$A$1", "$A$1:$B$2", "A1", "A1:B2"
        addr = addr.Replace("$", "").Split(':')[0];
        var digits = new string(addr.SkipWhile(char.IsLetter).TakeWhile(char.IsDigit).ToArray());
        return int.TryParse(digits, out var r) ? r : 0;
    }

    private static double ParseDoubleSafe(string? s) =>
        double.TryParse(s, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;
}

/// <summary>Per-template error report.</summary>
public class TemplateErrorReport
{
    public int DefTopId { get; set; }
    public string DefTopName { get; set; } = "";
    public int ClusterCount { get; set; }
    public string? ErrorMessage { get; set; }

    // Page setup from COM
    public double PageWidth { get; set; }
    public double PageHeight { get; set; }
    public double PrintedOriginX { get; set; }
    public double PrintedOriginY { get; set; }
    public double ClusterGapX { get; set; }
    public double ClusterGapY { get; set; }
    public bool CenterHorizontally { get; set; }
    public bool CenterVertically { get; set; }
    public string? PrintArea { get; set; }
    public int Zoom { get; set; }
    public int FitToPagesWide { get; set; }
    public int FitToPagesTall { get; set; }
    public int PaperSize { get; set; }
    public int Orientation { get; set; }

    // Workbook characteristics
    public WorkbookCharacteristics WorkbookChars { get; set; } = new();

    // Per-cluster errors
    public List<ClusterError> ClusterErrors { get; set; } = new();
    public List<DefCluster> UnmatchedClusters { get; set; } = new();

    // Statistics
    public double MeanErrorXPt { get; set; }
    public double MeanErrorYPt { get; set; }
    public double MeanErrorXRatio { get; set; }
    public double MeanErrorYRatio { get; set; }
    public double StdDevXPt { get; set; }
    public double StdDevYPt { get; set; }
    public double MinDeltaXPt { get; set; }
    public double MinDeltaYPt { get; set; }
    public double MaxDeltaXPt { get; set; }
    public double MaxDeltaYPt { get; set; }
    public double MaxAbsDeltaXPt { get; set; }
    public double MaxAbsDeltaYPt { get; set; }

    // Regression: X error vs Row
    public double RowRegSlope { get; set; }
    public double RowRegIntercept { get; set; }
    public double RowRegRSquared { get; set; }

    // Regression: X error vs Column
    public double ColRegSlope { get; set; }
    public double ColRegIntercept { get; set; }
    public double ColRegRSquared { get; set; }

    // Regression: Y error vs Row
    public double YRowRegSlope { get; set; }
    public double YRowRegIntercept { get; set; }
    public double YRowRegRSquared { get; set; }

    // Regression: Y error vs Distance from origin
    public double DistRegSlope { get; set; }
    public double DistRegIntercept { get; set; }
    public double DistRegRSquared { get; set; }

    public void ComputeStatistics()
    {
        if (ClusterErrors.Count == 0) return;

        var xs = ClusterErrors.Select(e => e.DeltaXPt).ToList();
        var ys = ClusterErrors.Select(e => e.DeltaYPt).ToList();
        var xRatios = ClusterErrors.Select(e => e.DeltaLeft).ToList();
        var yRatios = ClusterErrors.Select(e => e.DeltaTop).ToList();

        MeanErrorXPt = xs.Average();
        MeanErrorYPt = ys.Average();
        MeanErrorXRatio = xRatios.Average();
        MeanErrorYRatio = yRatios.Average();

        MinDeltaXPt = xs.Min();
        MinDeltaYPt = ys.Min();
        MaxDeltaXPt = xs.Max();
        MaxDeltaYPt = ys.Max();
        MaxAbsDeltaXPt = xs.Max(v => Math.Abs(v));
        MaxAbsDeltaYPt = ys.Max(v => Math.Abs(v));

        StdDevXPt = Math.Sqrt(xs.Average(v => Math.Pow(v - MeanErrorXPt, 2)));
        StdDevYPt = Math.Sqrt(ys.Average(v => Math.Pow(v - MeanErrorYPt, 2)));

        // Linear regression: X error vs Row
        {
            var r = DoRegress(
                ClusterErrors.Select(e => (double)e.Row).ToList(), xs);
            RowRegSlope = r.slope; RowRegIntercept = r.intercept; RowRegRSquared = r.rSquared;
        }
        // Linear regression: X error vs Column
        {
            var r = DoRegress(
                ClusterErrors.Select(e => (double)e.Column).ToList(), xs);
            ColRegSlope = r.slope; ColRegIntercept = r.intercept; ColRegRSquared = r.rSquared;
        }
        // Linear regression: Y error vs Row
        {
            var r = DoRegress(
                ClusterErrors.Select(e => (double)e.Row).ToList(), ys);
            YRowRegSlope = r.slope; YRowRegIntercept = r.intercept; YRowRegRSquared = r.rSquared;
        }
        // Linear regression: Y error vs Distance from origin
        {
            var r = DoRegress(
                ClusterErrors.Select(e => e.DistanceFromOrigin).ToList(), ys);
            DistRegSlope = r.slope; DistRegIntercept = r.intercept; DistRegRSquared = r.rSquared;
        }
    }

    private static (double slope, double intercept, double rSquared) DoRegress(List<double> x, List<double> y)
    {
        int n = x.Count;
        if (n < 2) return (0, 0, 0);

        double sx = x.Sum(), sy = y.Sum();
        double sxx = x.Sum(v => v * v);
        double sxy = 0;
        for (int i = 0; i < n; i++) sxy += x[i] * y[i];

        double denom = n * sxx - sx * sx;
        if (Math.Abs(denom) < 1e-15) return (0, 0, 0);

        double slope = (n * sxy - sx * sy) / denom;
        double intercept = (sy - slope * sx) / n;

        // R² = 1 - SS_res / SS_tot
        double meanY = sy / n;
        double ssRes = 0, ssTot = 0;
        for (int i = 0; i < n; i++)
        {
            double pred = slope * x[i] + intercept;
            ssRes += Math.Pow(y[i] - pred, 2);
            ssTot += Math.Pow(y[i] - meanY, 2);
        }
        double rSquared = ssTot > 1e-15 ? 1.0 - ssRes / ssTot : 0;
        return (slope, intercept, rSquared);
    }
}

/// <summary>Per-cluster error data.</summary>
public class ClusterError
{
    public string CellAddress { get; set; } = "";
    public int Row { get; set; }
    public int Column { get; set; }

    // DB values (ratios)
    public double DbLeft { get; set; }
    public double DbTop { get; set; }
    public double DbRight { get; set; }
    public double DbBottom { get; set; }

    // COM values (points)
    public double ComLeft { get; set; }
    public double ComTop { get; set; }
    public double ComWidth { get; set; }
    public double ComHeight { get; set; }

    // Generated values (ratios from our engine)
    public double GeneratedLeft { get; set; }
    public double GeneratedTop { get; set; }
    public double GeneratedRight { get; set; }
    public double GeneratedBottom { get; set; }

    // Delas (DB - Generated)
    public double DeltaLeft { get; set; }
    public double DeltaTop { get; set; }
    public double DeltaRight { get; set; }
    public double DeltaBottom { get; set; }

    // Delas in points
    public double DeltaXPt { get; set; }
    public double DeltaYPt { get; set; }

    // Distance (COM position + origin) from page origin
    public double DistanceFromOrigin { get; set; }

    // Flags
    public bool IsMerged { get; set; }
}

/// <summary>Workbook characteristics from OpenXML.</summary>
public class WorkbookCharacteristics
{
    public int ColumnCount { get; set; }
    public int CellCount { get; set; }
    public int MergeCount { get; set; }
    public int HiddenColumns { get; set; }
    public int HiddenRows { get; set; }
    public string? DefaultFontName { get; set; }
    public double? DefaultFontSize { get; set; }
    public double ContentWidth { get; set; }
    public bool HasPrintArea { get; set; }
    public double? DefaultRowHeight { get; set; }
    public int? PageSetupScale { get; set; }
    public int? FitToWidth { get; set; }
    public int? FitToHeight { get; set; }
    public double? PageWidth { get; set; }
    public double? PageHeight { get; set; }
    public string? Orientation { get; set; }
}
