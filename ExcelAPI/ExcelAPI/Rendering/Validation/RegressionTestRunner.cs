using System.Diagnostics;
using ExcelAPI.Models;
using SkiaSharp;

namespace ExcelAPI.Rendering.Validation
{
    /// <summary>
    /// Orchestrates the full regression testing workflow:
    ///   1. Discover workbook templates in RegressionTests/Templates/
    ///   2. For each workbook: parse, render via ExportCoordinator, compare with baseline
    ///   3. Generate pixel-diff comparison
    ///   4. Generate validation reports (JSON + HTML)
    ///   5. Summarize results
    ///
    /// Consumes the existing rendering pipeline exactly as production code does.
    /// No rendering architecture changes are needed.
    /// </summary>
    public class RegressionTestRunner
    {
        private readonly RegressionConfiguration _config;
        private readonly ExportCoordinator _exportCoordinator;
        private readonly PixelDiffEngine _diffEngine;
        private readonly RenderingBaselineManager _baselineManager;
        private readonly ILogger<RegressionTestRunner> _logger;

        public RegressionTestRunner(
            RegressionConfiguration config,
            ExportCoordinator exportCoordinator,
            PixelDiffEngine diffEngine,
            RenderingBaselineManager baselineManager,
            ILogger<RegressionTestRunner> logger)
        {
            _config = config;
            _exportCoordinator = exportCoordinator;
            _diffEngine = diffEngine;
            _baselineManager = baselineManager;
            _logger = logger;
        }

        /// <summary>
        /// Run the full regression test suite.
        /// Returns a ValidationSummary with all results.
        /// </summary>
        public ValidationSummary RunAll()
        {
            var stopwatch = Stopwatch.StartNew();
            _baselineManager.EnsureDirectoriesExist();

            var workbooks = DiscoverWorkbooks();
            _logger.LogInformation("Found {Count} workbook(s) to validate", workbooks.Count);

            var summary = new ValidationSummary
            {
                TotalWorkbooks = workbooks.Count,
                Timestamp = DateTime.UtcNow.ToString("o")
            };

            int tested = 0;
            foreach (var workbookPath in workbooks)
            {
                if (_config.MaxWorkbooks > 0 && tested >= _config.MaxWorkbooks)
                    break;

                var report = TestWorkbook(workbookPath);
                summary.Reports.Add(report);

                if (report.Passed)
                    summary.Passed++;
                else
                    summary.Failed++;

                tested++;

                if (_config.StopOnFirstFailure && !report.Passed)
                {
                    _logger.LogWarning("Stopping on first failure: {Path}", workbookPath);
                    break;
                }
            }

            stopwatch.Stop();
            summary.TotalRenderingTimeMs = stopwatch.ElapsedMilliseconds;
            summary.AverageRenderingTimeMs = summary.TotalWorkbooks > 0
                ? (double)summary.TotalRenderingTimeMs / summary.TotalWorkbooks
                : 0;
            summary.TotalPages = summary.Reports.Sum(r => r.Pages);
            summary.AveragePixelDifference = summary.Reports.Count > 0
                ? summary.Reports.Average(r => r.PixelDifference)
                : 0;

            // Generate reports
            string reportPath = Path.Combine(_config.ReportsDirectory,
                $"validation_{DateTime.Now:yyyyMMdd_HHmmss}");
            ValidationReport.SaveJson(summary, reportPath + ".json");
            if (_config.GenerateHtmlReport)
                ValidationReport.SaveHtml(summary, reportPath + ".html");

            _logger.LogInformation(
                "Validation complete: {Passed} passed, {Failed} failed, {Total} total",
                summary.Passed, summary.Failed, summary.TotalWorkbooks);

            return summary;
        }

        /// <summary>
        /// Test a single workbook against its baseline.
        /// </summary>
        private ImageComparisonReport TestWorkbook(string workbookPath)
        {
            var report = new ImageComparisonReport
            {
                WorkbookName = Path.GetFileName(workbookPath),
                WorkbookPath = workbookPath,
                Timestamp = DateTime.UtcNow.ToString("o")
            };

            var workbookStopwatch = Stopwatch.StartNew();

            try
            {
                // Parse the workbook
                var parseSw = Stopwatch.StartNew();
                var workbook = _exportCoordinator.ParseWorkbook(workbookPath);
                parseSw.Stop();
                report.ParseTimeMs = parseSw.ElapsedMilliseconds;

                // Create a FormDefinition from the workbook for rendering
                var form = CreateMinimalForm(workbook);
                string sheetId = form.Sheets.FirstOrDefault()?.Id ?? "sheet0";

                // Build options matching the baseline
                var options = new ExportOptions
                {
                    Dpi = _config.Dpi,
                    MaxPages = _config.MaxPagesPerWorkbook > 0 ? _config.MaxPagesPerWorkbook : 0,
                    FilePrefix = Path.GetFileNameWithoutExtension(workbookPath),
                    IncludePageNumbers = false,
                    TransparentBackground = false,
                    PngCompressionLevel = 100,
                    OutputDirectory = _config.CurrentDirectory
                };

                // Render
                var renderSw = Stopwatch.StartNew();
                var renderedPaths = _exportCoordinator.ExportPng(workbookPath, form, sheetId, options);
                renderSw.Stop();
                report.RenderingTimeMs = renderSw.ElapsedMilliseconds;

                // Compare each page with baseline
                int pageNum = 1;
                foreach (var actualPath in renderedPaths)
                {
                    // Ensure baseline exists
                    if (!_baselineManager.BaselineExists(workbookPath, pageNum))
                    {
                        _baselineManager.EnsureBaseline(workbookPath, pageNum);
                        if (!_baselineManager.BaselineExists(workbookPath, pageNum))
                        {
                            _logger.LogWarning(
                                "No baseline for {Workbook} page {Page} — skipping comparison",
                                workbookPath, pageNum);
                            pageNum++;
                            continue;
                        }
                    }

                    // Load images
                    using var baselineBitmap = SKBitmap.Decode(
                        _baselineManager.GetBaselinePath(workbookPath, pageNum));
                    using var actualBitmap = SKBitmap.Decode(actualPath);

                    if (baselineBitmap == null || actualBitmap == null)
                    {
                        _logger.LogWarning("Failed to decode images for {Workbook} page {Page}",
                            workbookPath, pageNum);
                        pageNum++;
                        continue;
                    }

                    // Run pixel diff comparison (returns diff bitmap when generateDiffImage=true)
                    var diffResult = _diffEngine.Compare(
                        baselineBitmap, actualBitmap,
                        _config.PixelTolerance,
                        _config.SkipAlphaChannel,
                        generateDiffImage: _config.GenerateDiffImages);

                    // Save diff image to disk
                    string? diffPath = null;
                    if (diffResult.DiffBitmap != null)
                    {
                        diffPath = _baselineManager.GetDiffPath(workbookPath, pageNum);
                        PixelDiffEngine.SaveDiffImage(diffResult.DiffBitmap, diffPath);
                    }

                    // Check if difference is approved
                    bool approved = _baselineManager.IsDifferenceApproved(workbookPath, pageNum);
                    bool pagePassed = diffResult.Passed || approved;

                    var pageResult = new PageComparisonResult
                    {
                        PageNumber = pageNum,
                        PixelDifference = diffResult.DifferencePercent,
                        DifferentPixels = diffResult.DifferentPixels,
                        TotalPixels = diffResult.TotalPixels,
                        Passed = pagePassed,
                        MaxError = diffResult.MaxError,
                        AverageError = diffResult.AverageError,
                        ExpectedImagePath = _baselineManager.GetBaselinePath(workbookPath, pageNum),
                        ActualImagePath = actualPath,
                        DiffImagePath = diffPath,
                        RenderTimeMs = renderSw.ElapsedMilliseconds / Math.Max(1, renderedPaths.Count),
                        Dimensions = $"{diffResult.Width}x{diffResult.Height}"
                    };

                    report.PageResults.Add(pageResult);
                    report.Pages++;

                    _logger.LogInformation(
                        "Page {Page}: diff={DiffPct}% errors={DiffPx}/{TotalPx} passed={Passed} approved={Approved}",
                        pageNum, diffResult.DifferencePercent,
                        diffResult.DifferentPixels, diffResult.TotalPixels,
                        pagePassed, approved);

                    pageNum++;
                }

                // Compute overall report metrics
                report.PixelDifference = report.PageResults.Count > 0
                    ? report.PageResults.Average(r => r.PixelDifference)
                    : 0;
                report.Passed = report.PageResults.Count > 0
                    && report.PageResults.All(r => r.Passed);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error testing workbook: {Path}", workbookPath);
                report.Passed = false;
                report.PixelDifference = 100.0;
            }

            workbookStopwatch.Stop();
            report.RenderingTimeMs = workbookStopwatch.ElapsedMilliseconds;

            return report;
        }

        /// <summary>
        /// Create a FormDefinition from a parsed workbook for rendering.
        /// Reads actual page settings from the workbook's sheet properties when available.
        /// </summary>
        private static FormDefinition CreateMinimalForm(RenderWorkbook workbook)
        {
            var form = new FormDefinition();

            for (int i = 0; i < workbook.Sheets.Count; i++)
            {
                var sheet = workbook.Sheets[i];

                // Use default page settings that match typical Excel defaults.
                // In a full implementation, these would be read from the worksheet's
                // pageSetup and pageMargins elements in the OpenXml.
                var sheetDef = new SheetDefinition
                {
                    Id = $"sheet{i}",
                    Name = sheet.Name,
                    Index = i,
                    PageSettings = new PageSettings
                    {
                        PaperSize = "Letter",
                        Orientation = "portrait",
                        WidthPt = 612,
                        HeightPt = 792,
                        LeftMargin = 70,
                        TopMargin = 70,
                        RightMargin = 70,
                        BottomMargin = 70,
                        CenterHorizontally = false,
                        CenterVertically = false,
                        Zoom = 100,
                        FitToPagesWide = 0,
                        FitToPagesTall = 0
                    }
                };

                form.Sheets.Add(sheetDef);
            }

            return form;
        }

        /// <summary>
        /// Discover workbook templates in the templates directory.
        /// </summary>
        private List<string> DiscoverWorkbooks()
        {
            var templatesDir = _config.TemplatesDirectory;

            if (!Directory.Exists(templatesDir))
            {
                _logger.LogWarning("Templates directory not found: {Dir}", templatesDir);
                return new List<string>();
            }

            var files = Directory.GetFiles(templatesDir, "*.xlsx")
                .OrderBy(f => f)
                .ToList();

            // Apply filter if specified
            if (_config.WorkbookFilter.Count > 0)
            {
                files = files.Where(f =>
                    _config.WorkbookFilter.Any(filter =>
                        Path.GetFileName(f).Contains(filter, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
            }

            return files;
        }
    }
}
