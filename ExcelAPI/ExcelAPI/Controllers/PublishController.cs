using ExcelAPI.Designer.Legacy.Debug;
using ExcelAPI.Designer.Legacy.Models;
using ExcelAPI.Designer.Legacy.PublishEngine;
using ExcelAPI.Designer.Legacy.VerificationEngine;
using Microsoft.AspNetCore.Mvc;

namespace ExcelAPI.Controllers;

[ApiController]
[Route("api/[controller]")]
public class PublishController : ControllerBase
{
    private readonly IPublishEngine _publishEngine;
    private readonly IVerificationEngine _verificationEngine;
    private readonly ILogger<PublishController> _logger;

    public PublishController(
        IPublishEngine publishEngine,
        IVerificationEngine verificationEngine,
        ILogger<PublishController> logger)
    {
        _publishEngine = publishEngine;
        _verificationEngine = verificationEngine;
        _logger = logger;
    }

    [HttpPost("publish")]
    public async Task<IActionResult> Publish(
        IFormFile file,
        [FromQuery] bool debug = false)
    {
        if (file == null || file.Length == 0)
            return BadRequest(new { error = "No file uploaded." });

        var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
        if (ext != ".xlsx" && ext != ".xls")
            return BadRequest(new { error = "Only .xlsx and .xls files are supported." });

        string tempDir = Path.Combine(Path.GetTempPath(), "PaperLessUpload",
            Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(tempDir);
        string filePath = Path.Combine(tempDir, file.FileName);

        try
        {
            await using (var stream = new FileStream(filePath, FileMode.Create))
                await file.CopyToAsync(stream);

            _logger.LogInformation("Publishing file: {File}, debug={Debug}", file.FileName, debug);

            var result = await _publishEngine.PublishAsync(filePath, tempDir);

            if (!result.Success)
            {
                _logger.LogError("Publish failed: {Error}", result.ErrorMessage);
                return StatusCode(500, new { error = result.ErrorMessage });
            }

            _logger.LogInformation(
                "Publish complete: {Clusters} clusters, {Sheets} sheets",
                result.Clusters.Count, result.Sheets.Count);

            return Ok(new
            {
                success = true,
                transform = _publishEngine.TransformName,
                sheets = result.Sheets.Select(s => new
                {
                    sheetNo = s.SheetNo,
                    sheetName = s.SheetName,
                    width = s.WidthPoints,
                    height = s.HeightPoints,
                    backgroundImage = s.BackgroundImagePath
                }),
                clusters = result.Clusters.Select(c => new
                {
                    c.ClusterId,
                    c.SheetNo,
                    c.CellAddress,
                    c.Left,
                    c.Top,
                    c.Right,
                    c.Bottom
                }),
                xml = result.XmlData,
                debugDir = debug ? Path.Combine(tempDir, "debug") : null
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Publish failed unexpectedly");
            return StatusCode(500, new { error = ex.Message });
        }
        finally
        {
            if (!debug)
            {
                try { Directory.Delete(tempDir, recursive: true); }
                catch { /* temp cleanup */ }
            }
        }
    }

    [HttpPost("publish-and-verify")]
    public async Task<IActionResult> PublishAndVerify(
        IFormFile file,
        [FromQuery] int defTopId,
        [FromQuery] bool debug = false)
    {
        if (file == null || file.Length == 0)
            return BadRequest(new { error = "No file uploaded." });

        var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
        if (ext != ".xlsx" && ext != ".xls")
            return BadRequest(new { error = "Only .xlsx and .xls files are supported." });

        string tempDir = Path.Combine(Path.GetTempPath(), "PaperLessPublishVerify",
            Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(tempDir);
        string filePath = Path.Combine(tempDir, file.FileName);

        try
        {
            await using (var stream = new FileStream(filePath, FileMode.Create))
                await file.CopyToAsync(stream);

            _logger.LogInformation("Publish+Verify: file={File}, defTopId={Id}, debug={Debug}",
                file.FileName, defTopId, debug);

            var publishResult = await _publishEngine.PublishAsync(filePath, tempDir);

            if (!publishResult.Success)
                return StatusCode(500, new { error = publishResult.ErrorMessage });

            var verificationReport = await _verificationEngine.VerifyAsync(
                defTopId, publishResult, tempDir);

            return Ok(new
            {
                success = true,
                transform = _publishEngine.TransformName,
                publish = new
                {
                    sheets = publishResult.Sheets.Select(s => new
                    {
                        sheetNo = s.SheetNo,
                        sheetName = s.SheetName,
                        width = s.WidthPoints,
                        height = s.HeightPoints,
                        backgroundImage = s.BackgroundImagePath
                    }),
                    clusters = publishResult.Clusters.Select(c => new
                    {
                        c.ClusterId,
                        c.SheetNo,
                        c.CellAddress,
                        c.Left,
                        c.Top,
                        c.Right,
                        c.Bottom
                    }),
                    xml = publishResult.XmlData
                },
                verification = verificationReport,
                debugDir = debug ? Path.Combine(tempDir, "debug") : null
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Publish+Verify failed unexpectedly");
            return StatusCode(500, new { error = ex.Message });
        }
        finally
        {
            if (!debug)
            {
                try { Directory.Delete(tempDir, recursive: true); }
                catch { /* temp cleanup */ }
            }
        }
    }

    [HttpGet("transform")]
    public IActionResult GetActiveTransform()
    {
        return Ok(new { transform = _publishEngine.TransformName });
    }

    [HttpPost("verify")]
    public async Task<IActionResult> Verify(
        IFormFile file,
        [FromQuery] int defTopId,
        [FromQuery] bool debug = false)
    {
        if (file == null || file.Length == 0)
            return BadRequest(new { error = "No file uploaded." });

        string tempDir = Path.Combine(Path.GetTempPath(), "PaperLessVerify",
            Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(tempDir);
        string filePath = Path.Combine(tempDir, file.FileName);

        try
        {
            await using (var stream = new FileStream(filePath, FileMode.Create))
                await file.CopyToAsync(stream);

            var verificationReport = await _verificationEngine.VerifyAsync(
                defTopId, filePath, tempDir);

            return Ok(new
            {
                success = true,
                transform = _publishEngine.TransformName,
                verification = verificationReport,
                debugDir = debug ? Path.Combine(tempDir, "debug") : null
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Verify failed unexpectedly");
            return StatusCode(500, new { error = ex.Message });
        }
        finally
        {
            if (!debug)
            {
                try { Directory.Delete(tempDir, recursive: true); }
                catch { /* temp cleanup */ }
            }
        }
    }

    [HttpPost("regression")]
    public async Task<IActionResult> Regression(
        [FromBody] RegressionRequest request)
    {
        if (request?.TemplateIds == null || request.TemplateIds.Count == 0)
            return BadRequest(new { error = "At least one template ID is required." });

        string outputDir = Path.Combine(Path.GetTempPath(), "PaperLessRegression",
            Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(outputDir);

        // Map template IDs to known file paths
        var templatePaths = new Dictionary<int, string>
        {
            { 546, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx" },
            { 547, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\[V3.1_Sample]アンケート用紙.xlsx" },
            { 548, @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Sample A.xlsx" }
        };

        var results = new List<object>();

        foreach (var templateId in request.TemplateIds)
        {
            _logger.LogInformation("Regression: processing template {Id}", templateId);

            if (!templatePaths.TryGetValue(templateId, out var excelPath) ||
                !System.IO.File.Exists(excelPath))
            {
                results.Add(new
                {
                    defTopId = templateId,
                    overall = "FAIL",
                    error = $"Template file not found for ID {templateId}"
                });
                continue;
            }

            try
            {
                var verificationReport = await _verificationEngine.VerifyAsync(
                    templateId, excelPath, outputDir);

                results.Add(new
                {
                    defTopId = templateId,
                    overall = verificationReport.Overall,
                    summary = new
                    {
                        totalChecks = verificationReport.Summary.TotalChecks,
                        passed = verificationReport.Summary.Passed,
                        failed = verificationReport.Summary.Failed,
                        matchPercentage = verificationReport.Summary.OverallMatchPercentage
                    },
                    xmlStatus = verificationReport.XmlComparison?.Status,
                    coordinateStatus = verificationReport.CoordinateComparison?.Status,
                    imageStatus = verificationReport.ImageComparison?.Status,
                    xmlMatchPct = verificationReport.XmlComparison?.MatchPercentage,
                    coordinateMatchPct = verificationReport.CoordinateComparison?.MatchPercentage,
                    error = (string?)null
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Regression failed for template {Id}", templateId);
                results.Add(new
                {
                    defTopId = templateId,
                    overall = "FAIL",
                    error = ex.Message
                });
            }
        }

        bool allPassed = results.All(r =>
        {
            var prop = r.GetType().GetProperty("overall");
            return prop?.GetValue(r)?.ToString() == "PASS";
        });

        return Ok(new
        {
            success = true,
            overall = allPassed ? "PASS" : "FAIL",
            results,
            outputDir
        });
    }
}

public class RegressionRequest
{
    public List<int> TemplateIds { get; set; } = new();
}
