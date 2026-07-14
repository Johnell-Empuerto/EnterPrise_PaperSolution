using Microsoft.AspNetCore.Mvc;
using ExcelAPI.LegacyEngine.OpenXml;
using ExcelAPI.Models;
using ExcelAPI.Rendering;
using ExcelAPI.Runtime;
using ExcelAPI.Services;
using SkiaSharp;

namespace ExcelAPI.Controllers;

[ApiController]
[Route("api/[controller]")]
public class LegacyRuntimeController : ControllerBase
{
    private readonly OpenXmlParser _parser;
    private readonly GeometryBuilder _geometry;
    private readonly FormRuntimeBuilder _runtimeBuilder;
    private readonly IPageSetupReader _pageSetupReader;
    private readonly RendererCoordinator _renderCoordinator;
    private readonly PythonRenderService _pythonRender;
    private readonly IWebHostEnvironment _env;
    private readonly ILogger<LegacyRuntimeController> _logger;

    private static readonly Dictionary<int, string> TemplatePaths = new()
    {
        [546] = @"..\..\Investigation_546\original.xlsx",
        [547] = @"..\..\[V3.1_Sample]アンケート用紙.xlsx",
        [548] = @"Forms\0f8e09082b5c4861abe5d92985ace548.xlsx"
    };

    public LegacyRuntimeController(
        OpenXmlParser parser,
        GeometryBuilder geometry,
        FormRuntimeBuilder runtimeBuilder,
        IPageSetupReader pageSetupReader,
        RendererCoordinator renderCoordinator,
        PythonRenderService pythonRender,
        IWebHostEnvironment env,
        ILogger<LegacyRuntimeController> logger)
    {
        _parser = parser;
        _geometry = geometry;
        _runtimeBuilder = runtimeBuilder;
        _pageSetupReader = pageSetupReader;
        _renderCoordinator = renderCoordinator;
        _pythonRender = pythonRender;
        _env = env;
        _logger = logger;
    }

    [HttpGet("{id}")]
    [ProducesResponseType(typeof(LegacyRuntimeDocument), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public async Task<IActionResult> GetRuntime(int id)
    {
        // Try Python renderer first (Phase 47)
        try
        {
            var pythonResult = await _pythonRender.RenderAsync(id);
            if (pythonResult is not null)
            {
                var doc = new LegacyRuntimeDocument
                {
                    BackgroundImage = pythonResult.BackgroundImage,
                    PageWidth = pythonResult.PageWidth,
                    PageHeight = pythonResult.PageHeight,
                    Fields = pythonResult.Fields.Select(f => new LegacyRuntimeField
                    {
                        Id = f.Id,
                        Label = f.Label,
                        LeftPx = f.LeftPx,
                        TopPx = f.TopPx,
                        WidthPx = f.WidthPx,
                        HeightPx = f.HeightPx,
                        Type = f.Type,
                        Required = f.Required
                    }).ToList()
                };
                return Ok(doc);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Python renderer unavailable for template {Id}, falling back to in-process renderer", id);
        }

        // Fallback: find the xlsx file
        string? xlsxPath = null;
        if (TemplatePaths.TryGetValue(id, out var relativePath))
        {
            var fullPath = Path.Combine(_env.ContentRootPath, relativePath);
            if (System.IO.File.Exists(fullPath))
                xlsxPath = fullPath;
        }
        if (xlsxPath is null)
        {
            string formsDir = Path.Combine(_env.ContentRootPath, "Forms");
            if (Directory.Exists(formsDir))
            {
                var files = Directory.GetFiles(formsDir, "*.xlsx");
                xlsxPath = files.FirstOrDefault(f =>
                    string.Equals(Path.GetFileNameWithoutExtension(f), id.ToString(), StringComparison.OrdinalIgnoreCase) ||
                    Path.GetFileNameWithoutExtension(f).StartsWith(id.ToString(), StringComparison.OrdinalIgnoreCase));
            }
        }
        if (xlsxPath is null)
            return NotFound(new { error = $"Template {id} not found" });

        try
        {
            return BuildRuntimeDocument(xlsxPath, id);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to build runtime for template {Id}", id);
            return StatusCode(500, new { error = $"Failed to build runtime: {ex.Message}" });
        }
    }

    private IActionResult BuildRuntimeDocument(string xlsxPath, int templateId)
    {
        const double dpi = 300;

        var workbook = _parser.Parse(xlsxPath);
        if (workbook.Sheets.Count == 0)
            return NotFound(new { error = "No sheets in workbook" });

        var sheet = workbook.Sheets[0];
        _geometry.ComputeGeometry(sheet);

        string sheetName = sheet.Name;
        int sheetId = sheet.Index + 1;
        var pageSetup = _pageSetupReader.Read(xlsxPath, sheetName, sheetId);

        double pageWidthPt = pageSetup.PageWidthPoints;
        double pageHeightPt = pageSetup.PageHeightPoints;
        double marginLeftPt = pageSetup.MarginLeft * 72.0;
        double marginRightPt = pageSetup.MarginRight * 72.0;
        double marginTopPt = pageSetup.MarginTop * 72.0;
        double marginBottomPt = pageSetup.MarginBottom * 72.0;
        double contentWidthPt = sheet.TotalWidthPt;
        double contentHeightPt = sheet.TotalHeightPt;

        double printableWidthPt = pageWidthPt - marginLeftPt - marginRightPt;
        double printableHeightPt = pageHeightPt - marginTopPt - marginBottomPt;

        double originXPt = marginLeftPt;
        double originYPt = marginTopPt;
        if (pageSetup.CenterHorizontally && contentWidthPt < printableWidthPt)
            originXPt += (printableWidthPt - contentWidthPt) / 2.0;
        if (pageSetup.CenterVertically && contentHeightPt < printableHeightPt)
            originYPt += (printableHeightPt - contentHeightPt) / 2.0;

        var runtimeForm = _runtimeBuilder.Build(workbook, (int)dpi, originXPt, originYPt);

        var ctx = new RenderingContext
        {
            Workbook = workbook,
            Sheet = sheet,
            PaperSize = "A4",
            Orientation = pageSetup.Orientation.ToLowerInvariant(),
            PageWidthPt = pageWidthPt,
            PageHeightPt = pageHeightPt,
            MarginLeftPt = marginLeftPt,
            MarginRightPt = marginRightPt,
            MarginTopPt = marginTopPt,
            MarginBottomPt = marginBottomPt,
            PrintableWidthPt = printableWidthPt,
            PrintableHeightPt = printableHeightPt,
            OriginXPt = originXPt,
            OriginYPt = originYPt,
            ScaleFactor = 1.0,
            Zoom = 100,
            IsScalingActive = false,
            ClipLeftPt = originXPt,
            ClipTopPt = originYPt,
            ClipRightPt = Math.Min(originXPt + contentWidthPt, marginLeftPt + printableWidthPt),
            ClipBottomPt = Math.Min(originYPt + contentHeightPt, marginTopPt + printableHeightPt),
            Dpi = dpi
        };

        int pageWidthPx = ctx.PixelWidth;
        int pageHeightPx = ctx.PixelHeight;

        string pngFileName = $"runtime_{templateId}_page1.png";
        string previewDir = Path.Combine(_env.ContentRootPath, "Preview");
        string pngPath = Path.Combine(previewDir, pngFileName);

        if (!Directory.Exists(previewDir))
            Directory.CreateDirectory(previewDir);

        using var bitmap = new SKBitmap(pageWidthPx, pageHeightPx);
        using var canvas = new SKCanvas(bitmap);
        _renderCoordinator.RenderToCanvas(canvas, ctx);

        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(SKEncodedImageFormat.Png, 90);
        System.IO.File.WriteAllBytes(pngPath, data.ToArray());

        var runtimeSheet = runtimeForm.Sheets.FirstOrDefault();
        var fields = new List<LegacyRuntimeField>();

        if (runtimeSheet is not null)
        {
            foreach (var rf in runtimeSheet.Fields)
            {
                fields.Add(new LegacyRuntimeField
                {
                    Id = rf.Id,
                    Label = rf.CellReference,
                    LeftPx = rf.LeftPx,
                    TopPx = rf.TopPx,
                    WidthPx = rf.WidthPx,
                    HeightPx = rf.HeightPx,
                    Type = rf.DataType ?? "text",
                    Required = rf.Required
                });
            }
        }

        var doc = new LegacyRuntimeDocument
        {
            BackgroundImage = $"/preview/{pngFileName}",
            PageWidth = pageWidthPx,
            PageHeight = pageHeightPx,
            Fields = fields
        };

        return Ok(doc);
    }
}
