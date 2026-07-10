using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using ExcelAPI.Models;
using ExcelAPI.Services;
using ExcelAPI.Runtime;
using ExcelAPI.Rendering;
using SkiaSharp;

namespace ExcelAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class FormController : ControllerBase
    {
        private readonly IFormSaveService _formSaveService;
        private readonly IWebHostEnvironment _env;
        private readonly IOptions<ExcelCaptureOptions> _options;
        private readonly ILogger<FormController> _logger;
        private readonly OpenXmlParser _xmlParser;
        private readonly FormRuntimeBuilder _runtimeBuilder;
        private readonly RuntimeSerializer _runtimeSerializer;
        private readonly RuntimeMetadataService _runtimeMetadata;

        public FormController(
            IFormSaveService formSaveService,
            IWebHostEnvironment env,
            IOptions<ExcelCaptureOptions> options,
            ILogger<FormController> logger,
            OpenXmlParser xmlParser,
            FormRuntimeBuilder runtimeBuilder,
            RuntimeSerializer runtimeSerializer,
            RuntimeMetadataService runtimeMetadata)
        {
            _formSaveService = formSaveService;
            _env = env;
            _options = options;
            _logger = logger;
            _xmlParser = xmlParser;
            _runtimeBuilder = runtimeBuilder;
            _runtimeSerializer = runtimeSerializer;
            _runtimeMetadata = runtimeMetadata;
        }

        [HttpPost("save")]
        [RequestSizeLimit(50 * 1024 * 1024)]
        [ProducesResponseType(typeof(ApiResponse<object>), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> Save([FromBody] FormDefinition form)
        {
            if (form == null)
            {
                return BadRequest(new ApiErrorResponse
                {
                    Message = "No form definition provided.",
                    ErrorCode = ErrorCodes.InvalidRequest
                });
            }

            try
            {
                string outputDir = Path.Combine(_env.ContentRootPath, "Forms");
                var result = await _formSaveService.SaveAsync(form, outputDir);

                return Ok(new ApiResponse<object>
                {
                    Success = true,
                    Message = $"Form saved successfully. {form.Sheets.Count} sheet(s), {form.Clusters.Count} cluster(s).",
                    Data = new
                    {
                        xmlPath = result.XmlPath,
                        workbookPath = result.WorkbookPath,
                        previewPath = result.PreviewPath,
                        pdfPath = result.PdfPath,
                        databaseObjects = result.DatabaseObjects,
                        relativeBase = "/forms"
                    }
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to save form");
                return StatusCode(500, new ApiErrorResponse
                {
                    Message = $"Failed to save form: {ex.Message}",
                    ErrorCode = ErrorCodes.InternalError
                });
            }
        }

        [HttpPost("from-excel")]
        [RequestSizeLimit(50 * 1024 * 1024)]
        [ProducesResponseType(typeof(ApiResponse<FormDefinition>), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status400BadRequest)]
        public async Task<IActionResult> FromExcel([FromForm] IFormFile? file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new ApiErrorResponse
                {
                    Message = "No file uploaded.",
                    ErrorCode = ErrorCodes.NoFileUploaded
                });
            }

            string extension = Path.GetExtension(file.FileName);
            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".xlsx", ".xls" };
            if (!allowed.Contains(extension))
            {
                return BadRequest(new ApiErrorResponse
                {
                    Message = "Only .xlsx and .xls files are supported.",
                    ErrorCode = ErrorCodes.InvalidFileExtension
                });
            }

            string uploadsFolder = Path.Combine(_env.ContentRootPath, _options.Value.UploadDirectory);
            Directory.CreateDirectory(uploadsFolder);

            string uniqueFileName = $"{Guid.NewGuid():N}{extension}";
            string filePath = Path.Combine(uploadsFolder, uniqueFileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            string templateId = Path.GetFileNameWithoutExtension(uniqueFileName);

            try
            {
                // Use existing service to capture, then convert result to FormDefinition
                // Pass templateId as fileId so the PNG filename matches — single source of truth.
                var captureService = HttpContext.RequestServices.GetRequiredService<Services.Interfaces.IExcelCaptureService>();
                var captureResult = await captureService.CapturePrintAreaAsync(filePath, templateId);

                // Convert capture result to FormDefinition
                var formDefinition = ConvertCaptureToForm(captureResult, file.FileName);

                // Persist the workbook in Forms/ so the Runtime endpoint can find it later
                string formsDir = Path.Combine(_env.ContentRootPath, "Forms");
                Directory.CreateDirectory(formsDir);
                string persistentPath = Path.Combine(formsDir, uniqueFileName);
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Move(filePath, persistentPath, overwrite: true);
                }

                // Preview URL — use the ImageUrl returned by CapturePrintAreaAsync.
                // Because we passed templateId as fileId, this is now /preview/page_{templateId}.png.
                string previewUrl = captureResult.ImageUrl ?? $"/preview/page_{templateId}.png";

                // Persist page setup metadata for Runtime coordinate alignment
                if (captureResult.PageSetup != null)
                {
                    PersistRuntimeMetadata(templateId, captureResult.PageSetup);
                }

                // Persist COM field rectangles as the single source of truth for Runtime overlay.
                // This eliminates OpenXML coordinate recalculation on every GET /api/form/runtime/{id}.
                _runtimeMetadata.Save(captureResult, templateId, formsDir, file.FileName);

                return Ok(new
                {
                    success = true,
                    message = $"Excel file processed. {captureResult.Fields.Count} field(s) detected.",
                    templateId,
                    previewUrl,
                    data = formDefinition
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to process Excel file for form definition");
                // Clean up temp file on error
                try { if (System.IO.File.Exists(filePath)) System.IO.File.Delete(filePath); } catch { }
                return BadRequest(new ApiErrorResponse
                {
                    Message = $"Failed to process Excel file: {ex.Message}",
                    ErrorCode = ErrorCodes.ExcelProcessingError
                });
            }
        }

        private FormDefinition ConvertCaptureToForm(Models.CaptureResult capture, string fileName)
        {
            const double dpi = 300.0;
            const double pointsToPixels = dpi / 72.0;

            int bgWidth = capture.Page?.Width ?? 0;
            int bgHeight = capture.Page?.Height ?? 0;

            // Build page settings from capture debug info, with sensible defaults
            var pageSettings = new PageSettings
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
                Zoom = 100
            };

            if (capture.PageSetup != null)
            {
                pageSettings.WidthPt = capture.PageSetup.PageWidthPt > 0
                    ? capture.PageSetup.PageWidthPt : pageSettings.WidthPt;
                pageSettings.HeightPt = capture.PageSetup.PageHeightPt > 0
                    ? capture.PageSetup.PageHeightPt : pageSettings.HeightPt;
                pageSettings.LeftMargin = capture.PageSetup.LeftMargin > 0
                    ? capture.PageSetup.LeftMargin : pageSettings.LeftMargin;
                pageSettings.TopMargin = capture.PageSetup.TopMargin > 0
                    ? capture.PageSetup.TopMargin : pageSettings.TopMargin;
                pageSettings.CenterHorizontally = capture.PageSetup.CenterHorizontally;
                pageSettings.CenterVertically = capture.PageSetup.CenterVertically;
                pageSettings.Zoom = capture.PageSetup.Zoom > 0
                    ? capture.PageSetup.Zoom : pageSettings.Zoom;
                pageSettings.FitToPagesWide = capture.PageSetup.FitToPagesWide;
                pageSettings.FitToPagesTall = capture.PageSetup.FitToPagesTall;
            }

            // Build print area info from the actual captured page dimensions
            var printArea = new PrintAreaInfo
            {
                Address = "",
                LeftPt = 0,
                TopPt = 0,
                WidthPt = bgWidth > 0 ? bgWidth / pointsToPixels : 0,
                HeightPt = bgHeight > 0 ? bgHeight / pointsToPixels : 0,
                Cols = 0,
                Rows = 0
            };

            // Generate a thumbnail as a base64 data URL (small downscaled version)
            string? thumbnail = null;
            if (!string.IsNullOrEmpty(capture.ImageUrl))
            {
                try
                {
                    string previewFolder = Path.Combine(_env.ContentRootPath,
                        _options.Value.PreviewDirectory);
                    string rawUrl = capture.ImageUrl.TrimStart('/');
                        if (rawUrl.StartsWith("preview/"))
                            rawUrl = rawUrl.Substring("preview/".Length);
                    string previewFileName = Path.GetFileName(rawUrl);
                    string previewPath = Path.Combine(previewFolder, previewFileName);

                    if (System.IO.File.Exists(previewPath))
                    {
                        using var fs = new FileStream(previewPath, FileMode.Open, FileAccess.Read);
                        using var ms = new MemoryStream();
                        fs.CopyTo(ms);
                        byte[] imageBytes = ms.ToArray();

                        // Downscale for thumbnail (max 200px wide)
                        using var srcBitmap = SkiaSharp.SKBitmap.Decode(imageBytes);
                        if (srcBitmap != null)
                        {
                            int thumbW = Math.Min(200, srcBitmap.Width);
                            int thumbH = (int)((double)thumbW / srcBitmap.Width * srcBitmap.Height);
                            using var resized = srcBitmap.Resize(
                                new SkiaSharp.SKImageInfo(thumbW, thumbH),
                                SkiaSharp.SKSamplingOptions.Default);
                            if (resized != null)
                            {
                                using var thumbImg = SkiaSharp.SKImage.FromBitmap(resized);
                                using var thumbData = thumbImg.Encode(
                                    SkiaSharp.SKEncodedImageFormat.Png, 80);
                                byte[] thumbBytes = thumbData.ToArray();
                                thumbnail = "data:image/png;base64," +
                                    Convert.ToBase64String(thumbBytes);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to generate thumbnail for preview image");
                }
            }

            // Persist the preview image to a permanent form-owned copy
            string? persistedBgUrl = capture.ImageUrl;
            if (!string.IsNullOrEmpty(capture.ImageUrl))
            {
                try
                {
                    string formsDir = Path.Combine(_env.ContentRootPath, "Forms");
                    Directory.CreateDirectory(formsDir);

                    string previewFolder = Path.Combine(_env.ContentRootPath,
                        _options.Value.PreviewDirectory);
                    string rawUrl = capture.ImageUrl.TrimStart('/');
                        if (rawUrl.StartsWith("preview/"))
                            rawUrl = rawUrl.Substring("preview/".Length);
                    string previewFileName2 = Path.GetFileName(rawUrl);
                    string srcPath = Path.Combine(previewFolder, previewFileName2);

                    if (System.IO.File.Exists(srcPath))
                    {
                        string bgFileName = $"bg_{Guid.NewGuid():N}.png";
                        string destPath = Path.Combine(formsDir, bgFileName);
                        System.IO.File.Copy(srcPath, destPath, true);
                        persistedBgUrl = $"/forms/{bgFileName}";
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to persist background image, using temporary");
                }
            }

            var form = new FormDefinition
            {
                Workbook = new WorkbookMetadata
                {
                    Title = Path.GetFileNameWithoutExtension(fileName),
                    Author = "",
                    Created = DateTime.Now.ToString("o"),
                    Modified = DateTime.Now.ToString("o"),
                    Version = "1.0",
                    Description = $"Imported from {fileName}"
                },
                Sheets = new List<SheetDefinition>
                {
                    new SheetDefinition
                    {
                        Id = "sheet_0",
                        Name = "Sheet1",
                        Index = 0,
                        PageSettings = pageSettings,
                        PrintArea = printArea,
                        BackgroundImage = persistedBgUrl,
                        BackgroundWidth = bgWidth,
                        BackgroundHeight = bgHeight,
                        Thumbnail = thumbnail,
                        RowHeights = new Dictionary<int, double>(),
                        ColumnWidths = new Dictionary<int, double>(),
                        MergedCells = new List<MergedCellInfo>(),
                        FreezePane = null,
                        CellStyles = new Dictionary<string, CellStyleInfo>(),
                        CellValues = new Dictionary<string, string>()
                    }
                },
                Clusters = capture.Fields.Select(f => new ClusterDefinition
                {
                    ClusterId = f.Id,
                    Name = $"field_{f.Cell}",
                    Type = f.Type.ToLower(),
                    SheetId = "sheet_0",
                    CellAddress = f.Cell,
                    Left = Math.Round(f.Left, 1),
                    Right = Math.Round(f.Left + f.Width, 1),
                    Top = Math.Round(f.Top, 1),
                    Bottom = Math.Round(f.Top + f.Height, 1),
                    LeftPt = f.ExcelLeft,
                    TopPt = f.ExcelTop,
                    WidthPt = f.ExcelWidthPt > 0 ? f.ExcelWidthPt : Math.Round(f.Width / pointsToPixels, 2),
                    HeightPt = f.ExcelHeightPt > 0 ? f.ExcelHeightPt : Math.Round(f.Height / pointsToPixels, 2),
                    InputParameters = new Dictionary<string, string>
                    {
                        ["type"] = f.Type,
                        ["comment"] = f.Comment
                    },
                    Visibility = "visible",
                    Readonly = false,
                    Remarks = f.Comment,
                    Functions = new List<string>(),
                    Metadata = new Dictionary<string, string>
                    {
                        ["isMerged"] = f.IsMerged.ToString(),
                        ["mergeAddress"] = f.MergeAddress ?? ""
                    }
                }).ToList(),
                Images = new List<ImageDefinition>(),
                Metadata = new Dictionary<string, string>
                {
                    ["sourceFile"] = fileName,
                    ["capturedAt"] = DateTime.Now.ToString("o")
                }
            };

            // Store calibration data for frontend debug use
            if (capture.PageSetup != null)
            {
                var debugMeta = form.Metadata;
                debugMeta["printedOriginX"] = capture.PageSetup.PrintedOriginX.ToString("F2");
                debugMeta["printedOriginY"] = capture.PageSetup.PrintedOriginY.ToString("F2");
                debugMeta["actualScaleX"] = capture.PageSetup.ActualScaleX.ToString("F6");
                debugMeta["actualScaleY"] = capture.PageSetup.ActualScaleY.ToString("F6");
                debugMeta["theoreticalScale"] = capture.PageSetup.Scale.ToString("F6");
                debugMeta["pageWidthPt"] = capture.PageSetup.PageWidthPt.ToString("F1");
                debugMeta["pageHeightPt"] = capture.PageSetup.PageHeightPt.ToString("F1");
            }

            // Generate debug overlay image
            GenerateDebugOverlay(form, bgWidth, bgHeight, capture);

            return form;
        }

        /// <summary>
        /// Get the runtime form definition for a previously saved template.
        /// Uses Excel COM geometry from persisted .runtime.json (single source of truth).
        /// Falls back to OpenXML coordinate pipeline only if COM metadata doesn't exist.
        /// Returns JSON consumable by the Next.js frontend for the Yellow Editable Overlay.
        /// </summary>
        /// <param name="templateId">Template ID / XLSX filename (without extension).</param>
        [HttpGet("runtime/{templateId}")]
        [ProducesResponseType(typeof(ApiResponse<RuntimeForm>), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status404NotFound)]
        public IActionResult GetRuntime([FromRoute] string templateId)
        {
            try
            {
                _logger.LogInformation("[RUNTIME] Stage 1: Looking for COM metadata for ID={TemplateId}", templateId);

                string formsDir = Path.Combine(_env.ContentRootPath, "Forms");

                // Try COM metadata first — single source of truth for field coordinates
                var comRuntime = _runtimeMetadata.Load(templateId, formsDir);
                if (comRuntime != null)
                {
                    _logger.LogInformation(
                        "[RUNTIME] COM metadata loaded — {SheetCount} sheet(s), {FieldCount} field(s)",
                        comRuntime.Sheets.Count,
                        comRuntime.Sheets.Sum(s => s.Fields.Count));
                    return Ok(new ApiResponse<RuntimeForm>
                    {
                        Success = true,
                        Message = $"Runtime form loaded from COM metadata: {comRuntime.Sheets.Count} sheet(s), {comRuntime.Sheets.Sum(s => s.Fields.Count)} field(s)",
                        Data = comRuntime
                    });
                }

                // Fallback: OpenXML coordinate pipeline (legacy templates without .runtime.json)
                _logger.LogInformation("[RUNTIME] Fallback: Finding template file for ID={TemplateId}", templateId);

                string xlsxPath = FindTemplateFile(templateId);
                if (xlsxPath == null)
                {
                    _logger.LogWarning("[RUNTIME] Template not found: {TemplateId}", templateId);
                    return NotFound(new ApiErrorResponse
                    {
                        Message = $"Template not found: {templateId}",
                        ErrorCode = ErrorCodes.InvalidRequest
                    });
                }

                _logger.LogInformation("[RUNTIME] Fallback: Parsing workbook at {Path}", xlsxPath);
                var workbook = _xmlParser.Parse(xlsxPath);

                _logger.LogInformation("[RUNTIME] Fallback: Parsed {Count} sheet(s)", workbook.Sheets.Count);
                if (workbook.Sheets.Count == 0)
                {
                    return NotFound(new ApiErrorResponse
                    {
                        Message = "No sheets found in workbook",
                        ErrorCode = ErrorCodes.InvalidRequest
                    });
                }

                _logger.LogInformation("[RUNTIME] Fallback: Loading runtime metadata for coordinate alignment");
                var (originXPt, originYPt, actualScaleX, actualScaleY) = LoadRuntimeMetadata(templateId);

                _logger.LogInformation(
                    "[RUNTIME] Fallback: Using origin=({OX:F1},{OY:F1})pt scale=({SX:F6},{SY:F6})px/pt",
                    originXPt, originYPt, actualScaleX, actualScaleY);

                int dpi = actualScaleX > 0 ? (int)Math.Round(actualScaleX * 72.0) : 300;
                var runtimeForm = _runtimeBuilder.Build(workbook, dpi, originXPt, originYPt);

                _logger.LogInformation("[RUNTIME] Fallback: Runtime built — {SheetCount} sheet(s), {FieldCount} field(s)",
                    runtimeForm.Sheets.Count,
                    runtimeForm.Sheets.Sum(s => s.Fields.Count));

                return Ok(new ApiResponse<RuntimeForm>
                {
                    Success = true,
                    Message = $"Runtime form built: {runtimeForm.Sheets.Count} sheet(s), {runtimeForm.Sheets.Sum(s => s.Fields.Count)} field(s)",
                    Data = runtimeForm
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[RUNTIME] FAILED at stage: {Message}\n{StackTrace}",
                    ex.Message, ex.ToString());
                return StatusCode(500, new
                {
                    success = false,
                    message = $"Failed to build runtime form: {ex.Message}",
                    errorCode = ErrorCodes.InternalError,
                    exceptionType = ex.GetType().FullName,
                    stackTrace = ex.ToString()
                });
            }
        }

        /// <summary>
        /// Find a template XLSX file by ID, searching known directories.
        /// </summary>
        private string? FindTemplateFile(string templateId)
        {
            string formsDir = Path.Combine(_env.ContentRootPath, "Forms");
            string uploadsDir = Path.Combine(_env.ContentRootPath, _options.Value.UploadDirectory);

            // Try exact filename match
            string[] searchDirs = [formsDir, uploadsDir];
            foreach (var dir in searchDirs)
            {
                if (!Directory.Exists(dir)) continue;

                var files = Directory.GetFiles(dir, "*.xlsx");
                foreach (var file in files)
                {
                    string name = Path.GetFileNameWithoutExtension(file);
                    if (string.Equals(name, templateId, StringComparison.OrdinalIgnoreCase))
                        return file;
                }
            }

            // Try prefix match (templateId might be a GUID prefix)
            foreach (var dir in searchDirs)
            {
                if (!Directory.Exists(dir)) continue;

                var files = Directory.GetFiles(dir, "*.xlsx");
                foreach (var file in files)
                {
                    string name = Path.GetFileNameWithoutExtension(file);
                    if (name.StartsWith(templateId, StringComparison.OrdinalIgnoreCase))
                        return file;
                }
            }

            return null;
        }

        /// <summary>
        /// Persists the page setup metadata as a JSON file alongside the template XLSX.
        /// This data is consumed by GetRuntime() to align Runtime overlays with the PNG background.
        /// </summary>
        private void PersistRuntimeMetadata(string templateId, PageSetupDebug pageSetup)
        {
            try
            {
                string formsDir = Path.Combine(_env.ContentRootPath, "Forms");
                Directory.CreateDirectory(formsDir);
                string metaPath = Path.Combine(formsDir, $"{templateId}.meta.json");

                var metadata = new
                {
                    pageSetup.PrintedOriginX,
                    pageSetup.PrintedOriginY,
                    pageSetup.LeftMargin,
                    pageSetup.TopMargin,
                    pageSetup.CenterHorizontally,
                    pageSetup.CenterVertically,
                    pageSetup.PageWidthPt,
                    pageSetup.PageHeightPt,
                    pageSetup.PrintAreaWidthPt,
                    pageSetup.PrintAreaHeightPt,
                    pageSetup.ActualScaleX,
                    pageSetup.ActualScaleY,
                    pageSetup.Scale,
                    pageSetup.Zoom,
                    pageSetup.FitToPagesWide,
                    pageSetup.FitToPagesTall
                };

                string json = System.Text.Json.JsonSerializer.Serialize(metadata,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = false });
                System.IO.File.WriteAllText(metaPath, json);
                _logger.LogInformation("[RUNTIME] Metadata saved: {Path}", metaPath);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[RUNTIME] Failed to persist runtime metadata");
            }
        }

        /// <summary>
        /// Loads runtime metadata saved during upload.
        /// Returns origin offsets in points and actual pixel scales.
        /// Falls back to (0,0,0,0) if no metadata file exists (legacy templates).
        /// </summary>
        private (double originXPt, double originYPt, double actualScaleX, double actualScaleY) LoadRuntimeMetadata(string templateId)
        {
            try
            {
                string formsDir = Path.Combine(_env.ContentRootPath, "Forms");
                string metaPath = Path.Combine(formsDir, $"{templateId}.meta.json");

                if (!System.IO.File.Exists(metaPath))
                {
                    _logger.LogInformation("[RUNTIME] No metadata file found at {Path}, using defaults", metaPath);
                    return (0, 0, 0, 0);
                }

                string json = System.IO.File.ReadAllText(metaPath);
                using var doc = System.Text.Json.JsonDocument.Parse(json);
                var root = doc.RootElement;

                double printedOriginX = root.TryGetProperty("printedOriginX", out var ox) ? ox.GetDouble() : 0;
                double printedOriginY = root.TryGetProperty("printedOriginY", out var oy) ? oy.GetDouble() : 0;
                double actualScaleX = root.TryGetProperty("actualScaleX", out var sx) ? sx.GetDouble() : 0;
                double actualScaleY = root.TryGetProperty("actualScaleY", out var sy) ? sy.GetDouble() : 0;

                // Convert pixel origin back to points using actual scale
                double originXPt = actualScaleX > 0 ? printedOriginX / actualScaleX : 0;
                double originYPt = actualScaleY > 0 ? printedOriginY / actualScaleY : 0;

                _logger.LogInformation(
                    "[RUNTIME] Loaded metadata: printedOrigin=({OX:F1},{OY:F1})px scale=({SX:F6},{SY:F6}) -> originPt=({XP:F1},{YP:F1})",
                    printedOriginX, printedOriginY, actualScaleX, actualScaleY, originXPt, originYPt);

                return (originXPt, originYPt, actualScaleX, actualScaleY);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[RUNTIME] Failed to load runtime metadata, using defaults");
                return (0, 0, 0, 0);
            }
        }

        /// <summary>
        /// Generates a debug PNG with overlay annotations to visually verify coordinate alignment.
        /// Red = computed cluster rectangle, green = detected cell boundary from pixel scan,
        /// blue dot = printed origin. Saved alongside the form background.
        /// </summary>
        private void GenerateDebugOverlay(FormDefinition form, int bgWidth, int bgHeight, Models.CaptureResult capture)
        {
            if (form.Sheets.Count == 0) return;
            var sheet = form.Sheets[0];
            if (string.IsNullOrEmpty(sheet.BackgroundImage)) return;

            try
            {
                string formsDir = Path.Combine(_env.ContentRootPath, "Forms");
                string bgFile = Path.GetFileName(sheet.BackgroundImage.TrimStart('/'));
                string bgPath = Path.Combine(formsDir, bgFile);
                if (!System.IO.File.Exists(bgPath)) return;

                byte[] imageBytes = System.IO.File.ReadAllBytes(bgPath);
                using var bitmap = SkiaSharp.SKBitmap.Decode(imageBytes);
                if (bitmap == null) return;

                using var canvas = new SkiaSharp.SKCanvas(bitmap);

                using var redPen = new SkiaSharp.SKPaint
                {
                    Color = new SkiaSharp.SKColor(255, 0, 0, 180),
                    Style = SkiaSharp.SKPaintStyle.Stroke,
                    StrokeWidth = 2
                };
                using var greenPen = new SkiaSharp.SKPaint
                {
                    Color = new SkiaSharp.SKColor(0, 200, 0, 180),
                    Style = SkiaSharp.SKPaintStyle.Stroke,
                    StrokeWidth = 2
                };
                using var blueDot = new SkiaSharp.SKPaint
                {
                    Color = new SkiaSharp.SKColor(0, 100, 255, 200),
                    Style = SkiaSharp.SKPaintStyle.Fill
                };
                using var labelFont = new SkiaSharp.SKFont(SkiaSharp.SKTypeface.Default, 14);
                using var labelPaint = new SkiaSharp.SKPaint
                {
                    Color = new SkiaSharp.SKColor(255, 0, 0, 220),
                    IsAntialias = true
                };

                // Draw printed origin (blue dot)
                double ox = capture.PageSetup?.PrintedOriginX ?? 0;
                double oy = capture.PageSetup?.PrintedOriginY ?? 0;
                canvas.DrawCircle((float)ox, (float)oy, 6, blueDot);

                // Draw each cluster
                foreach (var cluster in form.Clusters)
                {
                    float x = (float)cluster.Left;
                    float y = (float)cluster.Top;
                    float w = (float)(cluster.Right - cluster.Left);
                    float h = (float)(cluster.Bottom - cluster.Top);

                    // Red: computed rectangle
                    canvas.DrawRect(x, y, w, h, redPen);

                    // Label
                    string label = $"{cluster.CellAddress} ({cluster.Left:F0},{cluster.Top:F0}) {w:F0}x{h:F0}";
                    canvas.DrawText(label, x, y - 4, SkiaSharp.SKTextAlign.Left, labelFont, labelPaint);

                    // Log per-cluster diagnostic
                    _logger.LogInformation(
                        "[DEBUG] Cluster \"{Cell}\": computed=({L:F1},{T:F1},{W:F1},{H:F1}) " +
                        "origin=({OX:F1},{OY:F1}) " +
                        "cellPt=({CL:F1},{CT:F1}) wPt={WPt:F1} hPt={HPt:F1}",
                        cluster.CellAddress,
                        cluster.Left, cluster.Top, w, h,
                        ox, oy,
                        cluster.LeftPt, cluster.TopPt,
                        cluster.WidthPt, cluster.HeightPt);
                }

                // Save debug overlay
                string debugFile = $"debug_{bgFile}";
                string debugPath = Path.Combine(formsDir, debugFile);
                using var debugImage = SkiaSharp.SKImage.FromBitmap(bitmap);
                using var debugData = debugImage.Encode(SkiaSharp.SKEncodedImageFormat.Png, 95);
                System.IO.File.WriteAllBytes(debugPath, debugData.ToArray());
                _logger.LogInformation("[DEBUG] Overlay saved: {Path}", debugPath);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[DEBUG] Failed to generate debug overlay");
            }
        }
    }
}