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
        private readonly RuntimeCoordinateGenerator _runtimeGenerator;
        private readonly PythonRenderService _pythonRender;
        private readonly WorkbookReaderService _workbookReaderService;

        public FormController(
            IFormSaveService formSaveService,
            IWebHostEnvironment env,
            IOptions<ExcelCaptureOptions> options,
            ILogger<FormController> logger,
            OpenXmlParser xmlParser,
            FormRuntimeBuilder runtimeBuilder,
            RuntimeSerializer runtimeSerializer,
            RuntimeCoordinateGenerator runtimeGenerator,
            PythonRenderService pythonRender,
            WorkbookReaderService workbookReaderService)
        {
            _formSaveService = formSaveService;
            _env = env;
            _options = options;
            _logger = logger;
            _xmlParser = xmlParser;
            _runtimeBuilder = runtimeBuilder;
            _runtimeSerializer = runtimeSerializer;
            _runtimeGenerator = runtimeGenerator;
            _pythonRender = pythonRender;
            _workbookReaderService = workbookReaderService;
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

                // Determine the persistent background image URL
                // Priority: Forms copy (persistent) > preview (temporary)
                string bgUrl = formDefinition?.Sheets?.FirstOrDefault()?.BackgroundImage
                    ?? previewUrl;

                // Persist COM field rectangles as the single source of truth for Runtime overlay.
                // This eliminates OpenXML coordinate recalculation on every GET /api/form/runtime/{id}.
                // Pass the background image URL so the frontend can load the exact image used
                // during coordinate computation.
                _runtimeGenerator.SaveMetadata(captureResult, templateId, formsDir, file.FileName, bgUrl);

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

        [HttpPost("upload-preview")]
        [RequestSizeLimit(50 * 1024 * 1024)]
        public async Task<IActionResult> UploadPreview([FromForm] IFormFile? file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { success = false, message = "No file uploaded." });
            }

            string extension = Path.GetExtension(file.FileName);
            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".xlsx", ".xls" };
            if (!allowed.Contains(extension))
            {
                return BadRequest(new { success = false, message = "Only .xlsx and .xls files are supported." });
            }

            string uploadsDir = Path.Combine(_env.ContentRootPath, _options.Value.UploadDirectory);
            Directory.CreateDirectory(uploadsDir);
            string uniqueId = Guid.NewGuid().ToString("N");
            string xlsxPath = Path.Combine(uploadsDir, $"{uniqueId}{extension}");

            await using (var stream = new FileStream(xlsxPath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            try
            {
                string previewDir = Path.Combine(_env.ContentRootPath, _options.Value.PreviewDirectory);
                Directory.CreateDirectory(previewDir);

                var pythonResult = await _pythonRender.UploadPreviewAsync(xlsxPath, previewDir);

                if (pythonResult == null || !pythonResult.Success)
                {
                    return StatusCode(502, new { success = false, message = "Python preview service returned an error or is unavailable." });
                }

                // Map multi-page response
                var pageResults = (pythonResult.Pages ?? new List<PythonPreviewPageResult>())
                    .Select((p, pageIndex) =>
                    {
                        int fi = 0;
                        return new
                        {
                            sheetName = p.SheetName,
                            backgroundImage = $"/preview/{p.BackgroundImage}",
                            page = new { width = p.Page?.Width ?? 2550, height = p.Page?.Height ?? 3300 },
                            fields = (p.Fields ?? new List<PythonPreviewField>())
                                .Select(f =>
                                {
                                    string id = $"p{pageIndex + 1}f{++fi}";
                                    string name = string.IsNullOrWhiteSpace(f.Name) ? id : f.Name;
                                    return (object)new
                                    {
                                        id,
                                        name,
                                        type = f.Type,
                                        cellAddr = f.CellAddr,
                                        left_ratio = f.LeftRatio,
                                        top_ratio = f.TopRatio,
                                        right_ratio = f.RightRatio,
                                        bottom_ratio = f.BottomRatio,
                                    };
                                }).ToList()
                        };
                    })
                    .ToList();

                return Ok(new { success = true, pages = pageResults });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Upload preview failed");
                return StatusCode(500, new { success = false, message = $"Preview failed: {ex.Message}" });
            }
            finally
            {
                try { if (System.IO.File.Exists(xlsxPath)) System.IO.File.Delete(xlsxPath); } catch { }
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

            return form;
        }

        /// <summary>
        /// Upload an Output Excel workbook and reconstruct the FormDefinition.
        /// This enables the round-trip workflow: Generate → Upload → Edit → Regenerate.
        /// Reads _Fields sheet, cell comments, page setup, merged cells, styles, and cell values.
        /// </summary>
        [HttpPost("upload-excel")]
        [RequestSizeLimit(50 * 1024 * 1024)]
        [ProducesResponseType(typeof(ApiResponse<object>), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> UploadExcel([FromForm] IFormFile? file)
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
            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".xlsx" };
            if (!allowed.Contains(extension))
            {
                return BadRequest(new ApiErrorResponse
                {
                    Message = "Only .xlsx files are supported for upload-excel.",
                    ErrorCode = ErrorCodes.InvalidFileExtension
                });
            }

            // Save uploaded file to temp location
            string uploadsDir = Path.Combine(_env.ContentRootPath, "Uploads");
            Directory.CreateDirectory(uploadsDir);
            string tempPath = Path.Combine(uploadsDir, $"{Guid.NewGuid():N}.xlsx");

            await using (var stream = new FileStream(tempPath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            try
            {
                var formDefinition = _workbookReaderService.Read(tempPath, file.FileName);

                if (formDefinition == null)
                {
                    return BadRequest(new ApiErrorResponse
                    {
                        Message = "Failed to read workbook: file may be corrupt or not a valid Output Excel workbook.",
                        ErrorCode = ErrorCodes.ExcelProcessingError
                    });
                }

                if (formDefinition.Sheets.Count == 0)
                {
                    return BadRequest(new ApiErrorResponse
                    {
                        Message = "No content sheets found in the workbook.",
                        ErrorCode = ErrorCodes.ExcelProcessingError
                    });
                }

                if (formDefinition.Clusters.Count == 0)
                {
                    return BadRequest(new ApiErrorResponse
                    {
                        Message = "No field metadata found in workbook: missing both _Fields sheet and cell comments.",
                        ErrorCode = ErrorCodes.ExcelProcessingError
                    });
                }

                // Persist the workbook to Forms/ for round-trip compatibility
                string formsDir = Path.Combine(_env.ContentRootPath, "Forms");
                Directory.CreateDirectory(formsDir);
                string persistentFileName = $"{Guid.NewGuid():N}.xlsx";
                string persistentPath = Path.Combine(formsDir, persistentFileName);
                System.IO.File.Copy(tempPath, persistentPath, overwrite: true);

                return Ok(new ApiResponse<object>
                {
                    Success = true,
                    Message = $"Workbook read successfully: {formDefinition.Sheets.Count} sheet(s), {formDefinition.Clusters.Count} field(s).",
                    Data = new
                    {
                        formDefinition,
                        fieldCount = formDefinition.Clusters.Count,
                        sheetCount = formDefinition.Sheets.Count,
                        templateId = Path.GetFileNameWithoutExtension(persistentFileName),
                        workbookDownloadUrl = $"/forms/{persistentFileName}"
                    }
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to read uploaded workbook");
                return StatusCode(500, new ApiErrorResponse
                {
                    Message = $"Failed to read workbook: {ex.Message}",
                    ErrorCode = ErrorCodes.InternalError
                });
            }
            finally
            {
                try { if (System.IO.File.Exists(tempPath)) System.IO.File.Delete(tempPath); } catch { }
            }
        }

        /// <summary>
        /// Output Excel of Form — generates a complete Excel workbook and returns it
        /// directly as a browser download (legacy ConMas behavior).
        ///
        /// The workbook is generated with:
        ///   - Worksheet structure, merged cells, styles, page setup
        ///   - User-entered cell values (from SheetDefinition.CellValues)
        ///   - Cell comments with field metadata (legacy ConMas format)
        ///   - Hidden _Fields worksheet for republish compatibility
        ///
        /// Returns the workbook as a direct file download — no JSON wrapper, no
        /// intermediate download URL. The content-type is set so the browser
        /// automatically triggers its normal save/download behavior.
        ///
        /// On error, returns a JSON error response.
        /// </summary>
        [HttpPost("output-excel")]
        [RequestSizeLimit(50 * 1024 * 1024)]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> OutputExcel([FromBody] FormDefinition form)
        {
            if (form == null)
            {
                return BadRequest(new ApiErrorResponse
                {
                    Message = "No form definition provided.",
                    ErrorCode = ErrorCodes.InvalidRequest
                });
            }

            string outputDir = Path.Combine(_env.ContentRootPath, "Output");
            Directory.CreateDirectory(outputDir);

            string tempPath = Path.Combine(outputDir, $"temp_{Guid.NewGuid():N}.xlsx");

            try
            {
                // Generate workbook to temp file
                var result = await _formSaveService.OutputExcelAsync(form, outputDir);

                tempPath = result.WorkbookPath; // OutputExcelAsync already generates a unique path

                // Read bytes before any cleanup
                byte[] fileBytes = await System.IO.File.ReadAllBytesAsync(tempPath);

                // Extract a user-friendly filename from the workbook title
                string safeFileName = SanitizeFileName(form.Workbook?.Title ?? "output") + ".xlsx";

                _logger.LogInformation(
                    "Output Excel generated: {SheetCount} sheet(s), {ClusterCount} cluster(s), {Size} bytes. Returning as direct download: {FileName}",
                    form.Sheets.Count, form.Clusters.Count, fileBytes.Length, safeFileName);

                return File(
                    fileBytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    safeFileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to generate Output Excel");
                return StatusCode(500, new ApiErrorResponse
                {
                    Message = $"Failed to generate Output Excel: {ex.Message}",
                    ErrorCode = ErrorCodes.InternalError
                });
            }
            finally
            {
                // Clean up temp file
                try { if (System.IO.File.Exists(tempPath)) System.IO.File.Delete(tempPath); } catch { }
            }
        }

        /// <summary>
        /// Sanitize a string for use as a filename — remove or replace invalid characters.
        /// </summary>
        private static string SanitizeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var sanitized = new System.Text.StringBuilder(name.Length);
            foreach (char c in name)
            {
                sanitized.Append(invalid.Contains(c) ? '_' : c);
            }
            string result = sanitized.ToString().Trim();
            return string.IsNullOrWhiteSpace(result) ? "output" : result;
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
                var comRuntime = _runtimeGenerator.LoadMetadata(templateId, formsDir);
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
        /// Loads runtime metadata using the legacy coordinate algorithm.
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

                double originXPt = actualScaleX > 0 ? printedOriginX / actualScaleX : 0;
                double originYPt = actualScaleY > 0 ? printedOriginY / actualScaleY : 0;

                return (originXPt, originYPt, actualScaleX, actualScaleY);
            }
            catch
            {
                return (0, 0, 0, 0);
            }
        }
    }
}