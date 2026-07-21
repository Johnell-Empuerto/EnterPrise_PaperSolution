using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using ExcelAPI.Models;
using ExcelAPI.Application;
using ExcelAPI.Designer.Analysis;
using ExcelAPI.Runtime;
using ExcelAPI.Rendering;
using SkiaSharp;
using WbDef = ExcelAPI.Models.WorkbookDefinition;

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
        private readonly RuntimeCoordinateGenerator _runtimeGenerator;
        private readonly PythonRenderService _pythonRender;
        private readonly WorkbookReaderService _workbookReaderService;
        private readonly ISessionWorkbookStore _sessionStore;
        private readonly IDesignerModelReader _designerModelReader;

        public FormController(
            IFormSaveService formSaveService,
            IWebHostEnvironment env,
            IOptions<ExcelCaptureOptions> options,
            ILogger<FormController> logger,
            RuntimeCoordinateGenerator runtimeGenerator,
            PythonRenderService pythonRender,
            WorkbookReaderService workbookReaderService,
            ISessionWorkbookStore sessionStore,
            IDesignerModelReader designerModelReader)
        {
            _formSaveService = formSaveService;
            _env = env;
            _options = options;
            _logger = logger;
            _runtimeGenerator = runtimeGenerator;
            _pythonRender = pythonRender;
            _workbookReaderService = workbookReaderService;
            _sessionStore = sessionStore;
            _designerModelReader = designerModelReader;
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
                // Capture via COM — always produces InternalWorkbookDefinition (Stage 9)
                var captureService = HttpContext.RequestServices.GetRequiredService<Designer.Capture.IExcelCaptureService>();
                var captureResult = await captureService.CapturePrintAreaAsync(filePath, templateId);

                // Convert to FormDefinition via WbDef canonical converter (always available)
                var formDefinition = WbDef.WorkbookDefinitionConverter.ToFormDefinition(
                    captureResult.InternalWorkbookDefinition!);

                // Apply IO-specific operations (thumbnail, bg persistence)
                if (captureResult.InternalWorkbookDefinition != null)
                    ApplyFormDefinitionIO(formDefinition, captureResult);

                // Phase 5.2: Save the original workbook into the session store.
                // The browser will only receive a sessionId — no filenames.
                var session = _sessionStore.CreateSession(filePath, file.FileName);
                string sessionId = session.SessionId;

                // Save runtime.json alongside the original workbook in the session folder
                string sessionDir = Path.GetDirectoryName(session.WorkbookPath) ?? Path.Combine(_env.ContentRootPath, "Forms");
                int pngWidth = captureResult.Page?.Width ?? 0;
                int pngHeight = captureResult.Page?.Height ?? 0;
                _runtimeGenerator.SaveFromWbDef(captureResult, sessionId, sessionDir, file.FileName,
                    pngWidth, pngHeight, formDefinition?.Sheets?.FirstOrDefault()?.BackgroundImage ?? captureResult.ImageUrl ?? "");

                // Clean up the temp upload file — the session store has a copy
                try { if (System.IO.File.Exists(filePath)) System.IO.File.Delete(filePath); } catch { }

                // Preview URL
                string previewUrl = captureResult.ImageUrl ?? $"/preview/page_{templateId}.png";

                return Ok(new
                {
                    success = true,
                    message = $"Excel file processed. {captureResult.Fields.Count} field(s) detected.",
                    sessionId,
                    previewUrl,
                    data = formDefinition
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to process Excel file for form definition");
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

                // Phase 5.2: Save the original workbook into the session store.
                var session = _sessionStore.CreateSession(xlsxPath, file.FileName);
                string sessionId = session.SessionId;

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

                // ═════════════════════════════════════════════════════════
                // Read PaperLessConfig from the workbook (if present)
                // This restores field identity, style, and configuration
                // on re-upload without requiring designerModel.
                // ═════════════════════════════════════════════════════════
                object? paperLessConfig = null;
                try
                {
                    if (System.IO.File.Exists(xlsxPath))
                    {
                        using var plPkg = System.IO.Compression.ZipFile.OpenRead(xlsxPath);
                        var plEntry = plPkg.GetEntry("xl/worksheets/paperlessconfig.xml");
                        if (plEntry != null)
                        {
                            string plXml;
                            using (var plReader = new StreamReader(plEntry.Open()))
                                plXml = plReader.ReadToEnd();

                            var b1Match = System.Text.RegularExpressions.Regex.Match(plXml,
                                @"<c\s+r=""B1""[^>]*>.*?<is><t[^>]*>(.*?)</t></is></c>",
                                System.Text.RegularExpressions.RegexOptions.Singleline);
                            if (b1Match.Success)
                            {
                                string plJson = System.Net.WebUtility.HtmlDecode(b1Match.Groups[1].Value);
                                paperLessConfig = System.Text.Json.JsonSerializer.Deserialize<object>(plJson);
                                _logger.LogInformation("[UPLOAD-PREVIEW] PaperLessConfig read and returned in response");
                            }
                        }
                    }
                }
                catch (Exception plEx)
                {
                    _logger.LogWarning(plEx, "[UPLOAD-PREVIEW] Failed to read PaperLessConfig (non-fatal)");
                }

                return Ok(new
                {
                    success = true,
                    sessionId,
                    pages = pageResults,
                    paperLessConfig
                });
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

        /// <summary>
        /// Apply IO-specific operations on the FormDefinition: thumbnail generation
        /// and background image persistence. These are side-effect operations that
        /// the WbDef converter cannot perform (it is a pure data mapping).
        /// </summary>
        private void ApplyFormDefinitionIO(FormDefinition form, Models.CaptureResult capture)
        {
            if (form.Sheets.Count == 0) return;

            var sheet = form.Sheets[0];
            int bgWidth = capture.Page?.Width ?? 0;
            int bgHeight = capture.Page?.Height ?? 0;
            sheet.BackgroundWidth = bgWidth;
            sheet.BackgroundHeight = bgHeight;

            // Generate a thumbnail as a base64 data URL
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
            sheet.Thumbnail = thumbnail;

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
            sheet.BackgroundImage = persistedBgUrl;
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

                // Phase 5.2: Save the workbook into the session store.
                var session = _sessionStore.CreateSession(tempPath, file.FileName);
                string sessionId = session.SessionId;

                // Phase 21: Use DesignerModelReader to reconstruct the complete designer state
                // directly from the workbook. This is the canonical deserialization path:
                // Workbook → DesignerModelReader → DesignerModel → Frontend
                var designerModel = _designerModelReader.Read(tempPath, file.FileName, sessionId);

                // ═════════════════════════════════════════════════════════
                // PHASE 22.1 — STAGE 25: STYLE RE-READ (Upload → DesignerModel)
                // ═════════════════════════════════════════════════════════
                if (designerModel != null)
                {
                    _logger.LogInformation("========================================================");
                    _logger.LogInformation("  STAGE 25 — Style Re-read");
                    _logger.LogInformation("========================================================");

                    int reStageCount = 0;
                    foreach (var rePage in designerModel.Pages)
                    {
                        _logger.LogInformation("  Page: {Name}", rePage.Name);
                        foreach (var reField in rePage.Fields)
                        {
                            string reCell = reField.CellAddress ?? "?";
                            string fontName = reField.Style?.FontFamily ?? "(default)";
                            string fontSize = reField.Style?.FontSize?.ToString() ?? "(default)";
                            string boldStr = reField.Style?.Bold == true ? "true" : "false";
                            string italicStr = reField.Style?.Italic == true ? "true" : "false";
                            string fontColor = reField.Style?.FontColor ?? "(default)";
                            string fillColor = reField.Style?.FillColor ?? "(none)";
                            string hAlign = reField.Style?.HorizontalAlignment ?? "(default)";
                            string vAlign = reField.Style?.VerticalAlignment ?? "(default)";
                            string wrapStr = reField.Style?.WrapText == true ? "true" : "false";

                            _logger.LogInformation("  ------------------------------");
                            _logger.LogInformation($"  Cell:        {reCell}");
                            _logger.LogInformation($"  ✓ Font:       {fontName}");
                            _logger.LogInformation($"  ✓ Size:       {fontSize}");
                            _logger.LogInformation($"  ✓ Bold:       {boldStr}");
                            _logger.LogInformation($"  ✓ Italic:     {italicStr}");
                            _logger.LogInformation($"  ✓ Color:      {fontColor}");
                            _logger.LogInformation($"  ✓ Fill:       {fillColor}");
                            _logger.LogInformation($"  ✓ H-Align:    {hAlign}");
                            _logger.LogInformation($"  ✓ V-Align:    {vAlign}");
                            _logger.LogInformation($"  ✓ Wrap:       {wrapStr}");
                            _logger.LogInformation($"  Field ID:    {reField.Id}");
                            reStageCount++;
                        }
                    }

                    _logger.LogInformation("========================================================");
                    _logger.LogInformation($"  STAGE 25 COMPLETE: {reStageCount} fields re-read");
                    _logger.LogInformation("========================================================");
                }

                // ═════════════════════════════════════════════════════════
                // PAPERLESS DEBUG STAGE 14 — Final DesignerModel Response
                // ═════════════════════════════════════════════════════════
                _logger.LogInformation("========================================================");
                _logger.LogInformation("PAPERLESS DEBUG STAGE 14 — Final DesignerModel Response");
                _logger.LogInformation("========================================================");
                if (designerModel != null)
                {
                    foreach (var p in designerModel.Pages)
                    {
                        _logger.LogInformation("  Page: {Name}", p.Name);
                        foreach (var f in p.Fields)
                        {
                            _logger.LogInformation("    Field ID: {Id}", f.Id);
                            _logger.LogInformation("    Cell: {Cell}", f.CellAddress ?? "?");
                            _logger.LogInformation("    FontSize: {Sz}", f.Style?.FontSize ?? 0);
                            _logger.LogInformation("    FontFamily: {Ff}", f.Style?.FontFamily ?? "(default)");
                            _logger.LogInformation("    Bold: {B}", f.Style?.Bold ?? false);
                        }
                    }
                }
                else
                {
                    _logger.LogInformation("  designerModel IS NULL");
                }
                _logger.LogInformation("========================================================");

                return Ok(new ApiResponse<object>
                {
                    Success = true,
                    Message = $"Workbook read successfully: {formDefinition.Sheets.Count} sheet(s), {formDefinition.Clusters.Count} field(s).",
                    Data = new
                    {
                        formDefinition,
                        sessionId,
                        fieldCount = formDefinition.Clusters.Count,
                        sheetCount = formDefinition.Sheets.Count,
                        // Phase 21: Return DesignerModel for frontend reconstruction
                        designerModel
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
        /// directly as a browser download.
        ///
        /// Phase 4.6: DEPRECATED. The canonical save path is POST /api/form/save-edited
        /// which uses WorkbookValueWriter instead of the legacy COM WorkbookGenerator.
        /// This endpoint exists only for backward compatibility with old clients.
        /// </summary>
        [Obsolete("Use POST /api/form/save-edited instead. This endpoint is deprecated and will return 410 Gone.")]
        [HttpPost("output-excel")]
        [ApiExplorerSettings(IgnoreApi = true)]
        [ProducesResponseType(StatusCodes.Status410Gone)]
        public IActionResult OutputExcel()
        {
            _logger.LogWarning("[DEPRECATED] POST /api/form/output-excel called — use POST /api/form/save-edited instead.");
            return StatusCode(410, new ApiErrorResponse
            {
                Message = "This endpoint is deprecated. Use POST /api/form/save-edited instead.",
                ErrorCode = ErrorCodes.InvalidRequest
            });
        }

        /// <summary>
        /// Save edited field values back into the original Excel workbook (Phase 4.4 / 5.2).
        ///
        /// Accepts a WorkbookDefinition with edited field values and sessionId.
        /// The sessionId is used to resolve the original workbook from the server's
        /// session store (TempWorkbooks/{sessionId}/original.xlsx).
        ///
        /// The browser never needs to track filenames — the server owns the workbook.
        ///
        /// Pipeline:
        ///   WorkbookDefinition + sessionId (from frontend)
        ///     ↓
        ///   SessionWorkbookStore.ResolveWorkbookPath(sessionId)
        ///     ↓
        ///   WorkbookValueWriter.WriteValues()  ← CANONICAL PATH
        ///     ↓
        ///   WorkbookDiffValidator.Compare()    ← AUTO-VALIDATION
        ///     ↓
        ///   Edited workbook download
        /// </summary>
        [HttpPost("save-edited")]
        [RequestSizeLimit(50 * 1024 * 1024)]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> SaveEdited([FromBody] WbDef.WorkbookDefinition wbDef)
        {
            // ═════════════════════════════════════════════════════════
            // PHASE 21.4 — STAGE 1: RAW EXPORT REQUEST
            // ═════════════════════════════════════════════════════════
            _logger.LogInformation("=========================================================");
            _logger.LogInformation("RAW EXPORT REQUEST");
            _logger.LogInformation("=========================================================");
            _logger.LogInformation("Content-Type: {Ct}", Request.ContentType ?? "(not set)");
            _logger.LogInformation("Content-Length: {Len} bytes", Request.ContentLength ?? -1);
            _logger.LogInformation("Request Path: {Path}", Request.Path);
            _logger.LogInformation("Method: {Method}", Request.Method);
            _logger.LogInformation("Session cookie: {Sess}", Request.Cookies.ContainsKey("session") ? "present" : "not set");

            // PHASE 21.4 — STAGE 2: Raw JSON Body
            // NOTE: [FromBody] consumes the request stream before this method executes.
            // To capture the raw JSON before deserialization, an IAsyncActionFilter
            // must be registered. See note in Phase 21.4 Stage 2 of the report.
            _logger.LogInformation("PHASE 21.4 STAGE 2 — Raw JSON Body");
            _logger.LogInformation("  [INFO] Raw JSON cannot be captured here because [FromBody]");
            _logger.LogInformation("  [INFO] has already consumed the request stream.");
            _logger.LogInformation("  [INFO] To capture raw JSON, register an IAsyncActionFilter");
            _logger.LogInformation("  [INFO] that reads Request.Body before model binding.");
            _logger.LogInformation("  [INFO] Alternatively, enable client-side logging of the");
            _logger.LogInformation("  [INFO] POST body before the fetch() call.");
            _logger.LogInformation("=========================================================");

            if (wbDef == null)
            {
                _logger.LogError("PHASE 21.4 STAGE 3 — MODEL BINDING RESULT: wbDef IS NULL!");
                _logger.LogError("This means the JSON payload could not be deserialized.");
                _logger.LogError("Check: Content-Type must be application/json; charset=UTF-8");
                return BadRequest(new ApiErrorResponse
                {
                    Message = "No workbook definition provided.",
                    ErrorCode = ErrorCodes.InvalidRequest
                });
            }

            // ═════════════════════════════════════════════════════════
            // PAPERLESS DEBUG STAGE 2 — Request Payload Received (Model Binding)
            // ═════════════════════════════════════════════════════════
            _logger.LogInformation("=========================================================");
            _logger.LogInformation("PAPERLESS DEBUG STAGE 2 — Request Payload Received");
            _logger.LogInformation("=========================================================");
            _logger.LogInformation("wbDef != null: {NotNull}", wbDef != null);
            if (wbDef != null)
            {
                _logger.LogInformation("SessionId: {Sid}", wbDef.SessionId ?? "(null/empty)");
                _logger.LogInformation("SourceFileName: {File}", wbDef.SourceFileName ?? "(null/empty)");
                _logger.LogInformation("SourcePath: {Path}", wbDef.SourcePath ?? "(null/empty)");
                _logger.LogInformation("Info: {Info}", wbDef.Info != null ? "present" : "null");
                _logger.LogInformation("Sheets Count: {Count}", wbDef.Sheets?.Count ?? 0);

                if (wbDef.Sheets != null)
                {
                    for (int si = 0; si < wbDef.Sheets.Count; si++)
                    {
                        var sheet = wbDef.Sheets[si];
                        _logger.LogInformation("  Sheet[{Idx}]: Name='{Name}', Fields={Fields}",
                            si, sheet.Name ?? "(null)", sheet.Fields?.Count ?? 0);

                        // PHASE 21.4 — STAGE 4: DUMP EVERY PROPERTY
                        // Only logs first few fields to avoid extreme log spam
                        if (sheet.Fields != null)
                        {
                            int dumpCount = Math.Min(sheet.Fields.Count, 5);
                            for (int fi = 0; fi < dumpCount; fi++)
                            {
                                var f = sheet.Fields[fi];
                                _logger.LogInformation("    Field[{Idx}]:", fi);
                                _logger.LogInformation("      Id:             {Val}", f.Id ?? "(empty)");
                                _logger.LogInformation("      Name:           {Val}", f.Name ?? "(empty)");
                                _logger.LogInformation("      Cell.Address:   {Val}", f.Cell?.Address ?? "(empty/null)");
                                _logger.LogInformation("      Cell.RowIndex:  {Val}", f.Cell != null ? f.Cell.RowIndex.ToString() : "(null)");
                                _logger.LogInformation("      Type:           {Val}", f.Type.ToString());
                                _logger.LogInformation("      Value:          {Val}", string.IsNullOrWhiteSpace(f.Value) ? "(empty)" : f.Value);
                                _logger.LogInformation("      Required:       {Val}", f.Required);
                                _logger.LogInformation("      Locked:         {Val}", f.Locked);
                                _logger.LogInformation("      Formula:        {Val}", f.Formula ?? "(null)");
                                _logger.LogInformation("      Placeholder:    {Val}", f.Placeholder ?? "(null)");
                                _logger.LogInformation("      DefaultValue:   {Val}", f.DefaultValue ?? "(null)");
                                _logger.LogInformation("      MaxLength:      {Val}", f.MaxLength);
                                _logger.LogInformation("      TabIndex:       {Val}", f.TabIndex);
                                _logger.LogInformation("      Visible:        {Val}", f.Visible);
                                _logger.LogInformation("      DataValidation: {Val}", f.DataValidation != null ? f.DataValidation.Type : "(null)");
                                _logger.LogInformation("      Style.Font:     {Val}", f.Style?.Font != null ? f.Style.Font.Name ?? "(default)" : "(null)");
                                _logger.LogInformation("      Style.FontSize: {Val}", f.Style?.Font?.SizePt ?? 0);
                            }
                            if (sheet.Fields.Count > dumpCount)
                            {
                                _logger.LogInformation("      ... and {Remaining} more fields (first {Dumped} shown)",
                                    sheet.Fields.Count - dumpCount, dumpCount);
                            }
                        }
                    }
                }

                _logger.LogInformation("=========================================================");
                _logger.LogInformation("STAGE 3-4 COMPLETE: {Sheets} sheets, {Fields} fields total",
                    wbDef.Sheets?.Count ?? 0,
                    wbDef.Sheets?.Sum(s => s.Fields?.Count ?? 0) ?? 0);
                _logger.LogInformation("=========================================================");
            }

            // Phase 5.2: Resolve the original workbook from the session store.
            // The sessionId is the only thing the browser needs to remember.
            string sessionId = wbDef.SessionId;
            string? sourcePath = null;

            if (!string.IsNullOrEmpty(sessionId))
            {
                sourcePath = _sessionStore.ResolveWorkbookPath(sessionId);
                if (sourcePath == null)
                {
                    return StatusCode(410, new ApiErrorResponse
                    {
                        Message = "The editing session has expired. Please upload the workbook again.",
                        ErrorCode = ErrorCodes.InvalidRequest
                    });
                }
            }
            else
            {
                // Backward compatibility: fallback to SourceFileName (deprecated)
                if (!string.IsNullOrEmpty(wbDef.SourceFileName))
                {
                    string formsDir = Path.Combine(_env.ContentRootPath, "Forms");
                    sourcePath = Path.Combine(formsDir, wbDef.SourceFileName);
                    if (!System.IO.File.Exists(sourcePath))
                    {
                        sourcePath = null;
                    }
                }

                if (string.IsNullOrEmpty(sourcePath))
                {
                    return BadRequest(new ApiErrorResponse
                    {
                        Message = "No session or source workbook found. Please upload the workbook first.",
                        ErrorCode = ErrorCodes.InvalidRequest
                    });
                }

                _logger.LogWarning("[DEPRECATED] SaveEdited called without sessionId — fallback to SourceFileName '{File}'", wbDef.SourceFileName);
            }

            try
            {
                // ═════════════════════════════════════════════════════════
                // PHASE 21.3 — STAGE 1: EXPORT PAYLOAD RECEIVED
                // ═════════════════════════════════════════════════════════
                _logger.LogInformation("=========================================================");
                _logger.LogInformation("EXPORT PAYLOAD RECEIVED");
                _logger.LogInformation("=========================================================");
                _logger.LogInformation("Workbook: {Title}", wbDef.Info?.Title ?? "(untitled)");
                _logger.LogInformation("Session: {SessionId}", sessionId);
                _logger.LogInformation("SourcePath: {Path}", sourcePath);
                _logger.LogInformation("Pages (sheets): {Count}", wbDef.Sheets.Count);

                int st1TotalFields = wbDef.Sheets.Sum(s => s.Fields.Count);
                int st1WithValues = wbDef.Sheets.Sum(s => s.Fields.Count(f => !string.IsNullOrWhiteSpace(f.Value)));
                int st1Empty = st1TotalFields - st1WithValues;

                _logger.LogInformation("Total Fields: {Total}", st1TotalFields);
                _logger.LogInformation("Fields With Values: {WithVal}", st1WithValues);
                _logger.LogInformation("Empty Fields: {Empty}", st1Empty);

                // Print EVERY field
                int st1FieldIdx = 0;
                foreach (var st1Sheet in wbDef.Sheets)
                {
                    foreach (var st1Field in st1Sheet.Fields)
                    {
                        st1FieldIdx++;
                        string st1Id = st1Field.Id ?? $"field_{st1FieldIdx}";
                        _logger.LogInformation("  ---------------------------------------------------------");
                        _logger.LogInformation("  Field #{Idx}:", st1FieldIdx);
                        _logger.LogInformation("    Id             : {Id}", st1Id);
                        _logger.LogInformation("    Page           : {Sheet}", st1Sheet.Name);
                        _logger.LogInformation("    Cell           : {Cell}", st1Field.Cell?.Address ?? "?");
                        _logger.LogInformation("    Type           : {Type}", st1Field.Type.ToString());
                        _logger.LogInformation("    Value          : {Val}", string.IsNullOrWhiteSpace(st1Field.Value) ? "(empty)" : st1Field.Value);
                        _logger.LogInformation("    Required       : {Req}", st1Field.Required);
                        _logger.LogInformation("    ReadOnly       : {RO}", st1Field.Locked);
                        _logger.LogInformation("    Placeholder    : {PH}", string.IsNullOrEmpty(st1Field.Placeholder) ? "" : st1Field.Placeholder);
                        _logger.LogInformation("    Default Value  : {Def}", string.IsNullOrEmpty(st1Field.DefaultValue) ? "" : st1Field.DefaultValue);
                        _logger.LogInformation("    Validation     : {Val}", string.IsNullOrEmpty(st1Field.Formula) ? "" : st1Field.Formula);
                        _logger.LogInformation("    Max Length     : {Max}", st1Field.MaxLength > 0 ? st1Field.MaxLength.ToString() : "");
                        _logger.LogInformation("    DV Formula     : {Opt}", st1Field.DataValidation?.Formula1 ?? (st1Field.DataValidation != null ? "validation present" : ""));
                    }
                }

                _logger.LogInformation("=========================================================");
                _logger.LogInformation("EXPORT PAYLOAD END — {Total} fields received ({WithVal} with values, {Empty} empty)",
                    st1TotalFields, st1WithValues, st1Empty);
                _logger.LogInformation("=========================================================");

                // If all fields are empty, flag early
                if (st1WithValues == 0)
                {
                    _logger.LogError("[DIAG] CRITICAL: ALL fields have empty values. " +
                        "No cells will be written. Possible model binding mismatch: " +
                        "frontend runtimeFormToWorkbookDefinition may map fields with wrong property names. " +
                        "Check: values[f.id] mapping → must match backend FieldDefinition.Name");
                }

                foreach (var diagSheet in wbDef.Sheets)
                {
                    int nonEmpty = diagSheet.Fields.Count(f => !string.IsNullOrWhiteSpace(f.Value));
                    int empty = diagSheet.Fields.Count(f => string.IsNullOrWhiteSpace(f.Value));
                    _logger.LogInformation("  Sheet '{Name}': {Total} fields ({NonEmpty} with values, {Empty} empty)",
                        diagSheet.Name, diagSheet.Fields.Count, nonEmpty, empty);

                    // Log first 20 non-empty field details
                    foreach (var f in diagSheet.Fields.Where(f => !string.IsNullOrWhiteSpace(f.Value)).Take(20))
                    {
                        _logger.LogInformation("    Field: cell={Addr}, name='{Name}', type={Type}, value='{Val}', id={Id}",
                            f.Cell?.Address ?? "?", f.Name ?? "?", f.Type.ToString(), f.Value, f.Id ?? "?");
                    }

                    // Also log first 5 empty fields to verify model binding catches them
                    foreach (var f in diagSheet.Fields.Where(f => string.IsNullOrWhiteSpace(f.Value)).Take(5))
                    {
                        _logger.LogInformation("    EMPTY Field: cell={Addr}, name='{Name}', type={Type}, id={Id}",
                            f.Cell?.Address ?? "?", f.Name ?? "?", f.Type.ToString(), f.Id ?? "?");
                    }

                    if (diagSheet.Fields.Count(f => !string.IsNullOrWhiteSpace(f.Value)) == 0)
                    {
                        _logger.LogWarning("[DIAG] Sheet '{Name}' has ZERO non-empty field values!", diagSheet.Name);
                    }
                }

                if (wbDef.Sheets.Sum(s => s.Fields.Count(f => !string.IsNullOrWhiteSpace(f.Value))) == 0)
                {
                    _logger.LogError("[DIAG] CRITICAL: ALL fields have empty values. " +
                        "No cells will be written. Possible model binding mismatch: " +
                        "frontend runtimeFormToWorkbookDefinition may map fields with wrong property names. " +
                        "Check: values[f.id] mapping → must match backend FieldDefinition.Name");
                }

                _logger.LogInformation("========== PAYLOAD DIAGNOSTIC END ==========");

                // ═════════════════════════════════════════════════════════
                // PHASE 21.4 — STAGE 6: CALLING SAVE SERVICE
                // ═════════════════════════════════════════════════════════
                int stageFieldCountController = wbDef.Sheets?.Sum(s => s.Fields?.Count ?? 0) ?? 0;
                int stageWithValuesController = wbDef.Sheets?.Sum(s => s.Fields?.Count(f => !string.IsNullOrWhiteSpace(f.Value)) ?? 0) ?? 0;

                _logger.LogInformation("=========================================================");
                _logger.LogInformation("CALLING SAVE SERVICE");
                _logger.LogInformation("=========================================================");
                _logger.LogInformation("WorkbookDefinition");
                _logger.LogInformation("  Sheets: {Count}", wbDef.Sheets?.Count ?? 0);
                _logger.LogInformation("  Fields (total): {Count}", stageFieldCountController);
                _logger.LogInformation("  Fields (with values): {Count}", stageWithValuesController);
                _logger.LogInformation("  Fields (empty): {Count}", stageFieldCountController - stageWithValuesController);
                _logger.LogInformation("");
                string formsDir = Path.Combine(_env.ContentRootPath, "Forms");

                _logger.LogInformation("Calling: FormSaveService.SaveEditedValuesAsync()");
                _logger.LogInformation("  OutputDir: {Dir}", formsDir);
                _logger.LogInformation("  SourcePath: {Path}", sourcePath);
                _logger.LogInformation("=========================================================");
                Directory.CreateDirectory(formsDir);

                var result = await _formSaveService.SaveEditedValuesAsync(wbDef, formsDir, sourcePath);

                // ═════════════════════════════════════════════════════════
                // PHASE 21.4 — STAGE 10: FIELD LOSS DETECTOR
                // ═════════════════════════════════════════════════════════
                int stageCellsWritten = result.CellsWritten;
                int stageServiceWithValues = 0; // Will be overridden by FormSaveService diagnostic

                _logger.LogInformation("=========================================================");
                _logger.LogInformation("FIELD LOSS DETECTOR");
                _logger.LogInformation("=========================================================");
                _logger.LogInformation("Field count tracking:");
                _logger.LogInformation("  Stage 1-4 (Controller after binding):  {Cnt}", stageFieldCountController);
                _logger.LogInformation("  Stage 6 (Fields with values):          {Cnt}", stageWithValuesController);
                _logger.LogInformation("  Stage 7-8 (FormSaveService entry):     (see FormSaveService logs)");
                _logger.LogInformation("  Stage 9 (WorkbookValueWriter entry):   (see WorkbookValueWriter logs)");
                _logger.LogInformation("  Cells Written (result):               {Cnt}", stageCellsWritten);
                _logger.LogInformation("");

                // Detect where loss occurred
                if (stageFieldCountController == 0)
                {
                    _logger.LogError("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                    _logger.LogError("FIELD LOSS DETECTED: Controller has 0 fields after model binding");
                    _logger.LogError("LIKELY CAUSE: JSON property names don't match FieldDefinition model");
                    _logger.LogError("Check: Frontend sends {{ id, name, cell, type, value }}");
                    _logger.LogError("Backend expects: {{ Id, Name, Cell, Type, Value }}");
                    _logger.LogError("Possible mismatch: 'id' vs 'Id', 'cell' vs 'Cell', etc.");
                    _logger.LogError("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                else if (stageWithValuesController == 0 && stageFieldCountController > 0)
                {
                    _logger.LogError("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                    _logger.LogError("FIELD LOSS DETECTED: {Total} fields exist but ALL are empty", stageFieldCountController);
                    _logger.LogError("LIKELY CAUSE: Frontend values[f.id] mapping does not match");
                    _logger.LogError("Check: Frontend runtimeFormToWorkbookDefinition → values[f.id]");
                    _logger.LogError("  → f.id must match backend field.Id (case-sensitive)");
                    _logger.LogError("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                else if (stageCellsWritten == 0 && stageWithValuesController > 0)
                {
                    _logger.LogError("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                    _logger.LogError("FIELD LOSS DETECTED: {WithVal} fields with values, but 0 written", stageWithValuesController);
                    _logger.LogError("LIKELY CAUSE: Sheet name mismatch or cell address resolution failure");
                    _logger.LogError("Check: WorkbookValueWriter sheet name matching case-sensitivity");
                    _logger.LogError("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                else
                {
                    _logger.LogInformation("FIELD LOSS NOT DETECTED — counts are consistent");
                    _logger.LogInformation("  Controller: {Total} fields, {WithVal} with values",
                        stageFieldCountController, stageWithValuesController);
                    _logger.LogInformation("  Written: {Written} cells", stageCellsWritten);
                    _logger.LogInformation("  (Differences may be caused by empty-value skipping)");
                }

                byte[] fileBytes = await System.IO.File.ReadAllBytesAsync(result.WorkbookPath);
                string safeFileName = SanitizeFileName(wbDef.Info?.Title ?? "edited") + ".xlsx";

                // ═════════════════════════════════════════════════════════
                // PHASE 21.2: EXPORT STATE — log every field written
                // ═════════════════════════════════════════════════════════
                _logger.LogInformation("========================================================");
                _logger.LogInformation("  EXPORT STATE — Fields Written to Workbook");
                _logger.LogInformation("========================================================");

                int exportFieldCount = 0;
                foreach (var exportSheet in wbDef.Sheets)
                {
                    foreach (var exportField in exportSheet.Fields)
                    {
                        if (string.IsNullOrWhiteSpace(exportField.Value)) continue;
                        string cellAddr = exportField.Cell?.Address ?? "?";
                        string fieldId = exportField.Id ?? $"field_{exportFieldCount}";
                        _logger.LogInformation(
                            "  EXPORT Field: id='{Id}' cell='{Sheet}!{Cell}' value='{Val}' type={Type}",
                            fieldId, exportSheet.Name, cellAddr, exportField.Value, exportField.Type);
                        exportFieldCount++;
                    }
                }

                _logger.LogInformation("  EXPORT COMPLETE: {Count} fields written to workbook", exportFieldCount);
                _logger.LogInformation("========================================================");

                // ═════════════════════════════════════════════════════════
                // PHASE 21.4 — STAGE 11: EXPORT PIPELINE SUMMARY
                // ═════════════════════════════════════════════════════════
                _logger.LogInformation("=========================================================");
                _logger.LogInformation("EXPORT PIPELINE SUMMARY");
                _logger.LogInformation("=========================================================");
                _logger.LogInformation("Pipeline stage                         | Fields");
                _logger.LogInformation("---------------------------------------------------------");
                _logger.LogInformation("Frontend JSON (sent by browser)        | Unknown (see Stage 2)");
                _logger.LogInformation("ASP.NET Model Binding                  | {Total}", stageFieldCountController);
                _logger.LogInformation("Fields with non-empty values           | {WithVal}", stageWithValuesController);
                _logger.LogInformation("SaveService.SaveEditedValuesAsync()     | see service logs");
                _logger.LogInformation("WorkbookValueWriter.WriteValues()       | see writer logs");
                _logger.LogInformation("Cells Actually Written                 | {Written}", stageCellsWritten);
                _logger.LogInformation("Workbook Verification                  | see writer logs");
                _logger.LogInformation("");

                // Auto-diagnose: pinpoint the loss stage
                if (stageFieldCountController == 0)
                {
                    _logger.LogInformation("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                    _logger.LogInformation("BUG LOCATION: Between Frontend JSON → ASP.NET Model Binding");
                    _logger.LogInformation("  Frontend sends fields but binding produces 0.");
                    _logger.LogInformation("  Root cause: JSON property name mismatch.");
                    _logger.LogInformation("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                else if (stageWithValuesController == 0 && stageFieldCountController > 0)
                {
                    _logger.LogInformation("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                    _logger.LogInformation("BUG LOCATION: Model Binding → Export Pipeline");
                    _logger.LogInformation("  {Total} fields exist but ALL have empty values.", stageFieldCountController);
                    _logger.LogInformation("  Root cause: runtimeFormToWorkbookDefinition values mapping.");
                    _logger.LogInformation("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                else if (stageCellsWritten == 0 && stageWithValuesController > 0)
                {
                    _logger.LogInformation("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                    _logger.LogInformation("BUG LOCATION: FormSaveService → WorkbookValueWriter");
                    _logger.LogInformation("  {WithVal} fields reach the service but none are written.", stageWithValuesController);
                    _logger.LogInformation("  Root cause: Sheet name mismatch or cell resolution.");
                    _logger.LogInformation("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                else
                {
                    _logger.LogInformation("Pipeline trace: Normal (all counts consistent)");
                }
                _logger.LogInformation("=========================================================");

                _logger.LogInformation(
                    "Edited workbook saved: {CellsWritten} cells, {Size} bytes. Download: {FileName}",
                    result.CellsWritten, fileBytes.Length, safeFileName);

                // Phase 19: Add round-trip compatibility report as response headers
                if (result.RoundTripReport != null)
                {
                    var rt = result.RoundTripReport;

                    // Compatibility score header (0-100)
                    Response.Headers.Append("X-Compatibility-Score", rt.CompatibilityScore.ToString());
                    Response.Headers.Append("X-Compatibility-Warnings", rt.Warnings.Count.ToString());
                    Response.Headers.Append("X-Compatibility-Errors", rt.Errors.Count.ToString());

                    // Field comparison headers
                    Response.Headers.Append("X-Fields-Original", rt.FieldCountOriginal.ToString());
                    Response.Headers.Append("X-Fields-Edited", rt.FieldCountEdited.ToString());
                    Response.Headers.Append("X-Fields-Missing", rt.MissingFields.Count.ToString());
                    Response.Headers.Append("X-Fields-Added", rt.AddedFields.Count.ToString());
                    Response.Headers.Append("X-Fields-Changed", rt.ChangedFields.Count.ToString());

                    // Configuration headers
                    Response.Headers.Append("X-Config-FieldsSheet", rt.FieldsSheetExists ? "1" : "0");
                    Response.Headers.Append("X-Config-ExcelOutputSetting", rt.ExcelOutputSettingExists ? "1" : "0");

                    // Layout headers
                    Response.Headers.Append("X-Layout-Changes", rt.LayoutChanges.Count.ToString());

                    // Log if score is below threshold
                    if (rt.CompatibilityScore < 100)
                    {
                        _logger.LogInformation(
                            "Round-trip compatibility: Score={Score}/100, Fields={F}/{E}, Layout={L}, Config={C}",
                            rt.CompatibilityScore, rt.FieldCountOriginal, rt.FieldCountEdited,
                            rt.LayoutChanges.Count, rt.ConfigurationChanges.Count);
                    }
                }

                _logger.LogInformation("========== SAVE PIPELINE END ==========");

                return File(
                    fileBytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    safeFileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to save edited workbook");
                return StatusCode(500, new ApiErrorResponse
                {
                    Message = $"Failed to save edited workbook: {ex.Message}",
                    ErrorCode = ErrorCodes.InternalError
                });
            }
        }

        /// <summary>
        /// Phase 20: Build a DesignerModel from the reconstructed FormDefinition.
        /// Converts the legacy FormDefinition/ClusterDefinition model into the
        /// comprehensive DesignerModel used for frontend reconstruction.
        /// </summary>
        private static DesignerModel BuildDesignerModel(FormDefinition form, string sessionId)
        {
            var model = new DesignerModel
            {
                SessionId = sessionId,
                Info = new DesignerWorkbookInfo
                {
                    Title = form.Workbook?.Title ?? "Untitled",
                    Author = form.Workbook?.Author ?? "",
                    Description = form.Workbook?.Description ?? "",
                    Version = form.Workbook?.Version ?? "1.0"
                },
                Configuration = new DesignerConfiguration
                {
                    HasFieldsSheet = form.Sheets.Any(s =>
                        string.Equals(s.Name, "_Fields", StringComparison.OrdinalIgnoreCase)),
                    HasExcelOutputSetting = form.Sheets.Any(s =>
                        string.Equals(s.Name, "ExcelOutputSetting", StringComparison.OrdinalIgnoreCase)),
                    FieldsSheetRowCount = form.Clusters.Count
                }
            };

            // Populate FieldsSheetFieldIds from clusters
            model.Configuration.FieldsSheetFieldIds = form.Clusters
                .Select(c => c.ClusterId)
                .Where(id => !string.IsNullOrEmpty(id))
                .ToList();

            // Build pages from content sheets (skipping config sheets)
            int pageIndex = 0;
            foreach (var sheet in form.Sheets)
            {
                string sheetName = sheet.Name ?? "";
                if (string.Equals(sheetName, "_Fields", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(sheetName, "ExcelOutputSetting", StringComparison.OrdinalIgnoreCase))
                    continue;

                string pageId = $"page_{pageIndex}";

                var page = new DesignerPage
                {
                    Id = pageId,
                    Name = sheetName,
                    Index = pageIndex,
                    Layout = new DesignerPageLayout
                    {
                        PrintArea = sheet.PrintArea?.Address ?? "",
                        PaperSize = sheet.PageSettings?.PaperSize ?? "Letter",
                        Orientation = sheet.PageSettings?.Orientation ?? "portrait",
                        WidthPt = sheet.PageSettings?.WidthPt ?? 0,
                        HeightPt = sheet.PageSettings?.HeightPt ?? 0,
                        Zoom = sheet.PageSettings?.Zoom ?? 100,
                        FitToPagesWide = sheet.PageSettings?.FitToPagesWide ?? 0,
                        FitToPagesTall = sheet.PageSettings?.FitToPagesTall ?? 0,
                        CenterHorizontally = sheet.PageSettings?.CenterHorizontally ?? false,
                        CenterVertically = sheet.PageSettings?.CenterVertically ?? false,
                        Margins = new DesignerMargins
                        {
                            Top = sheet.PageSettings?.TopMargin ?? 54,
                            Bottom = sheet.PageSettings?.BottomMargin ?? 54,
                            Left = sheet.PageSettings?.LeftMargin ?? 50.4,
                            Right = sheet.PageSettings?.RightMargin ?? 50.4
                        },
                        Rows = sheet.RowHeights.Select(r => new DesignerRowInfo
                        {
                            Index = r.Key,
                            HeightPt = r.Value
                        }).ToList(),
                        Columns = sheet.ColumnWidths.Select(c => new DesignerColumnInfo
                        {
                            Index = c.Key,
                            WidthChars = c.Value
                        }).ToList(),
                        MergedCells = sheet.MergedCells.Select(m => m.Address).ToList()
                    },
                    BackgroundImage = sheet.BackgroundImage,
                    BackgroundWidth = sheet.BackgroundWidth,
                    BackgroundHeight = sheet.BackgroundHeight
                };

                // Build fields for this page from clusters matching this sheet
                foreach (var cluster in form.Clusters)
                {
                    string clusterSheetId = cluster.SheetId ?? "";
                    // Match by SheetId or by target sheet index
                    bool matchesSheet = string.Equals(clusterSheetId, sheet.Id, StringComparison.OrdinalIgnoreCase);
                    if (!matchesSheet && form.Sheets.Count > 0)
                    {
                        // Fallback: match by sheet index from SheetId
                        int sheetIdx = -1;
                        if (int.TryParse(clusterSheetId.Replace("sheet_", ""), out int parsed))
                            sheetIdx = parsed;
                        matchesSheet = sheetIdx == pageIndex;
                    }
                    if (!matchesSheet) continue;

                    var cellAddr = cluster.CellAddress ?? "";
                    cellAddr = cellAddr.Split(':')[0]; // Use first cell of merge range

                    var field = new DesignerField
                    {
                        Id = cluster.ClusterId ?? $"field_{form.Clusters.IndexOf(cluster)}",
                        Name = cluster.Name ?? "",
                        CellAddress = cellAddr,
                        Type = cluster.Type ?? "text",
                        Visible = string.Equals(cluster.Visibility ?? "visible", "visible", StringComparison.OrdinalIgnoreCase),
                        Value = cluster.Metadata?.GetValueOrDefault("currentValue") ?? "",
                        Label = cluster.Name ?? "",
                        Description = cluster.Remarks ?? ""
                    };

                    // Extract behavior from InputParameters or Metadata
                    if (cluster.InputParameters?.TryGetValue("required", out var required) == true)
                        field.Behavior.Required = required == "1" || string.Equals(required, "true", StringComparison.OrdinalIgnoreCase);
                    if (cluster.InputParameters?.TryGetValue("readonly", out var readOnly) == true)
                        field.Behavior.ReadOnly = readOnly == "1" || string.Equals(readOnly, "true", StringComparison.OrdinalIgnoreCase);
                    if (cluster.InputParameters?.TryGetValue("maxlength", out var maxLen) == true)
                    {
                        if (int.TryParse(maxLen, out int parsedMaxLen))
                            field.MaxLength = parsedMaxLen;
                    }
                    if (cluster.InputParameters?.TryGetValue("placeholder", out var placeholder) == true)
                        field.Placeholder = placeholder;
                    if (cluster.InputParameters?.TryGetValue("default", out var defaultVal) == true)
                        field.DefaultValue = defaultVal;

                    // Parse options for dropdown fields
                    if (string.Equals(field.Type, "dropdown", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(field.Type, "list", StringComparison.OrdinalIgnoreCase))
                    {
                        if (cluster.InputParameters?.TryGetValue("options", out var optionsStr) == true)
                        {
                            field.Options = optionsStr
                                .Split(new[] { ',', ';', '|' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(o => o.Trim())
                                .Where(o => !string.IsNullOrEmpty(o))
                                .ToList();
                        }
                    }

                    // Extract formatting from cell styles if available
                    if (sheet.CellStyles != null && cellAddr != null)
                    {
                        if (sheet.CellStyles.TryGetValue(cellAddr, out var styleInfo) && styleInfo != null)
                        {
                            field.Style = new DesignerFieldStyle
                            {
                                FontFamily = styleInfo.FontName,
                                FontSize = styleInfo.FontSize,
                                Bold = styleInfo.Bold,
                                Italic = styleInfo.Italic,
                                Underline = styleInfo.Underline,
                                FontColor = styleInfo.Color,
                                FillColor = styleInfo.FillColor,
                                HorizontalAlignment = styleInfo.HorizontalAlignment,
                                VerticalAlignment = styleInfo.VerticalAlignment,
                                WrapText = styleInfo.WrapText,
                                BorderTop = styleInfo.BorderTop,
                                BorderBottom = styleInfo.BorderBottom,
                                BorderLeft = styleInfo.BorderLeft,
                                BorderRight = styleInfo.BorderRight
                            };
                        }
                    }

                    // Extract validation if available
                    if (cluster.InputParameters?.TryGetValue("validation", out var valStr) == true)
                    {
                        field.Validation = new DesignerFieldValidation
                        {
                            Type = "custom",
                            Formula1 = valStr
                        };
                    }

                    page.Fields.Add(field);
                }

                model.Pages.Add(page);
                pageIndex++;
            }

            // Build comments list from all sheets
            foreach (var sheet in form.Sheets)
            {
                string sheetName = sheet.Name ?? "";
                foreach (var mc in sheet.MergedCells)
                {
                    // Check CellValues dictionary for comments (stored as metadata)
                    foreach (var kvp in sheet.CellStyles)
                    {
                        model.Comments.Add(new DesignerComment
                        {
                            CellAddress = kvp.Key,
                            Worksheet = sheetName,
                            Text = $"Field at {kvp.Key}",
                            Author = "PaperLess"
                        });
                    }
                    // Don't add duplicate comments for merged cells
                    break;
                }
            }

            // Add cluster remarks as comments
            foreach (var cluster in form.Clusters)
            {
                string remarks = cluster.Remarks ?? "";
                if (string.IsNullOrEmpty(remarks)) continue;

                string cellAddr = (cluster.CellAddress ?? "").Split(':')[0];
                if (string.IsNullOrEmpty(cellAddr)) continue;

                // Avoid duplicates
                if (!model.Comments.Any(c =>
                    string.Equals(c.CellAddress, cellAddr, StringComparison.OrdinalIgnoreCase)))
                {
                    model.Comments.Add(new DesignerComment
                    {
                        CellAddress = cellAddr,
                        Worksheet = "",
                        Text = remarks,
                        Author = "PaperLess"
                    });
                }
            }

            // Configuration duplicate detection
            int excelOutputCount = form.Sheets.Count(s =>
                string.Equals(s.Name, "ExcelOutputSetting", StringComparison.OrdinalIgnoreCase));
            model.Configuration.IsDuplicated = excelOutputCount > 1;

            return model;
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
        /// </summary>
        /// <param name="templateId">Template ID / session ID / XLSX filename (without extension).</param>
        [HttpGet("runtime/{templateId}")]
        [ProducesResponseType(typeof(ApiResponse<RuntimeForm>), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status404NotFound)]
        public IActionResult GetRuntime([FromRoute] string templateId)
        {
            try
            {
                _logger.LogInformation("[RUNTIME] Loading COM metadata for ID={TemplateId}", templateId);

                string formsDir = Path.Combine(_env.ContentRootPath, "Forms");

                // Load COM metadata — single source of truth for field coordinates
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

                _logger.LogWarning("[RUNTIME] No runtime metadata found for template {TemplateId}", templateId);
                return NotFound(new ApiErrorResponse
                {
                    Message = $"Runtime metadata not found for template: {templateId}. The template must be uploaded via POST /api/form/from-excel first.",
                    ErrorCode = ErrorCodes.InvalidRequest
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[RUNTIME] FAILED: {Message}\n{StackTrace}", ex.Message, ex.ToString());
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
    }
}
