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

        public FormController(
            IFormSaveService formSaveService,
            IWebHostEnvironment env,
            IOptions<ExcelCaptureOptions> options,
            ILogger<FormController> logger,
            RuntimeCoordinateGenerator runtimeGenerator,
            PythonRenderService pythonRender,
            WorkbookReaderService workbookReaderService,
            ISessionWorkbookStore sessionStore)
        {
            _formSaveService = formSaveService;
            _env = env;
            _options = options;
            _logger = logger;
            _runtimeGenerator = runtimeGenerator;
            _pythonRender = pythonRender;
            _workbookReaderService = workbookReaderService;
            _sessionStore = sessionStore;
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

                return Ok(new
                {
                    success = true,
                    sessionId,
                    pages = pageResults
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

                return Ok(new ApiResponse<object>
                {
                    Success = true,
                    Message = $"Workbook read successfully: {formDefinition.Sheets.Count} sheet(s), {formDefinition.Clusters.Count} field(s).",
                    Data = new
                    {
                        formDefinition,
                        sessionId,
                        fieldCount = formDefinition.Clusters.Count,
                        sheetCount = formDefinition.Sheets.Count
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
            if (wbDef == null)
            {
                return BadRequest(new ApiErrorResponse
                {
                    Message = "No workbook definition provided.",
                    ErrorCode = ErrorCodes.InvalidRequest
                });
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
                _logger.LogInformation("========== SAVE PIPELINE ==========");
                _logger.LogInformation("Controller: POST /api/form/save-edited");
                _logger.LogInformation("SessionId: {SessionId}", sessionId);
                _logger.LogInformation("SourcePath: {Path}", sourcePath);
                _logger.LogInformation("Workbook: {Title}", wbDef.Info?.Title ?? "(untitled)");
                _logger.LogInformation("Sheets: {Count}, Fields with values: {Fields}",
                    wbDef.Sheets.Count,
                    wbDef.Sheets.Sum(s => s.Fields.Count(f => !string.IsNullOrWhiteSpace(f.Value))));
                _logger.LogInformation("WorkbookGenerator invoked: FALSE");
                _logger.LogInformation("WorkbookValueWriter invoked: TRUE");

                // ── DIAGNOSTIC: Log incoming WbDef payload structure (Investigation 2/3) ──
                _logger.LogInformation("========== PAYLOAD DIAGNOSTIC ==========");
                _logger.LogInformation("SessionId present: {Has}", !string.IsNullOrEmpty(sessionId));
                _logger.LogInformation("SourceFileName present (deprecated): {Has}", !string.IsNullOrEmpty(wbDef.SourceFileName));
                _logger.LogInformation("Sheet count: {Count}", wbDef.Sheets.Count);
                _logger.LogInformation("Total fields: {Total}", wbDef.Sheets.Sum(s => s.Fields.Count));
                _logger.LogInformation("Total fields with non-empty values: {WithVal}",
                    wbDef.Sheets.Sum(s => s.Fields.Count(f => !string.IsNullOrWhiteSpace(f.Value))));

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

                string formsDir = Path.Combine(_env.ContentRootPath, "Forms");
                Directory.CreateDirectory(formsDir);

                var result = await _formSaveService.SaveEditedValuesAsync(wbDef, formsDir, sourcePath);

                byte[] fileBytes = await System.IO.File.ReadAllBytesAsync(result.WorkbookPath);
                string safeFileName = SanitizeFileName(wbDef.Info?.Title ?? "edited") + ".xlsx";

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
