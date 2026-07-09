using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using ExcelAPI.Models;
using ExcelAPI.Services;
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

        public FormController(
            IFormSaveService formSaveService,
            IWebHostEnvironment env,
            IOptions<ExcelCaptureOptions> options,
            ILogger<FormController> logger)
        {
            _formSaveService = formSaveService;
            _env = env;
            _options = options;
            _logger = logger;
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

            try
            {
                // Use existing service to capture, then convert result to FormDefinition
                var captureService = HttpContext.RequestServices.GetRequiredService<Services.Interfaces.IExcelCaptureService>();
                var captureResult = await captureService.CapturePrintAreaAsync(filePath);

                // Convert capture result to FormDefinition
                var formDefinition = ConvertCaptureToForm(captureResult, file.FileName);

                return Ok(new ApiResponse<FormDefinition>
                {
                    Success = true,
                    Message = $"Excel file processed. {captureResult.Fields.Count} field(s) detected.",
                    Data = formDefinition
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to process Excel file for form definition");
                return BadRequest(new ApiErrorResponse
                {
                    Message = $"Failed to process Excel file: {ex.Message}",
                    ErrorCode = ErrorCodes.ExcelProcessingError
                });
            }
            finally
            {
                try { if (System.IO.File.Exists(filePath)) System.IO.File.Delete(filePath); } catch { }
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
    }
}