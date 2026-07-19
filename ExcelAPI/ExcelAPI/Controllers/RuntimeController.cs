using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using ExcelAPI.Models;
using ExcelAPI.Application;
using ExcelAPI.Designer.Capture;
using ExcelAPI.Runtime;

namespace ExcelAPI.Controllers
{
    /// <summary>
    /// PaperLess Runtime API — the primary production endpoint for the frontend.
    ///
    /// Upload an Excel file → COM captures the print area → returns PNG URLs + overlay coordinates.
    /// No HTML rendering. No CSS Grid. Excel is the rendering engine.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class RuntimeController : ControllerBase
    {
        private readonly IWebHostEnvironment _env;
        private readonly IOptions<ExcelCaptureOptions> _options;
        private readonly ILogger<RuntimeController> _logger;
        private readonly ExcelCaptureService _captureService;
        private readonly RuntimeCoordinateGenerator _runtimeGenerator;

        public RuntimeController(
            IWebHostEnvironment env,
            IOptions<ExcelCaptureOptions> options,
            ILogger<RuntimeController> logger,
            ExcelCaptureService captureService,
            RuntimeCoordinateGenerator runtimeGenerator)
        {
            _env = env;
            _options = options;
            _logger = logger;
            _captureService = captureService;
            _runtimeGenerator = runtimeGenerator;
        }

        /// <summary>
        /// Upload an Excel workbook and receive rendered pages + overlay coordinates.
        ///
        /// The backend will:
        ///   1. Open the workbook via Excel COM
        ///   2. Read the configured Print Area
        ///   3. Export the Print Area as PDF → PNG at 300 DPI
        ///   4. Read cell comments and measure field positions via COM Range geometry
        ///   5. Persist the runtime metadata
        ///   6. Return page images and overlay definitions
        ///
        /// No HTML/CSS Excel re-rendering occurs. Excel is the rendering engine.
        /// </summary>
        [HttpPost("upload")]
        [RequestSizeLimit(50 * 1024 * 1024)]
        [ProducesResponseType(typeof(RuntimeUploadResponse), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> Upload([FromForm] IFormFile? file)
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
                // Capture the print area via COM
                var captureResult = await _captureService.CapturePrintAreaAsync(filePath, templateId);

                // Save the Excel file to Forms/ for persistence
                string formsDir = Path.Combine(_env.ContentRootPath, "Forms");
                Directory.CreateDirectory(formsDir);
                string persistentPath = Path.Combine(formsDir, uniqueFileName);
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Move(filePath, persistentPath, overwrite: true);
                }

                // Persist runtime metadata using WorkbookDefinition path (Phase 3.6)
                // SaveFromWbDef prefers InternalWorkbookDefinition when available.
                // Falls back to legacy CaptureResult path for pre-WbDef templates.
                string bgUrl = $"/preview/page_{templateId}.png";
                int pageW = captureResult.Page?.Width ?? 0;
                int pageH = captureResult.Page?.Height ?? 0;
                _runtimeGenerator.SaveFromWbDef(captureResult, templateId, formsDir, file.FileName,
                    pageW, pageH, bgUrl);

                // Build the clean response model
                var response = BuildResponse(captureResult, templateId, file.FileName);

                _logger.LogInformation(
                    "[Runtime] Upload processed: {FileName} → {PageCount} page(s), {OverlayCount} overlay(s)",
                    file.FileName, response.Pages.Count, response.Overlays.Count);

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[Runtime] Upload failed for: {FileName}", file.FileName);
                try { if (System.IO.File.Exists(filePath)) System.IO.File.Delete(filePath); } catch { }
                return BadRequest(new ApiErrorResponse
                {
                    Message = $"Failed to process Excel file: {ex.Message}",
                    ErrorCode = ErrorCodes.ExcelProcessingError
                });
            }
        }

        /// <summary>
        /// Convert a CaptureResult into the clean RuntimeUploadResponse model.
        /// </summary>
        private RuntimeUploadResponse BuildResponse(CaptureResult capture, string templateId, string fileName)
        {
            var response = new RuntimeUploadResponse();

            // Build page info
            response.Pages.Add(new RuntimePageInfo
            {
                SheetName = "Sheet1",
                BackgroundImage = $"/preview/page_{templateId}.png",
                Width = capture.Page?.Width ?? 0,
                Height = capture.Page?.Height ?? 0
            });

            // Build overlay definitions from captured fields
            foreach (var field in capture.Fields)
            {
                // Map the backend type to the frontend overlay type
                string overlayType = MapToOverlayType(field.Type);

                response.Overlays.Add(new RuntimeOverlayInfo
                {
                    Id = field.Id,
                    SheetName = "Sheet1",
                    Type = overlayType,
                    Left = field.Left,
                    Top = field.Top,
                    Width = field.Width,
                    Height = field.Height,
                    Cell = field.Cell,
                    IsMerged = field.IsMerged,
                    MergeAddress = field.MergeAddress
                });
            }

            return response;
        }

        /// <summary>
        /// Map backend field type to frontend overlay type.
        /// Backend uses PascalCase (Text, Checkbox), frontend uses lowercase (textbox, checkbox).
        /// </summary>
        private static string MapToOverlayType(string backendType)
        {
            return backendType.ToLowerInvariant() switch
            {
                "text" => "textbox",
                "checkbox" => "checkbox",
                "signature" => "signature",
                "date" => "date",
                "number" => "number",
                _ => "textbox"
            };
        }
    }
}
