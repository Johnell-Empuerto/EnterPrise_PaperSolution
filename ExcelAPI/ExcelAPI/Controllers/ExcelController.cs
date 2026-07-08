using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using ExcelAPI.Models;
using ExcelAPI.Services.Interfaces;

namespace ExcelAPI.Controllers
{
    /// <summary>
    /// Controller for Excel file upload and print area capture operations.
    /// This controller only handles HTTP concerns (validation, file saving, response).
    /// All Excel COM automation is delegated to IExcelCaptureService.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly IExcelCaptureService _excelCaptureService;
        private readonly IWebHostEnvironment _env;
        private readonly IOptions<ExcelCaptureOptions> _options;
        private readonly ILogger<ExcelController> _logger;

        // Supported Excel file extensions
        private static readonly HashSet<string> AllowedExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".xlsx",
            ".xls"
        };

        public ExcelController(
            IExcelCaptureService excelCaptureService,
            IWebHostEnvironment env,
            IOptions<ExcelCaptureOptions> options,
            ILogger<ExcelController> logger)
        {
            _excelCaptureService = excelCaptureService;
            _env = env;
            _options = options;
            _logger = logger;
        }

        /// <summary>
        /// Upload an Excel file and capture its print area as a PNG image.
        /// </summary>
        /// <param name="file">The Excel file to upload (multipart/form-data).</param>
        /// <returns>
        /// 200 OK with image URL on success.
        /// 400 Bad Request for validation errors.
        /// 500 Internal Server Error for processing failures.
        /// </returns>
        /// <response code="200">Print area captured successfully.</response>
        /// <response code="400">Invalid or missing file, unsupported extension, or no print area configured.</response>
        /// <response code="500">Excel processing or COM error.</response>
        [HttpPost("upload")]
        [RequestSizeLimit(50 * 1024 * 1024)] // Hard upper limit (50 MB) as middleware safety net
        [ProducesResponseType(typeof(ApiResponse<UploadResponse>), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(ApiErrorResponse), StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> Upload([FromForm] IFormFile? file)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();

            // Create a cancellation token with the configured timeout
            // This ensures Excel COM operations don't hang indefinitely
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(
                HttpContext.RequestAborted);
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(_options.Value.RequestTimeoutSeconds));

            // --- Step 1: Validate the uploaded file ---

            // Check if file was provided
            if (file == null || file.Length == 0)
            {
                _logger.LogWarning("Upload rejected: No file provided.");
                return BadRequest(new ApiErrorResponse
                {
                    Message = "No file uploaded. Please select an Excel file.",
                    ErrorCode = ErrorCodes.NoFileUploaded
                });
            }

            // Check file extension
            string extension = Path.GetExtension(file.FileName);
            if (!AllowedExtensions.Contains(extension))
            {
                _logger.LogWarning("Upload rejected: Invalid extension '{Ext}' for file {FileName}",
                    extension, file.FileName);
                return BadRequest(new ApiErrorResponse
                {
                    Message = $"Unsupported file extension '{extension}'. Only .xlsx and .xls files are supported.",
                    ErrorCode = ErrorCodes.InvalidFileExtension
                });
            }

            // Check file size against configured limit
            long maxSizeBytes = _options.Value.MaxUploadSizeMB * 1024L * 1024L;
            if (file.Length > maxSizeBytes)
            {
                _logger.LogWarning("Upload rejected: File {FileName} exceeds {MaxMB}MB limit ({Size} bytes)",
                    file.FileName, _options.Value.MaxUploadSizeMB, file.Length);
                return BadRequest(new ApiErrorResponse
                {
                    Message = $"File size exceeds the maximum allowed size of {_options.Value.MaxUploadSizeMB}MB.",
                    ErrorCode = ErrorCodes.FileTooLarge
                });
            }

            _logger.LogInformation("Upload started: {FileName}, {Size} bytes, type: {Ext}",
                file.FileName, file.Length, extension);

            // --- Step 2: Save the uploaded file ---

            // Ensure the Uploads directory exists
            string uploadsFolder = Path.Combine(_env.ContentRootPath, _options.Value.UploadDirectory);
            Directory.CreateDirectory(uploadsFolder);

            // Generate a unique filename to avoid collisions
            string uniqueFileName = $"{Guid.NewGuid():N}{extension}";
            string filePath = Path.Combine(uploadsFolder, uniqueFileName);

            _logger.LogDebug("Saving uploaded file to: {FilePath}", filePath);

            long saveTime = 0;
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
                saveTime = sw.ElapsedMilliseconds;
                _logger.LogInformation("File saved in {Duration}ms: {FilePath}", saveTime, filePath);
            }

            // --- Steps 3-9: Delegate to service, then clean up ---
            try
            {
                var result = await _excelCaptureService.CapturePrintAreaAsync(filePath, timeoutCts.Token);

                sw.Stop();
                _logger.LogInformation(
                    "Upload completed successfully: {FileName} -> {ImageUrl} ({FieldCount} fields, {Total}ms)",
                    file.FileName, result.ImageUrl, result.Fields.Count, sw.ElapsedMilliseconds);

                return Ok(new ApiResponse<CaptureResult>
                {
                    Success = true,
                    Message = "Print area captured successfully.",
                    Data = result
                });
            }
            catch (PrintAreaNotConfiguredException ex)
            {
                // User error: no print area configured -> return 400 Bad Request
                sw.Stop();
                _logger.LogWarning(ex, "Print area not configured for file: {FileName} ({Total}ms)",
                    file.FileName, sw.ElapsedMilliseconds);

                return BadRequest(new ApiErrorResponse
                {
                    Message = ex.Message,
                    ErrorCode = ErrorCodes.PrintAreaNotConfigured
                });
            }
            catch (InvalidOperationException ex)
            {
                // System/processing error -> return 500 Internal Server Error
                sw.Stop();
                _logger.LogWarning(ex, "Processing error for file: {FileName} ({Total}ms)",
                    file.FileName, sw.ElapsedMilliseconds);

                return StatusCode(500, new ApiErrorResponse
                {
                    Message = ex.Message,
                    ErrorCode = ErrorCodes.ExcelProcessingError
                });
            }
            catch (OperationCanceledException)
            {
                // Request timed out
                sw.Stop();
                _logger.LogError("Request timed out for file: {FileName} ({Total}ms)",
                    file.FileName, sw.ElapsedMilliseconds);

                return StatusCode(500, new ApiErrorResponse
                {
                    Message = "The request timed out while processing the Excel file.",
                    ErrorCode = ErrorCodes.RequestTimeout
                });
            }
            catch (Exception ex)
            {
                // Unexpected error -> return 500 Internal Server Error
                sw.Stop();
                _logger.LogError(ex, "Unexpected error for file: {FileName} ({Total}ms)",
                    file.FileName, sw.ElapsedMilliseconds);

                return StatusCode(500, new ApiErrorResponse
                {
                    Message = $"An unexpected error occurred: {ex.Message}",
                    ErrorCode = ErrorCodes.InternalError
                });
            }
            finally
            {
                // Clean up the uploaded file if configured
                if (_options.Value.DeleteUploadsAfterCapture)
                {
                    try
                    {
                        if (System.IO.File.Exists(filePath))
                        {
                            System.IO.File.Delete(filePath);
                            _logger.LogDebug("Cleaned up uploaded file: {FilePath}", filePath);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed to clean up uploaded file: {FilePath}", filePath);
                    }
                }
            }
        }
    }
}
