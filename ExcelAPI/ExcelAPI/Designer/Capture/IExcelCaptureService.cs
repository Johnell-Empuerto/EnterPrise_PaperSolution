using ExcelAPI.Models;

namespace ExcelAPI.Designer.Capture
{
    /// <summary>
    /// Service interface for capturing Excel worksheet print areas as PNG images.
    /// All Excel COM automation logic is encapsulated behind this interface.
    /// </summary>
    public interface IExcelCaptureService
    {
        /// <summary>
        /// Opens an Excel workbook, reads the configured print area from the first worksheet,
        /// captures it as a 300 DPI PNG image, and extracts field metadata from cell comments.
        /// </summary>
        /// <param name="excelFilePath">
        /// Full file path to the Excel workbook (.xlsx or .xls).
        /// </param>
        /// <param name="fileId">
        /// Optional deterministic file ID for the PNG filename.
        /// If provided, the PNG will be named page_{fileId}.png instead of a random GUID.
        /// This ensures the preview URL matches the template ID — single source of truth.
        /// </param>
        /// <param name="cancellationToken">
        /// Token to cancel the operation if it exceeds the configured timeout or the client disconnects.
        /// </param>
        /// <returns>
        /// A <see cref="CaptureResult"/> containing the image URL, page dimensions,
        /// and a list of form fields extracted from cell comments.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown when no print area is configured, Excel cannot be started,
        /// or any COM-related failure occurs.
        /// </exception>
        /// <exception cref="OperationCanceledException">
        /// Thrown when the operation is cancelled due to timeout or client disconnect.
        /// </exception>
        Task<CaptureResult> CapturePrintAreaAsync(string excelFilePath, string? fileId = null, CancellationToken cancellationToken = default);
    }
}
