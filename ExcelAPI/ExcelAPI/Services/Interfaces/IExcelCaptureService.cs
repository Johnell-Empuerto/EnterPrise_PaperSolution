namespace ExcelAPI.Services.Interfaces
{
    /// <summary>
    /// Service interface for capturing Excel worksheet print areas as PNG images.
    /// All Excel COM automation logic is encapsulated behind this interface.
    /// </summary>
    public interface IExcelCaptureService
    {
        /// <summary>
        /// Opens an Excel workbook, reads the configured print area from the first worksheet,
        /// and captures it as a PNG image.
        /// </summary>
        /// <param name="excelFilePath">
        /// Full file path to the Excel workbook (.xlsx or .xls).
        /// </param>
        /// <param name="cancellationToken">
        /// Token to cancel the operation if it exceeds the configured timeout or the client disconnects.
        /// </param>
        /// <returns>
        /// The relative URL path to the saved preview image (e.g., "/preview/page_abc.png").
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown when no print area is configured, Excel cannot be started,
        /// or any COM-related failure occurs.
        /// </exception>
        /// <exception cref="OperationCanceledException">
        /// Thrown when the operation is cancelled due to timeout or client disconnect.
        /// </exception>
        Task<string> CapturePrintAreaAsync(string excelFilePath, CancellationToken cancellationToken = default);
    }
}
