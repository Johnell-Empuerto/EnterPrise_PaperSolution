namespace ExcelAPI.Models
{
    /// <summary>
    /// Response model returned after a successful Excel file upload and print area capture.
    /// </summary>
    public class UploadResponse
    {
        /// <summary>
        /// Relative URL to the captured print area PNG image.
        /// Example: /preview/page_abc123.png
        /// </summary>
        public string? ImageUrl { get; set; }
    }
}
