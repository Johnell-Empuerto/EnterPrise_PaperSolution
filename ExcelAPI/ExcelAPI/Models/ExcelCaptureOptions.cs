namespace ExcelAPI.Models
{
    /// <summary>
    /// Configuration options for the Excel Capture feature.
    /// Bound to the "ExcelCapture" section of appsettings.json.
    /// </summary>
    public class ExcelCaptureOptions
    {
        public const string SectionName = "ExcelCapture";

        /// <summary>Directory for temporarily storing uploaded Excel files (relative to ContentRootPath).</summary>
        public string UploadDirectory { get; set; } = "Uploads";

        /// <summary>Directory for storing captured preview PNG images (relative to ContentRootPath).</summary>
        public string PreviewDirectory { get; set; } = "Preview";

        /// <summary>Maximum upload file size in megabytes.</summary>
        public int MaxUploadSizeMB { get; set; } = 25;

        /// <summary>Whether to delete uploaded Excel files after processing.</summary>
        public bool DeleteUploadsAfterCapture { get; set; } = true;

        /// <summary>Number of hours to retain preview images before cleanup.</summary>
        public int DeletePreviewAfterHours { get; set; } = 24;

        /// <summary>Maximum timeout in seconds for the entire capture operation.</summary>
        public int RequestTimeoutSeconds { get; set; } = 120;

        /// <summary>Cleanup interval in minutes for the background file cleanup service.</summary>
        public int CleanupIntervalMinutes { get; set; } = 60;
    }
}
