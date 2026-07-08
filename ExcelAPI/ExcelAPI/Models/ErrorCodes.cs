namespace ExcelAPI.Models
{
    /// <summary>
    /// Centralized error codes for consistent API error responses.
    /// Every error returned by the API should use one of these codes.
    /// </summary>
    public static class ErrorCodes
    {
        // --- Validation Errors (400) ---

        /// <summary>No file was provided in the upload request.</summary>
        public const string NoFileUploaded = "NO_FILE_UPLOADED";

        /// <summary>The uploaded file has an unsupported extension.</summary>
        public const string InvalidFileExtension = "INVALID_FILE_EXTENSION";

        /// <summary>The uploaded file exceeds the maximum allowed size.</summary>
        public const string FileTooLarge = "FILE_TOO_LARGE";

        /// <summary>The uploaded file appears to be corrupted or unreadable.</summary>
        public const string CorruptedWorkbook = "CORRUPTED_WORKBOOK";

        /// <summary>The worksheet does not have a print area configured.</summary>
        public const string PrintAreaNotConfigured = "PRINT_AREA_NOT_CONFIGURED";

        /// <summary>The requested worksheet was not found in the workbook.</summary>
        public const string WorksheetNotFound = "WORKSHEET_NOT_FOUND";

        // --- Processing Errors (500) ---

        /// <summary>Excel COM automation failed (Excel not installed, COM error, etc.).</summary>
        public const string ExcelProcessingError = "EXCEL_PROCESSING_ERROR";

        /// <summary>The capture operation timed out.</summary>
        public const string RequestTimeout = "REQUEST_TIMEOUT";

        /// <summary>An unexpected internal error occurred.</summary>
        public const string InternalError = "INTERNAL_ERROR";
    }
}
