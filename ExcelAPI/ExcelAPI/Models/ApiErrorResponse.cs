namespace ExcelAPI.Models
{
    /// <summary>
    /// Standardized error response with a machine-readable error code.
    /// Used for all API error responses to ensure consistency.
    /// </summary>
    public class ApiErrorResponse
    {
        /// <summary>Always false for error responses.</summary>
        public bool Success { get; set; } = false;

        /// <summary>Human-readable error message.</summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>Machine-readable error code for programmatic handling.</summary>
        public string ErrorCode { get; set; } = string.Empty;
    }
}
