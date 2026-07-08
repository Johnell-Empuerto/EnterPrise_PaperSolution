namespace ExcelAPI.Models
{
    /// <summary>
    /// Generic API response wrapper for consistent response format.
    /// </summary>
    /// <typeparam name="T">Type of the data payload.</typeparam>
    public class ApiResponse<T>
    {
        /// <summary>
        /// Indicates whether the request was successful.
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// A human-readable message describing the result.
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// The response payload.
        /// </summary>
        public T? Data { get; set; }
    }
}
