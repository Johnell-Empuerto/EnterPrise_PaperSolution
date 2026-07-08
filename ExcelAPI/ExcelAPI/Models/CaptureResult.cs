namespace ExcelAPI.Models
{
    /// <summary>
    /// Result of a print area capture operation.
    /// Contains the image URL, page dimensions, and any field metadata extracted from cell comments.
    /// </summary>
    public class CaptureResult
    {
        /// <summary>Relative URL to the captured PNG image.</summary>
        public string ImageUrl { get; set; } = string.Empty;

        /// <summary>Dimensions of the generated PNG image at 300 DPI.</summary>
        public PageInfo Page { get; set; } = new();

        /// <summary>List of form fields extracted from cell comments/notes.</summary>
        public List<ExcelField> Fields { get; set; } = new();

        /// <summary>Page setup debug information for verifying coordinate alignment (temporary).</summary>
        public PageSetupDebug? PageSetup { get; set; }
    }
}
