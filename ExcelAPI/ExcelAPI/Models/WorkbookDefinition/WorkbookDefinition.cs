// ────────────────────────────────────────────────────────────────────────────
// WorkbookDefinition — Canonical Excel Template Domain Model
//
// This is the single source of truth for describing an analyzed Excel workbook.
// It is produced by the Designer (COM Interop) and consumed by Runtime and
// Rendering layers. It describes ONLY the template — no user-entered values.
//
// Ownership: Shared (Designer → Runtime + Rendering)
// ────────────────────────────────────────────────────────────────────────────

namespace ExcelAPI.Models.WorkbookDefinition
{
    /// <summary>
    /// Root of the canonical workbook definition. Describes an analyzed Excel
    /// workbook template — its sheets, fields, images, shapes, styles, and layout.
    /// 
    /// Lifecycle:
    ///   1. Created by Designer COM analysis (ExcelCaptureService)
    ///   2. Stored alongside CaptureResult as the canonical template description
    ///   3. Consumed by FormRuntimeBuilder to produce RuntimeForm
    ///   4. Eventually consumed directly by Rendering (future phase)
    /// </summary>
    public class WorkbookDefinition
    {
        /// <summary>Workbook metadata (title, author, dates, version).</summary>
        public WorkbookInfo Info { get; set; } = new();

        /// <summary>All sheets in this workbook.</summary>
        public List<SheetDefinition> Sheets { get; set; } = new();

        /// <summary>
        /// The path to the original Excel file on disk (used during analysis only).
        /// Empty string after analysis is complete.
        /// </summary>
        public string SourcePath { get; set; } = string.Empty;

        /// <summary>
        /// Source file name (for display/logging purposes).
        /// Phase 5.2: DEPRECATED. Use SessionId instead.
        /// The browser should never track filenames — the server owns the workbook.
        /// </summary>
        [Obsolete("Use SessionId instead. The server owns the workbook, not the browser.")]
        public string SourceFileName { get; set; } = string.Empty;

        /// <summary>
        /// Server-side session identifier for the original workbook.
        /// Phase 5.2: Replaces SourceFileName. The browser only needs to remember
        /// this sessionId — the server resolves the original workbook from its
        /// session store (TempWorkbooks/{sessionId}/original.xlsx).
        /// </summary>
        public string SessionId { get; set; } = string.Empty;

        /// <summary>
        /// When this definition was captured.
        /// </summary>
        public DateTime CapturedAt { get; set; } = DateTime.UtcNow;

        /// <summary>
        /// Schema version, incremented when the model changes.
        /// </summary>
        public const string SchemaVersion = "1.0.0";
    }

    /// <summary>
    /// High-level metadata about the workbook.
    /// Populated from COM Workbook.BuiltinDocumentProperties or OOXML.
    /// </summary>
    public class WorkbookInfo
    {
        /// <summary>Workbook title.</summary>
        public string Title { get; set; } = "Untitled Workbook";

        /// <summary>Author name from document properties.</summary>
        public string Author { get; set; } = string.Empty;

        /// <summary>When the workbook was created.</summary>
        public DateTime Created { get; set; }

        /// <summary>When the workbook was last modified.</summary>
        public DateTime Modified { get; set; }

        /// <summary>Application version that last saved the file.</summary>
        public string Version { get; set; } = "1.0";

        /// <summary>Description / subject from document properties.</summary>
        public string Description { get; set; } = string.Empty;

        /// <summary>
        /// Keywords (tags) for categorization.
        /// </summary>
        public List<string> Keywords { get; set; } = new();

        /// <summary>
        /// Default tab index for keyboard navigation across the workbook.
        /// </summary>
        public int DefaultTabIndex { get; set; }
    }
}
