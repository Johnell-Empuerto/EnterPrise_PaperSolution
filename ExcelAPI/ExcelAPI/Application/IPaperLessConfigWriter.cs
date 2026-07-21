using WbDef = ExcelAPI.Models.WorkbookDefinition;

namespace ExcelAPI.Application
{
    /// <summary>
    /// Writes PaperLess configuration JSON into the PaperLessConfig worksheet
    /// inside the exported XLSX. Configuration is stored as VeryHidden metadata
    /// and restored during re-upload via DesignerModelReader.
    /// </summary>
    public interface IPaperLessConfigWriter
    {
        /// <summary>
        /// Write (or update) the PaperLessConfig worksheet in the output workbook.
        /// Idempotent — repeated calls always produce exactly one PaperLessConfig sheet.
        /// </summary>
        void WritePaperLessConfig(WbDef.WorkbookDefinition definition, string outputPath);
    }
}
