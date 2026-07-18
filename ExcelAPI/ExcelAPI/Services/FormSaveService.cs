using ExcelAPI.Generators;
using ExcelAPI.Models;
using Microsoft.Extensions.Logging;

namespace ExcelAPI.Services
{
    public class FormSaveResult
    {
        public string XmlPath { get; set; } = "";
        public string WorkbookPath { get; set; } = "";
        public string PreviewPath { get; set; } = "";
        public string PdfPath { get; set; } = "";
        public DatabaseResult? DatabaseObjects { get; set; }
    }

    public interface IFormSaveService
    {
        Task<FormSaveResult> SaveAsync(FormDefinition form, string outputDirectory, CancellationToken cancellationToken = default);
        Task<FormSaveResult> OutputExcelAsync(FormDefinition form, string outputDirectory, CancellationToken cancellationToken = default);
    }

    public class FormSaveService : IFormSaveService
    {
        private readonly XmlGenerator _xmlGenerator;
        private readonly DatabaseGenerator _databaseGenerator;
        private readonly WorkbookGenerator _workbookGenerator;
        private readonly PreviewGenerator _previewGenerator;
        private readonly PdfGenerator _pdfGenerator;
        private readonly ILogger<FormSaveService> _logger;

        public FormSaveService(
            XmlGenerator xmlGenerator,
            DatabaseGenerator databaseGenerator,
            WorkbookGenerator workbookGenerator,
            PreviewGenerator previewGenerator,
            PdfGenerator pdfGenerator,
            ILogger<FormSaveService> logger)
        {
            _xmlGenerator = xmlGenerator;
            _databaseGenerator = databaseGenerator;
            _workbookGenerator = workbookGenerator;
            _previewGenerator = previewGenerator;
            _pdfGenerator = pdfGenerator;
            _logger = logger;
        }

        public async Task<FormSaveResult> SaveAsync(FormDefinition definition, string outputDirectory, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Directory.CreateDirectory(outputDirectory);

            var formId = Guid.NewGuid().ToString("N");
            var result = new FormSaveResult();

            // 1. Generate XML
            cancellationToken.ThrowIfCancellationRequested();
            string xml = _xmlGenerator.Generate(definition);
            result.XmlPath = Path.Combine(outputDirectory, $"form_{formId}.xml");
            await File.WriteAllTextAsync(result.XmlPath, xml, cancellationToken);
            _logger.LogInformation("XML saved to {Path}", result.XmlPath);

            // 2. Generate Database Objects
            cancellationToken.ThrowIfCancellationRequested();
            result.DatabaseObjects = _databaseGenerator.Generate(definition);
            _logger.LogInformation("Database objects generated: {SheetCount} sheets, {ClusterCount} clusters",
                result.DatabaseObjects.Sheets.Count, result.DatabaseObjects.Clusters.Count);

            // 3. Generate Workbook (XLSX) — now includes cell values, comments, and _Fields sheet
            cancellationToken.ThrowIfCancellationRequested();
            result.WorkbookPath = Path.Combine(outputDirectory, $"form_{formId}.xlsx");
            result.WorkbookPath = _workbookGenerator.Generate(definition, result.WorkbookPath);
            _logger.LogInformation("Workbook saved to {Path}", result.WorkbookPath);

            // 4. Preview for each sheet
            cancellationToken.ThrowIfCancellationRequested();
            var firstSheet = definition.Sheets.FirstOrDefault();
            if (firstSheet != null)
            {
                result.PreviewPath = Path.Combine(outputDirectory, $"preview_{formId}.png");
                result.PreviewPath = _previewGenerator.Generate(definition, firstSheet.Id, result.PreviewPath);
                _logger.LogInformation("Preview saved to {Path}", result.PreviewPath);
            }

            // 5. PDF for each sheet
            cancellationToken.ThrowIfCancellationRequested();
            if (firstSheet != null)
            {
                result.PdfPath = Path.Combine(outputDirectory, $"form_{formId}.pdf");
                result.PdfPath = _pdfGenerator.Generate(definition, firstSheet.Id, result.PdfPath);
                _logger.LogInformation("PDF saved to {Path}", result.PdfPath);
            }

            return result;
        }

        /// <summary>
        /// Generate Output Excel of Form — creates a workbook with cell values,
        /// field metadata comments, and a hidden _Fields worksheet for republish compatibility.
        /// This is the legacy "Output Excel of Form" pipeline.
        /// </summary>
        /// <param name="definition">The form definition (SheetDefinition.CellValues must be populated).</param>
        /// <param name="outputDirectory">Directory to save the output workbook.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public async Task<FormSaveResult> OutputExcelAsync(FormDefinition definition, string outputDirectory, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Directory.CreateDirectory(outputDirectory);

            var formId = Guid.NewGuid().ToString("N");
            var result = new FormSaveResult();

            // Generate Output Excel workbook — includes structure, cell values, comments, and _Fields sheet
            cancellationToken.ThrowIfCancellationRequested();
            result.WorkbookPath = Path.Combine(outputDirectory, $"output_{formId}.xlsx");
            result.WorkbookPath = _workbookGenerator.Generate(definition, result.WorkbookPath);
            _logger.LogInformation("Output Excel saved to {Path}", result.WorkbookPath);

            return result;
        }
    }
}