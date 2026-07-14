using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services.Importer;

public class ExcelImporter
{
    private readonly IProgress<string> _progress;

    public ExcelImporter(IProgress<string> progress)
    {
        _progress = progress;
    }

    public ExtractionResult Import(string filePath)
    {
        _progress.Report("Opening workbook...");
        var result = new ExtractionResult
        {
            SourceFile = filePath,
            ExtractedAt = DateTime.UtcNow
        };

        using var doc = SpreadsheetDocument.Open(filePath, false);
        var wbPart = doc.WorkbookPart;
        if (wbPart == null) return result;

        _progress.Report("Parsing workbook metadata...");
        var wbParser = new WorkbookParser();
        result.Workbook = wbParser.Parse(wbPart);

        _progress.Report("Filtering sheets...");
        result.Workbook.SheetMetadata = result.Workbook.SheetMetadata
            .Where(sm => sm.State == "visible" && !sm.Name.StartsWith("_"))
            .ToList();

        _progress.Report("Parsing styles...");
        var styleParser = new StyleParser();
        result.Styles = styleParser.Parse(wbPart);

        _progress.Report("Loading shared strings...");
        var sharedStrings = LoadSharedStrings(wbPart);

        _progress.Report("Parsing sheets...");
        var sheetParser = new SheetParser(wbPart, result.Styles, sharedStrings);
        result.Workbook.Sheets = sheetParser.Parse(result.Workbook.SheetMetadata);

        _progress.Report($"Import complete: {result.Workbook.Sheets.Count} sheets, {result.Styles.Count} styles, {result.Workbook.Sheets.Sum(s => s.Cells.Count)} cells");

        return result;
    }

    private List<string> LoadSharedStrings(WorkbookPart wbPart)
    {
        var result = new List<string>();
        try
        {
            var ssp = wbPart.SharedStringTablePart;
            if (ssp?.GetStream() == null) return result;

            var sx = XDocument.Load(ssp.GetStream());
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            foreach (var si in sx.Descendants(ns + "si"))
            {
                var t = si.Element(ns + "t");
                result.Add(t?.Value ?? si.Value);
            }
        }
        catch { }
        return result;
    }
}
