using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelAPI.LegacyEngine.OpenXml;

public class PageSetupReader : IPageSetupReader
{
    public OpenXmlPageSetupModel Read(string filePath, string sheetName, int sheetId)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = doc.WorkbookPart;
        if (workbookPart is null) return new OpenXmlPageSetupModel();

        var sheet = workbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName);

        if (sheet is null) return new OpenXmlPageSetupModel();

        var wsPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart;
        var ws = wsPart?.Worksheet;
        if (ws is null) return new OpenXmlPageSetupModel();

        var model = new OpenXmlPageSetupModel();

        var ps = ws.Descendants<PageSetup>().FirstOrDefault();
        var pm = ws.Descendants<PageMargins>().FirstOrDefault();
        var po = ws.Descendants<PrintOptions>().FirstOrDefault();

        // Paper size
        if (ps?.PaperSize?.Value is not null)
        {
            var (w, h) = GetPaperDimensions((int)ps.PaperSize.Value);
            model.PaperWidthPoints = w;
            model.PaperHeightPoints = h;
        }

        // Orientation
        if (ps?.Orientation?.Value is not null)
        {
            string orient = ps.Orientation.Value.ToString();
            model.Orientation = orient.Equals("landscape", StringComparison.OrdinalIgnoreCase)
                ? "Landscape" : "Portrait";
        }

        // Page dimensions
        if (model.Orientation == "Landscape" && model.PaperWidthPoints < model.PaperHeightPoints)
        {
            model.PageWidthPoints = model.PaperHeightPoints;
            model.PageHeightPoints = model.PaperWidthPoints;
        }
        else
        {
            model.PageWidthPoints = model.PaperWidthPoints;
            model.PageHeightPoints = model.PaperHeightPoints;
        }

        // Margins (OpenXML stores them in inches — keep as-is for OriginEngine which multiplies by 72)
        if (pm is not null)
        {
            if (pm.Left?.Value is not null) model.MarginLeft = (double)pm.Left.Value;
            if (pm.Right?.Value is not null) model.MarginRight = (double)pm.Right.Value;
            if (pm.Top?.Value is not null) model.MarginTop = (double)pm.Top.Value;
            if (pm.Bottom?.Value is not null) model.MarginBottom = (double)pm.Bottom.Value;
            if (pm.Header?.Value is not null) model.MarginHeader = (double)pm.Header.Value;
            if (pm.Footer?.Value is not null) model.MarginFooter = (double)pm.Footer.Value;
        }

        // Centering (from PrintOptions, not PageSetup)
        if (po?.HorizontalCentered?.Value is not null) model.CenterHorizontally = po.HorizontalCentered.Value;
        if (po?.VerticalCentered?.Value is not null) model.CenterVertically = po.VerticalCentered.Value;

        return model;
    }

    private static (double width, double height) GetPaperDimensions(int paperSize) => paperSize switch
    {
        1 => (612, 792),     2 => (612, 792),     3 => (612, 1008),
        4 => (792, 1224),    5 => (612, 1008),    6 => (612, 792),
        7 => (612, 792),     8 => (595, 842),     9 => (842, 1191),
        10 => (595, 842),    11 => (420, 595),    12 => (595, 842),
        13 => (498, 708),    14 => (612, 792),    15 => (612, 1008),
        16 => (612, 792),    17 => (612, 792),    18 => (612, 792),
        19 => (612, 792),    20 => (612, 792),    21 => (612, 792),
        22 => (612, 792),    23 => (612, 792),    27 => (612, 1008),
        28 => (595, 842),    29 => (498, 638),    30 => (420, 595),
        31 => (595, 842),    32 => (612, 1008),   42 => (842, 1191),
        43 => (595, 842),    50 => (420, 595),    51 => (612, 1008),
        52 => (595, 842),    53 => (612, 1008),   54 => (612, 792),
        _ => (612, 792)
    };
}
