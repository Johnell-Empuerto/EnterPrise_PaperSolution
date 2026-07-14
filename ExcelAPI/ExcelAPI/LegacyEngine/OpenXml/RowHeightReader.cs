using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelAPI.LegacyEngine.OpenXml;

public class RowHeightReader : IRowHeightReader
{
    public IReadOnlyList<double> Read(string filePath, string sheetName, int sheetId)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = doc.WorkbookPart;
        if (workbookPart is null) return Array.Empty<double>();

        var sheet = workbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName);
        if (sheet is null) return Array.Empty<double>();

        var wsPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart;
        var ws = wsPart?.Worksheet;
        if (ws is null) return Array.Empty<double>();

        // Determine max row from sheet dimension
        int maxRow = ParseMaxRow(ws.SheetDimension?.Reference?.Value);
        if (maxRow == 0) return Array.Empty<double>();

        var heights = new double[maxRow];

        // Default row height
        double defaultHeight = 15.0;
        var sheetProps = ws.Descendants<SheetFormatProperties>().FirstOrDefault();
        if (sheetProps?.DefaultRowHeight?.Value is not null)
            defaultHeight = (double)sheetProps.DefaultRowHeight.Value;

        for (int i = 0; i < maxRow; i++) heights[i] = defaultHeight;

        // Apply custom heights
        foreach (var row in ws.Descendants<Row>())
        {
            uint index = row.RowIndex?.Value ?? 0;
            if (index == 0 || index > maxRow) continue;
            if (row.Height?.Value is not null)
                heights[index - 1] = (double)row.Height.Value;
        }

        return heights;
    }

    private static int ParseMaxRow(string? dimensionRef)
    {
        if (string.IsNullOrEmpty(dimensionRef)) return 0;
        var parts = dimensionRef.Split(':');
        if (parts.Length < 2) return 0;
        string endRef = parts[^1];
        string digits = new(endRef.SkipWhile(char.IsLetter).ToArray());
        return int.TryParse(digits, out var r) ? r : 0;
    }
}
