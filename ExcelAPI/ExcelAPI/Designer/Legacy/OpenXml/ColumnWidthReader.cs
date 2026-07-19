using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelAPI.Designer.Legacy.OpenXml;

public class ColumnWidthReader : IColumnWidthReader
{
    public IReadOnlyList<double> Read(string filePath, string sheetName, int sheetId, double maxDigitWidth)
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

        // Determine max column from sheet dimension reference
        int maxCol = ParseMaxColumn(ws.SheetDimension?.Reference?.Value);
        if (maxCol == 0) return Array.Empty<double>();

        var widths = new double[maxCol];

        // Default column width
        double defaultCharWidth = 8.43;
        var sheetProps = ws.ChildElements.FirstOrDefault() as SheetFormatProperties;
        if (sheetProps is null)
        {
            // Try via Descendants
            sheetProps = ws.Descendants<SheetFormatProperties>().FirstOrDefault();
        }
        if (sheetProps?.DefaultColumnWidth?.Value is not null)
            defaultCharWidth = (double)sheetProps.DefaultColumnWidth.Value;

        double defaultWidthPoints = ColumnWidthToPoints(defaultCharWidth, maxDigitWidth);
        for (int i = 0; i < maxCol; i++) widths[i] = defaultWidthPoints;

        // Apply custom widths from <cols>
        var cols = ws.Descendants<Columns>().FirstOrDefault();
        if (cols is not null)
        {
            foreach (var col in cols.Cast<Column>())
            {
                uint min = col.Min?.Value ?? 1;
                uint max = col.Max?.Value ?? min;
                double? customWidth = col.Width?.Value;
                if (customWidth is null) continue;

                double colW = ColumnWidthToPoints((double)customWidth, maxDigitWidth);
                for (uint c = min; c <= max && c <= maxCol; c++)
                    widths[c - 1] = colW;
            }
        }

        return widths;
    }

    public static double ColumnWidthToPoints(double charWidth, double maxDigitWidth)
    {
        double pixels = charWidth * maxDigitWidth + 5;
        return pixels * 72.0 / 96.0;
    }

    private static int ParseMaxColumn(string? dimensionRef)
    {
        if (string.IsNullOrEmpty(dimensionRef)) return 0;
        var parts = dimensionRef.Split(':');
        if (parts.Length < 2) return 0;
        string endRef = parts[^1];
        int col = 0;
        string letters = new(endRef.TakeWhile(char.IsLetter).ToArray());
        foreach (char c in letters.ToUpperInvariant())
            col = col * 26 + (c - 'A' + 1);
        return col;
    }
}
