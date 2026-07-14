using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelAPI.LegacyEngine.OpenXml;

public class PrintAreaReader : IPrintAreaReader
{
    public OpenXmlPrintAreaModel Read(string filePath, string sheetName)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = doc.WorkbookPart;
        if (workbookPart is null) return new OpenXmlPrintAreaModel();

        var definedNames = workbookPart.Workbook.DefinedNames;
        if (definedNames is null) return new OpenXmlPrintAreaModel();

        foreach (var dn in definedNames.Cast<DefinedName>())
        {
            string? name = dn.Name?.Value;
            if (name is null || !name.Equals("_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase))
                continue;

            string? text = dn.Text;
            if (string.IsNullOrEmpty(text)) continue;

            // Parse: 'SheetName'!$A$1:$L$45 or SheetName!$A$1:$L$45
            text = text.Trim('\'', '=');
            int exclIndex = text.IndexOf('!');
            if (exclIndex < 0) continue;

            string refSheet = text[..exclIndex].Trim('\'');
            string range = text[(exclIndex + 1)..];

            if (!string.Equals(refSheet, sheetName, StringComparison.OrdinalIgnoreCase))
                continue;

            var parts = range.Replace("$", "").Split(':');
            if (parts.Length == 0) continue;

            var start = ParseCellRef(parts[0]);
            if (parts.Length == 1)
            {
                return new OpenXmlPrintAreaModel
                {
                    StartColumn = start.col, StartRow = start.row,
                    EndColumn = start.col, EndRow = start.row
                };
            }

            var end = ParseCellRef(parts[1]);
            return new OpenXmlPrintAreaModel
            {
                StartColumn = start.col, StartRow = start.row,
                EndColumn = end.col, EndRow = end.row
            };
        }

        return new OpenXmlPrintAreaModel();
    }

    private static (int col, int row) ParseCellRef(string ref_)
    {
        string letters = new(ref_.TakeWhile(char.IsLetter).ToArray());
        string digits = new(ref_.SkipWhile(char.IsLetter).ToArray());
        int col = 0;
        foreach (char c in letters.ToUpperInvariant())
            col = col * 26 + (c - 'A' + 1);
        int row = int.TryParse(digits, out var r) ? r : 1;
        return (col, row);
    }
}
