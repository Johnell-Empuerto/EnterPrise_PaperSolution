using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelAPI.Designer.Legacy.OpenXml;

public class MergeCellReader : IMergeCellReader
{
    public IReadOnlyList<OpenXmlMergeAreaModel> Read(string filePath, string sheetName, int sheetId)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = doc.WorkbookPart;
        if (workbookPart is null) return Array.Empty<OpenXmlMergeAreaModel>();

        var sheet = workbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName);
        if (sheet is null) return Array.Empty<OpenXmlMergeAreaModel>();

        var wsPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart;
        var ws = wsPart?.Worksheet;
        if (ws is null) return Array.Empty<OpenXmlMergeAreaModel>();

        var mergeCells = ws.Descendants<MergeCells>().FirstOrDefault();
        if (mergeCells is null) return Array.Empty<OpenXmlMergeAreaModel>();

        var result = new List<OpenXmlMergeAreaModel>();

        foreach (var mc in mergeCells.Cast<MergeCell>())
        {
            string? reference = mc.Reference?.Value;
            if (string.IsNullOrEmpty(reference)) continue;

            var parts = reference.Split(':');
            if (parts.Length < 1) continue;

            var start = ParseCellRef(parts[0]);
            if (parts.Length == 1)
            {
                result.Add(new OpenXmlMergeAreaModel
                {
                    StartRow = start.row, StartColumn = start.col,
                    EndRow = start.row, EndColumn = start.col
                });
            }
            else
            {
                var end = ParseCellRef(parts[1]);
                result.Add(new OpenXmlMergeAreaModel
                {
                    StartRow = start.row, StartColumn = start.col,
                    EndRow = end.row, EndColumn = end.col
                });
            }
        }

        return result;
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
