using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelAPI.Designer.Legacy.OpenXml;

public class CellDataReader : ICellDataReader
{
    public IReadOnlyList<OpenXmlCellModel> Read(string filePath, string sheetName, int sheetId)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = doc.WorkbookPart;
        if (workbookPart is null) return Array.Empty<OpenXmlCellModel>();

        var sheet = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName);
        if (sheet is null) return Array.Empty<OpenXmlCellModel>();

        var wsPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart;
        if (wsPart is null) return Array.Empty<OpenXmlCellModel>();

        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;

        var result = new List<OpenXmlCellModel>();
        var rows = wsPart.Worksheet.Descendants<Row>();
        foreach (var row in rows)
        {
            foreach (var cell in row.Descendants<Cell>())
            {
                var (col, rowNum) = ParseCellRef(cell.CellReference?.Value ?? "");
                if (rowNum == 0 || col == 0) continue;

                string? value = GetCellValue(cell, sharedStringTable);
                if (string.IsNullOrEmpty(value)) continue;

                int styleIndex = cell.StyleIndex is not null ? (int)cell.StyleIndex.Value : -1;

                result.Add(new OpenXmlCellModel
                {
                    Row = rowNum,
                    Column = col,
                    Value = value,
                    StyleIndex = styleIndex
                });
            }
        }

        return result;
    }

    private static string? GetCellValue(Cell cell, SharedStringTable? sharedStringTable)
    {
        if (cell.CellValue is null && cell.InlineString is null) return null;

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            if (cell.CellValue?.Text is null) return null;
            if (!int.TryParse(cell.CellValue.Text, out int sstIndex)) return null;
            if (sharedStringTable is null || sstIndex < 0 || sstIndex >= sharedStringTable.Count())
                return null;
            return sharedStringTable.ElementAt(sstIndex).InnerText;
        }

        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString?.Text?.Text ?? "";
        }

        return cell.CellValue?.Text ?? "";
    }

    private static (int col, int row) ParseCellRef(string ref_)
    {
        if (string.IsNullOrEmpty(ref_)) return (0, 0);
        string letters = new(ref_.TakeWhile(char.IsLetter).ToArray());
        string digits = new(ref_.SkipWhile(char.IsLetter).ToArray());
        if (letters.Length == 0) return (0, 0);
        int col = 0;
        foreach (char c in letters.ToUpperInvariant())
            col = col * 26 + (c - 'A' + 1);
        int row = int.TryParse(digits, out var r) ? r : 0;
        return (col, row);
    }
}
