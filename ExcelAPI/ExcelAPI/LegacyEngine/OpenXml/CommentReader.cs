using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace ExcelAPI.LegacyEngine.OpenXml;

public class CommentReader : ICommentReader
{
    public IReadOnlyList<OpenXmlCommentModel> Read(string filePath, string sheetName, int sheetId)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = doc.WorkbookPart;
        if (workbookPart is null) return Array.Empty<OpenXmlCommentModel>();

        var sheet = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName);
        if (sheet is null) return Array.Empty<OpenXmlCommentModel>();

        var wsPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart;
        if (wsPart is null) return Array.Empty<OpenXmlCommentModel>();

        var commentsPart = wsPart.WorksheetCommentsPart;
        if (commentsPart is null) return Array.Empty<OpenXmlCommentModel>();

        using var stream = commentsPart.GetStream();
        var xDoc = XDocument.Load(stream);
        XNamespace ns = xDoc.Root?.GetDefaultNamespace() ?? XNamespace.None;

        var result = new List<OpenXmlCommentModel>();
        var commentList = xDoc.Descendants(ns + "commentList").FirstOrDefault();
        if (commentList is null) return result;

        foreach (var comment in commentList.Elements(ns + "comment"))
        {
            string? refStr = comment.Attribute("ref")?.Value;
            if (string.IsNullOrEmpty(refStr)) continue;

            var (col, row) = ParseCellRef(refStr);

            var textBuilder = new System.Text.StringBuilder();
            foreach (var t in comment.Descendants(ns + "t"))
                textBuilder.Append(t.Value);

            string commentText = textBuilder.ToString().Trim();
            if (string.IsNullOrEmpty(commentText)) continue;

            result.Add(new OpenXmlCommentModel
            {
                Row = row,
                Column = col,
                RowCount = 1,
                ColumnCount = 1,
                Text = commentText
            });
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
