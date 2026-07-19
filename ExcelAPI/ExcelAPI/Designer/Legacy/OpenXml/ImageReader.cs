using DocumentFormat.OpenXml.Packaging;
using DSS = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using D = DocumentFormat.OpenXml.Drawing;

namespace ExcelAPI.Designer.Legacy.OpenXml;

public class ImageReader : IImageReader
{
    public IReadOnlyList<OpenXmlImageModel> Read(string filePath, string sheetName, int sheetId)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = doc.WorkbookPart;
        if (workbookPart is null) return Array.Empty<OpenXmlImageModel>();

        var sheet = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName);
        if (sheet is null) return Array.Empty<OpenXmlImageModel>();

        var wsPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart;
        if (wsPart is null) return Array.Empty<OpenXmlImageModel>();

        var dp = wsPart.DrawingsPart;
        if (dp is null) return Array.Empty<OpenXmlImageModel>();

        var wsDr = dp.WorksheetDrawing;
        if (wsDr is null) return Array.Empty<OpenXmlImageModel>();

        var result = new List<OpenXmlImageModel>();

        foreach (var anchor in wsDr.ChildElements)
        {
            int colFrom = 0, rowFrom = 0, colOffPx = 0, rowOffPx = 0;

            if (anchor.LocalName == "oneCellAnchor" || anchor.LocalName == "twoCellAnchor")
            {
                var from = anchor.GetFirstChild<DSS.FromMarker>();
                if (from is not null)
                {
                    uint.TryParse(from.ColumnId?.Text, out uint ac);
                    uint.TryParse(from.RowId?.Text, out uint ar);
                    long.TryParse(from.ColumnOffset?.Text, out long co);
                    long.TryParse(from.RowOffset?.Text, out long ro);

                    colFrom = (int)ac + 1;
                    rowFrom = (int)ar + 1;
                    colOffPx = (int)(co / 914400.0 * 96);
                    rowOffPx = (int)(ro / 914400.0 * 96);
                }
            }

            // Check if this anchor has a picture
            var picture = anchor.GetFirstChild<DSS.Picture>();
            if (picture is null) continue;

            var blipFill = picture.GetFirstChild<DSS.BlipFill>();
            var blip = blipFill?.GetFirstChild<D.Blip>();
            string? embedId = blip?.Embed?.Value;
            if (string.IsNullOrEmpty(embedId)) continue;

            var imagePart = dp.GetPartById(embedId) as ImagePart;
            if (imagePart is null) continue;

            long bytesLength = 0;
            try { using var s = imagePart.GetStream(); bytesLength = s.Length; } catch { }

            result.Add(new OpenXmlImageModel
            {
                Row = rowFrom,
                Column = colFrom,
                RowOffsetPx = rowOffPx,
                ColumnOffsetPx = colOffPx,
                ImageBytesLength = bytesLength,
                ContentType = imagePart.ContentType ?? "",
                FileName = imagePart.Uri?.OriginalString ?? ""
            });
        }

        return result;
    }
}
