using DocumentFormat.OpenXml.Packaging;

namespace ExcelAPI.Designer.Legacy.OpenXml;

public class OpenXmlWorksheetLoader : IOpenXmlWorksheetLoader
{
    private readonly IStyleReader _styleReader;
    private readonly IPageSetupReader _pageSetupReader;
    private readonly IPrintAreaReader _printAreaReader;
    private readonly IColumnWidthReader _columnWidthReader;
    private readonly IRowHeightReader _rowHeightReader;
    private readonly IMergeCellReader _mergeCellReader;
    private readonly ICommentReader _commentReader;
    private readonly IImageReader _imageReader;
    private readonly ICellDataReader _cellDataReader;
    private readonly ICellStyleReader _cellStyleReader;

    public OpenXmlWorksheetLoader(
        IStyleReader styleReader,
        IPageSetupReader pageSetupReader,
        IPrintAreaReader printAreaReader,
        IColumnWidthReader columnWidthReader,
        IRowHeightReader rowHeightReader,
        IMergeCellReader mergeCellReader,
        ICommentReader commentReader,
        IImageReader imageReader,
        ICellDataReader cellDataReader,
        ICellStyleReader cellStyleReader)
    {
        _styleReader = styleReader;
        _pageSetupReader = pageSetupReader;
        _printAreaReader = printAreaReader;
        _columnWidthReader = columnWidthReader;
        _rowHeightReader = rowHeightReader;
        _mergeCellReader = mergeCellReader;
        _commentReader = commentReader;
        _imageReader = imageReader;
        _cellDataReader = cellDataReader;
        _cellStyleReader = cellStyleReader;
    }

    public List<OpenXmlSheetModel> LoadSheets(string filePath)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = doc.WorkbookPart;
        if (workbookPart is null) return new List<OpenXmlSheetModel>();

        var styles = _styleReader.Read(filePath);
        var allCellStyles = _cellStyleReader.Read(filePath);

        var result = new List<OpenXmlSheetModel>();
        int sheetIndex = 0;

        foreach (var sheet in workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>())
        {
            // Skip hidden sheets
            if (sheet.State is not null &&
                sheet.State?.Value == DocumentFormat.OpenXml.Spreadsheet.SheetStateValues.Hidden)
                continue;

            sheetIndex++;
            string name = sheet.Name?.Value ?? $"Sheet{sheetIndex}";
            int id = 0;
            if (sheet.SheetId?.Value is not null)
                id = (int)sheet.SheetId.Value;
            if (id == 0) id = sheetIndex;
            string relId = sheet.Id ?? "";

            var printArea = _printAreaReader.Read(filePath, name);
            var pageSetup = _pageSetupReader.Read(filePath, name, id);
            var columnWidths = _columnWidthReader.Read(filePath, name, id, styles.MaxDigitWidth);
            var rowHeights = _rowHeightReader.Read(filePath, name, id);
            var mergeAreas = _mergeCellReader.Read(filePath, name, id);
            var comments = _commentReader.Read(filePath, name, id);
            var images = _imageReader.Read(filePath, name, id);
            var cells = _cellDataReader.Read(filePath, name, id);

            // Build address-to-style map
            var cellStylesByAddress = new Dictionary<string, OpenXmlCellStyleModel>();
            foreach (var cell in cells)
            {
                if (cell.StyleIndex >= 0 && allCellStyles.TryGetValue(cell.StyleIndex, out var cellStyle))
                {
                    string address = GetCellAddress(cell.Column, cell.Row);
                    cellStylesByAddress[address] = cellStyle;
                }
            }

            result.Add(new OpenXmlSheetModel
            {
                Name = name,
                SheetIndex = sheetIndex,
                SheetId = id,
                RelationshipId = relId,
                PrintArea = printArea,
                PageSetup = pageSetup,
                ColumnWidths = columnWidths,
                RowHeights = rowHeights,
                MergeAreas = mergeAreas,
                Comments = comments,
                Images = images,
                Cells = cells,
                CellStyles = cellStylesByAddress
            });
        }

        return result;
    }

    private static string GetCellAddress(int col, int row)
    {
        string colLetter = "";
        int c = col;
        while (c > 0)
        {
            c--;
            colLetter = (char)('A' + c % 26) + colLetter;
            c /= 26;
        }
        return $"{colLetter}{row}";
    }
}
