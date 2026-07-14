namespace ExcelAPI.LegacyEngine.OpenXml;

public class OpenXmlWorkbookLoader : IOpenXmlWorkbookLoader
{
    private readonly IOpenXmlWorksheetLoader _sheetLoader;

    public OpenXmlWorkbookLoader(IOpenXmlWorksheetLoader sheetLoader)
    {
        _sheetLoader = sheetLoader;
    }

    public OpenXmlWorkbookModel Load(string filePath)
    {
        var sheets = _sheetLoader.LoadSheets(filePath);

        return new OpenXmlWorkbookModel
        {
            FilePath = filePath,
            FileName = Path.GetFileName(filePath),
            Sheets = sheets
        };
    }
}
