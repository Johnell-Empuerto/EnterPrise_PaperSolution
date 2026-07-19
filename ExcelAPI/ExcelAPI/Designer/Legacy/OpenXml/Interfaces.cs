namespace ExcelAPI.Designer.Legacy.OpenXml;

public interface IOpenXmlWorkbookLoader
{
    OpenXmlWorkbookModel Load(string filePath);
}

public interface IOpenXmlWorksheetLoader
{
    List<OpenXmlSheetModel> LoadSheets(string filePath);
}

public interface IStyleReader
{
    OpenXmlStyleModel Read(string filePath);
}

public interface IPageSetupReader
{
    OpenXmlPageSetupModel Read(string filePath, string sheetName, int sheetId);
}

public interface IPrintAreaReader
{
    OpenXmlPrintAreaModel Read(string filePath, string sheetName);
}

public interface IColumnWidthReader
{
    IReadOnlyList<double> Read(string filePath, string sheetName, int sheetId, double maxDigitWidth);
}

public interface IRowHeightReader
{
    IReadOnlyList<double> Read(string filePath, string sheetName, int sheetId);
}

public interface IMergeCellReader
{
    IReadOnlyList<OpenXmlMergeAreaModel> Read(string filePath, string sheetName, int sheetId);
}

public interface ICommentReader
{
    IReadOnlyList<OpenXmlCommentModel> Read(string filePath, string sheetName, int sheetId);
}

public interface IImageReader
{
    IReadOnlyList<OpenXmlImageModel> Read(string filePath, string sheetName, int sheetId);
}

public interface ICellDataReader
{
    IReadOnlyList<OpenXmlCellModel> Read(string filePath, string sheetName, int sheetId);
}

public interface ICellStyleReader
{
    IReadOnlyDictionary<int, OpenXmlCellStyleModel> Read(string filePath);
}
