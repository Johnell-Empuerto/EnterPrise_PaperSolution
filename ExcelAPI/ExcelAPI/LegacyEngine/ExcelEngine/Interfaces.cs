using ExcelAPI.LegacyEngine.Models;

namespace ExcelAPI.LegacyEngine.ExcelEngine;

public interface IWorkbookLoader
{
    WorkbookModel Load(string filePath);
}

public interface IWorksheetLoader
{
    SheetModel Load(dynamic worksheet, int sheetIndex);
}

public interface IPrintAreaLoader
{
    PrintAreaInfo Load(dynamic worksheet);
    PageSetupModel LoadPageSetup(dynamic worksheet);
}
