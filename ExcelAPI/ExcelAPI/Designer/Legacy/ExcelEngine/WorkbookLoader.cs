using ExcelAPI.Designer.Legacy.Models;
using Microsoft.Office.Interop.Excel;

namespace ExcelAPI.Designer.Legacy.ExcelEngine;

public class WorkbookLoader : IWorkbookLoader
{
    private readonly IWorksheetLoader _sheetLoader;

    public WorkbookLoader(IWorksheetLoader sheetLoader)
    {
        _sheetLoader = sheetLoader;
    }

    public WorkbookModel Load(string filePath)
    {
        App excel = new();
        Workbooks? workbooks = null;
        _Workbook? workbook = null;

        try
        {
            workbooks = excel.Workbooks;
            workbook = workbooks.Open(filePath);

            var model = new WorkbookModel
            {
                FilePath = filePath,
                FileName = Path.GetFileName(filePath)
            };

            int sheetIndex = 0;
            foreach (Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Visible == XlSheetVisibility.xlSheetVisible)
                {
                    sheetIndex++;
                    var sheetModel = _sheetLoader.Load(sheet, sheetIndex);
                    model.Sheets.Add(sheetModel);
                }
            }

            return model;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to load workbook: {ex.Message}", ex);
        }
        finally
        {
            workbook?.Close(false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            workbooks?.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
    }
}
