namespace ExcelAPI.Designer.Legacy.Models;

public class WorkbookModel
{
    public string FilePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public List<SheetModel> Sheets { get; set; } = new();
    public int SheetCount => Sheets.Count;
}
