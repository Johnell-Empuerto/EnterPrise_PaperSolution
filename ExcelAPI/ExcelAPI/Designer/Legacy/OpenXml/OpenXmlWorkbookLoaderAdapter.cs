using ExcelAPI.Designer.Legacy.ExcelEngine;
using ExcelAPI.Designer.Legacy.Models;

namespace ExcelAPI.Designer.Legacy.OpenXml;

public class OpenXmlWorkbookLoaderAdapter : IWorkbookLoader
{
    private readonly IOpenXmlWorkbookLoader _openXmlLoader;

    public OpenXmlWorkbookLoaderAdapter(IOpenXmlWorkbookLoader openXmlLoader)
    {
        _openXmlLoader = openXmlLoader;
    }

    public WorkbookModel Load(string filePath)
    {
        var openXml = _openXmlLoader.Load(filePath);

        var model = new WorkbookModel
        {
            FilePath = openXml.FilePath,
            FileName = openXml.FileName
        };

        foreach (var openSheet in openXml.Sheets)
        {
            var sheet = new SheetModel
            {
                Name = openSheet.Name,
                SheetIndex = openSheet.SheetIndex,
                PrintArea = new PrintAreaInfo
                {
                    Address = openSheet.PrintArea.IsValid
                        ? $"${CellModel.GetColumnLetter(openSheet.PrintArea.StartColumn)}${openSheet.PrintArea.StartRow}:${CellModel.GetColumnLetter(openSheet.PrintArea.EndColumn)}${openSheet.PrintArea.EndRow}"
                        : "",
                    StartColumn = openSheet.PrintArea.StartColumn,
                    StartRow = openSheet.PrintArea.StartRow,
                    EndColumn = openSheet.PrintArea.EndColumn,
                    EndRow = openSheet.PrintArea.EndRow
                },
                PageSetup = new PageSetupModel
                {
                    PageWidthPoints = openSheet.PageSetup.PageWidthPoints,
                    PageHeightPoints = openSheet.PageSetup.PageHeightPoints,
                    PaperWidthPoints = openSheet.PageSetup.PaperWidthPoints,
                    PaperHeightPoints = openSheet.PageSetup.PaperHeightPoints,
                    MarginLeft = openSheet.PageSetup.MarginLeft,
                    MarginRight = openSheet.PageSetup.MarginRight,
                    MarginTop = openSheet.PageSetup.MarginTop,
                    MarginBottom = openSheet.PageSetup.MarginBottom,
                    MarginHeader = openSheet.PageSetup.MarginHeader,
                    MarginFooter = openSheet.PageSetup.MarginFooter,
                    CenterHorizontally = openSheet.PageSetup.CenterHorizontally,
                    CenterVertically = openSheet.PageSetup.CenterVertically,
                    Orientation = openSheet.PageSetup.Orientation
                }
            };

            // Columns
            for (int i = 0; i < openSheet.ColumnWidths.Count; i++)
            {
                sheet.Columns.Add(new ColumnModel
                {
                    Index = i + 1,
                    WidthPoints = openSheet.ColumnWidths[i]
                });
            }

            // Rows
            for (int i = 0; i < openSheet.RowHeights.Count; i++)
            {
                sheet.Rows.Add(new RowModel
                {
                    Index = i + 1,
                    HeightPoints = openSheet.RowHeights[i]
                });
            }

            // Comments
            foreach (var c in openSheet.Comments)
            {
                sheet.Comments.Add(new CommentModel
                {
                    Row = c.Row,
                    Column = c.Column,
                    RowCount = c.RowCount,
                    ColumnCount = c.ColumnCount,
                    Text = c.Text
                });
            }

            // Merged cells (only those that have associated comments for cluster detection)
            foreach (var m in openSheet.MergeAreas)
            {
                bool hasComment = openSheet.Comments.Any(c =>
                    c.Row == m.StartRow && c.Column == m.StartColumn);

                sheet.MergedCells.Add(new CellModel
                {
                    Row = m.StartRow,
                    Column = m.StartColumn,
                    EndRow = m.EndRow,
                    EndColumn = m.EndColumn
                });

                // Non-cluster merged cells still need to be in the Cells list
                if (!hasComment)
                {
                    sheet.Cells.Add(new CellModel
                    {
                        Row = m.StartRow,
                        Column = m.StartColumn,
                        EndRow = m.EndRow,
                        EndColumn = m.EndColumn
                    });
                }
            }

            model.Sheets.Add(sheet);
        }

        return model;
    }
}
