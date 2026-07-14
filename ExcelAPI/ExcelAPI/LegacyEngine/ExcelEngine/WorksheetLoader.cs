using ExcelAPI.LegacyEngine.Models;

namespace ExcelAPI.LegacyEngine.ExcelEngine;

public class WorksheetLoader : IWorksheetLoader
{
    private readonly IPrintAreaLoader _printAreaLoader;

    public WorksheetLoader(IPrintAreaLoader printAreaLoader)
    {
        _printAreaLoader = printAreaLoader;
    }

    public SheetModel Load(dynamic worksheet, int sheetIndex)
    {
        var printArea = _printAreaLoader.Load(worksheet);
        if (!printArea.IsValid)
        {
            throw new InvalidOperationException(
                $"Sheet '{worksheet.Name}' has no PrintArea configured.");
        }

        var pageSetup = _printAreaLoader.LoadPageSetup(worksheet);
        string[,]? cellValues = null;
        try { cellValues = worksheet.UsedRange.Value as string[,]; } catch { }

        var model = new SheetModel
        {
            Name = worksheet.Name,
            SheetIndex = sheetIndex,
            PrintArea = printArea,
            PageSetup = pageSetup
        };

        for (int c = printArea.StartColumn; c <= printArea.EndColumn; c++)
        {
            double width;
            try { width = (double)(worksheet.Columns[c] as dynamic).Width; }
            catch { width = 48.0; }
            model.Columns.Add(new ColumnModel { Index = c, WidthPoints = width });
        }

        for (int r = printArea.StartRow; r <= printArea.EndRow; r++)
        {
            double height;
            try { height = (double)(worksheet.Rows[r] as dynamic).Height; }
            catch { height = 14.4; }
            model.Rows.Add(new RowModel { Index = r, HeightPoints = height });
        }

        // Read comments (legacy cluster definition source)
        try
        {
            dynamic worksheetComments = worksheet.Comments;
            if (worksheetComments != null)
            {
                int commentCount = 0;
                try { commentCount = worksheetComments.Count; } catch { }
                if (commentCount > 0)
                {
                    for (int ci = 0; ci < commentCount; ci++)
                    {
                        dynamic comment = worksheetComments[ci + 1];
                        dynamic parent = comment.Parent;
                        dynamic mergeArea;
                        try { mergeArea = parent.MergeArea; } catch { mergeArea = null; }
                        if (mergeArea == null) mergeArea = parent;

                        string commentText = "";
                        try { commentText = comment.Text() ?? ""; } catch { }

                        model.Comments.Add(new CommentModel
                        {
                            Row = mergeArea.Row,
                            Column = mergeArea.Column,
                            RowCount = mergeArea.Rows.Count,
                            ColumnCount = mergeArea.Columns.Count,
                            Text = commentText
                        });
                    }
                }
            }
        }
        catch
        {
            // Comments are optional - proceed if unavailable
        }

        for (int r = printArea.StartRow; r <= printArea.EndRow; r++)
        {
            for (int c = printArea.StartColumn; c <= printArea.EndColumn; c++)
            {
                dynamic cell = worksheet.Cells[r, c];
                string? val = null;
                try { val = cell.Value?.ToString(); } catch { }

                bool isMerged = false;
                try { isMerged = (bool)cell.MergeCells; } catch { }

                if (isMerged)
                {
                    dynamic mergeArea = cell.MergeArea;
                    int mRow = mergeArea.Row;
                    int mCol = mergeArea.Column;
                    int mRows = mergeArea.Rows.Count;
                    int mCols = mergeArea.Columns.Count;

                    bool alreadyAdded = model.MergedCells.Any(m =>
                        m.Row == mRow && m.Column == mCol);
                    if (!alreadyAdded)
                    {
                        var mc = new CellModel
                        {
                            Row = mRow,
                            Column = mCol,
                            EndRow = mRow + mRows - 1,
                            EndColumn = mCol + mCols - 1,
                            Value = val
                        };
                        model.MergedCells.Add(mc);
                        model.Cells.Add(mc);
                    }
                }
                else
                {
                    model.Cells.Add(new CellModel
                    {
                        Row = r,
                        Column = c,
                        EndRow = r,
                        EndColumn = c,
                        Value = val
                    });
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(cell);
            }
        }

        return model;
    }
}
