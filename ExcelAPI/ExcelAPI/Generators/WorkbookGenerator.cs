using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelAPI.Models;
using Microsoft.Extensions.Logging;

namespace ExcelAPI.Generators
{
    public class WorkbookGenerator
    {
        private readonly ILogger<WorkbookGenerator> _logger;

        public WorkbookGenerator(ILogger<WorkbookGenerator> logger)
        {
            _logger = logger;
        }

        public string Generate(FormDefinition form, string outputPath)
        {
            Application? excelApp = null;
            Workbook? workbook = null;

            try
            {
                excelApp = new Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                workbook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                // Remove default extra sheets if any
                while (workbook.Worksheets.Count > form.Sheets.Count)
                {
                    ((Excel.Worksheet)workbook.Worksheets[workbook.Worksheets.Count]).Delete();
                }

                // Add sheets if needed
                while (workbook.Worksheets.Count < form.Sheets.Count)
                {
                    var newSheet = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                    Marshal.ReleaseComObject(newSheet);
                }

                for (int i = 0; i < form.Sheets.Count; i++)
                {
                    var sheetDef = form.Sheets[i];
                    var ws = (Excel.Worksheet)workbook.Worksheets[i + 1];

                    ws.Name = sheetDef.Name;

                    ApplyPageSettings(ws, sheetDef.PageSettings);
                    ApplyPrintArea(ws, sheetDef.PrintArea);
                    ApplyRowHeights(ws, sheetDef.RowHeights);
                    ApplyColumnWidths(ws, sheetDef.ColumnWidths);
                    ApplyMergedCells(ws, sheetDef.MergedCells, sheetDef.CellStyles);
                    ApplyFreezePane(ws, sheetDef.FreezePane);
                    ApplyCellStyles(ws, sheetDef.CellStyles);
                }

                workbook.SaveAs(outputPath);
                _logger.LogInformation("Workbook saved to {Path}", outputPath);

                return outputPath;
            }
            finally
            {
                if (workbook != null)
                {
                    try { workbook.Close(SaveChanges: false); } catch { }
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    try { excelApp.Quit(); } catch { }
                    Marshal.ReleaseComObject(excelApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void ApplyPageSettings(Excel.Worksheet ws, PageSettings settings)
        {
            ws.PageSetup.Orientation = settings.Orientation == "landscape"
                ? XlPageOrientation.xlLandscape
                : XlPageOrientation.xlPortrait;
            ws.PageSetup.LeftMargin = settings.LeftMargin;
            ws.PageSetup.TopMargin = settings.TopMargin;
            ws.PageSetup.RightMargin = settings.RightMargin;
            ws.PageSetup.BottomMargin = settings.BottomMargin;
            ws.PageSetup.CenterHorizontally = settings.CenterHorizontally;
            ws.PageSetup.CenterVertically = settings.CenterVertically;

            if (settings.Zoom > 0)
                ws.PageSetup.Zoom = settings.Zoom;

            if (settings.FitToPagesWide > 0)
                ws.PageSetup.FitToPagesWide = settings.FitToPagesWide;
            if (settings.FitToPagesTall > 0)
                ws.PageSetup.FitToPagesTall = settings.FitToPagesTall;
        }

        private void ApplyPrintArea(Excel.Worksheet ws, PrintAreaInfo? printArea)
        {
            if (printArea != null && !string.IsNullOrEmpty(printArea.Address))
            {
                ws.PageSetup.PrintArea = printArea.Address;
            }
        }

        private void ApplyRowHeights(Excel.Worksheet ws, Dictionary<int, double> rowHeights)
        {
            foreach (var kv in rowHeights)
            {
                try
                {
                    var row = (Excel.Range)ws.Rows[kv.Key];
                    row.RowHeight = kv.Value;
                    Marshal.ReleaseComObject(row);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to set row {Row} height", kv.Key);
                }
            }
        }

        private void ApplyColumnWidths(Excel.Worksheet ws, Dictionary<int, double> columnWidths)
        {
            foreach (var kv in columnWidths)
            {
                try
                {
                    var col = (Excel.Range)ws.Columns[kv.Key];
                    col.ColumnWidth = kv.Value;
                    Marshal.ReleaseComObject(col);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to set column {Col} width", kv.Key);
                }
            }
        }

        private void ApplyMergedCells(Excel.Worksheet ws, List<MergedCellInfo> mergedCells, Dictionary<string, CellStyleInfo> cellStyles)
        {
            foreach (var mc in mergedCells)
            {
                try
                {
                    var range = ws.Range[mc.Address];
                    range.Merge();

                    if (cellStyles.TryGetValue(mc.CellAddress, out var style))
                    {
                        ApplyStyleToRange(range, style);
                    }

                    Marshal.ReleaseComObject(range);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to merge cells {Address}", mc.Address);
                }
            }
        }

        private void ApplyFreezePane(Excel.Worksheet ws, string? freezePane)
        {
            if (!string.IsNullOrEmpty(freezePane))
            {
                try
                {
                    ws.Activate();
                    var range = ws.Range[freezePane];
                    range.Activate();
                    ws.Application.ActiveWindow.FreezePanes = true;
                    Marshal.ReleaseComObject(range);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to set freeze pane at {Pane}", freezePane);
                }
            }
        }

        private void ApplyCellStyles(Excel.Worksheet ws, Dictionary<string, CellStyleInfo> cellStyles)
        {
            foreach (var kv in cellStyles)
            {
                try
                {
                    var range = ws.Range[kv.Key];
                    ApplyStyleToRange(range, kv.Value);
                    Marshal.ReleaseComObject(range);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to apply style to cell {Cell}", kv.Key);
                }
            }
        }

        private static void ApplyStyleToRange(Excel.Range range, CellStyleInfo style)
        {
            if (style.FontName != null)
                range.Font.Name = style.FontName;
            if (style.FontSize.HasValue)
                range.Font.Size = style.FontSize.Value;
            if (style.Bold.HasValue)
                range.Font.Bold = style.Bold.Value;
            if (style.Italic.HasValue)
                range.Font.Italic = style.Italic.Value;
            if (style.Underline.HasValue && style.Underline.Value)
                range.Font.Underline = true;
            if (style.Color != null)
            {
                try
                {
                    var (r, g, b) = ParseColor(style.Color);
                    range.Font.Color = System.Drawing.Color.FromArgb(r, g, b);
                }
                catch { }
            }
            if (style.FillColor != null)
            {
                try
                {
                    var (r, g, b) = ParseColor(style.FillColor);
                    range.Interior.Color = System.Drawing.Color.FromArgb(r, g, b);
                }
                catch { }
            }
            if (style.HorizontalAlignment != null)
            {
                range.HorizontalAlignment = style.HorizontalAlignment switch
                {
                    "left" => XlHAlign.xlHAlignLeft,
                    "center" => XlHAlign.xlHAlignCenter,
                    "right" => XlHAlign.xlHAlignRight,
                    _ => XlHAlign.xlHAlignGeneral
                };
            }
            if (style.VerticalAlignment != null)
            {
                range.VerticalAlignment = style.VerticalAlignment switch
                {
                    "top" => XlVAlign.xlVAlignTop,
                    "center" => XlVAlign.xlVAlignCenter,
                    "bottom" => XlVAlign.xlVAlignBottom,
                    _ => XlVAlign.xlVAlignBottom
                };
            }
            if (style.WrapText.HasValue)
                range.WrapText = style.WrapText.Value;
        }

        private static (int r, int g, int b) ParseColor(string color)
        {
            if (color.StartsWith("#"))
                color = color[1..];

            if (color.Length == 6 && int.TryParse(color, System.Globalization.NumberStyles.HexNumber, null, out var rgb))
            {
                return ((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
            }

            return (0, 0, 0);
        }
    }
}
