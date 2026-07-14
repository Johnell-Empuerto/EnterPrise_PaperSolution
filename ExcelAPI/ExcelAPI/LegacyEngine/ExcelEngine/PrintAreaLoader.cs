using ExcelAPI.LegacyEngine.Models;

namespace ExcelAPI.LegacyEngine.ExcelEngine;

public class PrintAreaLoader : IPrintAreaLoader
{
    private const double DefaultMarginInch = 0.75;

    public PrintAreaInfo Load(dynamic worksheet)
    {
        string? printArea = null;
        try { printArea = worksheet.PageSetup.PrintArea as string; } catch { }

        if (string.IsNullOrEmpty(printArea))
        {
            string? usedAddress = null;
            try { usedAddress = worksheet.UsedRange.Address as string; } catch { }
            if (!string.IsNullOrEmpty(usedAddress))
            {
                printArea = usedAddress;
            }
        }

        if (string.IsNullOrEmpty(printArea))
            return new PrintAreaInfo();

        var info = new PrintAreaInfo { Address = printArea };
        var parts = printArea.Replace("$", "").Split(':');

        if (parts.Length == 2)
        {
            var start = ParseCellRef(parts[0]);
            var end = ParseCellRef(parts[1]);
            info.StartColumn = start.col;
            info.StartRow = start.row;
            info.EndColumn = end.col;
            info.EndRow = end.row;
        }
        else if (parts.Length == 1)
        {
            var p = ParseCellRef(parts[0]);
            info.StartColumn = info.EndColumn = p.col;
            info.StartRow = info.EndRow = p.row;
        }

        return info;
    }

    public PageSetupModel LoadPageSetup(dynamic worksheet)
    {
        var ps = new PageSetupModel();

        try
        {
            dynamic pageSetup = worksheet.PageSetup;

            ps.PageWidthPoints = (double)pageSetup.PageSetup.PageWidth;
            ps.PageHeightPoints = (double)pageSetup.PageSetup.PageHeight;
            ps.PaperWidthPoints = (double)pageSetup.PageSetup.PaperWidth;
            ps.PaperHeightPoints = (double)pageSetup.PageSetup.PaperHeight;

            ps.MarginLeft = (double)pageSetup.LeftMargin / 72.0;
            ps.MarginRight = (double)pageSetup.RightMargin / 72.0;
            ps.MarginTop = (double)pageSetup.TopMargin / 72.0;
            ps.MarginBottom = (double)pageSetup.BottomMargin / 72.0;
            ps.MarginHeader = (double)pageSetup.HeaderMargin / 72.0;
            ps.MarginFooter = (double)pageSetup.FooterMargin / 72.0;

            ps.CenterHorizontally = (bool)pageSetup.CenterHorizontally;
            ps.CenterVertically = (bool)pageSetup.CenterVertically;

            int? zoom = null;
            try { zoom = (int)pageSetup.Zoom; } catch { }
            if (zoom == 0) zoom = null;
            ps.Zoom = zoom;

            int? fitW = null;
            try { fitW = (int)pageSetup.FitToPagesWide; } catch { }
            if (fitW == 0) fitW = null;
            ps.FitToPagesWide = fitW;

            int? fitH = null;
            try { fitH = (int)pageSetup.FitToPagesTall; } catch { }
            if (fitH == 0) fitH = null;
            ps.FitToPagesTall = fitH;

            int orient = 1;
            try { orient = (int)pageSetup.Orientation; } catch { }
            ps.Orientation = orient == 2 ? "Landscape" : "Portrait";
        }
        catch
        {
            // Defaults are fine
        }

        return ps;
    }

    private static (int col, int row) ParseCellRef(string ref_)
    {
        string letters = new(ref_.TakeWhile(char.IsLetter).ToArray());
        string digits = new(ref_.SkipWhile(char.IsLetter).ToArray());

        int col = 0;
        foreach (char c in letters.ToUpperInvariant())
            col = col * 26 + (c - 'A' + 1);

        int row = int.TryParse(digits, out var r) ? r : 1;
        return (col, row);
    }
}
