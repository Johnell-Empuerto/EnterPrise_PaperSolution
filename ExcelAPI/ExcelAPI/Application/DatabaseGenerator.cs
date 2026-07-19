using ExcelAPI.Models;

namespace ExcelAPI.Application
{
    public class DefTop
    {
        public int Id { get; set; }
        public string Title { get; set; } = "";
        public string Author { get; set; } = "";
        public string Created { get; set; } = "";
        public string Modified { get; set; } = "";
        public string Version { get; set; } = "";
        public string Description { get; set; } = "";
        public string Metadata { get; set; } = "{}";
    }

    public class DefSheet
    {
        public int Id { get; set; }
        public int TopId { get; set; }
        public string SheetId { get; set; } = "";
        public string Name { get; set; } = "";
        public int Index { get; set; }
        public string PaperSize { get; set; } = "";
        public string Orientation { get; set; } = "";
        public double WidthPt { get; set; }
        public double HeightPt { get; set; }
        public double LeftMargin { get; set; }
        public double TopMargin { get; set; }
        public double RightMargin { get; set; }
        public double BottomMargin { get; set; }
        public bool CenterHorizontally { get; set; }
        public bool CenterVertically { get; set; }
        public int Zoom { get; set; }
        public int FitToPagesWide { get; set; }
        public int FitToPagesTall { get; set; }
        public string? PrintAreaAddress { get; set; }
        public double? PrintAreaLeftPt { get; set; }
        public double? PrintAreaTopPt { get; set; }
        public double? PrintAreaWidthPt { get; set; }
        public double? PrintAreaHeightPt { get; set; }
        public int? PrintAreaCols { get; set; }
        public int? PrintAreaRows { get; set; }
        public string? BackgroundImage { get; set; }
        public string? FreezePane { get; set; }
        public string RowHeights { get; set; } = "{}";
        public string ColumnWidths { get; set; } = "{}";
        public string MergedCells { get; set; } = "[]";
    }

    public class DefCluster
    {
        public int Id { get; set; }
        public int TopId { get; set; }
        public string ClusterId { get; set; } = "";
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public string SheetId { get; set; } = "";
        public string CellAddress { get; set; } = "";
        public double Left { get; set; }
        public double Right { get; set; }
        public double Top { get; set; }
        public double Bottom { get; set; }
        public double LeftPt { get; set; }
        public double TopPt { get; set; }
        public double WidthPt { get; set; }
        public double HeightPt { get; set; }
        public string InputParameters { get; set; } = "{}";
        public string Visibility { get; set; } = "visible";
        public bool Readonly { get; set; }
        public string Remarks { get; set; } = "";
        public string Functions { get; set; } = "[]";
        public string Metadata { get; set; } = "{}";
    }

    public class DatabaseResult
    {
        public DefTop Top { get; set; } = new();
        public List<DefSheet> Sheets { get; set; } = new();
        public List<DefCluster> Clusters { get; set; } = new();
    }

    public class DatabaseGenerator
    {
        private static string SerializeDict(Dictionary<string, string> dict)
        {
            var items = dict.Select(kv => $"\"{EscapeJson(kv.Key)}\":\"{EscapeJson(kv.Value)}\"");
            return "{" + string.Join(",", items) + "}";
        }

        private static string SerializeList(List<string> list)
        {
            var items = list.Select(i => $"\"{EscapeJson(i)}\"");
            return "[" + string.Join(",", items) + "]";
        }

        private static string EscapeJson(string s)
        {
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t");
        }

        private static string SerializeRowHeights(Dictionary<int, double> dict)
        {
            var items = dict.Select(kv => $"\"{kv.Key}\":{kv.Value}");
            return "{" + string.Join(",", items) + "}";
        }

        private static string SerializeColumnWidths(Dictionary<int, double> dict)
        {
            var items = dict.Select(kv => $"\"{kv.Key}\":{kv.Value}");
            return "{" + string.Join(",", items) + "}";
        }

        private static string SerializeMergedCells(List<MergedCellInfo> cells)
        {
            var items = cells.Select(mc =>
                "{" +
                $"\"address\":\"{EscapeJson(mc.Address)}\"," +
                $"\"cellAddress\":\"{EscapeJson(mc.CellAddress)}\"," +
                $"\"leftPt\":{mc.LeftPt}," +
                $"\"topPt\":{mc.TopPt}," +
                $"\"widthPt\":{mc.WidthPt}," +
                $"\"heightPt\":{mc.HeightPt}" +
                "}"
            );
            return "[" + string.Join(",", items) + "]";
        }

        public DatabaseResult Generate(FormDefinition form)
        {
            var top = new DefTop
            {
                Title = form.Workbook.Title,
                Author = form.Workbook.Author,
                Created = form.Workbook.Created,
                Modified = form.Workbook.Modified,
                Version = form.Workbook.Version,
                Description = form.Workbook.Description,
                Metadata = SerializeDict(form.Metadata)
            };

            var sheets = form.Sheets.Select(s => new DefSheet
            {
                TopId = top.Id,
                SheetId = s.Id,
                Name = s.Name,
                Index = s.Index,
                PaperSize = s.PageSettings.PaperSize,
                Orientation = s.PageSettings.Orientation,
                WidthPt = s.PageSettings.WidthPt,
                HeightPt = s.PageSettings.HeightPt,
                LeftMargin = s.PageSettings.LeftMargin,
                TopMargin = s.PageSettings.TopMargin,
                RightMargin = s.PageSettings.RightMargin,
                BottomMargin = s.PageSettings.BottomMargin,
                CenterHorizontally = s.PageSettings.CenterHorizontally,
                CenterVertically = s.PageSettings.CenterVertically,
                Zoom = s.PageSettings.Zoom,
                FitToPagesWide = s.PageSettings.FitToPagesWide,
                FitToPagesTall = s.PageSettings.FitToPagesTall,
                PrintAreaAddress = s.PrintArea?.Address,
                PrintAreaLeftPt = s.PrintArea?.LeftPt,
                PrintAreaTopPt = s.PrintArea?.TopPt,
                PrintAreaWidthPt = s.PrintArea?.WidthPt,
                PrintAreaHeightPt = s.PrintArea?.HeightPt,
                PrintAreaCols = s.PrintArea?.Cols,
                PrintAreaRows = s.PrintArea?.Rows,
                BackgroundImage = s.BackgroundImage,
                FreezePane = s.FreezePane,
                RowHeights = SerializeRowHeights(s.RowHeights),
                ColumnWidths = SerializeColumnWidths(s.ColumnWidths),
                MergedCells = SerializeMergedCells(s.MergedCells)
            }).ToList();

            var clusters = form.Clusters.Select(c => new DefCluster
            {
                TopId = top.Id,
                ClusterId = c.ClusterId,
                Name = c.Name,
                Type = c.Type,
                SheetId = c.SheetId,
                CellAddress = c.CellAddress,
                Left = c.Left,
                Right = c.Right,
                Top = c.Top,
                Bottom = c.Bottom,
                LeftPt = c.LeftPt,
                TopPt = c.TopPt,
                WidthPt = c.WidthPt,
                HeightPt = c.HeightPt,
                InputParameters = SerializeDict(c.InputParameters),
                Visibility = c.Visibility,
                Readonly = c.Readonly,
                Remarks = c.Remarks,
                Functions = SerializeList(c.Functions),
                Metadata = SerializeDict(c.Metadata)
            }).ToList();

            return new DatabaseResult
            {
                Top = top,
                Sheets = sheets,
                Clusters = clusters
            };
        }
    }
}
