using ExcelAPI.Designer.Legacy.LayoutEngine;
using ExcelAPI.Designer.Legacy.Models;

namespace ExcelAPI.Designer.Legacy.ClusterEngine;

public class ClusterBuilder : IClusterBuilder
{
    private readonly IColumnEngine _columnEngine;
    private readonly IRowEngine _rowEngine;

    public ClusterBuilder(IColumnEngine columnEngine, IRowEngine rowEngine)
    {
        _columnEngine = columnEngine;
        _rowEngine = rowEngine;
    }

    public ClusterModel Build(CellModel cell, SheetModel sheet, int clusterId, int sheetIndex)
    {
        double widthPt = _columnEngine.GetWidthPoints(sheet, cell.Column, cell.EndColumn);
        double heightPt = _rowEngine.GetHeightPoints(sheet, cell.Row, cell.EndRow);
        double leftPt = _columnEngine.GetLeftPoints(sheet, cell.Column);
        double topPt = _rowEngine.GetTopPoints(sheet, cell.Row);

        double sheetWidth = _columnEngine.GetTotalWidthPoints(sheet);
        double sheetHeight = _rowEngine.GetTotalHeightPoints(sheet);
        double totalArea = sheetWidth * sheetHeight;
        double areaPercent = totalArea > 0 ? widthPt * heightPt * 100.0 / totalArea : 0;
        double aspectRatio = heightPt > 0 ? widthPt / heightPt : 0;

        return new ClusterModel
        {
            ClusterId = clusterId,
            SheetIndex = sheetIndex,
            StartRow = cell.Row,
            EndRow = cell.EndRow,
            StartColumn = cell.Column,
            EndColumn = cell.EndColumn,
            CellAddress = cell.Reference,
            Comment = cell.Value,
            WidthPoints = widthPt,
            HeightPoints = heightPt,
            LeftPoints = leftPt,
            TopPoints = topPt,
            AreaPercent = Math.Round(areaPercent, 7),
            AspectRatio = Math.Round(aspectRatio, 7)
        };
    }

    public List<ClusterModel> BuildAll(SheetModel sheet, int sheetIndex)
    {
        var clusters = new List<ClusterModel>();
        int clusterId = 0;

        foreach (var comment in sheet.Comments)
        {
            string name = "";
            string typeKey = "";
            string[] lines = comment.Text.Replace("\r", "").Split('\n');
            if (lines.Length > 0) name = lines[0];
            if (lines.Length > 1) typeKey = lines[1];

            double widthPt = _columnEngine.GetWidthPoints(sheet, comment.Column, comment.EndColumn);
            double heightPt = _rowEngine.GetHeightPoints(sheet, comment.Row, comment.EndRow);
            double leftPt = _columnEngine.GetLeftPoints(sheet, comment.Column);
            double topPt = _rowEngine.GetTopPoints(sheet, comment.Row);

            double sheetWidth = _columnEngine.GetTotalWidthPoints(sheet);
            double sheetHeight = _rowEngine.GetTotalHeightPoints(sheet);
            double totalArea = sheetWidth * sheetHeight;
            double areaPercent = totalArea > 0 ? widthPt * heightPt * 100.0 / totalArea : 0;
            double aspectRatio = heightPt > 0 ? widthPt / heightPt : 0;

            clusters.Add(new ClusterModel
            {
                ClusterId = clusterId++,
                SheetIndex = sheetIndex,
                StartRow = comment.Row,
                EndRow = comment.EndRow,
                StartColumn = comment.Column,
                EndColumn = comment.EndColumn,
                CellAddress = comment.CellAddress,
                Comment = comment.Text,
                CommentName = name,
                CommentType = typeKey,
                WidthPoints = widthPt,
                HeightPoints = heightPt,
                LeftPoints = leftPt,
                TopPoints = topPt,
                AreaPercent = Math.Round(areaPercent, 7),
                AspectRatio = Math.Round(aspectRatio, 7)
            });
        }

        return clusters;
    }
}
