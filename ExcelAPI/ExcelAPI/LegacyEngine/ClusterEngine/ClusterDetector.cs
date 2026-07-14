using ExcelAPI.LegacyEngine.Models;

namespace ExcelAPI.LegacyEngine.ClusterEngine;

public class ClusterDetector : IClusterDetector
{
    public List<ClusterModel> Detect(SheetModel sheet)
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

            clusters.Add(new ClusterModel
            {
                ClusterId = clusterId++,
                SheetIndex = sheet.SheetIndex,
                StartRow = comment.Row,
                EndRow = comment.EndRow,
                StartColumn = comment.Column,
                EndColumn = comment.EndColumn,
                CellAddress = comment.CellAddress,
                Comment = comment.Text,
                CommentName = name,
                CommentType = typeKey
            });
        }

        return clusters;
    }
}
