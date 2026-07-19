using System.Text;
using ExcelAPI.Designer.Legacy.Models;

namespace ExcelAPI.Designer.Legacy.PublishEngine;

public class XmlGenerator : IXmlGenerator
{
    private const string Indent = "  ";

    public string Generate(WorkbookModel workbook, List<ClusterModel> clusters,
        Dictionary<(int sheetIndex, int clusterId), CoordinateRect> finalCoords)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<conmas>");
        AppendHeader(sb);
        AppendTop(sb, workbook, clusters, finalCoords);
        sb.AppendLine("</conmas>");
        return sb.ToString();
    }

    private void AppendHeader(StringBuilder sb)
    {
        sb.AppendLine(Indent + "<header>");
        sb.AppendLine(Indent + Indent + "<version></version>");
        sb.AppendLine(Indent + Indent + "<command></command>");
        sb.AppendLine(Indent + Indent + "<resultCode></resultCode>");
        sb.AppendLine(Indent + Indent + "<message></message>");
        sb.AppendLine(Indent + Indent + "<requestUser></requestUser>");
        sb.AppendLine(Indent + Indent + "<createTime></createTime>");
        sb.AppendLine(Indent + "</header>");
    }

    private void AppendTop(StringBuilder sb, WorkbookModel workbook,
        List<ClusterModel> clusters,
        Dictionary<(int sheetIndex, int clusterId), CoordinateRect> finalCoords)
    {
        sb.AppendLine(Indent + "<top>");
        sb.AppendLine(Indent + Indent + $"<definitionFile>");
        sb.AppendLine(Indent + Indent + Indent + $"<name>{EscapeXml(workbook.FileName)}</name>");
        sb.AppendLine(Indent + Indent + Indent + "<type>xlsx</type>");
        sb.AppendLine(Indent + Indent + Indent + "<value></value>");
        sb.AppendLine(Indent + Indent + $"</definitionFile>");
        sb.AppendLine(Indent + Indent + "<backgroundImage></backgroundImage>");
        sb.AppendLine(Indent + Indent + "<sheets>");

        foreach (var sheet in workbook.Sheets)
        {
            AppendSheet(sb, sheet, clusters, finalCoords);
        }

        sb.AppendLine(Indent + Indent + "</sheets>");
        sb.AppendLine(Indent + "</top>");
    }

    private void AppendSheet(StringBuilder sb, SheetModel sheet,
        List<ClusterModel> clusters,
        Dictionary<(int sheetIndex, int clusterId), CoordinateRect> finalCoords)
    {
        sb.AppendLine(Indent + Indent + Indent + "<sheet>");
        sb.AppendLine(Indent + Indent + Indent + Indent +
            $"<defSheetName>{EscapeXml(sheet.Name)}</defSheetName>");
        sb.AppendLine(Indent + Indent + Indent + Indent +
            $"<sheetNo>{sheet.SheetIndex}</sheetNo>");
        sb.AppendLine(Indent + Indent + Indent + Indent +
            $"<width>{sheet.PageSetup.PageWidthPoints}</width>");
        sb.AppendLine(Indent + Indent + Indent + Indent +
            $"<height>{sheet.PageSetup.PageHeightPoints}</height>");
        sb.AppendLine(Indent + Indent + Indent + Indent + "<clusters>");

        var sheetClusters = clusters
            .Where(c => c.SheetIndex == sheet.SheetIndex)
            .OrderBy(c => c.StartRow)
            .ThenBy(c => c.StartColumn)
            .ToList();

        foreach (var cluster in sheetClusters)
        {
            var key = (sheet.SheetIndex, cluster.ClusterId);
            if (!finalCoords.TryGetValue(key, out var coord))
                continue;

            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + "<cluster>");
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent +
                $"<clusterId>{cluster.ClusterId}</clusterId>");
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent +
                $"<sheetNo>{cluster.SheetIndex}</sheetNo>");
            string clusterName = cluster.CommentName ?? "";
            string clusterType = cluster.CommentType ?? "";
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent +
                $"<name>{EscapeXml(clusterName)}</name>");
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent +
                $"<type>{EscapeXml(clusterType)}</type>");
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent +
                $"<left>{FormatRatio(coord.Left)}</left>");
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent +
                $"<top>{FormatRatio(coord.Top)}</top>");
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent +
                $"<right>{FormatRatio(coord.Right)}</right>");
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent +
                $"<bottom>{FormatRatio(coord.Bottom)}</bottom>");
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent +
                $"<cellAddress>{EscapeXml(cluster.CellAddress)}</cellAddress>");
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent +
                "<inputParameters>Required=0;Lines=1;InputRestriction=None;MaxLength=0;" +
                "Align=Center;Font=Arial;FontSize=11;Weight=Normal;" +
                "Color=0,0,0;VerticalAlignment=2;DefaultFontSize=11</inputParameters>");
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + "</cluster>");
        }

        sb.AppendLine(Indent + Indent + Indent + Indent + "</clusters>");
        sb.AppendLine(Indent + Indent + Indent + "</sheet>");
    }

    private static string FormatRatio(double ratio)
    {
        return ratio.ToString("0.0#######");
    }

    private static string EscapeXml(string s)
    {
        return s.Replace("&", "&amp;").Replace("<", "&lt;")
            .Replace(">", "&gt;").Replace("\"", "&quot;")
            .Replace("'", "&apos;");
    }
}
