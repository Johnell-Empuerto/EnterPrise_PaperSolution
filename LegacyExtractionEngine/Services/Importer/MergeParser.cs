using System.Xml.Linq;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services.Importer;

public class MergeParser
{
    public List<MergedCellData> Parse(XElement wsXml)
    {
        var merges = new List<MergedCellData>();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        foreach (var mc in wsXml.Descendants(ns + "mergeCell"))
        {
            var refStr = (string)mc.Attribute("ref") ?? "";
            var parts = refStr.Split(':');
            merges.Add(new MergedCellData
            {
                Reference = refStr,
                StartRow = CellParser.ParseRowIndex(parts.Length > 0 ? parts[0] : ""),
                StartColumn = CellParser.ParseColumnIndex(parts.Length > 0 ? parts[0] : ""),
                EndRow = CellParser.ParseRowIndex(parts.Length > 1 ? parts[1] : ""),
                EndColumn = CellParser.ParseColumnIndex(parts.Length > 1 ? parts[1] : "")
            });
        }

        return merges;
    }
}
