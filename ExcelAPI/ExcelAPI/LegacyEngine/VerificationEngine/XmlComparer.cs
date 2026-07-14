using System.Xml.Linq;

namespace ExcelAPI.LegacyEngine.VerificationEngine;

public class XmlComparer : IXmlComparer
{
    private static readonly string[] CoordinateElements = { "left", "top", "right", "bottom", "width", "height" };
    private const double Tolerance = 0.0001;

    public XmlComparisonResult Compare(string generatedXml, string databaseXml)
    {
        var result = new XmlComparisonResult();

        if (string.IsNullOrEmpty(databaseXml))
        {
            result.Status = "MISSING";
            result.MatchPercentage = 0;
            return result;
        }

        try
        {
            var genDoc = XDocument.Parse(generatedXml);
            var dbDoc = XDocument.Parse(databaseXml);

            var genClusters = genDoc.Descendants("cluster").ToList();
            var dbClusters = dbDoc.Descendants("cluster").ToList();

            result.NodeCountGenerated = genClusters.Count;
            result.NodeCountDatabase = dbClusters.Count;

            int maxCount = Math.Max(genClusters.Count, dbClusters.Count);
            for (int i = 0; i < maxCount; i++)
            {
                var gc = i < genClusters.Count ? genClusters[i] : null;
                var dc = i < dbClusters.Count ? dbClusters[i] : null;

                if (gc != null && dc != null)
                {
                    var gId = gc.Element("clusterId")?.Value;
                    var dId = dc.Element("clusterId")?.Value;
                    if (gId == dId)
                    {
                        CompareClusterElements(result, gc, dc, i);
                        continue;
                    }
                }

                if (gc == null)
                {
                    result.MissingNodes++;
                    result.Differences.Add(new XmlNodeDiff
                    {
                        Path = $"cluster[{i}]",
                        Status = "MISSING",
                        DatabaseValue = $"cluster id {dc?.Element("clusterId")?.Value}"
                    });
                }
                else if (dc == null)
                {
                    result.ExtraNodes++;
                    result.Differences.Add(new XmlNodeDiff
                    {
                        Path = $"cluster[{i}]",
                        Status = "EXTRA",
                        GeneratedValue = $"cluster id {gc.Element("clusterId")?.Value}"
                    });
                }
                else
                {
                    CompareClusterElements(result, gc, dc, i);
                }
            }

            int totalComparisons = result.MatchingNodes + result.DifferingNodes + result.MissingNodes + result.ExtraNodes;
            result.MatchPercentage = totalComparisons > 0
                ? (double)result.MatchingNodes / totalComparisons * 100.0 : 0;

            result.Status = result.MatchPercentage >= 100 ? "PASS"
                : result.MatchPercentage > 0 ? "PARTIAL" : "FAIL";
        }
        catch (Exception ex)
        {
            result.Status = "FAIL";
            result.Differences.Add(new XmlNodeDiff
            {
                Path = "xml",
                Status = "FAIL",
                DatabaseValue = "Parse error",
                GeneratedValue = ex.Message
            });
        }

        return result;
    }

    private static void CompareClusterElements(XmlComparisonResult result, XElement genCluster, XElement dbCluster, int index)
    {
        var genElements = genCluster.Elements().ToList();
        var dbElements = dbCluster.Elements().ToList();

        bool allMatch = true;

        foreach (var dbEl in dbElements)
        {
            var genEl = genElements.FirstOrDefault(e => e.Name == dbEl.Name);
            if (genEl == null)
            {
                result.MissingNodes++;
                result.Differences.Add(new XmlNodeDiff
                {
                    Path = $"cluster[{index}]/{dbEl.Name}",
                    Status = "MISSING",
                    DatabaseValue = dbEl.Value
                });
                allMatch = false;
                continue;
            }

            if (IsCoordinateElement(dbEl.Name.LocalName))
            {
                if (!CompareCoordinateValues(dbEl.Value, genEl.Value))
                {
                    result.DifferingNodes++;
                    result.Differences.Add(new XmlNodeDiff
                    {
                        Path = $"cluster[{index}]/{dbEl.Name}",
                        Status = "DIFFERENT",
                        DatabaseValue = dbEl.Value,
                        GeneratedValue = genEl.Value
                    });
                    allMatch = false;
                }
            }
            else if (dbEl.Value.Trim() != genEl.Value.Trim())
            {
                result.DifferingNodes++;
                result.Differences.Add(new XmlNodeDiff
                {
                    Path = $"cluster[{index}]/{dbEl.Name}",
                    Status = "DIFFERENT",
                    DatabaseValue = dbEl.Value,
                    GeneratedValue = genEl.Value
                });
                allMatch = false;
            }

            foreach (var dbAttr in dbEl.Attributes())
            {
                var genAttr = genEl.Attribute(dbAttr.Name);
                if (genAttr == null)
                {
                    result.DifferingNodes++;
                    result.Differences.Add(new XmlNodeDiff
                    {
                        Path = $"cluster[{index}]/{dbEl.Name}",
                        Attribute = dbAttr.Name.LocalName,
                        Status = "MISSING",
                        DatabaseValue = dbAttr.Value
                    });
                    allMatch = false;
                }
                else if (genAttr.Value != dbAttr.Value)
                {
                    result.DifferingNodes++;
                    result.Differences.Add(new XmlNodeDiff
                    {
                        Path = $"cluster[{index}]/{dbEl.Name}",
                        Attribute = dbAttr.Name.LocalName,
                        Status = "DIFFERENT",
                        DatabaseValue = dbAttr.Value,
                        GeneratedValue = genAttr.Value
                    });
                    allMatch = false;
                }
            }
        }

        foreach (var genEl in genElements)
        {
            if (!dbElements.Any(e => e.Name == genEl.Name))
            {
                result.ExtraNodes++;
                result.Differences.Add(new XmlNodeDiff
                {
                    Path = $"cluster[{index}]/{genEl.Name}",
                    Status = "EXTRA",
                    GeneratedValue = genEl.Value
                });
                allMatch = false;
            }
        }

        if (allMatch)
            result.MatchingNodes++;
    }

    private static bool IsCoordinateElement(string name)
    {
        return CoordinateElements.Contains(name, StringComparer.OrdinalIgnoreCase);
    }

    private static bool CompareCoordinateValues(string dbVal, string genVal)
    {
        if (double.TryParse(dbVal, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var dbNum)
            && double.TryParse(genVal, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var genNum))
            return Math.Abs(dbNum - genNum) < Tolerance;
        return dbVal.Trim() == genVal.Trim();
    }
}
