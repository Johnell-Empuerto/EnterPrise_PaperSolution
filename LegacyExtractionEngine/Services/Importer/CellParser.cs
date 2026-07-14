using System.Xml.Linq;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services.Importer;

public class CellParser
{
    private readonly List<string> _sharedStrings;

    public CellParser(List<string> sharedStrings)
    {
        _sharedStrings = sharedStrings;
    }

    public List<CellData> ParseCells(XElement sheetData, List<StyleData> styles)
    {
        var cells = new List<CellData>();

        foreach (var row in sheetData.Elements())
        {
            var rowIndex = int.Parse(row.Attribute("r")?.Value ?? "0");

            foreach (var cell in row.Elements())
            {
                var cellRef = (string)cell.Attribute("r") ?? "";
                var cellType = (string)cell.Attribute("t");
                var styleIdx = (int?)cell.Attribute("s");

                var cd = new CellData
                {
                    Reference = cellRef,
                    Row = rowIndex,
                    Column = ParseColumnIndex(cellRef),
                    StyleIndex = styleIdx,
                    DataType = cellType
                };

                var f = cell.Element(GetNs() + "f");
                if (f != null) cd.Formula = f.Value;

                var v = cell.Element(GetNs() + "v");
                if (v != null)
                {
                    var cellValue = v.Value;
                    if (cellType == "s" && int.TryParse(cellValue, out var ssid)
                        && ssid >= 0 && ssid < _sharedStrings.Count)
                    {
                        cd.Value = _sharedStrings[ssid];
                        cd.Type = "shared_string";
                    }
                    else if (cellType == "b")
                    {
                        cd.Value = cellValue;
                        cd.Type = "boolean";
                    }
                    else if (cellType == "e")
                    {
                        cd.Value = cellValue;
                        cd.Type = "error";
                    }
                    else
                    {
                        if (double.TryParse(cellValue,
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var num))
                        {
                            cd.Value = num;
                            cd.Type = "number";
                        }
                        else
                        {
                            cd.Value = cellValue;
                            cd.Type = "string";
                        }
                    }
                }

                if (cd.StyleIndex.HasValue && cd.StyleIndex.Value < styles.Count)
                {
                    cd.NumberFormat = styles[cd.StyleIndex.Value]?.NumberFormat?.FormatCode;
                }

                cells.Add(cd);
            }
        }

        return cells;
    }

    private static XNamespace GetNs()
    {
        return "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    }

    public static int ParseColumnIndex(string reference)
    {
        if (string.IsNullOrEmpty(reference)) return 0;
        int col = 0;
        foreach (var c in reference.TakeWhile(char.IsLetter))
        {
            col = col * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        }
        return col;
    }

    public static int ParseRowIndex(string reference)
    {
        if (string.IsNullOrEmpty(reference)) return 0;
        var rowPart = new string(reference.SkipWhile(char.IsLetter).ToArray());
        return int.TryParse(rowPart, out var row) ? row : 0;
    }
}
