using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services.Importer;

public class StyleParser
{
    public List<StyleData> Parse(WorkbookPart wbPart)
    {
        var styles = new List<StyleData>();
        var stylesPart = wbPart.WorkbookStylesPart;
        if (stylesPart?.GetStream() == null) return styles;

        try
        {
            var sx = XDocument.Load(stylesPart.GetStream());
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            var numFmtMap = new Dictionary<int, string>();
            foreach (var nf in sx.Descendants(ns + "numFmt"))
            {
                var id = (int?)nf.Attribute("numFmtId");
                var code = (string)nf.Attribute("formatCode");
                if (id.HasValue && code != null)
                    numFmtMap[id.Value] = code;
            }

            var fonts = new List<FontData?>();
            foreach (var font in sx.Descendants(ns + "font"))
            {
                var fd = new FontData();
                var sz = font.Element(ns + "sz");
                if (sz != null) fd.Size = (double?)sz.Attribute("val");

                var name = font.Element(ns + "name");
                if (name != null) fd.FontName = (string)name.Attribute("val");

                var family = font.Element(ns + "family");
                if (family != null) fd.FontFamily = (int?)family.Attribute("val");

                var scheme = font.Element(ns + "scheme");
                if (scheme != null) fd.FontScheme = (string)scheme.Attribute("val");

                var charset = font.Element(ns + "charset");
                if (charset != null) fd.Charset = (int?)charset.Attribute("val");

                if (font.Element(ns + "b") != null) fd.Bold = true;
                if (font.Element(ns + "i") != null) fd.Italic = true;

                var u = font.Element(ns + "u");
                if (u != null) fd.Underline = (string)u.Attribute("val") ?? "single";

                if (font.Element(ns + "strike") != null) fd.Strikethrough = true;

                var color = font.Element(ns + "color");
                if (color != null) fd.Color = ReadColorAttr(color);

                fonts.Add(fd);
            }

            var fills = new List<FillData?>();
            foreach (var fill in sx.Descendants(ns + "fill"))
            {
                var fd = new FillData();
                var pf = fill.Element(ns + "patternFill");
                if (pf != null)
                {
                    fd.PatternType = (string)pf.Attribute("patternType");
                    var fg = pf.Element(ns + "fgColor");
                    if (fg != null) fd.ForegroundColor = ReadColorAttr(fg);
                    var bg = pf.Element(ns + "bgColor");
                    if (bg != null) fd.BackgroundColor = ReadColorAttr(bg);
                }
                fills.Add(fd);
            }

            var borders = new List<BorderData?>();
            foreach (var border in sx.Descendants(ns + "border"))
            {
                var bd = new BorderData
                {
                    DiagonalUp = (bool?)border.Attribute("diagonalUp"),
                    DiagonalDown = (bool?)border.Attribute("diagonalDown")
                };

                foreach (var side in border.Elements())
                {
                    var edge = new BorderEdgeData
                    {
                        Style = (string)side.Attribute("style")
                    };
                    var c = side.Element(ns + "color");
                    if (c != null) edge.Color = ReadColorAttr(c);

                    switch (side.Name.LocalName)
                    {
                        case "left": bd.Left = edge; break;
                        case "right": bd.Right = edge; break;
                        case "top": bd.Top = edge; break;
                        case "bottom": bd.Bottom = edge; break;
                        case "diagonal": bd.Diagonal = edge; break;
                    }
                }
                borders.Add(bd);
            }

            int idx = 0;
            foreach (var xf in sx.Descendants(ns + "xf"))
            {
                var sd = new StyleData { StyleIndex = idx };
                sd.NumberFormat = xf.Attribute("numFmtId") != null
                    && numFmtMap.TryGetValue((int)xf.Attribute("numFmtId")!, out var fmtCode)
                    ? new NumberFormatData { FormatCode = fmtCode } : null;

                var fontId = (int?)xf.Attribute("fontId");
                if (fontId.HasValue && fontId.Value < fonts.Count)
                    sd.Font = fonts[fontId.Value];

                var fillId = (int?)xf.Attribute("fillId");
                if (fillId.HasValue && fillId.Value < fills.Count)
                    sd.Fill = fills[fillId.Value];

                var borderId = (int?)xf.Attribute("borderId");
                if (borderId.HasValue && borderId.Value < borders.Count)
                    sd.Border = borders[borderId.Value];

                var alignEl = xf.Element(ns + "alignment");
                if (alignEl != null)
                {
                    sd.Alignment = new AlignmentData
                    {
                        Horizontal = (string)alignEl.Attribute("horizontal"),
                        Vertical = (string)alignEl.Attribute("vertical"),
                        TextRotation = (int?)alignEl.Attribute("textRotation"),
                        WrapText = (bool?)alignEl.Attribute("wrapText"),
                        Indent = (int?)alignEl.Attribute("indent")
                    };
                }

                styles.Add(sd);
                idx++;
            }
        }
        catch { }

        return styles;
    }

    private string? ReadColorAttr(XElement color)
    {
        if (color == null) return null;
        var rgb = (string)color.Attribute("rgb");
        if (rgb != null) return rgb;
        var theme = (string)color.Attribute("theme");
        if (theme != null) return $"theme:{theme}";
        var indexed = (string)color.Attribute("indexed");
        if (indexed != null) return $"index:{indexed}";
        var auto = (string)color.Attribute("auto");
        if (auto != null) return "auto";
        return null;
    }
}
