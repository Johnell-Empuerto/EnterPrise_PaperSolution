using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelAPI.LegacyEngine.OpenXml;

public class CellStyleReader : ICellStyleReader
{
    public IReadOnlyDictionary<int, OpenXmlCellStyleModel> Read(string filePath)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var styles = doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        if (styles is null) return new Dictionary<int, OpenXmlCellStyleModel>();

        var fonts = styles.Fonts?.Cast<Font>().ToList() ?? new List<Font>();
        var fills = styles.Fills?.Cast<Fill>().ToList() ?? new List<Fill>();
        var borders = styles.Borders?.Cast<Border>().ToList() ?? new List<Border>();
        var cellFormats = styles.CellFormats?.Cast<CellFormat>().ToList() ?? new List<CellFormat>();

        var result = new Dictionary<int, OpenXmlCellStyleModel>();

        for (int i = 0; i < cellFormats.Count; i++)
        {
            var xf = cellFormats[i];

            var style = new OpenXmlCellStyleModel();

            // Font
            if (xf.FontId?.Value is not null)
            {
                int fontId = (int)xf.FontId.Value;
                if (fontId >= 0 && fontId < fonts.Count)
                {
                    var font = fonts[fontId];
                    style.FontName = font.FontName?.Val?.Value ?? "";
                    if (font.FontSize?.Val?.Value is not null)
                        style.FontSize = (double)font.FontSize.Val.Value;
                    style.Bold = font.Bold?.Val?.Value ?? false;
                    style.Italic = font.Italic?.Val?.Value ?? false;
                    style.Underline = font.Underline?.Val?.Value is not null;

                    // Color is accessed via GetFirstChild<Color>() on Font
                    var fontColor = font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>();
                    style.FontColor = GetColor(fontColor);
                }
            }

            // Fill
            if (xf.FillId?.Value is not null)
            {
                int fillId = (int)xf.FillId.Value;
                if (fillId >= 0 && fillId < fills.Count)
                {
                    var fill = fills[fillId];
                    var patternFill = fill.PatternFill;
                    if (patternFill is not null && patternFill.PatternType?.Value is not null)
                    {
                        var patternType = patternFill.PatternType.Value;
                        if (patternType != PatternValues.None && patternType != PatternValues.Gray125)
                        {
                            string? fgColor = ExtractPatternFillColor(patternFill.ForegroundColor);
                            if (!string.IsNullOrEmpty(fgColor))
                                style.FillColor = fgColor;
                            else
                            {
                                string? bgColor = ExtractPatternFillColor(patternFill.BackgroundColor);
                                if (!string.IsNullOrEmpty(bgColor))
                                    style.FillColor = bgColor;
                            }
                        }
                    }
                }
            }

            // Border
            if (xf.BorderId?.Value is not null)
            {
                int borderId = (int)xf.BorderId.Value;
                if (borderId >= 0 && borderId < borders.Count)
                {
                    var border = borders[borderId];
                    style.BorderTop = GetBorderCss(border.TopBorder);
                    style.BorderBottom = GetBorderCss(border.BottomBorder);
                    style.BorderLeft = GetBorderCss(border.LeftBorder);
                    style.BorderRight = GetBorderCss(border.RightBorder);
                }
            }

            // Alignment
            if (xf.Alignment is not null)
            {
                var alignment = xf.Alignment;
                if (alignment.Horizontal is not null)
                    style.HorizontalAlignment = MapHorizontalAlignment(alignment.Horizontal.Value.ToString());
                if (alignment.Vertical is not null)
                    style.VerticalAlignment = MapVerticalAlignment(alignment.Vertical.Value.ToString());
                style.WrapText = alignment.WrapText?.Value ?? false;
                style.Indent = (int)(alignment.Indent?.Value ?? 0);
            }

            result[i] = style;
        }

        return result;
    }

    private static string GetColor(DocumentFormat.OpenXml.Spreadsheet.Color? color)
    {
        if (color is null) return "";

        if (color.Rgb?.Value is not null)
        {
            string rgb = color.Rgb.Value;
            if (rgb.Length >= 8 && (rgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase) ||
                                    rgb.StartsWith("ff", StringComparison.OrdinalIgnoreCase)))
                return "#" + rgb[2..];
            if (rgb.Length >= 6)
                return "#" + rgb;
        }

        if (color.Indexed?.Value is not null)
            return IndexedColorToHex((int)color.Indexed.Value);

        if (color.Auto?.Value is not null && color.Auto.Value)
            return "#000000";

        return "";
    }

    private static string? ExtractPatternFillColor(ForegroundColor? fgColor)
    {
        if (fgColor is null) return null;
        if (fgColor.Rgb?.Value is not null)
        {
            string rgb = fgColor.Rgb.Value;
            if (rgb.Length >= 8 && (rgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase) ||
                                    rgb.StartsWith("ff", StringComparison.OrdinalIgnoreCase)))
                return "#" + rgb[2..];
            if (rgb.Length >= 6)
                return "#" + rgb;
        }
        if (fgColor.Indexed?.Value is not null)
            return IndexedColorToHex((int)fgColor.Indexed.Value);
        if (fgColor.Auto?.Value is not null && fgColor.Auto.Value)
            return "#000000";
        return null;
    }

    private static string? ExtractPatternFillColor(BackgroundColor? bgColor)
    {
        if (bgColor is null) return null;
        if (bgColor.Rgb?.Value is not null)
        {
            string rgb = bgColor.Rgb.Value;
            if (rgb.Length >= 8 && (rgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase) ||
                                    rgb.StartsWith("ff", StringComparison.OrdinalIgnoreCase)))
                return "#" + rgb[2..];
            if (rgb.Length >= 6)
                return "#" + rgb;
        }
        if (bgColor.Indexed?.Value is not null)
            return IndexedColorToHex((int)bgColor.Indexed.Value);
        if (bgColor.Auto?.Value is not null && bgColor.Auto.Value)
            return "#000000";
        return null;
    }

    private static string GetBorderCss(BorderPropertiesType? border)
    {
        if (border is null) return "";
        if (border.Style is null) return "";

        var borderStyle = border.Style.Value;
        if (borderStyle == BorderStyleValues.None) return "";

        string width = "1px";
        string lineStyle = "solid";

        if (borderStyle == BorderStyleValues.Thin) width = "1px";
        else if (borderStyle == BorderStyleValues.Medium) width = "2px";
        else if (borderStyle == BorderStyleValues.Thick) width = "3px";
        else if (borderStyle == BorderStyleValues.Double) { width = "3px"; lineStyle = "double"; }
        else if (borderStyle == BorderStyleValues.Hair) width = "1px";
        else if (borderStyle == BorderStyleValues.Dotted) { width = "1px"; lineStyle = "dotted"; }
        else if (borderStyle == BorderStyleValues.Dashed ||
                 borderStyle == BorderStyleValues.DashDot ||
                 borderStyle == BorderStyleValues.DashDotDot) { width = "1px"; lineStyle = "dashed"; }
        else if (borderStyle == BorderStyleValues.MediumDashed ||
                 borderStyle == BorderStyleValues.MediumDashDot ||
                 borderStyle == BorderStyleValues.MediumDashDotDot) { width = "2px"; lineStyle = "dashed"; }
        else if (borderStyle == BorderStyleValues.SlantDashDot) { width = "2px"; lineStyle = "dashed"; }

        string color = GetColor(border.Color);
        if (string.IsNullOrEmpty(color))
            color = "#000000";

        return $"{width} {lineStyle} {color}";
    }

    private static string MapHorizontalAlignment(string? align)
    {
        return align switch
        {
            "left" => "left",
            "center" => "center",
            "right" => "right",
            "centerContinuous" => "center",
            "fill" => "left",
            "justify" => "justify",
            "distributed" => "justify",
            _ => ""
        };
    }

    private static string MapVerticalAlignment(string? align)
    {
        return align switch
        {
            "top" => "top",
            "center" => "middle",
            "bottom" => "bottom",
            "justify" => "middle",
            "distributed" => "middle",
            _ => ""
        };
    }

    private static string IndexedColorToHex(int idx)
    {
        return idx switch
        {
            0 => "#000000",
            1 => "#FFFFFF",
            2 => "#FF0000",
            3 => "#00FF00",
            4 => "#0000FF",
            5 => "#FFFF00",
            6 => "#FF00FF",
            7 => "#00FFFF",
            8 => "#000000",
            9 => "#FFFFFF",
            10 => "#FF0000",
            11 => "#00FF00",
            12 => "#0000FF",
            13 => "#FFFF00",
            14 => "#FF00FF",
            15 => "#00FFFF",
            16 => "#800000",
            17 => "#008000",
            18 => "#000080",
            19 => "#808000",
            20 => "#800080",
            21 => "#008080",
            22 => "#C0C0C0",
            23 => "#808080",
            24 => "#9999FF",
            25 => "#993366",
            26 => "#FFFFCC",
            27 => "#CCFFFF",
            28 => "#660066",
            29 => "#FF8080",
            30 => "#0066CC",
            31 => "#CCCCFF",
            32 => "#000080",
            33 => "#FF00FF",
            34 => "#FFFF00",
            35 => "#00FFFF",
            36 => "#800080",
            37 => "#800000",
            38 => "#008080",
            39 => "#0000FF",
            40 => "#6666CC",
            41 => "#CCCCFF",
            42 => "#000099",
            43 => "#99CCFF",
            44 => "#993366",
            45 => "#FFFFCC",
            46 => "#CCFFFF",
            47 => "#660066",
            48 => "#FF8080",
            49 => "#0066CC",
            50 => "#CCCCFF",
            51 => "#000080",
            52 => "#FF00FF",
            53 => "#FFFF00",
            54 => "#00FFFF",
            55 => "#800080",
            56 => "#800000",
            57 => "#008080",
            58 => "#0000FF",
            59 => "#6666CC",
            60 => "#CCCCFF",
            61 => "#000099",
            62 => "#99CCFF",
            63 => "#993366",
            _ => "#000000"
        };
    }
}
