using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelAPI.Designer.Legacy.OpenXml;

public class StyleReader : IStyleReader
{
    public OpenXmlStyleModel Read(string filePath)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        var styles = doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        if (styles is null)
            return new OpenXmlStyleModel();

        Font? defaultFont = styles.Fonts?.ChildElements[0] as Font;
        string fontName = "Calibri";
        double fontSize = 11;

        if (defaultFont is not null)
        {
            fontName = defaultFont.FontName?.Val?.Value ?? "Calibri";
            if (defaultFont.FontSize?.Val?.Value is not null)
                fontSize = (double)defaultFont.FontSize.Val.Value;
        }

        double maxDigitWidth = EstimateMaxDigitWidth(fontName, fontSize);

        return new OpenXmlStyleModel
        {
            DefaultFontName = fontName,
            DefaultFontSize = fontSize,
            MaxDigitWidth = maxDigitWidth
        };
    }

    private static double EstimateMaxDigitWidth(string fontName, double fontSize)
    {
        return (fontName.ToLowerInvariant(), fontSize) switch
        {
            ("calibri", >= 10 and <= 12) => 7,
            ("calibri", _) => fontSize / 11.0 * 7,
            ("arial", >= 10 and <= 12) => 7,
            ("arial", _) => fontSize / 11.0 * 7,
            ("ms mincho", _) => fontSize / 11.0 * 10,
            ("ms gothic", _) => fontSize / 11.0 * 10,
            ("times new roman", >= 10 and <= 12) => 6,
            ("times new roman", _) => fontSize / 11.0 * 6,
            ("ｍｓ 明朝", _) => fontSize / 11.0 * 10,
            ("ｍｓ ゴシック", _) => fontSize / 11.0 * 10,
            _ => fontSize / 11.0 * 7
        };
    }
}
