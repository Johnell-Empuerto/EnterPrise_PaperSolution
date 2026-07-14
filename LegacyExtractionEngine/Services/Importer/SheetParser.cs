using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services.Importer;

public class SheetParser
{
    private readonly WorkbookPart _wbPart;
    private readonly List<StyleData> _styles;
    private readonly CellParser _cellParser;
    private readonly MergeParser _mergeParser;
    private readonly PageSetupParser _pageSetupParser;

    public SheetParser(WorkbookPart wbPart, List<StyleData> styles, List<string> sharedStrings)
    {
        _wbPart = wbPart;
        _styles = styles;
        _cellParser = new CellParser(sharedStrings);
        _mergeParser = new MergeParser();
        _pageSetupParser = new PageSetupParser();
    }

    public List<SheetData> Parse(List<SheetMetadataItem> sheetMetadata)
    {
        var sheets = new List<SheetData>();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        foreach (var sm in sheetMetadata)
        {
            var wsPart = _wbPart.GetPartById(sm.RelId) as WorksheetPart;
            if (wsPart == null) continue;

            var wsXml = XDocument.Load(wsPart.GetStream());

            var sd = new SheetData
            {
                Name = sm.Name,
                SheetId = sm.Id,
                Visible = sm.State
            };

            sd.SheetFormatProperties = _pageSetupParser.ParseSheetFormatProperties(wsXml.Root!);
            sd.PageSetup = _pageSetupParser.ParsePageSetup(wsXml.Root!);
            sd.PageMargins = _pageSetupParser.ParsePageMargins(wsXml.Root!);
            _pageSetupParser.ParsePrintOptions(wsXml.Root!, sd.PageSetup);

            var colsEl = wsXml.Descendants(ns + "cols").FirstOrDefault();
            if (colsEl != null)
            {
                foreach (var col in colsEl.Elements(ns + "col"))
                {
                    sd.Columns.Add(new ColumnData
                    {
                        Min = (int?)col.Attribute("min") ?? 1,
                        Max = (int?)col.Attribute("max") ?? 1,
                        Width = (double?)col.Attribute("width") ?? 8.43,
                        CustomWidth = (bool?)col.Attribute("customWidth"),
                        Hidden = (bool?)col.Attribute("hidden") ?? false,
                        BestFit = (bool?)col.Attribute("bestFit"),
                        Style = (int?)col.Attribute("style")
                    });
                }
            }

            var rowsData = wsXml.Descendants(ns + "sheetData").FirstOrDefault();
            if (rowsData != null)
            {
                foreach (var row in rowsData.Elements(ns + "row"))
                {
                    var rowIndex = (int)row.Attribute("r");
                    sd.Rows.Add(new RowData
                    {
                        RowIndex = rowIndex,
                        Height = (double?)row.Attribute("ht"),
                        CustomHeight = (bool?)row.Attribute("customHeight"),
                        Hidden = (bool?)row.Attribute("hidden") ?? false,
                        Collapsed = (bool?)row.Attribute("collapsed"),
                        OutlineLevel = (int?)row.Attribute("outlineLevel")
                    });
                }
            }

            if (rowsData != null)
                sd.Cells = _cellParser.ParseCells(rowsData, _styles);

            sd.MergedCells = _mergeParser.Parse(wsXml.Root!);

            foreach (var cf in wsXml.Descendants(ns + "conditionalFormatting"))
            {
                foreach (var rule in cf.Elements(ns + "cfRule"))
                {
                    sd.ConditionalFormats.Add(new ConditionalFormatData
                    {
                        RuleType = (string)rule.Attribute("type"),
                        Priority = (int?)rule.Attribute("priority") ?? 0,
                        Formula = rule.Element(ns + "formula")?.Value,
                        Text = (string)rule.Attribute("text"),
                        Operator = (string)rule.Attribute("operator"),
                        DxfId = (int?)rule.Attribute("dxfId")
                    });
                }
            }

            foreach (var dv in wsXml.Descendants(ns + "dataValidation"))
            {
                sd.DataValidations.Add(new DataValidationData
                {
                    Type = (string)dv.Attribute("type"),
                    ErrorStyle = (string)dv.Attribute("errorStyle"),
                    Operator = (string)dv.Attribute("operator"),
                    Formula1 = dv.Element(ns + "formula1")?.Value,
                    Formula2 = dv.Element(ns + "formula2")?.Value,
                    AllowBlank = (bool?)dv.Attribute("allowBlank"),
                    ShowInputMessage = (bool?)dv.Attribute("showInputMessage"),
                    ShowErrorMessage = (bool?)dv.Attribute("showErrorMessage"),
                    ErrorTitle = (string)dv.Attribute("errorTitle"),
                    Error = (string)dv.Attribute("error"),
                    PromptTitle = (string)dv.Attribute("promptTitle"),
                    Prompt = (string)dv.Attribute("prompt"),
                    SqRef = (string)dv.Attribute("sqref")
                });
            }

            foreach (var hl in wsXml.Descendants(ns + "hyperlink"))
            {
                sd.Hyperlinks.Add(new HyperlinkData
                {
                    Reference = (string)hl.Attribute("ref") ?? "",
                    Location = (string)hl.Attribute("location"),
                    Tooltip = (string)hl.Attribute("tooltip"),
                    Display = (string)hl.Attribute("display")
                });
            }

            if (wsXml.Descendants(ns + "sheetProtection").Any())
                sd.Protection = new ProtectionData { Protected = true };

            var afXml = wsXml.Descendants(ns + "autoFilter").FirstOrDefault();
            if (afXml != null)
            {
                sd.AutoFilter = new AutoFilterData
                {
                    Reference = (string)afXml.Attribute("ref")
                };
            }

            foreach (var sv in wsXml.Descendants(ns + "sheetView"))
            {
                var pane = sv.Element(ns + "pane");
                if (pane != null)
                {
                    sd.FreezePanes = new FreezePanesData
                    {
                        Row = (int?)pane.Attribute("ySplit"),
                        Column = (int?)pane.Attribute("xSplit")
                    };
                }
            }

            foreach (var brk in wsXml.Descendants(ns + "rowBreaks").Elements(ns + "brk"))
            {
                sd.PageBreaks.Add(new PageBreakData { Type = "row", Row = (int?)brk.Attribute("id") });
            }
            foreach (var brk in wsXml.Descendants(ns + "colBreaks").Elements(ns + "brk"))
            {
                sd.PageBreaks.Add(new PageBreakData { Type = "column", Column = (int?)brk.Attribute("id") });
            }

            ExtractComments(wsPart, sd);
            ExtractDrawings(wsPart, sd);

            sheets.Add(sd);
        }

        return sheets;
    }

    private void ExtractComments(WorksheetPart wsPart, SheetData sd)
    {
        try
        {
            var commentsPart = wsPart.WorksheetCommentsPart;
            if (commentsPart?.GetStream() == null) return;

            var cx = XDocument.Load(commentsPart.GetStream());
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            foreach (var comment in cx.Descendants(ns + "comment"))
            {
                var textEl = comment.Descendants(ns + "t").FirstOrDefault();
                sd.Comments.Add(new CommentData
                {
                    Reference = (string)comment.Attribute("ref") ?? "",
                    Author = (string)comment.Attribute("authorId"),
                    Text = textEl?.Value
                });
            }
        }
        catch { }
    }

    private void ExtractDrawings(WorksheetPart wsPart, SheetData sd)
    {
        try
        {
            var drawingsPart = wsPart.DrawingsPart;
            if (drawingsPart == null) return;

            foreach (var imgPart in drawingsPart.ImageParts)
            {
                try
                {
                    var img = new ImageData
                    {
                        ContentType = imgPart.ContentType,
                        Format = imgPart.Uri?.OriginalString
                    };
                    using var stream = imgPart.GetStream();
                    using var ms = new MemoryStream();
                    stream.CopyTo(ms);
                    img.DataBase64 = Convert.ToBase64String(ms.ToArray());
                    sd.Images.Add(img);
                }
                catch { }
            }
        }
        catch { }
    }
}
