using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

public class ExcelExtractionEngine
{
    private readonly IProgress<string> _progress;

    public ExcelExtractionEngine(IProgress<string> progress)
    {
        _progress = progress;
    }

    public ExtractionResult Extract(string filePath)
    {
        _progress.Report("[8/15] Reading workbook...");
        var result = new ExtractionResult
        {
            SourceFile = filePath,
            ExtractedAt = DateTime.UtcNow
        };

        using var doc = SpreadsheetDocument.Open(filePath, false);
        var wbPart = doc.WorkbookPart;
        if (wbPart == null) return result;

        result.Workbook = ExtractWorkbook(wbPart);
        result.Styles = ExtractStyles(wbPart);

        _progress.Report("[9/15] Reading sheets...");
        ExtractSheets(wbPart, result);

        return result;
    }

    private WorkbookData ExtractWorkbook(WorkbookPart wbPart)
    {
        var wb = new WorkbookData();
        var wbXml = XDocument.Load(wbPart.GetStream());

        // Namespace
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace nsr = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // Defined names
        var definedNames = wbXml.Descendants(ns + "definedNames").Elements(ns + "definedName");
        foreach (var dn in definedNames)
        {
            wb.DefinedNames.Add(new DefinedNameData
            {
                Name = (string)dn.Attribute("name") ?? "",
                Comment = (string)dn.Attribute("comment"),
                RefersTo = dn.Value,
                LocalSheetId = (string)dn.Attribute("localSheetId"),
                Hidden = (bool?)dn.Attribute("hidden") ?? false
            });
        }

        // Extended file properties
        var extPropsPart = wbPart.GetPartsOfType<ExtendedFilePropertiesPart>().FirstOrDefault();
        if (extPropsPart?.GetStream() != null)
        {
            try
            {
                var extXml = XDocument.Load(extPropsPart.GetStream());
                XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
                XNamespace dc = "http://purl.org/dc/elements/1.1/";
                XNamespace dcterms = "http://purl.org/dc/terms/";

                wb.Properties.Title = (string)extXml.Descendants(dc + "title").FirstOrDefault();
                wb.Properties.Subject = (string)extXml.Descendants(dc + "subject").FirstOrDefault();
                wb.Properties.Creator = (string)extXml.Descendants(dc + "creator").FirstOrDefault();
                wb.Properties.Description = (string)extXml.Descendants(dc + "description").FirstOrDefault();
                wb.Properties.LastModifiedBy = (string)extXml.Descendants(cp + "lastModifiedBy").FirstOrDefault();
                wb.Properties.Revision = (string)extXml.Descendants(cp + "revision").FirstOrDefault();
                wb.Properties.Created = (DateTime?)extXml.Descendants(dcterms + "created").FirstOrDefault();
                wb.Properties.Modified = (DateTime?)extXml.Descendants(dcterms + "modified").FirstOrDefault();
                wb.Properties.Category = (string)extXml.Descendants(cp + "category").FirstOrDefault();
                wb.Properties.Keywords = (string)extXml.Descendants(cp + "keywords").FirstOrDefault();
                wb.Properties.Application = (string)extXml.Descendants(cp + "Application").FirstOrDefault();
                wb.Properties.AppVersion = (string)extXml.Descendants(cp + "appVersion").FirstOrDefault();
                wb.Properties.Company = (string)extXml.Descendants(cp + "company").FirstOrDefault();
                wb.Properties.Manager = (string)extXml.Descendants(cp + "manager").FirstOrDefault();
                wb.Properties.HyperlinkBase = (string)extXml.Descendants(cp + "hyperlinkBase").FirstOrDefault();
                wb.Properties.ContentType = (string)extXml.Descendants(cp + "contentType").FirstOrDefault();
            }
            catch { }
        }

        // Custom properties
        var customPart = wbPart.GetPartsOfType<CustomFilePropertiesPart>().FirstOrDefault();
        if (customPart?.GetStream() != null)
        {
            try
            {
                var custXml = XDocument.Load(customPart.GetStream());
                XNamespace nsct = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
                XNamespace vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

                foreach (var prop in custXml.Descendants(nsct + "property"))
                {
                    wb.CustomProperties.Add(new CustomPropertyData
                    {
                        Name = (string)prop.Attribute("name") ?? "",
                        Value = prop.Value
                    });
                }
            }
            catch { }
        }

        // Workbook protection
        var wbProtection = wbXml.Descendants(ns + "workbookProtection").FirstOrDefault();
        if (wbProtection != null)
        {
            wb.WorkbookProtection = new ProtectionData { Protected = true };
        }

        return wb;
    }

    private List<StyleData> ExtractStyles(WorkbookPart wbPart)
    {
        _progress.Report("[6/15] Reading styles...");
        var styles = new List<StyleData>();

        var stylesPart = wbPart.WorkbookStylesPart;
        if (stylesPart?.GetStream() == null) return styles;

        try
        {
            var sx = XDocument.Load(stylesPart.GetStream());
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            // Number formats
            var numFmtMap = new Dictionary<int, string>();
            foreach (var nf in sx.Descendants(ns + "numFmt"))
            {
                var id = (int?)nf.Attribute("numFmtId");
                var code = (string)nf.Attribute("formatCode");
                if (id.HasValue && code != null)
                    numFmtMap[id.Value] = code;
            }

            // Fonts
            var fonts = new List<FontData?>();
            foreach (var font in sx.Descendants(ns + "font"))
            {
                var fd = new FontData();
                var sz = font.Element(ns + "sz");
                if (sz != null) fd.Size = (double?)sz.Attribute("val");

                var name = font.Element(ns + "name");
                if (name != null) fd.FontName = (string)name.Attribute("val");

                if (font.Element(ns + "b") != null) fd.Bold = true;
                if (font.Element(ns + "i") != null) fd.Italic = true;

                var u = font.Element(ns + "u");
                if (u != null) fd.Underline = (string)u.Attribute("val") ?? "single";

                if (font.Element(ns + "strike") != null) fd.Strikethrough = true;

                var color = font.Element(ns + "color");
                if (color != null) fd.Color = ReadColorAttr(color);

                fonts.Add(fd);
            }

            // Fills
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

            // Borders
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

            // Cell formats
            int idx = 0;
            foreach (var xf in sx.Descendants(ns + "xf"))
            {
                var sd = new StyleData { StyleIndex = idx };
                sd.NumberFormat = xf.Attribute("numFmtId") != null
                    && numFmtMap.TryGetValue((int)xf.Attribute("numFmtId"), out var fmtCode)
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
        catch (Exception ex)
        {
            _progress.Report($"  Warning: Style parsing error: {ex.Message}");
        }

        return styles;
    }

    private void ExtractSheets(WorkbookPart wbPart, ExtractionResult result)
    {
        // Load shared strings
        var sharedStrings = LoadSharedStrings(wbPart);

        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace nsr = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var wbXml = XDocument.Load(wbPart.GetStream());
        int sheetIndex = 0;

        foreach (var sheetEl in wbXml.Descendants(ns + "sheet"))
        {
            var name = (string)sheetEl.Attribute("name") ?? "";
            var sid = (string)sheetEl.Attribute("sheetId") ?? "0";
            var stateAttr = (string)sheetEl.Attribute("state") ?? "visible";
            var relId = (string)sheetEl.Attribute(nsr + "id") ?? "";

            _progress.Report($"  Processing sheet: '{name}' (ID={sid})");

            var wsPart = wbPart.GetPartById(relId) as WorksheetPart;
            if (wsPart == null) continue;

            var wsXml = XDocument.Load(wsPart.GetStream());

            var sd = new SheetData
            {
                Name = name,
                SheetId = int.TryParse(sid, out var sidInt) ? sidInt : 0,
                Visible = stateAttr
            };

            // Sheet format properties
            var sheetFormatPr = wsXml.Descendants(ns + "sheetFormatPr").FirstOrDefault();
            if (sheetFormatPr != null)
            {
                sd.SheetFormatProperties = new SheetFormatPropertiesData
                {
                    DefaultRowHeight = (double?)sheetFormatPr.Attribute("defaultRowHeight"),
                    DefaultColumnWidth = (double?)sheetFormatPr.Attribute("defaultColWidth"),
                    OutlineLevelRow = (int?)sheetFormatPr.Attribute("outlineLevelRow"),
                    OutlineLevelColumn = (int?)sheetFormatPr.Attribute("outlineLevelCol")
                };
            }

            // Columns
            foreach (var col in wsXml.Descendants(ns + "col"))
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

            // Rows and cells
            foreach (var row in wsXml.Descendants(ns + "row"))
            {
                var rowIndex = (int)row.Attribute("r");
                var rd = new RowData
                {
                    RowIndex = rowIndex,
                    Height = (double?)row.Attribute("ht"),
                    CustomHeight = (bool?)row.Attribute("customHeight"),
                    Hidden = (bool?)row.Attribute("hidden") ?? false,
                    Collapsed = (bool?)row.Attribute("collapsed"),
                    OutlineLevel = (int?)row.Attribute("outlineLevel")
                };
                sd.Rows.Add(rd);

                foreach (var cell in row.Elements(ns + "c"))
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

                    // Formula
                    var f = cell.Element(ns + "f");
                    if (f != null) cd.Formula = f.Value;

                    // Value
                    var v = cell.Element(ns + "v");
                    if (v != null)
                    {
                        var cellValue = v.Value;
                        if (cellType == "s" && int.TryParse(cellValue, out var ssid)
                            && ssid >= 0 && ssid < sharedStrings.Count)
                        {
                            cd.Value = sharedStrings[ssid];
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

                    // Number format
                    if (cd.StyleIndex.HasValue && cd.StyleIndex.Value < result.Styles.Count)
                    {
                        cd.NumberFormat = result.Styles[cd.StyleIndex.Value]?.NumberFormat?.FormatCode;
                    }

                    sd.Cells.Add(cd);
                }
            }

            // Merged cells
            foreach (var mc in wsXml.Descendants(ns + "mergeCell"))
            {
                var refStr = (string)mc.Attribute("ref") ?? "";
                var parts = refStr.Split(':');
                sd.MergedCells.Add(new MergedCellData
                {
                    Reference = refStr,
                    StartRow = ParseRowIndex(parts.Length > 0 ? parts[0] : ""),
                    StartColumn = ParseColumnIndex(parts.Length > 0 ? parts[0] : ""),
                    EndRow = ParseRowIndex(parts.Length > 1 ? parts[1] : ""),
                    EndColumn = ParseColumnIndex(parts.Length > 1 ? parts[1] : "")
                });
            }

            // Conditional formatting
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
                        DxfId = (int?)rule.Attribute("dxfId"),
                        AllAttributes = ReadAllAttrs(rule)
                    });
                }
            }

            // Data validations
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

            // Hyperlinks
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

            // Page setup
            var psXml = wsXml.Descendants(ns + "pageSetup").FirstOrDefault();
            if (psXml != null)
            {
                sd.PageSetup = new PageSetupData
                {
                    PaperSize = (int?)psXml.Attribute("paperSize"),
                    Orientation = (string)psXml.Attribute("orientation"),
                    Scale = (int?)psXml.Attribute("scale"),
                    FitToWidth = (int?)psXml.Attribute("fitToWidth"),
                    FitToHeight = (int?)psXml.Attribute("fitToHeight"),
                    BlackAndWhite = (bool?)psXml.Attribute("blackAndWhite"),
                    Draft = (bool?)psXml.Attribute("draft"),
                    CellComments = (string)psXml.Attribute("cellComments"),
                    Errors = (string)psXml.Attribute("errors"),
                    HorizontalDpi = (int?)psXml.Attribute("horizontalDpi"),
                    VerticalDpi = (int?)psXml.Attribute("verticalDpi"),
                    Copies = (int?)psXml.Attribute("copies"),
                    FirstPageNumber = (int?)psXml.Attribute("firstPageNumber"),
                    UseFirstPageNumber = (bool?)psXml.Attribute("useFirstPageNumber"),
                    PageOrder = (string)psXml.Attribute("pageOrder"),
                    AllAttributes = ReadAllAttrs(psXml)
                };
            }

            // Page margins
            var pmXml = wsXml.Descendants(ns + "pageMargins").FirstOrDefault();
            if (pmXml != null)
            {
                sd.PageMargins = new PageMarginsData
                {
                    Top = (double?)pmXml.Attribute("top") ?? 0.75,
                    Bottom = (double?)pmXml.Attribute("bottom") ?? 0.75,
                    Left = (double?)pmXml.Attribute("left") ?? 0.70,
                    Right = (double?)pmXml.Attribute("right") ?? 0.70,
                    Header = (double?)pmXml.Attribute("header") ?? 0.30,
                    Footer = (double?)pmXml.Attribute("footer") ?? 0.30
                };
            }

            // Sheet protection
            if (wsXml.Descendants(ns + "sheetProtection").Any())
            {
                sd.Protection = new ProtectionData { Protected = true };
            }

            // Auto filter
            var afXml = wsXml.Descendants(ns + "autoFilter").FirstOrDefault();
            if (afXml != null)
            {
                sd.AutoFilter = new AutoFilterData
                {
                    Reference = (string)afXml.Attribute("ref"),
                    AllAttributes = ReadAllAttrs(afXml)
                };
            }

            // Sheet views (freeze panes)
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

            // Row breaks
            var rowBreaks = wsXml.Descendants(ns + "rowBreaks").FirstOrDefault();
            if (rowBreaks != null)
            {
                foreach (var brk in rowBreaks.Elements(ns + "brk"))
                {
                    sd.PageBreaks.Add(new PageBreakData
                    {
                        Type = "row",
                        Row = (int?)brk.Attribute("id")
                    });
                }
            }

            // Column breaks
            var colBreaks = wsXml.Descendants(ns + "colBreaks").FirstOrDefault();
            if (colBreaks != null)
            {
                foreach (var brk in colBreaks.Elements(ns + "brk"))
                {
                    sd.PageBreaks.Add(new PageBreakData
                    {
                        Type = "column",
                        Column = (int?)brk.Attribute("id")
                    });
                }
            }

            // Comments
            ExtractComments(wsPart, sd);

            // Drawings (images/shapes)
            ExtractDrawings(wsPart, sd);

            result.Workbook.Sheets.Add(sd);
            sheetIndex++;
        }
    }

    private void ExtractComments(WorksheetPart wsPart, Models.SheetData sd)
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

    private void ExtractDrawings(WorksheetPart wsPart, Models.SheetData sd)
    {
        try
        {
            var drawingsPart = wsPart.DrawingsPart;
            if (drawingsPart == null) return;

            // Images from ImageParts
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

    private List<string> LoadSharedStrings(WorkbookPart wbPart)
    {
        var result = new List<string>();
        try
        {
            var ssp = wbPart.SharedStringTablePart;
            if (ssp?.GetStream() == null) return result;

            var sx = XDocument.Load(ssp.GetStream());
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            foreach (var si in sx.Descendants(ns + "si"))
            {
                var t = si.Element(ns + "t");
                result.Add(t?.Value ?? si.Value);
            }
        }
        catch { }
        return result;
    }

    private string? ReadColorAttr(XElement color)
    {
        if (color == null) return null;
        var rgb = (string)color.Attribute("rgb");
        if (rgb != null) return "#" + rgb;
        var theme = (string)color.Attribute("theme");
        if (theme != null) return $"theme:{theme}";
        var indexed = (string)color.Attribute("indexed");
        if (indexed != null) return $"index:{indexed}";
        var auto = (string)color.Attribute("auto");
        if (auto != null) return "auto";
        return null;
    }

    private int ParseColumnIndex(string reference)
    {
        if (string.IsNullOrEmpty(reference)) return 0;
        int col = 0;
        foreach (var c in reference.TakeWhile(char.IsLetter))
        {
            col = col * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        }
        return col;
    }

    private int ParseRowIndex(string reference)
    {
        if (string.IsNullOrEmpty(reference)) return 0;
        var rowPart = new string(reference.SkipWhile(char.IsLetter).ToArray());
        return int.TryParse(rowPart, out var row) ? row : 0;
    }

    private Dictionary<string, object?> ReadAllAttrs(XElement el)
    {
        var dict = new Dictionary<string, object?>();
        foreach (var attr in el.Attributes())
            dict[attr.Name.LocalName] = attr.Value;
        return dict;
    }
}
