using System.Xml.Linq;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services.Importer;

public class PageSetupParser
{
    public PageSetupData ParsePageSetup(XElement wsXml)
    {
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var ps = new PageSetupData();

        var psXml = wsXml.Descendants(ns + "pageSetup").FirstOrDefault();
        if (psXml != null)
        {
            ps.PaperSize = (int?)psXml.Attribute("paperSize");
            ps.Orientation = (string)psXml.Attribute("orientation");
            ps.Scale = (int?)psXml.Attribute("scale");
            ps.FitToWidth = (int?)psXml.Attribute("fitToWidth");
            ps.FitToHeight = (int?)psXml.Attribute("fitToHeight");
            ps.PageHeight = (double?)psXml.Attribute("pageHeight");
            ps.PageWidth = (double?)psXml.Attribute("pageWidth");
            ps.BlackAndWhite = (bool?)psXml.Attribute("blackAndWhite");
            ps.Draft = (bool?)psXml.Attribute("draft");
            ps.CellComments = (string)psXml.Attribute("cellComments");
            ps.Errors = (string)psXml.Attribute("errors");
            ps.HorizontalDpi = (int?)psXml.Attribute("horizontalDpi");
            ps.VerticalDpi = (int?)psXml.Attribute("verticalDpi");
            ps.Copies = (int?)psXml.Attribute("copies");
            ps.FirstPageNumber = (int?)psXml.Attribute("firstPageNumber");
            ps.UseFirstPageNumber = (bool?)psXml.Attribute("useFirstPageNumber");
            ps.PageOrder = (string)psXml.Attribute("pageOrder");
        }

        return ps;
    }

    public PageMarginsData ParsePageMargins(XElement wsXml)
    {
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var pm = new PageMarginsData();

        var pmXml = wsXml.Descendants(ns + "pageMargins").FirstOrDefault();
        if (pmXml != null)
        {
            pm.Top = (double?)pmXml.Attribute("top") ?? 0.75;
            pm.Bottom = (double?)pmXml.Attribute("bottom") ?? 0.75;
            pm.Left = (double?)pmXml.Attribute("left") ?? 0.70;
            pm.Right = (double?)pmXml.Attribute("right") ?? 0.70;
            pm.Header = (double?)pmXml.Attribute("header") ?? 0.30;
            pm.Footer = (double?)pmXml.Attribute("footer") ?? 0.30;
        }

        return pm;
    }

    public void ParsePrintOptions(XElement wsXml, PageSetupData ps)
    {
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var po = wsXml.Descendants(ns + "printOptions").FirstOrDefault();
        if (po == null) return;

        foreach (var attr in po.Attributes())
        {
            ps.AllAttributes[attr.Name.LocalName] = attr.Value;
        }
    }

    public SheetFormatPropertiesData? ParseSheetFormatProperties(XElement wsXml)
    {
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var sfp = wsXml.Descendants(ns + "sheetFormatPr").FirstOrDefault();
        if (sfp == null) return null;

        return new SheetFormatPropertiesData
        {
            DefaultRowHeight = (double?)sfp.Attribute("defaultRowHeight"),
            DefaultColumnWidth = (double?)sfp.Attribute("defaultColWidth"),
            OutlineLevelRow = (int?)sfp.Attribute("outlineLevelRow"),
            OutlineLevelColumn = (int?)sfp.Attribute("outlineLevelCol")
        };
    }
}
