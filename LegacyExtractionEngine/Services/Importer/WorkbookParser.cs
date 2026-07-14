using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services.Importer;

public class WorkbookParser
{
    public WorkbookData Parse(WorkbookPart wbPart)
    {
        var wb = new WorkbookData();
        var wbXml = XDocument.Load(wbPart.GetStream());
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace nsr = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var sheets = new List<SheetMetadataItem>();
        foreach (var sheetEl in wbXml.Descendants(ns + "sheet"))
        {
            sheets.Add(new SheetMetadataItem
            {
                Name = (string)sheetEl.Attribute("name") ?? "",
                Id = int.TryParse((string)sheetEl.Attribute("sheetId") ?? "0", out var sid) ? sid : 0,
                State = (string)sheetEl.Attribute("state") ?? "visible",
                RelId = (string)sheetEl.Attribute(nsr + "id") ?? ""
            });
        }
        wb.SheetMetadata = sheets;

        foreach (var dn in wbXml.Descendants(ns + "definedNames").Elements(ns + "definedName"))
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

        var customPart = wbPart.GetPartsOfType<CustomFilePropertiesPart>().FirstOrDefault();
        if (customPart?.GetStream() != null)
        {
            try
            {
                var custXml = XDocument.Load(customPart.GetStream());
                XNamespace nsct = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
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

        var wbProtection = wbXml.Descendants(ns + "workbookProtection").FirstOrDefault();
        if (wbProtection != null)
        {
            wb.WorkbookProtection = new ProtectionData { Protected = true };
        }

        return wb;
    }
}
