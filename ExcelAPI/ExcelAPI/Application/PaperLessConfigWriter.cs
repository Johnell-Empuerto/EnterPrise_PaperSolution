using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Microsoft.Extensions.Logging;
using WbDef = ExcelAPI.Models.WorkbookDefinition;

namespace ExcelAPI.Application
{
    /// <summary>
    /// Embeds PaperLess field configuration as a VeryHidden worksheet inside the exported XLSX.
    ///
    /// The PaperLessConfig sheet stores a minified JSON blob containing stable field IDs,
    /// styles, and configuration that survive export → re-upload round trips.
    ///
    /// Pattern: follows ConMasCompatibleWorkbookWriter.PostProcessZipForConMas — ZIP-level
    /// manipulation after COM save.
    /// </summary>
    public class PaperLessConfigWriter : IPaperLessConfigWriter
    {
        private const string SheetName = "PaperLessConfig";
        private const string WorksheetPath = "xl/worksheets/paperlessconfig.xml";
        private const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        private const string RelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

        private readonly ILogger<PaperLessConfigWriter> _logger;
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = false,
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
        };

        public PaperLessConfigWriter(ILogger<PaperLessConfigWriter> logger)
        {
            _logger = logger;
        }

        public void WritePaperLessConfig(WbDef.WorkbookDefinition definition, string outputPath)
        {
            if (!File.Exists(outputPath))
            {
                _logger.LogWarning("[PAPERLESS CONFIG] Workbook not found: {Path}", outputPath);
                return;
            }

            _logger.LogInformation("[PAPERLESS CONFIG] Writing configuration");

            try
            {
                // ═════════════════════════════════════════════════════════
                // PAPERLESS DEBUG STAGE 4 — PaperLessConfigWriter Input
                // ═════════════════════════════════════════════════════════
                _logger.LogInformation("========================================================");
                _logger.LogInformation("PAPERLESS DEBUG STAGE 4 — PaperLessConfigWriter Input");
                _logger.LogInformation("========================================================");
                var config = BuildConfig(definition);
                if (config.Sheets != null)
                {
                    foreach (var cs in config.Sheets)
                    {
                        _logger.LogInformation("  Config Sheet: {Name}", cs.Name);
                        if (cs.Fields != null)
                        {
                            foreach (var cf in cs.Fields.Take(3))
                            {
                                _logger.LogInformation("    Field ID: {Id}", cf.Id);
                                _logger.LogInformation("    Cell: {Cell}", cf.Cell);
                                _logger.LogInformation("    FontSize: {Sz}", cf.Style?.Font?.SizePt ?? 0);
                                _logger.LogInformation("    FontName: {Fn}", cf.Style?.Font?.Name ?? "(null)");
                                _logger.LogInformation("    Bold: {B}", cf.Style?.Font?.Bold ?? false);
                            }
                        }
                    }
                }
                _logger.LogInformation("========================================================");

                // ═════════════════════════════════════════════════════════
                // PAPERLESS DEBUG STAGE 5 — Serialized JSON
                // ═════════════════════════════════════════════════════════
                string json = JsonSerializer.Serialize(config, JsonOptions);
                _logger.LogInformation("========================================================");
                _logger.LogInformation("PAPERLESS DEBUG STAGE 5 — Serialized JSON");
                _logger.LogInformation("========================================================");
                // Find p1f1 fragment in JSON
                int p1f1Idx = json.IndexOf("\"id\":\"p1f1\"", StringComparison.OrdinalIgnoreCase);
                if (p1f1Idx >= 0)
                {
                    int start = Math.Max(0, p1f1Idx - 60);
                    int len = Math.Min(json.Length - start, 300);
                    _logger.LogInformation("p1f1 JSON context: {Ctx}", json.Substring(start, len));
                }
                else
                {
                    _logger.LogInformation("p1f1 not found in JSON — logging first 300 chars");
                    _logger.LogInformation("JSON preview: {Preview}",
                        json.Length <= 300 ? json : json[..300] + "...");
                }
                _logger.LogInformation("Total JSON length: {Len} bytes", json.Length);
                _logger.LogInformation("========================================================");

                using var pkg = ZipFile.Open(outputPath, ZipArchiveMode.Update);

                int fieldCount = config.Sheets?.Sum(s => s.Fields?.Count ?? 0) ?? 0;
                _logger.LogInformation("[PAPERLESS CONFIG] Fields serialized: {Count}", fieldCount);

                // 2. Remove existing PaperLessConfig worksheet (idempotent)
                RemoveExistingConfig(pkg);

                // 3. Create the worksheet XML with inline string (no shared string dependency)
                string worksheetXml = BuildWorksheetXml(json);
                var worksheetEntry = pkg.CreateEntry(WorksheetPath);
                using (var w = new StreamWriter(worksheetEntry.Open()))
                    w.Write(worksheetXml);

                // 4. Compute next available relationship ID
                string relId = ComputeNextRelId(pkg);

                // 5. Update workbook.xml — add sheet entry with veryHidden state
                UpdateWorkbookXml(pkg, relId);

                // 6. Update workbook.xml.rels — add relationship
                UpdateWorkbookRels(pkg, relId);

                // 7. Update [Content_Types].xml — add override
                UpdateContentTypes(pkg);

                _logger.LogInformation("[PAPERLESS CONFIG] Configuration sheet created ({Sheet})", WorksheetPath);
                _logger.LogInformation("[PAPERLESS CONFIG] Configuration persisted successfully");

                // ═════════════════════════════════════════════════════════
                // PAPERLESS DEBUG STAGE 6 — XLSX ZIP Verification
                // ═════════════════════════════════════════════════════════
                _logger.LogInformation("========================================================");
                _logger.LogInformation("PAPERLESS DEBUG STAGE 6 — XLSX ZIP Verification");
                _logger.LogInformation("========================================================");
                try
                {
                    // Reopen the XLSX as ZIP (read-only) and verify the PaperLessConfig worksheet
                    using var verifyPkg = ZipFile.OpenRead(outputPath);
                    var verifyEntry = verifyPkg.GetEntry(WorksheetPath);
                    if (verifyEntry != null)
                    {
                        string verifyXml;
                        using (var vr = new StreamReader(verifyEntry.Open()))
                            verifyXml = vr.ReadToEnd();

                        // Extract JSON from B1 inline string: <c r="B1" t="inlineStr"><is><t>...</t></is></c>
                        var b1Match = System.Text.RegularExpressions.Regex.Match(verifyXml,
                            @"<c\s+r=""B1""[^>]*>.*?<is><t[^>]*>(.*?)</t></is></c>",
                            System.Text.RegularExpressions.RegexOptions.Singleline);
                        if (b1Match.Success)
                        {
                            string storedJson = System.Net.WebUtility.HtmlDecode(b1Match.Groups[1].Value);
                            _logger.LogInformation("STAGE 6 — Config worksheet found. Stored JSON length: {Len}", storedJson.Length);
                            int p1f1Pos = storedJson.IndexOf("\"id\":\"p1f1\"", StringComparison.OrdinalIgnoreCase);
                            if (p1f1Pos >= 0)
                            {
                                int ctxStart = Math.Max(0, p1f1Pos - 60);
                                int ctxLen = Math.Min(storedJson.Length - ctxStart, 300);
                                _logger.LogInformation("STAGE 6 — p1f1 in stored JSON: {Ctx}", storedJson.Substring(ctxStart, ctxLen));
                            }
                            else
                            {
                                _logger.LogInformation("STAGE 6 — p1f1 NOT FOUND in stored JSON");
                                _logger.LogInformation("STAGE 6 — First 300 chars: {Ctx}",
                                    storedJson.Length <= 300 ? storedJson : storedJson[..300] + "...");
                            }
                        }
                        else
                        {
                            _logger.LogWarning("STAGE 6 — Could not extract JSON from B1 cell");
                            _logger.LogWarning("STAGE 6 — Worksheet XML fragment: {Xml}",
                                verifyXml.Length <= 500 ? verifyXml : verifyXml[..500] + "...");
                        }
                    }
                    else
                    {
                        _logger.LogWarning("STAGE 6 — PaperLessConfig worksheet NOT FOUND in ZIP at {Path}", WorksheetPath);
                    }
                }
                catch (Exception exStage6)
                {
                    _logger.LogWarning(exStage6, "STAGE 6 — ZIP verification failed: {Msg}", exStage6.Message);
                }
                _logger.LogInformation("========================================================");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[PAPERLESS CONFIG] Failed to write configuration: {Msg}", ex.Message);
            }
        }

        private static WbDef.PaperLessConfig BuildConfig(WbDef.WorkbookDefinition definition)
        {
            var config = new WbDef.PaperLessConfig
            {
                SchemaVersion = 1,
                PaperLess = new WbDef.PaperLessInfo { Version = "1.0" }
            };

            if (definition.Sheets == null) return config;

            foreach (var sheet in definition.Sheets)
            {
                var ps = new WbDef.PaperLessSheet { Name = sheet.Name ?? "" };

                if (sheet.Fields != null)
                {
                    foreach (var field in sheet.Fields)
                    {
                        var pf = new WbDef.PaperLessField
                        {
                            Id = field.Id ?? "",
                            Cell = field.Cell?.Address ?? "",
                            Type = field.Type.ToString(),
                            Style = CloneCellStyle(field.Style),
                            Config = new WbDef.PaperLessFieldConfig
                            {
                                Required = field.Required,
                                MinLength = 0,
                                MaxLength = field.MaxLength,
                                InputRestriction = field.Type switch
                                {
                                    WbDef.FieldType.Number => "Numeric",
                                    WbDef.FieldType.Date => "Date",
                                    _ => "None"
                                },
                                Lines = 1,
                                ValidateOnEditing = field.ValidateOnEditing,
                                ReadOnly = field.Locked,
                                Hidden = !field.Visible,
                                Placeholder = field.Placeholder,
                                DefaultValue = field.DefaultValue
                            }
                        };
                        ps.Fields.Add(pf);
                    }
                }

                config.Sheets.Add(ps);
            }

            return config;
        }

        /// <summary>
        /// Clone a CellStyle with only non-default properties to minimize JSON size.
        /// Returns null if the style has no meaningful properties.
        /// </summary>
        private static WbDef.CellStyle? CloneCellStyle(WbDef.CellStyle? style)
        {
            if (style == null) return null;

            bool hasFont = style.Font != null &&
                (!string.IsNullOrEmpty(style.Font.Name) || style.Font.SizePt > 0 ||
                 style.Font.Bold || style.Font.Italic || style.Font.Underline ||
                 !string.IsNullOrEmpty(style.Font.ColorArgb));

            bool hasFill = style.Fill != null &&
                !string.IsNullOrEmpty(style.Fill.ColorArgb) &&
                style.Fill.PatternType != "none";

            bool hasAlign = style.Alignment != null &&
                (!string.IsNullOrEmpty(style.Alignment.Horizontal) ||
                 !string.IsNullOrEmpty(style.Alignment.Vertical));

            if (!hasFont && !hasFill && !hasAlign && !style.WrapText)
                return null;

            return style;
        }

        private static string BuildWorksheetXml(string json)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
            sb.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<sheetData>");
            sb.Append("<row r=\"1\">");
            sb.Append("<c r=\"A1\" t=\"inlineStr\"><is><t>ConfigurationJson</t></is></c>");
            sb.Append("<c r=\"B1\" t=\"inlineStr\"><is><t>");
            sb.Append(System.Security.SecurityElement.Escape(json));
            sb.Append("</t></is></c>");
            sb.Append("</row>");
            sb.Append("</sheetData>");
            sb.Append("<pageMargins left=\"0\" right=\"0\" top=\"0\" bottom=\"0\" header=\"0\" footer=\"0\"/>");
            sb.Append("</worksheet>");
            return sb.ToString();
        }

        private void RemoveExistingConfig(ZipArchive pkg)
        {
            var existing = pkg.GetEntry(WorksheetPath);
            if (existing != null)
            {
                existing.Delete();
                _logger.LogInformation("[PAPERLESS CONFIG] Removed existing configuration sheet");
            }
        }

        private static string ComputeNextRelId(ZipArchive pkg)
        {
            int maxRelId = 0;
            var relsEntry = pkg.GetEntry("xl/_rels/workbook.xml.rels");
            if (relsEntry != null)
            {
                string relsXml;
                using (var r = new StreamReader(relsEntry.Open()))
                    relsXml = r.ReadToEnd();
                var relsDoc = XDocument.Parse(relsXml);
                XNamespace relNs = relsDoc.Root.Name.Namespace;
                foreach (var rel in relsDoc.Root.Elements(relNs + "Relationship"))
                {
                    var idAttr = (string)rel.Attribute("Id");
                    if (idAttr != null && idAttr.StartsWith("rId") && int.TryParse(idAttr[3..], out int rid))
                        if (rid > maxRelId) maxRelId = rid;
                }
            }
            return "rId" + (maxRelId + 1);
        }

        private static void UpdateWorkbookXml(ZipArchive pkg, string relId)
        {
            var entry = pkg.GetEntry("xl/workbook.xml");
            if (entry == null) return;

            string xml;
            using (var r = new StreamReader(entry.Open()))
                xml = r.ReadToEnd();

            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;

            var sheets = doc.Root.Descendants(ns + "sheets").FirstOrDefault();
            if (sheets == null) return;

            // Remove existing PaperLessConfig sheet entry
            var existingSheet = sheets.Elements(ns + "sheet")
                .FirstOrDefault(s => string.Equals((string)s.Attribute("name"), SheetName, StringComparison.OrdinalIgnoreCase));
            if (existingSheet != null)
                existingSheet.Remove();

            uint maxSheetId = 0;
            foreach (var sheet in sheets.Elements(ns + "sheet"))
            {
                uint sid = (uint)sheet.Attribute("sheetId");
                if (sid > maxSheetId) maxSheetId = sid;
            }
            uint newSheetId = maxSheetId + 1;

            // Add with veryHidden state
            var newSheet = new XElement(ns + "sheet",
                new XAttribute("name", SheetName),
                new XAttribute("sheetId", newSheetId),
                new XAttribute("state", "veryHidden"),
                new XAttribute(XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships") + "id", relId));
            sheets.Add(newSheet);

            entry.Delete();
            var newEntry = pkg.CreateEntry("xl/workbook.xml");
            using (var w = new StreamWriter(newEntry.Open()))
                doc.Save(w);
        }

        private static void UpdateWorkbookRels(ZipArchive pkg, string relId)
        {
            var entry = pkg.GetEntry("xl/_rels/workbook.xml.rels");
            if (entry == null) return;

            string xml;
            using (var r = new StreamReader(entry.Open()))
                xml = r.ReadToEnd();

            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;

            // Remove existing PaperLessConfig relationship
            var existingRel = doc.Root.Elements(ns + "Relationship")
                .FirstOrDefault(rel => string.Equals((string)rel.Attribute("Target"), "worksheets/paperlessconfig.xml", StringComparison.OrdinalIgnoreCase));
            if (existingRel != null)
                existingRel.Remove();

            var newRel = new XElement(ns + "Relationship",
                new XAttribute("Id", relId),
                new XAttribute("Type", RelationshipType),
                new XAttribute("Target", "worksheets/paperlessconfig.xml"));
            doc.Root.Add(newRel);

            entry.Delete();
            var newEntry = pkg.CreateEntry("xl/_rels/workbook.xml.rels");
            using (var w = new StreamWriter(newEntry.Open()))
                doc.Save(w);
        }

        private static void UpdateContentTypes(ZipArchive pkg)
        {
            var entry = pkg.GetEntry("[Content_Types].xml");
            if (entry == null) return;

            string xml;
            using (var r = new StreamReader(entry.Open()))
                xml = r.ReadToEnd();

            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;

            // Remove existing PaperLessConfig override
            var existingOverride = doc.Root.Elements(ns + "Override")
                .FirstOrDefault(o => string.Equals((string)o.Attribute("PartName"), "/xl/worksheets/paperlessconfig.xml", StringComparison.OrdinalIgnoreCase));
            if (existingOverride != null)
                existingOverride.Remove();

            var newOverride = new XElement(ns + "Override",
                new XAttribute("PartName", "/xl/worksheets/paperlessconfig.xml"),
                new XAttribute("ContentType", ContentType));
            doc.Root.Add(newOverride);

            entry.Delete();
            var newEntry = pkg.CreateEntry("[Content_Types].xml");
            using (var w = new StreamWriter(newEntry.Open()))
                doc.Save(w);
        }
    }
}
