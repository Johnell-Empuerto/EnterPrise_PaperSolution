using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using System.Security.Cryptography;
using System.Text;

// Resolve Hyperlink ambiguity — prefer Spreadsheet hyperlink
using Hyperlink = DocumentFormat.OpenXml.Spreadsheet.Hyperlink;

namespace ExcelAPI.Application
{
    // ═══════════════════════════════════════════════════════════════════════
    // WorkbookDiffResult — comprehensive validation outcome
    // ═══════════════════════════════════════════════════════════════════════

    public class WorkbookDiffResult
    {
        // ── Workbook structure ──
        public int SheetCountChanges { get; set; }
        public int SheetNameChanges { get; set; }
        public int SheetVisibilityChanges { get; set; }
        public int DefinedNameChanges { get; set; }
        public int WorkbookProtectionChanges { get; set; }
        public int VBAChanges { get; set; }

        // ── Worksheet structure ──
        public int RowCountChanges { get; set; }
        public int HiddenRowChanges { get; set; }
        public int ColumnCountChanges { get; set; }
        public int HiddenColumnChanges { get; set; }
        public int RowHeightChanges { get; set; }
        public int ColumnWidthChanges { get; set; }
        public int FreezePaneChanges { get; set; }

        // ── Formatting ──
        public int StyleChanges { get; set; }
        public int FontChanges { get; set; }
        public int FillChanges { get; set; }
        public int BorderChanges { get; set; }
        public int AlignmentChanges { get; set; }
        public int NumberFormatChanges { get; set; }

        // ── Layout ──
        public int MergeChanges { get; set; }
        public int PrintAreaChanges { get; set; }
        public int PageMarginChanges { get; set; }
        public int PageSetupChanges { get; set; }
        public int HeaderFooterChanges { get; set; }

        // ── Objects ──
        public int DrawingChanges { get; set; }
        public int ImageChanges { get; set; }
        public int HyperlinkChanges { get; set; }
        public int CommentChanges { get; set; }
        public int DataValidationChanges { get; set; }
        public int ConditionalFormattingChanges { get; set; }
        public int TableChanges { get; set; }

        // ── Formulas and values ──
        public int FormulaChanges { get; set; }
        public int EditableValueChanges { get; set; }
        public int NewCellChanges { get; set; }
        public int MissingCellChanges { get; set; }

        // ── Phase 5.0: Additional parts ──
        public int SharedStringsChanges { get; set; }
        public int WorkbookXmlChanges { get; set; }
        public int ExternalLinksChanges { get; set; }
        public int CustomXmlChanges { get; set; }
        public int PrinterSettingsChanges { get; set; }
        public int RelationshipsChanges { get; set; }

        // ── Binary hashes ──
        public List<XmlHashResult> PartHashes { get; set; } = new();

        // ── Detailed mismatches ──
        public List<string> Details { get; set; } = new();
        public List<string> EditableChangedCells { get; set; } = new();

        // ── Computed ──
        public int TotalDifferences =>
            SheetCountChanges + SheetNameChanges + SheetVisibilityChanges +
            DefinedNameChanges + WorkbookProtectionChanges + VBAChanges +
            RowCountChanges + HiddenRowChanges + ColumnCountChanges +
            HiddenColumnChanges + RowHeightChanges + ColumnWidthChanges +
            FreezePaneChanges +
            StyleChanges + FontChanges + FillChanges + BorderChanges +
            AlignmentChanges + NumberFormatChanges +
            MergeChanges + PrintAreaChanges + PageMarginChanges +
            PageSetupChanges + HeaderFooterChanges +
            DrawingChanges + ImageChanges + HyperlinkChanges +
            DataValidationChanges +
            ConditionalFormattingChanges + TableChanges +
            FormulaChanges + NewCellChanges + MissingCellChanges +
            SharedStringsChanges + WorkbookXmlChanges + ExternalLinksChanges +
            CustomXmlChanges + PrinterSettingsChanges + RelationshipsChanges;

        public bool Passed => TotalDifferences == 0;

        // ═══════════════════════════════════════════════════════════════════
        // GetFormattedReport — Structured Diagnostic Output
        // Produces a clean, human-readable report of every validation category.
        // Non-zero categories are highlighted with ❌; clean categories with ✅.
        // Includes part hashes, mismatch details, and cell value changes.
        // ═══════════════════════════════════════════════════════════════════

        public string GetFormattedReport()
        {
            var sb = new StringBuilder();
            sb.AppendLine("╔══════════════════════════════════════════════════════════╗");
            sb.AppendLine("║          WORKBOOK FIDELITY REPORT                     ║");
            sb.AppendLine("╚══════════════════════════════════════════════════════════╝");
            sb.AppendLine();

            AppendCategory(sb, "Workbook Structure", new (string, int)[]
            {
                ("SheetCountChanges", SheetCountChanges),
                ("SheetNameChanges", SheetNameChanges),
                ("SheetVisibilityChanges", SheetVisibilityChanges),
                ("DefinedNameChanges", DefinedNameChanges),
                ("WorkbookProtectionChanges", WorkbookProtectionChanges),
                ("VBAChanges", VBAChanges),
            });

            AppendCategory(sb, "Worksheet Geometry", new (string, int)[]
            {
                ("RowCountChanges", RowCountChanges),
                ("HiddenRowChanges", HiddenRowChanges),
                ("ColumnCountChanges", ColumnCountChanges),
                ("HiddenColumnChanges", HiddenColumnChanges),
                ("RowHeightChanges", RowHeightChanges),
                ("ColumnWidthChanges", ColumnWidthChanges),
                ("FreezePaneChanges", FreezePaneChanges),
            });

            AppendCategory(sb, "Formatting", new (string, int)[]
            {
                ("StyleChanges", StyleChanges),
                ("FontChanges", FontChanges),
                ("FillChanges", FillChanges),
                ("BorderChanges", BorderChanges),
                ("AlignmentChanges", AlignmentChanges),
                ("NumberFormatChanges", NumberFormatChanges),
            });

            AppendCategory(sb, "Layout", new (string, int)[]
            {
                ("MergeChanges", MergeChanges),
                ("PrintAreaChanges", PrintAreaChanges),
                ("PageMarginChanges", PageMarginChanges),
                ("PageSetupChanges", PageSetupChanges),
                ("HeaderFooterChanges", HeaderFooterChanges),
            });

            AppendCategory(sb, "Objects", new (string, int)[]
            {
                ("DrawingChanges", DrawingChanges),
                ("ImageChanges", ImageChanges),
                ("HyperlinkChanges", HyperlinkChanges),
                ("CommentChanges", CommentChanges),
                ("DataValidationChanges", DataValidationChanges),
                ("ConditionalFormattingChanges", ConditionalFormattingChanges),
                ("TableChanges", TableChanges),
            });

            AppendCategory(sb, "Formulas & Cells", new (string, int)[]
            {
                ("FormulaChanges", FormulaChanges),
                ("NewCellChanges", NewCellChanges),
                ("MissingCellChanges", MissingCellChanges),
            });

            AppendCategory(sb, "Phase 5.0", new (string, int)[]
            {
                ("SharedStringsChanges", SharedStringsChanges),
                ("WorkbookXmlChanges", WorkbookXmlChanges),
                ("ExternalLinksChanges", ExternalLinksChanges),
                ("CustomXmlChanges", CustomXmlChanges),
                ("PrinterSettingsChanges", PrinterSettingsChanges),
                ("RelationshipsChanges", RelationshipsChanges),
            });

            sb.AppendLine();
            sb.AppendLine("─────────────────────────────────────────────────────");
            sb.AppendLine($"  EditableValueChanges............{EditableValueChanges,3}");
            sb.AppendLine($"  TOTAL STRUCTURAL DIFFERENCES....{TotalDifferences,3}");
            sb.AppendLine($"  VALIDATION......................{(Passed ? "PASS" : "FAIL")}");
            sb.AppendLine("─────────────────────────────────────────────────────");

            // Part Hashes
            if (PartHashes.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("── Part Hashes ──");
                foreach (var hash in PartHashes)
                {
                    string status = hash.Match ? "✅" : "❌";
                    sb.AppendLine($"  {status} {hash.PartName}");
                    sb.AppendLine($"     Original: {hash.OriginalHash}");
                    sb.AppendLine($"     Edited:   {hash.EditedHash}");
                }
            }

            // Detail Mismatches
            if (Details.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("── Mismatch Details ──");
                for (int i = 0; i < Details.Count; i++)
                {
                    sb.AppendLine($"  {i + 1}. {Details[i]}");
                }
            }

            // Editable cell changes
            if (EditableChangedCells.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("── Cell Value Changes ──");
                foreach (var cell in EditableChangedCells)
                {
                    sb.AppendLine($"  {cell}");
                }
            }

            sb.AppendLine();
            sb.AppendLine("╚══════════════════════════════════════════════════════════╝");
            return sb.ToString();
        }

        private static void AppendCategory(StringBuilder sb, string heading, (string Name, int Value)[] fields)
        {
            sb.AppendLine($"── {heading} ──");
            bool hasAny = false;
            foreach (var (name, value) in fields)
            {
                if (value != 0)
                {
                    sb.AppendLine($"  ❌ {name}...{value}");
                    hasAny = true;
                }
                else
                {
                    sb.AppendLine($"  ✅ {name}...0");
                }
            }
            sb.AppendLine();
        }
    }

    public class XmlHashResult
    {
        public string PartName { get; set; } = "";
        public string OriginalHash { get; set; } = "";
        public string EditedHash { get; set; } = "";
        public bool Match { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // WorkbookFidelityException — thrown when validation fails
    // ═══════════════════════════════════════════════════════════════════════

    public class WorkbookFidelityException : Exception
    {
        public WorkbookDiffResult ValidationResult { get; }

        public WorkbookFidelityException(WorkbookDiffResult result)
            : base(result.GetFormattedReport())
        {
            ValidationResult = result;
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // WorkbookDiffValidator — comprehensive fidelity checker
    // ═══════════════════════════════════════════════════════════════════════

    public class WorkbookDiffValidator
    {
        private readonly ILogger<WorkbookDiffValidator> _logger;

        private static readonly HashSet<string> KnownAdditionalSheets =
            new(StringComparer.OrdinalIgnoreCase) { "ExcelOutputSetting" };

        public WorkbookDiffValidator(ILogger<WorkbookDiffValidator> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Compare original and edited workbooks. Returns a comprehensive diff report.
        /// Checks every category specified in Phase 4.5.
        /// </summary>
        public WorkbookDiffResult Compare(string originalPath, string editedPath)
        {
            var result = new WorkbookDiffResult();

            using var origDoc = SpreadsheetDocument.Open(originalPath, false);
            using var editDoc = SpreadsheetDocument.Open(editedPath, false);

            var origWbPart = origDoc.WorkbookPart;
            var editWbPart = editDoc.WorkbookPart;

            if (origWbPart == null || editWbPart == null)
            {
                result.Details.Add("One or both files have no WorkbookPart.");
                return result;
            }

            // ── 0. Binary XML Hash Comparison ──
            ComparePartHashes(origDoc, editDoc, result);

            // ── 1. Workbook-level Checks ──
            CompareWorkbook(origWbPart, editWbPart, result);

            // ── 2. Stylesheet Check (resolved) ──
            CompareStyles(origWbPart, editWbPart, result);

            // ── 3. Per-Sheet Checks ──
            var origSheets = origWbPart.Workbook.Descendants<Sheet>().ToList();
            var editSheets = editWbPart.Workbook.Descendants<Sheet>().ToList();

            foreach (var editSheet in editSheets)
            {
                if (editSheet.Id == null || editSheet.Name == null) continue;

                var editWsPart = editWbPart.GetPartById(editSheet.Id) as WorksheetPart;
                var origSheet = origSheets.FirstOrDefault(s =>
                    s.Name == editSheet.Name || s.Id?.Value == editSheet.Id?.Value);
                if (origSheet?.Id == null) continue;
                var origWsPart = origWbPart.GetPartById(origSheet.Id) as WorksheetPart;

                if (origWsPart == null || editWsPart == null) continue;

                CompareWorksheet(origWsPart, editWsPart, editSheet.Name, editWbPart, result, origDoc);
            }

            // ── 4. Object Parts (drawings, images, comments, etc.) ──
            CompareObjectParts(origWbPart, editWbPart, result, origSheets, editSheets);

            // ── 5. Additional Parts (Phase 5.0) ──
            CompareAdditionalParts(origDoc, editDoc, origWbPart, editWbPart, result);

            // ── 6. Defined Names ──
            CompareDefinedNames(origWbPart, editWbPart, result);

            // ── Log Summary (compact) ──
            _logger.LogInformation(
                "========== WORKBOOK VALIDATION ==========\n" +
                "Original: {Orig}\nEdited: {Edit}\n" +
                "SheetCount: {V0}, Style: {V1}, Merge: {V2}, PageSetup: {V5}\n" +
                "ValueChanges: {VC}\nTotalDiffs: {TD}\nValidation: {VP}\n" +
                "========================================",
                originalPath, editedPath,
                result.SheetCountChanges, result.StyleChanges, result.MergeChanges,
                result.PageSetupChanges,
                result.EditableValueChanges, result.TotalDifferences,
                result.Passed ? "PASS" : "FAIL");

            if (!result.Passed)
            {
                // Log the full formatted report on failure
                _logger.LogWarning("Workbook validation FAILED. Full report:\n{Report}", result.GetFormattedReport());
            }

            return result;
        }

        // ═══════════════════════════════════════════════════════════════════
        // Binary XML Hash Comparison
        // ═══════════════════════════════════════════════════════════════════

        private void ComparePartHashes(
            SpreadsheetDocument origDoc, SpreadsheetDocument editDoc, WorkbookDiffResult result)
        {
            var origParts = origDoc.Parts.ToDictionary(p => p.OpenXmlPart.Uri.ToString(), p => p.OpenXmlPart);
            var editParts = editDoc.Parts.ToDictionary(p => p.OpenXmlPart.Uri.ToString(), p => p.OpenXmlPart);

            var partsToCheck = new[] { "xl/styles.xml", "xl/theme/theme1.xml" };

            foreach (var partName in partsToCheck)
            {
                string origHash = "";
                string editHash = "";

                if (origParts.TryGetValue(partName, out var origPart) && origPart is not WorksheetPart)
                {
                    using var ms = new MemoryStream();
                    origPart.GetStream().CopyTo(ms);
                    origHash = ComputeSha256(ms.ToArray());
                }

                if (editParts.TryGetValue(partName, out var editPart) && editPart is not WorksheetPart)
                {
                    using var ms = new MemoryStream();
                    editPart.GetStream().CopyTo(ms);
                    editHash = ComputeSha256(ms.ToArray());
                }

                bool match = string.IsNullOrEmpty(origHash) && string.IsNullOrEmpty(editHash)
                    || origHash == editHash;
                if (!match)
                {
                    result.Details.Add($"Part hash mismatch: {partName} (orig={origHash}, edit={editHash})");
                    result.StyleChanges++; // styles hash mismatch = style change
                }

                result.PartHashes.Add(new XmlHashResult
                {
                    PartName = partName,
                    OriginalHash = origHash,
                    EditedHash = editHash,
                    Match = match
                });
            }

            // Check worksheet XML hashes — they SHOULD differ (cell values changed)
            foreach (var kvp in origParts)
            {
                if (kvp.Value is WorksheetPart)
                {
                    string hashKey = kvp.Key;
                    string origHash = "";
                    string editHash = "";

                    using var ms = new MemoryStream();
                    kvp.Value.GetStream().CopyTo(ms);
                    origHash = ComputeSha256(ms.ToArray());

                    if (editParts.TryGetValue(hashKey, out var ePart))
                    {
                        using var ms2 = new MemoryStream();
                        ePart.GetStream().CopyTo(ms2);
                        editHash = ComputeSha256(ms2.ToArray());
                    }

                    result.PartHashes.Add(new XmlHashResult
                    {
                        PartName = hashKey,
                        OriginalHash = origHash,
                        EditedHash = editHash,
                        Match = origHash == editHash
                    });
                }
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        // Workbook-level Checks
        // ═══════════════════════════════════════════════════════════════════

        private void CompareWorkbook(
            WorkbookPart origWbPart, WorkbookPart editWbPart, WorkbookDiffResult result)
        {
            var origSheets = origWbPart.Workbook.Descendants<Sheet>().ToList();
            var editSheets = editWbPart.Workbook.Descendants<Sheet>().ToList();

            // Phase 5.5.2: Filter out intentionally-added sheets (e.g. ExcelOutputSetting for ConMas)
            var filteredEditSheets = editSheets
                .Where(s => !KnownAdditionalSheets.Contains(s.Name?.Value ?? ""))
                .ToList();

            if (origSheets.Count != filteredEditSheets.Count)
            {
                result.SheetCountChanges = Math.Abs(origSheets.Count - filteredEditSheets.Count);
                result.Details.Add($"Sheet count: original={origSheets.Count}, edited={editSheets.Count} (filtered={filteredEditSheets.Count})");
            }

            for (int i = 0; i < Math.Min(origSheets.Count, filteredEditSheets.Count); i++)
            {
                var o = origSheets[i];
                var e = filteredEditSheets[i];
                if ((o.Name ?? "") != (e.Name ?? ""))
                {
                    result.SheetNameChanges++;
                    result.Details.Add($"Sheet #{i}: name '{o.Name}' vs '{e.Name}'");
                }
                string oState = (o.State?.Value ?? SheetStateValues.Visible).ToString();
                string eState = (e.State?.Value ?? SheetStateValues.Visible).ToString();
                if (oState != eState)
                {
                    result.SheetVisibilityChanges++;
                    result.Details.Add($"Sheet #{i} '{o.Name}': visibility '{oState}' vs '{eState}'");
                }
            }

            // Workbook protection
            var origProt = origWbPart.Workbook.WorkbookProtection;
            var editProt = editWbPart.Workbook.WorkbookProtection;
            if (!XmlEquals(origProt, editProt))
            {
                result.WorkbookProtectionChanges++;
                result.Details.Add($"WorkbookProtection differs\n  Orig: {FormatXmlSnippet(origProt?.OuterXml)}\n  Edit: {FormatXmlSnippet(editProt?.OuterXml)}");
            }

            // VBA presence
            bool origHasVba = origWbPart.VbaProjectPart != null;
            bool editHasVba = editWbPart.VbaProjectPart != null;
            if (origHasVba != editHasVba)
            {
                result.VBAChanges++;
                result.Details.Add($"VBA project: original={(origHasVba ? "present" : "absent")}, " +
                    $"edited={(editHasVba ? "present" : "absent")}");
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        // Stylesheet Comparison
        // ═══════════════════════════════════════════════════════════════════

        private void CompareStyles(
            WorkbookPart origWbPart, WorkbookPart editWbPart, WorkbookDiffResult result)
        {
            var origStylePart = origWbPart.WorkbookStylesPart;
            var editStylePart = editWbPart.WorkbookStylesPart;

            if (origStylePart == null && editStylePart == null) return;
            if (origStylePart == null || editStylePart == null)
            {
                result.StyleChanges++;
                result.Details.Add("One workbook is missing StylesheetPart");
                return;
            }

            var origSst = origStylePart.Stylesheet;
            var editSst = editStylePart.Stylesheet;

            if (!XmlEquals(origSst, editSst))
            {
                result.StyleChanges++;
                result.Details.Add("Stylesheet XML differs (hash should catch this)");
            }

            CompareStyleCollection(origSst?.Fonts, editSst?.Fonts, "Font", result);
            CompareStyleCollection(origSst?.Fills, editSst?.Fills, "Fill", result);
            CompareStyleCollection(origSst?.Borders, editSst?.Borders, "Border", result);
            CompareStyleCollection(origSst?.CellFormats, editSst?.CellFormats, "CellFormat", result);
        }

        private void CompareStyleCollection(
            OpenXmlElement? orig, OpenXmlElement? edit, string label, WorkbookDiffResult result)
        {
            if (orig == null && edit == null) return;
            if (orig == null || edit == null)
            {
                IncrementStyleCounter(label, result);
                result.Details.Add($"Stylesheet {label} count: one is null");
                return;
            }

            var origItems = orig.ChildElements.ToList();
            var editItems = edit.ChildElements.ToList();

            if (origItems.Count != editItems.Count)
            {
                IncrementStyleCounter(label, result);
                result.Details.Add($"Stylesheet {label} count: original={origItems.Count}, edited={editItems.Count}");
                return;
            }

            for (int i = 0; i < origItems.Count; i++)
            {
                if (origItems[i].OuterXml != editItems[i].OuterXml)
                {
                    IncrementStyleCounter(label, result);
                    result.Details.Add($"Stylesheet {label}[{i}] differs\n  Orig: {FormatXmlSnippet(origItems[i].OuterXml)}\n  Edit: {FormatXmlSnippet(editItems[i].OuterXml)}");
                }
            }
        }

        private void IncrementStyleCounter(string label, WorkbookDiffResult result)
        {
            switch (label)
            {
                case "Font": result.FontChanges++; break;
                case "Fill": result.FillChanges++; break;
                case "Border": result.BorderChanges++; break;
                case "CellFormat": result.AlignmentChanges++; break;
                default: result.StyleChanges++; break;
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        // Per-Worksheet Comparison
        // ═══════════════════════════════════════════════════════════════════

        private void CompareWorksheet(
            WorksheetPart origWsPart, WorksheetPart editWsPart, string sheetName,
            WorkbookPart editWbPart, WorkbookDiffResult result,
            SpreadsheetDocument origDoc)
        {
            var origWs = origWsPart.Worksheet;
            var editWs = editWsPart.Worksheet;

            // ── Rows ──
            var origRows = origWs.Descendants<Row>().ToList();
            var editRows = editWs.Descendants<Row>().ToList();

            if (origRows.Count != editRows.Count)
            {
                result.RowCountChanges++;
                result.Details.Add($"Sheet '{sheetName}': row count {origRows.Count} vs {editRows.Count}");
            }

            var origRowMap = origRows.ToDictionary(r => r.RowIndex, r => r);
            var editRowMap = editRows.ToDictionary(r => r.RowIndex, r => r);

            foreach (var (index, origRow) in origRowMap)
            {
                if (!editRowMap.ContainsKey(index))
                {
                    result.RowCountChanges++;
                    result.Details.Add($"Sheet '{sheetName}': Row {index} missing in edited");
                    continue;
                }
                var editRow = editRowMap[index];

                bool origHidden = origRow.Hidden?.Value ?? false;
                bool editHidden = editRow.Hidden?.Value ?? false;
                if (origHidden != editHidden)
                {
                    result.HiddenRowChanges++;
                    result.Details.Add($"Sheet '{sheetName}': Row {index} hidden {origHidden} vs {editHidden}");
                }

                double origHt = origRow.Height?.Value ?? 0.0;
                double editHt = editRow.Height?.Value ?? 0.0;
                if (Math.Abs(origHt - editHt) > 0.01)
                {
                    result.RowHeightChanges++;
                    result.Details.Add($"Sheet '{sheetName}': Row {index} height {origHt:F2} vs {editHt:F2}");
                }
            }

            foreach (var (index, _) in editRowMap)
            {
                if (!origRowMap.ContainsKey(index))
                {
                    result.RowCountChanges++;
                    result.Details.Add($"Sheet '{sheetName}': New row {index} in edited file");
                }
            }

            // ── Columns ──
            var origCols = origWs.Descendants<Column>().ToList();
            var editCols = editWs.Descendants<Column>().ToList();

            var origColMap = new Dictionary<int, Column>();
            foreach (var c in origCols)
            {
                uint minVal = c.Min?.Value ?? 0u;
                uint maxVal = c.Max?.Value ?? minVal;
                for (uint ci = minVal; ci <= maxVal; ci++) origColMap[(int)ci] = c;
            }
            var editColMap = new Dictionary<int, Column>();
            foreach (var c in editCols)
            {
                uint minVal = c.Min?.Value ?? 0u;
                uint maxVal = c.Max?.Value ?? minVal;
                for (uint ci = minVal; ci <= maxVal; ci++) editColMap[(int)ci] = c;
            }

            foreach (var (colIdx, origCol) in origColMap)
            {
                if (!editColMap.TryGetValue(colIdx, out var editCol))
                {
                    result.ColumnCountChanges++;
                    result.Details.Add($"Sheet '{sheetName}': Column {colIdx} missing in edited");
                    continue;
                }

                bool origHidden = origCol.Hidden?.Value ?? false;
                bool editHidden = editCol.Hidden?.Value ?? false;
                if (origHidden != editHidden)
                {
                    result.HiddenColumnChanges++;
                    result.Details.Add($"Sheet '{sheetName}': Column {colIdx} hidden {origHidden} vs {editHidden}");
                }

                double origW = origCol.Width?.Value ?? 0.0;
                double editW = editCol.Width?.Value ?? 0.0;
                if (Math.Abs(origW - editW) > 0.01)
                {
                    result.ColumnWidthChanges++;
                    result.Details.Add($"Sheet '{sheetName}': Column {colIdx} width {origW:F2} vs {editW:F2}");
                }
            }

            // ── Freeze Panes ──
            var origFreeze = origWs.Descendants<Pane>().FirstOrDefault();
            var editFreeze = editWs.Descendants<Pane>().FirstOrDefault();
            if (!XmlEquals(origFreeze, editFreeze))
            {
                result.FreezePaneChanges++;
                result.Details.Add($"Sheet '{sheetName}': FreezePane differs\n  Orig: {FormatXmlSnippet(origFreeze?.OuterXml)}\n  Edit: {FormatXmlSnippet(editFreeze?.OuterXml)}");
            }

            // ── Merged Cells ──
            var origMerges = origWs.Descendants<MergeCell>().ToList();
            var editMerges = editWs.Descendants<MergeCell>().ToList();

            if (origMerges.Count != editMerges.Count)
            {
                result.MergeChanges += Math.Abs(origMerges.Count - editMerges.Count);
                result.Details.Add($"Sheet '{sheetName}': Merge count {origMerges.Count} vs {editMerges.Count}");
                // Show which merges are missing
                var origRefs = origMerges.Select(m => m.Reference?.Value ?? "").ToHashSet();
                var editRefs = editMerges.Select(m => m.Reference?.Value ?? "").ToHashSet();
                foreach (var refVal in origRefs)
                {
                    if (!editRefs.Contains(refVal))
                        result.Details.Add($"  Missing merge: {refVal}");
                }
                foreach (var refVal in editRefs)
                {
                    if (!origRefs.Contains(refVal))
                        result.Details.Add($"  Extra merge: {refVal}");
                }
            }
            else
            {
                for (int m = 0; m < origMerges.Count; m++)
                {
                    if ((origMerges[m].Reference?.Value ?? "") != (editMerges[m].Reference?.Value ?? ""))
                    {
                        result.MergeChanges++;
                        result.Details.Add($"Sheet '{sheetName}': Merge #{m} reference differs: '{origMerges[m].Reference?.Value}' vs '{editMerges[m].Reference?.Value}'");
                    }
                }
            }

            // ── Print Area ──
            string origPa = origWs.Descendants<DefinedName>().FirstOrDefault()?.Text ?? "";
            string editPa = editWs.Descendants<DefinedName>().FirstOrDefault()?.Text ?? "";
            if (!string.IsNullOrEmpty(origPa) && origPa != editPa)
            {
                result.PrintAreaChanges++;
                result.Details.Add($"Sheet '{sheetName}': PrintArea differs\n  Orig: {origPa}\n  Edit: {editPa}");
            }

            // ── PrintOptions & PageSetup (with XML snippets) ──
            var origPrintOpt = origWs.Descendants<PrintOptions>().FirstOrDefault();
            var editPrintOpt = editWs.Descendants<PrintOptions>().FirstOrDefault();
            if (!XmlEquals(origPrintOpt, editPrintOpt))
            {
                result.PageSetupChanges++;
                result.Details.Add($"Sheet '{sheetName}': PrintOptions differs\n  Orig: {FormatXmlSnippet(origPrintOpt?.OuterXml)}\n  Edit: {FormatXmlSnippet(editPrintOpt?.OuterXml)}");
            }

            var origPgSetup = origWs.Descendants<DocumentFormat.OpenXml.Spreadsheet.PageSetup>().FirstOrDefault();
            var editPgSetup = editWs.Descendants<DocumentFormat.OpenXml.Spreadsheet.PageSetup>().FirstOrDefault();
            if (!XmlEquals(origPgSetup, editPgSetup))
            {
                result.PageSetupChanges++;
                result.Details.Add($"Sheet '{sheetName}': PageSetup differs\n  Orig: {FormatXmlSnippet(origPgSetup?.OuterXml)}\n  Edit: {FormatXmlSnippet(editPgSetup?.OuterXml)}");
            }

            // ── Page Margins ──
            var origMargins = origWs.Descendants<PageMargins>().FirstOrDefault();
            var editMargins = editWs.Descendants<PageMargins>().FirstOrDefault();
            if (!XmlEquals(origMargins, editMargins))
            {
                result.PageMarginChanges++;
                result.Details.Add($"Sheet '{sheetName}': PageMargins differs\n  Orig: {FormatXmlSnippet(origMargins?.OuterXml)}\n  Edit: {FormatXmlSnippet(editMargins?.OuterXml)}");
            }

            // ── HeaderFooter ──
            var origHF = origWs.Descendants<HeaderFooter>().FirstOrDefault();
            var editHF = editWs.Descendants<HeaderFooter>().FirstOrDefault();
            if (!XmlEquals(origHF, editHF))
            {
                result.HeaderFooterChanges++;
                result.Details.Add($"Sheet '{sheetName}': HeaderFooter differs\n  Orig: {FormatXmlSnippet(origHF?.OuterXml)}\n  Edit: {FormatXmlSnippet(editHF?.OuterXml)}");
            }

            // ── Data Validations ──
            var origDv = origWs.Descendants<DataValidations>().FirstOrDefault();
            var editDv = editWs.Descendants<DataValidations>().FirstOrDefault();
            if (!XmlEquals(origDv, editDv))
            {
                result.DataValidationChanges++;
                result.Details.Add($"Sheet '{sheetName}': DataValidations differs\n  Orig: {FormatXmlSnippet(origDv?.OuterXml)}\n  Edit: {FormatXmlSnippet(editDv?.OuterXml)}");
            }

            // ── Conditional Formatting ──
            var origCond = origWs.Descendants<ConditionalFormatting>().ToList();
            var editCond = editWs.Descendants<ConditionalFormatting>().ToList();
            for (int c = 0; c < Math.Min(origCond.Count, editCond.Count); c++)
            {
                if (origCond[c].OuterXml != editCond[c].OuterXml)
                {
                    result.ConditionalFormattingChanges++;
                    result.Details.Add($"Sheet '{sheetName}': ConditionalFormatting #{c} differs\n  Orig: {FormatXmlSnippet(origCond[c].OuterXml)}\n  Edit: {FormatXmlSnippet(editCond[c].OuterXml)}");
                }
            }
            if (origCond.Count != editCond.Count)
            {
                result.ConditionalFormattingChanges += Math.Abs(origCond.Count - editCond.Count);
                result.Details.Add($"Sheet '{sheetName}': ConditionalFormatting count {origCond.Count} vs {editCond.Count}");
            }

            // ── Hyperlinks ──
            var origHl = origWs.Descendants<Hyperlink>().ToList();
            var editHl = editWs.Descendants<Hyperlink>().ToList();
            for (int h = 0; h < Math.Min(origHl.Count, editHl.Count); h++)
            {
                string oRef = origHl[h].Reference?.Value ?? "";
                string eRef = editHl[h].Reference?.Value ?? "";
                string oId = origHl[h].Id?.Value ?? "";
                string eId = editHl[h].Id?.Value ?? "";
                if (oRef != eRef || oId != eId)
                {
                    result.HyperlinkChanges++;
                    result.Details.Add($"Sheet '{sheetName}': Hyperlink #{h} ({oRef}) differs");
                }
            }
            if (origHl.Count != editHl.Count)
            {
                result.HyperlinkChanges += Math.Abs(origHl.Count - editHl.Count);
            }

            // ── Sheet Protection ──
            var origProt = origWs.Descendants<SheetProtection>().FirstOrDefault();
            var editProt = editWs.Descendants<SheetProtection>().FirstOrDefault();
            if (!XmlEquals(origProt, editProt))
            {
                result.Details.Add($"Sheet '{sheetName}': SheetProtection differs (logged, not counted)\n  Orig: {FormatXmlSnippet(origProt?.OuterXml)}\n  Edit: {FormatXmlSnippet(editProt?.OuterXml)}");
            }

            // ── Cells (detailed) ──
            var origCells = origWs.Descendants<Cell>().ToList();
            var editCells = editWs.Descendants<Cell>().ToList();

            var origCellMap = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in origCells)
            {
                var key = c.CellReference?.Value ?? "";
                if (!string.IsNullOrEmpty(key)) origCellMap[key] = c;
            }

            var editCellMap = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in editCells)
            {
                var key = c.CellReference?.Value ?? "";
                if (!string.IsNullOrEmpty(key)) editCellMap[key] = c;
            }

            foreach (var (refStr, origCell) in origCellMap)
            {
                if (!editCellMap.TryGetValue(refStr, out var editCell))
                {
                    result.MissingCellChanges++;
                    result.Details.Add($"Sheet '{sheetName}': Cell {refStr} missing in edited file");
                    continue;
                }

                uint origStyle = origCell.StyleIndex?.Value ?? 0u;
                uint editStyle = editCell.StyleIndex?.Value ?? 0u;
                if (origStyle != editStyle)
                {
                    result.StyleChanges++;
                    result.Details.Add($"Sheet '{sheetName}': Cell {refStr} StyleIndex {origStyle} vs {editStyle}");
                }

                string origFormula = origCell.CellFormula?.Text ?? "";
                string editFormula = editCell.CellFormula?.Text ?? "";
                if (origFormula != editFormula)
                {
                    result.FormulaChanges++;
                    result.Details.Add($"Sheet '{sheetName}': Cell {refStr} formula differs");
                }

                string origVal = origCell.CellValue?.Text ?? "";
                string editVal = editCell.CellValue?.Text ?? "";
                if (origVal != editVal)
                {
                    result.EditableValueChanges++;
                    result.EditableChangedCells.Add($"{sheetName}!{refStr}: '{TruncateValue(origVal)}' → '{TruncateValue(editVal)}'");
                }
            }

            foreach (var (refStr, _) in editCellMap)
            {
                if (!origCellMap.ContainsKey(refStr))
                {
                    result.NewCellChanges++;
                    result.Details.Add($"Sheet '{sheetName}': New cell {refStr} in edited file");
                }
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        // Object Parts (drawings, images, comments)
        // ═══════════════════════════════════════════════════════════════════

        private void CompareObjectParts(
            WorkbookPart origWbPart, WorkbookPart editWbPart, WorkbookDiffResult result,
            List<Sheet> origSheets, List<Sheet> editSheets)
        {
            for (int i = 0; i < Math.Min(origSheets.Count, editSheets.Count); i++)
            {
                var oSheet = origSheets[i];
                var eSheet = editSheets[i];
                if (oSheet.Id == null || eSheet.Id == null) continue;

                var oWsPart = origWbPart.GetPartById(oSheet.Id) as WorksheetPart;
                var eWsPart = editWbPart.GetPartById(eSheet.Id) as WorksheetPart;
                if (oWsPart == null || eWsPart == null) continue;

                // Drawings
                var oDrawPart = oWsPart.DrawingsPart;
                var eDrawPart = eWsPart.DrawingsPart;

                if ((oDrawPart != null) != (eDrawPart != null))
                {
                    result.DrawingChanges++;
                    result.Details.Add($"Sheet '{oSheet.Name}': DrawingsPart presence differs");
                }
                else if (oDrawPart != null && eDrawPart != null)
                {
                    var oDrawXml = oDrawPart.WorksheetDrawing?.OuterXml ?? "";
                    var eDrawXml = eDrawPart.WorksheetDrawing?.OuterXml ?? "";
                    if (oDrawXml != eDrawXml)
                    {
                        result.DrawingChanges++;
                        result.Details.Add($"Sheet '{oSheet.Name}': Drawings XML differs\n  Orig: {FormatXmlSnippet(oDrawXml)}\n  Edit: {FormatXmlSnippet(eDrawXml)}");
                    }

                    var oImageParts = oDrawPart.ImageParts.Count();
                    var eImageParts = eDrawPart.ImageParts.Count();
                    if (oImageParts != eImageParts)
                    {
                        result.ImageChanges++;
                        result.Details.Add($"Sheet '{oSheet.Name}': Image count {oImageParts} vs {eImageParts}");
                    }
                }

                // Comments
                var oCommentPart = oWsPart.WorksheetCommentsPart;
                var eCommentPart = eWsPart.WorksheetCommentsPart;
                if ((oCommentPart != null) != (eCommentPart != null))
                {
                    result.CommentChanges++;
                    result.Details.Add($"Sheet '{oSheet.Name}': CommentsPart presence differs");
                }
                else if (oCommentPart != null && eCommentPart != null)
                {
                    var oComments = oCommentPart.Comments?.Descendants<Comment>().ToList() ?? new();
                    var eComments = eCommentPart.Comments?.Descendants<Comment>().ToList() ?? new();
                    if (oComments.Count != eComments.Count)
                    {
                        result.CommentChanges++;
                        result.Details.Add($"Sheet '{oSheet.Name}': Comment count {oComments.Count} vs {eComments.Count}");
                    }
                    else
                    {
                        for (int c = 0; c < oComments.Count; c++)
                        {
                            var oRef = oComments[c].Reference?.Value ?? "";
                            var eRef = eComments[c].Reference?.Value ?? "";
                            if (oRef != eRef)
                            {
                                result.CommentChanges++;
                                result.Details.Add($"Sheet '{oSheet.Name}': Comment #{c} reference differs");
                            }
                        }
                    }
                }
            }

            // Tables
            var origTableParts = origWbPart.Parts.Where(p =>
                p.OpenXmlPart.Uri.ToString().IndexOf("table", StringComparison.OrdinalIgnoreCase) >= 0
                || p.OpenXmlPart.Uri.ToString().IndexOf("pivot", StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            var editTableParts = editWbPart.Parts.Where(p =>
                p.OpenXmlPart.Uri.ToString().IndexOf("table", StringComparison.OrdinalIgnoreCase) >= 0
                || p.OpenXmlPart.Uri.ToString().IndexOf("pivot", StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            if (origTableParts.Count != editTableParts.Count)
            {
                result.TableChanges++;
                result.Details.Add($"Table/PivotTable count: original={origTableParts.Count}, edited={editTableParts.Count}");
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        // Additional Parts (Phase 5.0)
        // ═══════════════════════════════════════════════════════════════════

        private void CompareAdditionalParts(
            SpreadsheetDocument origDoc, SpreadsheetDocument editDoc,
            WorkbookPart origWbPart, WorkbookPart editWbPart, WorkbookDiffResult result)
        {
            // SharedStringTable
            var origSstPart = origWbPart.SharedStringTablePart;
            var editSstPart = editWbPart.SharedStringTablePart;
            if ((origSstPart != null) != (editSstPart != null))
            {
                result.SharedStringsChanges++;
                result.Details.Add("SharedStringTablePart presence differs");
            }
            else if (origSstPart != null && editSstPart != null)
            {
                var origSst = origSstPart.SharedStringTable;
                var editSst = editSstPart.SharedStringTable;
                int origCount = origSst?.Count() ?? 0;
                int editCount = editSst?.Count() ?? 0;
                if (editCount < origCount)
                {
                    result.SharedStringsChanges++;
                    result.Details.Add($"SharedStringTable shrank: {origCount} -> {editCount}");
                }
                for (int i = 0; i < Math.Min(origCount, editCount); i++)
                {
                    var oItem = origSst?.ElementAt(i) as SharedStringItem;
                    var eItem = editSst?.ElementAt(i) as SharedStringItem;
                    if (oItem?.OuterXml != eItem?.OuterXml)
                    {
                        result.SharedStringsChanges++;
                        result.Details.Add($"SharedStringTable entry #{i} changed");
                        break;
                    }
                }
            }

            // Workbook.xml root element — ELEMENT-BY-ELEMENT XPath Comparison
            // Phase 5.2.4: Instead of comparing the entire XML blob, we compare
            // each child element individually and log the EXACT XPath of any difference.
            var origWbXml = origWbPart.Workbook;
            var editWbXml = editWbPart.Workbook;
            var origWbClone = origWbXml?.CloneNode(true) as Workbook;
            var editWbClone = editWbXml?.CloneNode(true) as Workbook;
            if (origWbClone != null && editWbClone != null)
            {
                origWbClone.RemoveAllChildren<Sheets>();
                editWbClone.RemoveAllChildren<Sheets>();
                origWbClone.RemoveAllChildren<DefinedNames>();
                editWbClone.RemoveAllChildren<DefinedNames>();
                origWbClone.RemoveAllChildren<WorkbookProtection>();
                editWbClone.RemoveAllChildren<WorkbookProtection>();

                // Phase 5.2.4: Element-by-element comparison with XPath logging
                var origChildren = origWbClone.ChildElements.ToList();
                var editChildren = editWbClone.ChildElements.ToList();

                bool anyWbDiff = false;
                int maxChildren = Math.Max(origChildren.Count, editChildren.Count);
                for (int ci = 0; ci < maxChildren; ci++)
                {
                    string xpath = "/workbook/" + (ci < origChildren.Count ? origChildren[ci].LocalName
                        : ci < editChildren.Count ? editChildren[ci].LocalName : $"child[{ci}]");
                    string origOuter = ci < origChildren.Count ? origChildren[ci].OuterXml : "(missing)";
                    string editOuter = ci < editChildren.Count ? editChildren[ci].OuterXml : "(missing)";

                    if (origOuter != editOuter)
                    {
                        anyWbDiff = true;
                        result.Details.Add($"Workbook.xml difference:\n" +
                            $"  XPath: {xpath}\n" +
                            $"  Original: {origOuter}\n" +
                            $"  Edited:   {editOuter}");
                    }
                }

                if (anyWbDiff)
                    result.WorkbookXmlChanges++;
            }

            // External Links
            var origExtParts = origWbPart.ExternalRelationships;
            var editExtParts = editWbPart.ExternalRelationships;
            int origExtCount = origExtParts?.Count() ?? 0;
            int editExtCount = editExtParts?.Count() ?? 0;
            if (origExtCount != editExtCount)
            {
                result.ExternalLinksChanges++;
                result.Details.Add($"External links: original={origExtCount}, edited={editExtCount}");
            }

            // Custom XML parts
            int origCustomParts = origWbPart.Parts.Count(p =>
                p.OpenXmlPart.Uri.ToString().IndexOf("customXml", StringComparison.OrdinalIgnoreCase) >= 0);
            int editCustomParts = editWbPart.Parts.Count(p =>
                p.OpenXmlPart.Uri.ToString().IndexOf("customXml", StringComparison.OrdinalIgnoreCase) >= 0);
            if (origCustomParts != editCustomParts)
            {
                result.CustomXmlChanges++;
                result.Details.Add($"Custom XML parts: original={origCustomParts}, edited={editCustomParts}");
            }

            // Printer settings per worksheet
            var checkSheets = origWbPart.Workbook.Descendants<Sheet>().ToList();
            var checkEditSheets = editWbPart.Workbook.Descendants<Sheet>().ToList();
            for (int i = 0; i < Math.Min(checkSheets.Count, checkEditSheets.Count); i++)
            {
                if (checkSheets[i].Id == null || checkEditSheets[i].Id == null) continue;
                var oWsPart = origWbPart.GetPartById(checkSheets[i].Id!) as WorksheetPart;
                var eWsPart = editWbPart.GetPartById(checkEditSheets[i].Id!) as WorksheetPart;
                if (oWsPart == null || eWsPart == null) continue;

                bool oHasPrinter = oWsPart.Worksheet?.Descendants<PageSetup>().Any(p => p.UsePrinterDefaults?.Value == false) == true;
                bool eHasPrinter = eWsPart.Worksheet?.Descendants<PageSetup>().Any(p => p.UsePrinterDefaults?.Value == false) == true;
                if (oHasPrinter != eHasPrinter)
                {
                    result.PrinterSettingsChanges++;
                    result.Details.Add($"Sheet '{checkSheets[i].Name}': Printer settings differs");
                }
            }

            // Relationships (Phase 5.5.2: account for known additional sheet parts)
            var origRelCount = origWbPart.Parts.Count();
            var editRelCount = editWbPart.Parts.Count();
            int knownExtraParts = checkEditSheets.Count(s => KnownAdditionalSheets.Contains(s.Name?.Value ?? ""))
                - checkSheets.Count(s => KnownAdditionalSheets.Contains(s.Name?.Value ?? ""));
            if (origRelCount + knownExtraParts != editRelCount)
            {
                result.RelationshipsChanges++;
                result.Details.Add($"Relationship count: original={origRelCount}, edited={editRelCount} (expected extra={knownExtraParts})");
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        // Defined Names
        // ═══════════════════════════════════════════════════════════════════

        private void CompareDefinedNames(
            WorkbookPart origWbPart, WorkbookPart editWbPart, WorkbookDiffResult result)
        {
            var origNames = origWbPart.Workbook.Descendants<DefinedName>().ToList();
            var editNames = editWbPart.Workbook.Descendants<DefinedName>().ToList();

            var origNameMap = new Dictionary<string, DefinedName>(StringComparer.OrdinalIgnoreCase);
            foreach (var n in origNames)
            {
                var key = n.Name?.Value ?? "";
                if (!string.IsNullOrEmpty(key)) origNameMap[key] = n;
            }

            var editNameMap = new Dictionary<string, DefinedName>(StringComparer.OrdinalIgnoreCase);
            foreach (var n in editNames)
            {
                var key = n.Name?.Value ?? "";
                if (!string.IsNullOrEmpty(key)) editNameMap[key] = n;
            }

            foreach (var (name, origN) in origNameMap)
            {
                if (!editNameMap.TryGetValue(name, out var editN))
                {
                    result.DefinedNameChanges++;
                    result.Details.Add($"DefinedName '{name}' missing in edited");
                    continue;
                }
                if ((origN.Text ?? "") != (editN.Text ?? ""))
                {
                    result.DefinedNameChanges++;
                    result.Details.Add($"DefinedName '{name}': '{origN.Text}' vs '{editN.Text}'");
                }
            }

            foreach (var (name, _) in editNameMap)
            {
                if (!origNameMap.ContainsKey(name))
                {
                    result.DefinedNameChanges++;
                    result.Details.Add($"New DefinedName '{name}' in edited file");
                }
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        // Helpers
        // ═══════════════════════════════════════════════════════════════════

        private static bool XmlEquals(OpenXmlElement? a, OpenXmlElement? b)
        {
            if (a == null && b == null) return true;
            if (a == null || b == null) return false;
            return a.OuterXml == b.OuterXml;
        }

        private static string ComputeSha256(byte[] data)
        {
            using var sha = SHA256.Create();
            var hash = sha.ComputeHash(data);
            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }

        private static string TruncateValue(string val, int maxLen = 60)
        {
            if (string.IsNullOrEmpty(val)) return "(empty)";
            return val.Length <= maxLen ? val : val[..maxLen] + "...";
        }

        private static string FormatXmlSnippet(string? xml, int maxLen = 400)
        {
            if (string.IsNullOrEmpty(xml)) return "(null)";
            return xml.Length <= maxLen ? xml : xml[..maxLen] + "...";
        }
    }
}
