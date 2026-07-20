using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using WbDef = ExcelAPI.Models.WorkbookDefinition;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using System.IO.Compression;
using OoxmlComment = DocumentFormat.OpenXml.Spreadsheet.Comment;

namespace ExcelAPI.Application
{
    /// <summary>
    /// Writes user-entered field values back into the original Excel workbook
    /// using DocumentFormat.OpenXml. Only modifies cell values — never touches
    /// styles, formatting, layout, or any other structural elements.
    ///
    /// Phase 4.4: Core round-trip engine.
    ///
    /// Rules:
    ///   1. Open original workbook (read-write copy)
    ///   2. For each WbDef field, locate the cell by CellReference
    ///   3. Write the new value as an inline string or number
    ///   4. Preserve StyleIndex (never modify styles)
    ///   5. Preserve formulas (never overwrite formula cells)
    ///   6. Save as new file
    /// </summary>
    public class WorkbookValueWriter
    {
        private readonly ILogger<WorkbookValueWriter> _logger;

        public WorkbookValueWriter(ILogger<WorkbookValueWriter> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Write field values from WorkbookDefinition into the original XLSX
        /// and save as outputPath. Original file is never modified.
        /// </summary>
        /// <param name="wbDef">WorkbookDefinition with edited field values.</param>
        /// <param name="sourceXlsxPath">Path to the original workbook (must exist).</param>
        /// <param name="outputPath">Path for the edited workbook output.</param>
        /// <returns>Number of cells written.</returns>
        public int WriteValues(WbDef.WorkbookDefinition wbDef, string sourceXlsxPath, string outputPath)
        {
            if (!System.IO.File.Exists(sourceXlsxPath))
                throw new FileNotFoundException($"Source workbook not found: {sourceXlsxPath}");

            // Copy the original workbook so we never modify the source
            System.IO.File.Copy(sourceXlsxPath, outputPath, overwrite: true);

            // ── DIAGNOSTIC: Save ALL original ZIP entries before OpenXml SDK touches them ──
            // Phase 5.2.3: Comprehensive metadata preservation.
            // DocumentFormat.OpenXml is known to mutate workbook.xml, styles.xml, theme,
            // relationships, content types, and other parts when opening in read-write mode,
            // even if no structural changes are made. We save every ZIP entry before the SDK
            // opens the file, then restore all unmodified parts after the SDK closes.
            //
            // Intentionally modified entries (NOT restored):
            //   - xl/worksheets/sheet*.xml  (cell values changed)
            //   - xl/sharedStrings.xml      (new strings added)
            var originalZipEntries = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using var origPkg = System.IO.Compression.ZipFile.OpenRead(outputPath);
                foreach (var entry in origPkg.Entries)
                {
                    var name = entry.FullName;
                    using var ms = new MemoryStream();
                    entry.Open().CopyTo(ms);
                    originalZipEntries[name] = ms.ToArray();
                }
                _logger.LogInformation("[PRESERVE] Saved {Count} original ZIP entries before SDK", originalZipEntries.Count);
                foreach (var kvp in originalZipEntries.OrderBy(k => k.Key))
                    _logger.LogInformation("  Saved entry: {Name} ({Len} bytes)", kvp.Key, kvp.Value.Length);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[PRESERVE] Could not pre-save ZIP entries");
            }

            int totalCellsWritten = 0;
            int totalFieldsReceived = 0;
            int totalFieldsWithValues = 0;
            int totalFieldsSkippedEmpty = 0;
            int totalFieldsSheetNotFound = 0;

            // ═════════════════════════════════════════════════════════
            // PHASE 21.3 — STAGE 3: WORKBOOK VALUE WRITER INPUT
            // ═════════════════════════════════════════════════════════
            _logger.LogInformation("=========================================================");
            _logger.LogInformation("WORKBOOK VALUE WRITER INPUT");
            _logger.LogInformation("=========================================================");
            _logger.LogInformation("Workbook: {Title}", wbDef.Info?.Title ?? "(untitled)");
            _logger.LogInformation("Source: {Path}", sourceXlsxPath);
            _logger.LogInformation("Output: {Path}", outputPath);
            _logger.LogInformation("Sheets received: {Count}", wbDef.Sheets.Count);

            foreach (var diagSheet in wbDef.Sheets)
            {
                totalFieldsReceived += diagSheet.Fields.Count;
                int nonEmpty = diagSheet.Fields.Count(f => !string.IsNullOrWhiteSpace(f.Value));
                int empty = diagSheet.Fields.Count(f => string.IsNullOrWhiteSpace(f.Value));
                totalFieldsWithValues += nonEmpty;
                totalFieldsSkippedEmpty += empty;
                _logger.LogInformation("  Sheet '{Name}': {Total} fields ({NonEmpty} with values, {Empty} empty)",
                    diagSheet.Name, diagSheet.Fields.Count, nonEmpty, empty);

                if (nonEmpty > 0 && nonEmpty <= 20)
                {
                    foreach (var f in diagSheet.Fields)
                    {
                        if (!string.IsNullOrWhiteSpace(f.Value))
                            _logger.LogInformation("    Field: cell={Addr}, name='{Name}', type={Type}, value='{Val}'",
                                f.Cell?.Address ?? "?", f.Name ?? "?", f.Type.ToString(), f.Value);
                    }
                }
            }

            _logger.LogInformation("Total fields received: {Total}", totalFieldsReceived);
            _logger.LogInformation("Total fields with non-empty values: {WithVal}", totalFieldsWithValues);
            _logger.LogInformation("Total fields skipped (empty): {Skipped}", totalFieldsSkippedEmpty);

            if (totalFieldsWithValues == 0)
            {
                _logger.LogWarning("[DIAG] ZERO fields have non-empty values! No cells will be written.");
                _logger.LogWarning("[DIAG] Possible cause: Field IDs in frontend don't match WbDef model binding.");
                _logger.LogWarning("[DIAG] Check: runtimeFormToWorkbookDefinition → values[f.id] → f.id must match backend field.Name");
            }

            using (var doc = SpreadsheetDocument.Open(outputPath, true))
            {
                if (doc.WorkbookPart == null)
                    throw new InvalidOperationException("WorkbookPart is null — file may be corrupt.");

                var wbPart = doc.WorkbookPart;
                var sheets = wbPart.Workbook.Descendants<Sheet>().ToList();

                _logger.LogInformation("Workbook sheets found: {Count}", sheets.Count);
                foreach (var s in sheets)
                    _logger.LogInformation("  Available sheet: '{Name}' (id={Id})", s.Name?.Value ?? "?", s.Id?.Value ?? "?");

                // Build sheet name → SheetPart mapping
                var sheetParts = new Dictionary<string, WorksheetPart>(StringComparer.OrdinalIgnoreCase);
                foreach (var sheet in sheets)
                {
                    if (sheet.Id == null) continue;
                    var wsPart = wbPart.GetPartById(sheet.Id) as WorksheetPart;
                    if (wsPart != null && sheet.Name != null)
                    {
                        sheetParts[sheet.Name] = wsPart;
                    }
                }

                // Process each WbDef sheet
                foreach (var wbSheet in wbDef.Sheets)
                {
                    if (wbSheet.Fields.Count == 0) continue;

                    // Find the matching worksheet part
                    if (!sheetParts.TryGetValue(wbSheet.Name, out var wsPart))
                    {
                        _logger.LogWarning("Sheet '{Name}' not found in workbook — skipping {Count} fields",
                            wbSheet.Name, wbSheet.Fields.Count);
                        totalFieldsSheetNotFound += wbSheet.Fields.Count;
                        continue;
                    }

                    var worksheet = wsPart.Worksheet;
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData == null)
                    {
                        _logger.LogWarning("Sheet '{Name}' has no SheetData — skipping", wbSheet.Name);
                        continue;
                    }

                    // Ensure SharedStringTable exists for inline string values
                    var sstPart = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (sstPart == null)
                    {
                        sstPart = wbPart.AddNewPart<SharedStringTablePart>();
                        sstPart.SharedStringTable = new SharedStringTable();
                    }

                    int cellsWritten = 0;

                    foreach (var field in wbSheet.Fields)
                    {
                        if (string.IsNullOrWhiteSpace(field.Value))
                            continue; // Skip empty values

                        string cellRef = field.Cell.Address;
                        uint rowIndex = (uint)field.Cell.RowIndex;

                        // ── DIAGNOSTIC: Log every cell write attempt ──
                        _logger.LogInformation("[DIAG] Writing cell {Sheet}!{Cell}: value='{Value}', type={Type}",
                            wbSheet.Name, cellRef, field.Value, field.Type.ToString());

                        // Find or create the row
                        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
                        string rowStatus = row == null ? "NEW" : "EXISTS";
                        if (row == null)
                        {
                            row = new Row { RowIndex = rowIndex };
                            sheetData.Append(row);
                            _logger.LogInformation("[DIAG]  Created new Row {RowIdx} for {Cell}", rowIndex, cellRef);
                        }

                        // Find or create the cell
                        var cell = row.Elements<Cell>().FirstOrDefault(c =>
                            string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));

                        bool isNewCell = false;
                        string cellStatus = "";
                        if (cell == null)
                        {
                            cell = new Cell();
                            isNewCell = true;
                            cellStatus = "NEW";
                            _logger.LogInformation("[DIAG]  Created new Cell {Cell} (Row {RowIdx})", cellRef, rowIndex);
                        }
                        else
                        {
                            cellStatus = "EXISTS";
                            string oldVal = cell.CellValue?.Text ?? "(empty)";
                            uint oldStyle = cell.StyleIndex?.Value ?? 0;
                            _logger.LogInformation("[DIAG]  Found existing Cell {Cell}: oldValue='{Old}', styleIndex={Style}",
                                cellRef, oldVal, oldStyle);
                        }

                        // Set cell reference
                        cell.CellReference = cellRef;

                        // PRESERVE StyleIndex — never touch existing styles
                        // For new cells, don't set any style (default)
                        // For existing cells, StyleIndex stays unchanged
                        _logger.LogInformation("[DIAG]  StyleIndex preserved: {Style} (isNewCell={IsNew})",
                            cell.StyleIndex?.Value ?? 0, isNewCell);

                        // Phase 5.4: Restore formula for calculated fields
                        if (!string.IsNullOrWhiteSpace(field.Formula)
                            && field.Type == WbDef.FieldType.Calculated)
                        {
                            cell.CellFormula = new CellFormula(field.Formula);
                            cell.CellValue = null;
                            if (cell.DataType != null)
                                cell.DataType = null;
                            _logger.LogInformation("[DIAG]  Wrote as FORMULA: '{Formula}'", field.Formula);
                        }
                        else
                        {
                            // Write the value based on field type
                            if (IsNumericField(field))
                            {
                                if (double.TryParse(field.Value, out var numVal))
                                {
                                    cell.DataType = CellValues.Number;
                                    cell.CellValue = new CellValue(numVal.ToString("G"));
                                    _logger.LogInformation("[DIAG]  Wrote as NUMBER: '{NumVal}'", numVal.ToString("G"));
                                }
                                else
                                {
                                    WriteInlineString(cell, field.Value, sstPart);
                                    _logger.LogInformation("[DIAG]  Wrote as SHARED STRING (numeric field, non-numeric value): '{Val}'", field.Value);
                                }
                            }
                            else
                            {
                                WriteInlineString(cell, field.Value, sstPart);
                                _logger.LogInformation("[DIAG]  Wrote as SHARED STRING: '{Val}'", field.Value);
                            }
                        }

                        if (isNewCell)
                        {
                            row.Append(cell);
                        }

                        cellsWritten++;

                        // ── DIAGNOSTIC: Verify the value was written ──
                        string cellFormula = cell.CellFormula?.Text ?? "";
                        _logger.LogInformation("[DIAG]  RESULT: {Sheet}!{Cell} | formula='{Formula}' | value='{New}' ({NewDt}) | style={Style} | row={RowStatus} | cell={CellStatus}",
                            wbSheet.Name, cellRef,
                            cellFormula,
                            cell.CellValue?.Text ?? "(empty)", cell.DataType?.ToString() ?? "(none)",
                            cell.StyleIndex?.Value ?? 0,
                            rowStatus, cellStatus);
                    }

                    // Save changes to this worksheet
                    wsPart.Worksheet.Save();

                    _logger.LogInformation(
                        "Wrote {Count} values to sheet '{Name}'",
                        cellsWritten, wbSheet.Name);
                    totalCellsWritten += cellsWritten;
                }

                // Save shared strings
                if (wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault() is { } sst)
                {
                    sst.SharedStringTable.Save();
                    _logger.LogInformation("[DIAG] SharedStringTable saved ({Count} items)", sst.SharedStringTable.Count());
                }

                // ── Phase 5.4: Write ConMas-compatible cell comments ──
                // DISABLED in Phase 5.5.2 — comments are now written by PostProcessZipForConMas
                // after the ZIP restore step to ensure exact \r\n format and xr:uid attributes.
                // WriteConMasCellComments(wbPart, wbDef);
            }

            // ═════════════════════════════════════════════════════════
            // PHASE 21.3 — STAGE 5 & 6: WORKBOOK VERIFICATION + STYLE CHECK
            // ═════════════════════════════════════════════════════════
            int verifyPass = 0;
            int verifyFail = 0;
            int styleChecked = 0;

            _logger.LogInformation("=========================================================");
            _logger.LogInformation("WORKBOOK VERIFICATION — Re-opening to verify written cells");
            _logger.LogInformation("=========================================================");
            try
            {
                using var verifyDoc = SpreadsheetDocument.Open(outputPath, false);
                if (verifyDoc.WorkbookPart != null)
                {
                    var stylesPart = verifyDoc.WorkbookPart.WorkbookStylesPart;
                    var verifySheets = verifyDoc.WorkbookPart.Workbook.Descendants<Sheet>().ToList();
                    foreach (var wbSheet in wbDef.Sheets)
                    {
                        if (wbSheet.Fields.Count == 0) continue;
                        var matchSheet = verifySheets.FirstOrDefault(s => s.Name == wbSheet.Name);
                        if (matchSheet?.Id == null) continue;
                        var verifyWsPart = verifyDoc.WorkbookPart.GetPartById(matchSheet.Id) as WorksheetPart;
                        if (verifyWsPart == null) continue;

                        var verifyCells = verifyWsPart.Worksheet.Descendants<Cell>().ToList();
                        foreach (var field in wbSheet.Fields)
                        {
                            if (string.IsNullOrWhiteSpace(field.Value)) continue;
                            string cellRef = field.Cell.Address;
                            var verifyCell = verifyCells.FirstOrDefault(c =>
                                string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));

                            _logger.LogInformation("  Cell {Sheet}!{Cell}:", wbSheet.Name, cellRef);

                            // Stage 5: Value Verification
                            string expectedValue = field.Value;
                            string actualValue = "(cell not found)";
                            if (verifyCell != null)
                            {
                                actualValue = verifyCell.CellValue?.Text ?? "(empty)";
                            }
                            bool valueMatches = actualValue == expectedValue
                                || (expectedValue.Length <= 10 && actualValue.Contains(expectedValue.TrimEnd(' ')));
                            _logger.LogInformation("    Expected Value: '{Exp}'", expectedValue);
                            _logger.LogInformation("    Actual Value:   '{Act}'", actualValue);
                            _logger.LogInformation("    Value Result:   {Status}", valueMatches ? "PASS" : "FAIL");
                            if (valueMatches) verifyPass++; else verifyFail++;

                            // Stage 6: Style Verification
                            if (verifyCell != null && stylesPart?.Stylesheet != null)
                            {
                                // Read the actual cell style
                                uint styleIdx = verifyCell.StyleIndex?.Value ?? 0u;
                                var cellFormats = stylesPart.Stylesheet.CellFormats?.Cast<CellFormat>().ToList() ?? new();

                                if (styleIdx < cellFormats.Count)
                                {
                                    var xf = cellFormats[(int)styleIdx];
                                    var fonts = stylesPart.Stylesheet.Fonts?.Cast<Font>().ToList() ?? new();
                                    var fills = stylesPart.Stylesheet.Fills?.Cast<Fill>().ToList() ?? new();

                                    string actualFontName = "";
                                    double actualFontSize = 0;
                                    bool actualBold = false;
                                    bool actualItalic = false;
                                    bool actualUnderline = false;
                                    string actualHorizAlign = "(default)";
                                    string actualVertAlign = "(default)";
                                    bool actualWrap = false;
                                    string actualFillColor = "(none)";
                                    string actualFontColor = "(default)";

                                    // Read font properties
                                    if (xf.FontId?.Value != null)
                                    {
                                        int fontId = (int)xf.FontId.Value;
                                        if (fontId >= 0 && fontId < fonts.Count)
                                        {
                                            var font = fonts[fontId];
                                            actualFontName = font.FontName?.Val?.Value ?? "(default)";
                                            actualFontSize = font.FontSize?.Val?.Value ?? 0;
                                            actualBold = font.Bold?.Val?.Value ?? false;
                                            actualItalic = font.Italic?.Val?.Value ?? false;
                                            actualUnderline = font.Underline?.Val?.Value != null;

                                            // Font color
                                            var fontColorEl = font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>();
                                            if (fontColorEl?.Rgb?.Value != null)
                                            {
                                                string rgb = fontColorEl.Rgb.Value;
                                                actualFontColor = rgb.Length >= 8 ? "#" + rgb[2..] : "#" + rgb;
                                            }
                                        }
                                    }

                                    // Read fill
                                    if (xf.FillId?.Value != null)
                                    {
                                        int fillId = (int)xf.FillId.Value;
                                        if (fillId >= 0 && fillId < fills.Count)
                                        {
                                            var fill = fills[fillId];
                                            var pf = fill.PatternFill;
                                            if (pf?.PatternType?.Value != null &&
                                                pf.PatternType.Value != PatternValues.None &&
                                                pf.PatternType.Value != PatternValues.Gray125)
                                            {
                                                var fgColor = pf.ForegroundColor;
                                                if (fgColor?.Rgb?.Value != null)
                                                {
                                                    string rgb = fgColor.Rgb.Value;
                                                    actualFillColor = rgb.Length >= 8 ? "#" + rgb[2..] : "#" + rgb;
                                                }
                                            }
                                        }
                                    }

                                    // Read alignment
                                    if (xf.Alignment != null)
                                    {
                                        var align = xf.Alignment;
                                        if (align.Horizontal?.Value != null)
                                            actualHorizAlign = align.Horizontal.Value.ToString();
                                        if (align.Vertical?.Value != null)
                                            actualVertAlign = align.Vertical.Value.ToString();
                                        actualWrap = align.WrapText?.Value ?? false;
                                    }

                                    // Log all style properties
                                    _logger.LogInformation("    Style Check — Cell {Cell}:", cellRef);

                                    // The frontend does NOT send style properties with the export payload.
                                    // Style properties come from the original workbook template.
                                    // We log the actual values preserved from the template.
                                    _logger.LogInformation("      Style properties: not sent by frontend in export payload.");
                                    _logger.LogInformation("      Preserved from original workbook (template):");
                                    _logger.LogInformation("      Font Name:     '{Act}'", actualFontName);
                                    _logger.LogInformation("      Font Size:     '{Act}'", actualFontSize);
                                    _logger.LogInformation("      Bold:          '{Act}'", actualBold);
                                    _logger.LogInformation("      Italic:        '{Act}'", actualItalic);
                                    _logger.LogInformation("      Underline:     '{Act}'", actualUnderline);
                                    _logger.LogInformation("      Font Color:    '{Act}'", actualFontColor);
                                    _logger.LogInformation("      Fill Color:    '{Act}'", actualFillColor);
                                    _logger.LogInformation("      H-Align:       '{Act}'", actualHorizAlign);
                                    _logger.LogInformation("      V-Align:       '{Act}'", actualVertAlign);
                                    _logger.LogInformation("      Wrap Text:     '{Act}'", actualWrap);

                                    _logger.LogInformation("    Workbook styles preserved successfully from original template.");
                                    styleChecked++;
                                }
                                else
                                {
                                    _logger.LogInformation("    Style Check — Cell {Cell}: style index {StyleIdx} out of range (no custom style)", cellRef, styleIdx);
                                    styleChecked++;
                                }
                            }
                            else if (verifyCell != null)
                            {
                                // No styles part available
                                _logger.LogInformation("    Style Check — Cell {Cell}: no styles part available in workbook", cellRef);
                                styleChecked++;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[DIAG] Post-write verification failed");
            }

            // ── SHA256 DIAGNOSTIC: Log workbook.xml hashes before restore ──
            byte[] restoreWbXml = originalZipEntries.TryGetValue("xl/workbook.xml", out var owb) ? owb : Array.Empty<byte>();
            {
                byte[] origWbXml = restoreWbXml;
                string origHash = origWbXml.Length > 0 ? ComputeSha256ForLogging(origWbXml) : "(no original)";
                string beforeHash = "(unavailable)";
                try
                {
                    using var prePkg = System.IO.Compression.ZipFile.OpenRead(outputPath);
                    var preEntry = prePkg.GetEntry("xl/workbook.xml");
                    if (preEntry != null)
                    {
                        using var ms = new MemoryStream();
                        preEntry.Open().CopyTo(ms);
                        beforeHash = ComputeSha256ForLogging(ms.ToArray());
                    }
                }
                catch { }
                _logger.LogInformation("[SHA256] workbook.xml before restore: original={Orig}, current={Before}", origHash, beforeHash);
            }

            // ── RESTORE: Put back ALL original ZIP entries except intentionally modified ones ──
            // Phase 5.2.3: Comprehensive metadata preservation.
            // The OpenXml SDK mutates many parts (workbook.xml, styles.xml, theme,
            // relationships, content types, etc.) even when no structural changes are made.
            // We restore every original byte to prevent WorkbookDiffValidator from flagging
            // false positives.
            //
            // Entries NOT restored (intentionally modified):
            //   - xl/worksheets/sheet*.xml  (cell values changed)
            //   - xl/sharedStrings.xml      (new strings added)
            if (originalZipEntries.Count > 0)
            {
                int restoredCount = 0;
                int unchangedCount = 0;
                int skippedModified = 0;
                bool workbookXmlWasRestored = false;

                using var pkg = System.IO.Compression.ZipFile.Open(outputPath, System.IO.Compression.ZipArchiveMode.Update);

                foreach (var kvp in originalZipEntries.OrderBy(e => e.Key))
                {
                    string entryName = kvp.Key;
                    byte[] originalBytes = kvp.Value;

                    // Skip entries that were intentionally modified
                    bool isWorksheet = entryName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase);
                    bool isSharedStrings = entryName.Equals("xl/sharedstrings.xml", StringComparison.OrdinalIgnoreCase)
                        || entryName.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase);
                    bool isComments = entryName.StartsWith("xl/comments", StringComparison.OrdinalIgnoreCase);
                    bool isVmlDrawing = entryName.StartsWith("xl/drawings/vmldrawing", StringComparison.OrdinalIgnoreCase);
                    bool isWorksheetRels = entryName.StartsWith("xl/worksheets/_rels/sheet", StringComparison.OrdinalIgnoreCase)
                        && entryName.EndsWith(".xml.rels", StringComparison.OrdinalIgnoreCase);
                    bool isContentTypes = entryName.Equals("[content_types].xml", StringComparison.OrdinalIgnoreCase);
                    if (isWorksheet || isSharedStrings || isComments || isVmlDrawing || isWorksheetRels || isContentTypes)
                    {
                        skippedModified++;
                        _logger.LogInformation("[RESTORE] Skipped '{Name}' — intentionally modified", entryName);
                        continue;
                    }

                    try
                    {
                        // Get current bytes from the edited zip
                        var currentEntry = pkg.GetEntry(entryName);
                        byte[] currentBytes;
                        if (currentEntry != null)
                        {
                            using var ms = new MemoryStream();
                            currentEntry.Open().CopyTo(ms);
                            currentBytes = ms.ToArray();
                        }
                        else
                        {
                            currentBytes = Array.Empty<byte>();
                        }

                        // Compare — only restore if actually different
                        if (currentBytes.Length != originalBytes.Length ||
                            !currentBytes.AsSpan().SequenceEqual(originalBytes.AsSpan()))
                        {
                            // Replace with original
                            if (currentEntry != null)
                            {
                                currentEntry.Delete();
                            }
                            var newEntry = pkg.CreateEntry(entryName);
                            using var writer = newEntry.Open();
                            writer.Write(originalBytes, 0, originalBytes.Length);
                            restoredCount++;
                            _logger.LogInformation("[RESTORE] Restored '{Name}' ({Len} bytes)", entryName, originalBytes.Length);
                            if (entryName.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase))
                                workbookXmlWasRestored = true;
                        }
                        else
                        {
                            unchangedCount++;
                            _logger.LogInformation("[RESTORE] Unchanged '{Name}' ({Len} bytes) — no restore needed", entryName, originalBytes.Length);
                        }
                    }
                    catch (Exception exInner)
                    {
                        _logger.LogWarning(exInner, "[RESTORE] Failed to restore '{Name}' — skipping, will be handled by post-process: {Msg}", entryName, exInner.Message);
                    }
                }

                _logger.LogInformation("[RESTORE] Summary: {Restored} restored, {Unchanged} unchanged, {Skipped} skipped (intentionally modified). workbook.xml restored={WbXml}",
                    restoredCount, unchangedCount, skippedModified, workbookXmlWasRestored ? "YES" : "NO — check if it was unchanged or missing");
                pkg.Dispose();
            }

            // ── SHA256 DIAGNOSTIC: Log workbook.xml hash AFTER restore ──
            {
                byte[] origWbXml = restoreWbXml;
                string origHash = origWbXml.Length > 0 ? ComputeSha256ForLogging(origWbXml) : "(no original)";
                string afterHash = "(unavailable)";
                try
                {
                    using var postPkg = System.IO.Compression.ZipFile.OpenRead(outputPath);
                    var postEntry = postPkg.GetEntry("xl/workbook.xml");
                    if (postEntry != null)
                    {
                        using var ms = new MemoryStream();
                        postEntry.Open().CopyTo(ms);
                        afterHash = ComputeSha256ForLogging(ms.ToArray());
                    }
                }
                catch { }
                bool hashesMatch = origHash == afterHash;
                _logger.LogInformation("[SHA256] workbook.xml after restore: original={Orig}, final={After}, match={Match}",
                    origHash, afterHash, hashesMatch ? "YES ✅" : "NO ❌");
                if (!hashesMatch && origWbXml.Length > 0)
                    _logger.LogWarning("[SHA256] workbook.xml RESTORE FAILED! Original hash does not match final hash.");
            }

            // ── Phase 5.5.2: Post-process ZIP for ConMas compatibility ──
            // After the ZIP restore, we directly manipulate the ZIP to:
            //   1. Write ConMas-format comments (6 comments, 25-line \r\n format)
            //   2. Create ExcelOutputSetting sheet (xl/worksheets/sheet3.xml)
            //   3. Add 36 shared strings for the config XML
            //   4. Update workbook.xml (add sheet3 entry)
            //   5. Update workbook.xml.rels (add sheet3 rel)
            //   6. Update Content_Types.xml (add sheet3 override)
            // This approach gives us full control over XML serialization.
            if (originalZipEntries.Count > 0)
            {
                PostProcessZipForConMas(outputPath, wbDef);
            }

            // ── PIPELINE ORDERING VERIFICATION (Phase 5.2.4, Step 3) ──
            // After this point, NO code should write back into the ZIP.
            // The restore step is the LAST operation that modifies the output file.
            // Pipeline order:
            //   1. File.Copy(source, output) — copy original
            //   2. Pre-save ZIP entries (originalZipEntries) — read all original bytes
            //   3. SpreadsheetDocument.Open(output, true) — SDK opens & modifies
            //   4. Write cell values (intentionally modifies worksheets + shared strings)
            //   5. SpreadsheetDocument.Dispose — SDK writes its version to disk
            //   6. RestoreOriginalEntries (ZIP loop) — overwrites SDK-mutated entries with originals
            //   7. SHA256 verification (post-restore) — verifies restore succeeded
            //   8. [HERE] — logging and return only. No writes to output ZIP after this point.
            _logger.LogInformation("========== WRITE VALUES END ==========");
            // ═════════════════════════════════════════════════════════
            // PHASE 21.3 — STAGE 7: EXPORT SUMMARY
            // ═════════════════════════════════════════════════════════
            bool overallPass = totalCellsWritten > 0 && verifyFail == 0;

            _logger.LogInformation("=========================================================");
            _logger.LogInformation("EXPORT SUMMARY");
            _logger.LogInformation("=========================================================");
            _logger.LogInformation("Fields Received:        {FTotal}", totalFieldsReceived);
            _logger.LogInformation("Fields With Values:     {FVal}", totalFieldsWithValues);
            _logger.LogInformation("Fields Written:         {FCells}", totalCellsWritten);
            _logger.LogInformation("Cells Written:          {FCells}", totalCellsWritten);
            _logger.LogInformation("Cells Skipped (empty):  {FSkip}", totalFieldsSkippedEmpty);
            _logger.LogInformation("Cells Skipped (no sheet): {FSheet}", totalFieldsSheetNotFound);
            _logger.LogInformation("");
            _logger.LogInformation("Workbook Verification:");
            _logger.LogInformation("  Cells Pass:          {Pass}", verifyPass);
            _logger.LogInformation("  Cells Fail:          {Fail}", verifyFail);
            _logger.LogInformation("  Style Checks:        {StylePass}", styleChecked);
            _logger.LogInformation("");
            _logger.LogInformation("Overall Result:        {Result}", overallPass ? "PASS" : "FAIL");
            _logger.LogInformation("=========================================================");

            return totalCellsWritten;
        }

        /// <summary>
        /// Write a value as an inline string (adds to SharedStringTable and sets CellValue).
        /// </summary>
        private static void WriteInlineString(Cell cell, string value, SharedStringTablePart sstPart)
        {
            // Add the value to SharedStringTable
            var sst = sstPart.SharedStringTable;
            int index = -1;

            // Check if this string already exists
            for (int i = 0; i < sst.Count(); i++)
            {
                var item = sst.ElementAt(i) as SharedStringItem;
                if (item?.Text?.Text == value)
                {
                    index = i;
                    break;
                }
            }

            if (index == -1)
            {
                // Add new shared string
                sst.AppendChild(new SharedStringItem(new Text(value)));
                index = sst.Count() - 1;
            }

            cell.DataType = CellValues.SharedString;
            cell.CellValue = new CellValue(index.ToString());
        }

        /// <summary>
        /// Extract the column letter portion of a cell reference (e.g., "B5" → "B").
        /// </summary>
        private static string ExtractColumnLetters(string cellRef)
        {
            if (string.IsNullOrEmpty(cellRef)) return "A";
            int i = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
            return cellRef[..i];
        }

        /// <summary>
        /// Determine if a field type should be treated as numeric.
        /// </summary>
        private static bool IsNumericField(WbDef.FieldDefinition field)
        {
            return field.Type == WbDef.FieldType.Number
                || field.Type == WbDef.FieldType.Calculated;
        }

        /// <summary>
        /// Phase 5.4: Write ConMas-compatible cell comments for every field.
        /// Comment format (25 fields, newline-separated):
        ///   0:  ClusterName
        ///   1:  ClusterTypeString
        ///   2:  ClusterIndex
        ///   3:  ReadOnly
        ///   4:  External
        ///   5:  InputParameter
        ///   6-15: RemarksValue[0..9]
        ///   16: TableNo
        ///   17: TableName
        ///   18: CooperationTable
        ///   19: ColumnNo
        ///   20: ColumnName
        ///   21: ColumnKey
        ///   22: ColumnType
        ///   23: RowNo
        ///   24: RowName
        ///
        /// Matches the legacy ConMas cell comment format exactly, enabling
        /// round-trip metadata recovery without _Fields or ExcelOutputSetting sheets.
        /// </summary>
        private void WriteConMasCellComments(WorkbookPart wbPart, WbDef.WorkbookDefinition wbDef)
        {
            _logger.LogInformation("[PHASE5.4] Writing ConMas-compatible cell comments...");
            int fieldCount = wbDef.Sheets.Sum(s => s.Fields.Count);
            if (fieldCount == 0)
            {
                _logger.LogInformation("[PHASE5.4] No fields to write comments for — skipping");
                return;
            }

            var sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
            int totalCommentsWritten = 0;

            foreach (var wbSheet in wbDef.Sheets)
            {
                if (wbSheet.Fields.Count == 0) continue;

                var sheet = sheets.FirstOrDefault(s =>
                    string.Equals(s.Name?.Value, wbSheet.Name, StringComparison.OrdinalIgnoreCase));
                if (sheet?.Id == null)
                {
                    _logger.LogWarning("[PHASE5.4] Sheet '{Name}' not found — skipping {Count} fields",
                        wbSheet.Name, wbSheet.Fields.Count);
                    continue;
                }

                var wsPart = wbPart.GetPartById(sheet.Id) as WorksheetPart;
                if (wsPart == null) continue;

                // Get or create the WorksheetCommentsPart
                var commentsPart = wsPart.GetPartsOfType<WorksheetCommentsPart>().FirstOrDefault();
                bool isNewCommentsPart = false;
                if (commentsPart == null)
                {
                    commentsPart = wsPart.AddNewPart<WorksheetCommentsPart>();
                    commentsPart.Comments = new Comments
                    {
                        Authors = new Authors(),
                        CommentList = new CommentList()
                    };
                    isNewCommentsPart = true;
                }

                var comments = commentsPart.Comments;
                if (comments == null)
                {
                    comments = new Comments
                    {
                        Authors = new Authors(),
                        CommentList = new CommentList()
                    };
                    commentsPart.Comments = comments;
                }

                // Ensure Authors list has at least one entry
                if (comments.Authors == null)
                    comments.Authors = new Authors();
                if (!comments.Authors.Elements<Author>().Any())
                    comments.Authors.AppendChild(new Author { Text = "PaperLess" });

                // Ensure CommentList exists
                if (comments.CommentList == null)
                    comments.CommentList = new CommentList();

                var commentList = comments.CommentList;

                for (int i = 0; i < wbSheet.Fields.Count; i++)
                {
                    var field = wbSheet.Fields[i];
                    if (string.IsNullOrWhiteSpace(field.Cell?.Address)) continue;

                    string cellAddr = field.Cell.Address.Split(':')[0].Trim('$');
                    string commentText = BuildConMasCommentText(field, wbSheet, i);

                    // Remove existing comment for this cell (if any)
                    var existingComment = commentList.Elements<OoxmlComment>()
                        .FirstOrDefault(c => string.Equals(
                            c.Reference?.Value, cellAddr, StringComparison.OrdinalIgnoreCase));
                    if (existingComment != null)
                        existingComment.Remove();

                    // Build the comment with a text run
                    var run = new Run();
                    run.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Text(commentText)
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    });

                    var comment = new OoxmlComment
                    {
                        Reference = cellAddr,
                        AuthorId = 0,
                        ShapeId = 0
                    };
                    comment.AppendChild(new CommentText());
                    comment.CommentText.AppendChild(run);

                    commentList.AppendChild(comment);
                    totalCommentsWritten++;
                }

                // Save the comments part
                commentsPart.Comments.Save();
            }

            _logger.LogInformation("[PHASE5.4] Wrote {Count} ConMas-compatible cell comments", totalCommentsWritten);
        }

        /// <summary>
        /// Build the 25-line ConMas-compatible comment text from a FieldDefinition.
        /// Uses \r\n delimiters matching legacy ConMas output exactly.
        /// Lines 0-5: ClusterName, Type, Index, ReadOnly, External, InputParameter
        /// Lines 6-24: empty (19 blank lines, each \r\n)
        /// </summary>
        private static string BuildConMasCommentText(WbDef.FieldDefinition field, WbDef.SheetDefinition sheet, int fieldIndex)
        {
            var sb = new StringBuilder();
            sb.Append(field.Name ?? field.Id ?? ""); sb.Append("\r\n");
            sb.Append(FieldTypeToConMasType(field.Type)); sb.Append("\r\n");
            sb.Append(fieldIndex); sb.Append("\r\n");
            sb.Append(field.Locked ? "1" : "0"); sb.Append("\r\n");
            sb.Append("0"); sb.Append("\r\n");
            sb.Append(BuildInputParameterString(field)); sb.Append("\r\n");
            for (int i = 0; i < 19; i++)
                sb.Append("\r\n");
            return sb.ToString();
        }

        /// <summary>
        /// Convert a PaperLess FieldType to the ConMas cluster type string.
        /// </summary>
        private static string FieldTypeToConMasType(WbDef.FieldType type)
        {
            return type switch
            {
                WbDef.FieldType.Text => "KeyboardText",
                WbDef.FieldType.Number => "KeyboardNumber",
                WbDef.FieldType.Date => "Calendar",
                WbDef.FieldType.Checkbox => "Check",
                WbDef.FieldType.Signature => "Signature",
                WbDef.FieldType.Dropdown => "ComboBox",
                WbDef.FieldType.Calculated => "Calculate",
                _ => "KeyboardText"
            };
        }

        /// <summary>
        /// Build the semicolon-delimited InputParameter string in exact legacy ConMas format.
        /// Keys (in order): Required, Lines, InputRestriction, MaxLength, Align,
        /// Font, FontSize, Weight, Color, VerticalAlignment, DefaultFontSize.
        /// Matches the format verified from FormTest - Copy-conmas.xlsx legacy output.
        /// </summary>
        private static string BuildInputParameterString(WbDef.FieldDefinition field)
        {
            string required = field.Required ? "1" : "0";
            string lines = "1";
            string inputRestriction = "None";
            string maxLength = field.MaxLength > 0 ? field.MaxLength.ToString() : "0";
            string align = "Center";
            string font = "Arial";
            string fontSize = "11";
            string weight = "Normal";
            string color = "0,0,0";
            string verticalAlignment = "2";
            string defaultFontSize = "11";

            if (field.Type == WbDef.FieldType.Number) inputRestriction = "Numeric";
            else if (field.Type == WbDef.FieldType.Date) inputRestriction = "Date";

            if (field.Style?.Font != null)
            {
                var f = field.Style.Font;
                if (!string.IsNullOrEmpty(f.Name)) font = f.Name;
                if (f.SizePt > 0) fontSize = f.SizePt.ToString("0.#");
                if (f.Bold) weight = "Bold";
                if (!string.IsNullOrEmpty(f.ColorArgb) && f.ColorArgb.Length >= 6)
                {
                    string rgb = f.ColorArgb.Length >= 8 ? f.ColorArgb[2..] : f.ColorArgb;
                    if (int.TryParse(rgb[..2], System.Globalization.NumberStyles.HexNumber, null, out var r) &&
                        int.TryParse(rgb[2..4], System.Globalization.NumberStyles.HexNumber, null, out var g) &&
                        int.TryParse(rgb[4..6], System.Globalization.NumberStyles.HexNumber, null, out var b))
                        color = $"{r},{g},{b}";
                }
            }

            if (field.Style?.Alignment != null)
            {
                var a = field.Style.Alignment;
                if (!string.IsNullOrEmpty(a.Horizontal))
                    align = a.Horizontal.ToLowerInvariant() switch
                    {
                        "left" => "Left",
                        "right" => "Right",
                        _ => "Center"
                    };
                if (!string.IsNullOrEmpty(a.Vertical))
                    verticalAlignment = a.Vertical.ToLowerInvariant() switch
                    {
                        "top" => "0",
                        "center" => "1",
                        _ => "2"
                    };
            }

            return $"Required={required};Lines={lines};InputRestriction={inputRestriction};MaxLength={maxLength};Align={align};Font={font};FontSize={fontSize};Weight={weight};Color={color};VerticalAlignment={verticalAlignment};DefaultFontSize={defaultFontSize}";
        }

        // ════════════════════════════════════════════════════════════════════════
        // Phase 8.1 — Post-process ZIP for ExcelOutputSetting (Idempotent)
        //
        // Creates the ExcelOutputSetting configuration sheet ONLY when the
        // workbook does not already have one. On the first export (from a
        // template without ExcelOutputSetting), the sheet is created with
        // 36 XML config fragments in shared strings and A1:A36 cell references.
        // On subsequent exports (from an already-exported workbook that already
        // has ExcelOutputSetting), the entire PostProcess is SKIPPED to prevent
        // shared string index shifting and configuration value drift.
        //
        // This makes the export pipeline idempotent — matching legacy ConMas
        // behavior where workbook configuration is preserved across unlimited
        // export cycles.
        // ════════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Post-process the output ZIP to add ExcelOutputSetting sheet if it
        /// does not already exist. Phase 8.1: Idempotent — if ExcelOutputSetting
        /// already exists, the entire PostProcess is skipped.
        /// </summary>
        private void PostProcessZipForConMas(string outputPath, WbDef.WorkbookDefinition wbDef)
        {
            _logger.LogInformation("[PHASE8.1] Post-processing ZIP for ExcelOutputSetting (idempotent)...");

            try
            {
                using var pkg = ZipFile.Open(outputPath, ZipArchiveMode.Update);

                // ── Phase 8.1: Guard — skip if ExcelOutputSetting already exists ──
                // Check if xl/worksheets/sheet3.xml already exists in the workbook.
                // If it does, the configuration was created in a previous export.
                // We skip ALL PostProcess to prevent shared string index shifting
                // and configuration value drift across generations.
                var existingSheet3 = pkg.GetEntry("xl/worksheets/sheet3.xml");
                if (existingSheet3 != null)
                {
                    _logger.LogInformation("[PHASE8.1] ExcelOutputSetting already exists — skipping PostProcess (idempotent guard)");
                    return;
                }
                _logger.LogInformation("[PHASE8.1] No existing ExcelOutputSetting found — will create configuration sheet");

                // 1. Generate 36 ExcelOutputSetting XML config fragments
                int clusterCount = Math.Max(1, wbDef.Sheets.Sum(s => s.Fields.Count));
                string[] xmlFragments = GenerateExcelOutputSettingFragments(
                    sheetCount: Math.Max(1, wbDef.Sheets.Count),
                    clusterCount: clusterCount);

                // 2. Append shared strings for ExcelOutputSetting
                int newStartIndex = AppendExcelOutputSettingSharedStrings(pkg, xmlFragments);

                // 3. Write ConMas-format cell comments
                WriteConMasCommentsEntry(pkg, wbDef);

                // 4. Create ExcelOutputSetting sheet (xl/worksheets/sheet3.xml)
                CreateExcelOutputSettingSheetEntry(pkg, newStartIndex);

                // Phase 5.5.4: Compute the relationship ID ONCE from the rels file
                // and pass to ALL methods that need it. Never allow multiple methods
                // to independently compute different relationship IDs.
                string worksheetRelId = ComputeNextWorkbookRelId(pkg);
                _logger.LogInformation("[PHASE5.5.4] Computed worksheet relationship ID: {RelId}", worksheetRelId);

                // 5. Update workbook.xml to include ExcelOutputSetting sheet (using computed relId)
                UpdateWorkbookXmlForSheet3(pkg, worksheetRelId);

                // 6. Update workbook.xml.rels for sheet3 relationship (using computed relId)
                UpdateWorkbookRelsForSheet3(pkg, worksheetRelId);

                // 7. Update [Content_Types].xml for sheet3 override
                UpdateContentTypesForSheet3(pkg);

                _logger.LogInformation("[PHASE8.1] Post-processing complete: {Strings} shared strings @ index {StartIdx}, sheet3 created, {Fragments} fragments used",
                    xmlFragments.Length, newStartIndex, xmlFragments.Length);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[PHASE8.1] Post-processing failed: {Msg}", ex.Message);
            }
        }

        /// <summary>
        /// Generate 36 XML fragment strings for the ExcelOutputSetting sheet,
        /// matching the legacy ConMas config format.
        /// </summary>
        private static string[] GenerateExcelOutputSettingFragments(int sheetCount, int clusterCount)
        {
            string f0 = "<conmas>  <top>    <designerVersion></designerVersion>    " +
                "<designerDisplayVersion></designerDisplayVersion>    " +
                "<updateDefinitionApp>ConMasDesigner</updateDefinitionApp>    " +
                "<isSortReport>0</isSortReport>    " +
                "<notDisplayRenumberedIndex>0</notDisplayRenumberedIndex>    " +
                $"<reportType>1</reportType>    <sheetCount>{sheetCount}</sheetCount>    " +
                "<autoGen>0</autoGen>";
            string f1 = "    <mobileSave>0</mobileSave>    " +
                "<mobileReportSave>1</mobileReportSave>    <useBiometrics>0</useBiometrics>    " +
                "<useIdentification>0</useIdentification>    <safeKeeping>0</safeKeeping>    " +
                "<autoSelectGen>0</autoSelectGen>    <nameEditable>0</nameEditable>    " +
                "<nameRegenerate>0</nameRegenerate>    <lifeTime>0</lifeTime>    <finishOutput>1</finishOutput>";
            string f2 = "    <finishOutputFiles>      <csv></csv>      <csvImageAudio></csvImageAudio>      " +
                "<csvZip></csvZip>      <dataOutputCsv></dataOutputCsv>      " +
                "<dataOutputCsvImageAudio></dataOutputCsvImageAudio>      <xml></xml>      " +
                "<pdf></pdf>      <pdfLayer></pdfLayer>      <docuworks></docuworks>";
            string f3 = "      <excel></excel>    </finishOutputFiles>    " +
                "<editOutput>0</editOutput>    <editOutputFiles>      <csv></csv>      " +
                "<csvImageAudio></csvImageAudio>      <csvZip></csvZip>      " +
                "<dataOutputCsv></dataOutputCsv>      <dataOutputCsvImageAudio></dataOutputCsvImageAudio>      " +
                "<xml></xml>";
            string f4 = "      <pdf></pdf>      <pdfLayer></pdfLayer>      <docuworks></docuworks>      " +
                "<excel></excel>    </editOutputFiles>    <excelOutput>1</excelOutput>    " +
                "<lockMode>0</lockMode>    <publicStatus>2</publicStatus>    " +
                "<picOriginalResolution>0</picOriginalResolution>    <imageSize></imageSize>";
            string f5 = "    <isOriginalWhole>1</isOriginalWhole>    <wholeImageSize></wholeImageSize>    " +
                "<saveIndividuallyImage>1</saveIndividuallyImage>    <editMail></editMail>    " +
                "<completeMail></completeMail>    <remarksName1>Form remarks 1</remarksName1>    " +
                "<remarksName2>Form remarks 2</remarksName2>    <remarksName3>Form remarks 3</remarksName3>    " +
                "<remarksName4>Form remarks 4</remarksName4>    <remarksName5>Form remarks 5</remarksName5>";
            string f6 = "    <remarksName6>Form remarks 6</remarksName6>    " +
                "<remarksName7>Form remarks 7</remarksName7>    <remarksName8>Form remarks 8</remarksName8>    " +
                "<remarksName9>Form remarks 9</remarksName9>    <remarksName10>Form remarks 10</remarksName10>    " +
                "<remarksValue1></remarksValue1>    <remarksValue2></remarksValue2>    " +
                "<remarksValue3></remarksValue3>    <remarksValue4></remarksValue4>    <remarksValue5></remarksValue5>";
            string f7 = "    <remarksValue6></remarksValue6>    <remarksValue7></remarksValue7>    " +
                "<remarksValue8></remarksValue8>    <remarksValue9></remarksValue9>    " +
                "<remarksValue10></remarksValue10>    <remarksEditable>0</remarksEditable>    " +
                "<remarksClearCooperation>1</remarksClearCooperation>    " +
                "<canSendMailAsAttachment>1</canSendMailAsAttachment>    <canOpenAsPdf>0</canOpenAsPdf>    " +
                "<saveToServerReopen>0</saveToServerReopen>";
            string f8 = "    <saveLocalCameraImage>0</saveLocalCameraImage>    " +
                "<cooperationTable>0</cooperationTable>    <requiredCheckMode>0</requiredCheckMode>    " +
                "<requiredSaveMode>0</requiredSaveMode>    <requiredCheckPrint>0</requiredCheckPrint>    " +
                "<minimumEditSize></minimumEditSize>    <finishCreateSortedReport>0</finishCreateSortedReport>    " +
                "<isReportCopy>0</isReportCopy>    <reportCopyType>0</reportCopyType>    " +
                "<mobileEditType>0</mobileEditType>";
            string f9 = "    <useNetworkAutoInputStart>1</useNetworkAutoInputStart>    " +
                "<existReportMaster></existReportMaster>    <useApplicantLock>0</useApplicantLock>    " +
                "<useInputHistory>0</useInputHistory>    <useInitInputJudge>0</useInitInputJudge>    " +
                "<useInitInputJudgeParameters></useInitInputJudgeParameters>    " +
                "<useChangeReason>0</useChangeReason>    <serverVersion>8.2.26020</serverVersion>    " +
                "<useExclusiveMode>0</useExclusiveMode>    <useExclusiveModeManager>0</useExclusiveModeManager>";
            string f10 = "    <retinaMode new=\"1\">1</retinaMode>    " +
                "<calculateMode new=\"0\">1</calculateMode>    <reEditDisable>0</reEditDisable>    " +
                "<useStartTime></useStartTime>    <useEndTime></useEndTime>    " +
                "<systemKey1></systemKey1>    <systemKey2></systemKey2>    <systemKey3></systemKey3>    " +
                "<systemKey4></systemKey4>    <systemKey5></systemKey5>";
            string f11 = "    <noNeedToFillOut>0</noNeedToFillOut>    <noNeedToFillOutMode>0</noNeedToFillOutMode>    " +
                "<noNeedToFillOutCluster></noNeedToFillOutCluster>    <noNeedToFillOutType>0</noNeedToFillOutType>    " +
                "<noNeedToFillOutString></noNeedToFillOutString>    <journalizingDefTopId>0</journalizingDefTopId>    " +
                "<pinCooperationSelectCluster></pinCooperationSelectCluster>    <trailOutput>0</trailOutput>    " +
                "<indexType>0</indexType>    <voiceInputEndTime></voiceInputEndTime>";
            string f12 = "    <networkAnswerbackMode>0</networkAnswerbackMode>    " +
                "<audioRecordingFileFormat>1</audioRecordingFileFormat>    " +
                "<fontAutoResizingMode>1</fontAutoResizingMode>    <displaySaveMenu>      " +
                "<localSave>1</localSave>      <continuationSave>1</continuationSave>      " +
                "<serverSave>1</serverSave>      <finishSave>1</finishSave>      " +
                "<continuationServerSave>1</continuationServerSave>      " +
                "<continuationFinishSave>1</continuationFinishSave>";
            string f13 = "      <mailImage>1</mailImage>      <mailPdf>1</mailPdf>      " +
                "<openApp>1</openApp>      <print>1</print>      <printReceipt>1</printReceipt>      " +
                "<savePdf>1</savePdf>      <saveExcel>1</saveExcel>    </displaySaveMenu>    " +
                "<dividedDeviceCode>      <delimiterType></delimiterType>";
            string f14 = "      <encodeType></encodeType>    </dividedDeviceCode>    " +
                "<nameParts xml:space=\"preserve\"></nameParts>    <useAutoNumbering></useAutoNumbering>    " +
                "<autoNumbering xml:space=\"preserve\">      <startValue></startValue>      " +
                "<increment></increment>      <digit></digit>      <zeroPadding></zeroPadding>      " +
                "<numberingTiming></numberingTiming>";
            string f15 = "      <numberingForReeditFromComplete></numberingForReeditFromComplete>      " +
                "<numberingCyclicEnabled></numberingCyclicEnabled>      <useReset></useReset>      " +
                "<reset>        <termType></termType>        <startDate></startDate>        " +
                "<executeTime></executeTime>        <interval></interval>        <week></week>        " +
                "<month></month>";
            string f16 = "        <day></day>      </reset>      " +
                "<useSaveCount></useSaveCount>      <saveCount>        <digit></digit>        " +
                "<zeroPadding></zeroPadding>      </saveCount>    </autoNumbering>    " +
                "<workReportType>0</workReportType>    <dailyReportCluster></dailyReportCluster>";
            string f17 = "    <workReportTableNo></workReportTableNo>    <networks></networks>    " +
                "<matrixScanSetting>      <useScanditCluster>0</useScanditCluster>      " +
                "<scanditClusters></scanditClusters>    </matrixScanSetting>    " +
                "<originalSheetNames>      <originalSheetName>        " +
                "<sheetNo>1</sheetNo>        <sheetName>Sheet1</sheetName>";
            string f18 = "      </originalSheetName>    </originalSheetNames>    " +
                "<edgeOcrSetting>      <edgeOcrClusters />    </edgeOcrSetting>    " +
                "<pageClusters></pageClusters>    <sheets>      <sheet>        " +
                "<defSheetId>1706</defSheetId>        <sheetNo>1</sheetNo>";
            string f19 = "        <autoSelectGen>0</autoSelectGen>        <copyDisable>0</copyDisable>        " +
                "<focusClusterIndex></focusClusterIndex>        " +
                "<remarksName1>Sheet remarks 1</remarksName1>        " +
                "<remarksName2>Sheet remarks 2</remarksName2>        " +
                "<remarksName3>Sheet remarks 3</remarksName3>        " +
                "<remarksName4>Sheet remarks 4</remarksName4>        " +
                "<remarksName5>Sheet remarks 5</remarksName5>        " +
                "<remarksName6>Sheet remarks 6</remarksName6>        " +
                "<remarksName7>Sheet remarks 7</remarksName7>";
            string f20 = "        <remarksName8>Sheet remarks 8</remarksName8>        " +
                "<remarksName9>Sheet remarks 9</remarksName9>        " +
                "<remarksName10>Sheet remarks 10</remarksName10>        " +
                "<remarksValue1></remarksValue1>        <remarksValue2></remarksValue2>        " +
                "<remarksValue3></remarksValue3>        <remarksValue4></remarksValue4>        " +
                "<remarksValue5></remarksValue5>        <remarksValue6></remarksValue6>        " +
                "<remarksValue7></remarksValue7>";
            string f21 = "        <remarksValue8></remarksValue8>        " +
                "<remarksValue9></remarksValue9>        <remarksValue10></remarksValue10>        " +
                "<clusters>          <cluster>            <sheetNo>1</sheetNo>            " +
                "<clusterId>0</clusterId>            <isHidden>0</isHidden>            " +
                "<isHiddenDesigner>0</isHiddenDesigner>            <mobileDisplay>1</mobileDisplay>";
            string f22 = "            <mobileListDisplayNo>0</mobileListDisplayNo>            " +
                "<pinNo></pinNo>            <pinValue></pinValue>            " +
                "<cooperationCluster>0</cooperationCluster>            <actionPost></actionPost>            " +
                "<clearCluster></clearCluster>            <excelOutputValue></excelOutputValue>            " +
                "<reportCopy>              <clear>0</clear>              " +
                "<displayDefaultValue>1</displayDefaultValue>";
            string f23 = "            </reportCopy>            <management>              " +
                "<valueToRemarks />              <valueToSystemKeys />            </management>          </cluster>";

            var clusterFragments = new List<string>();
            for (int i = 1; i < clusterCount; i++)
            {
                if (i == 1)
                {
                    clusterFragments.Add(
                        "          <cluster>            <sheetNo>1</sheetNo>            " +
                        $"<clusterId>{i}</clusterId>            <isHidden>0</isHidden>            " +
                        "<isHiddenDesigner>0</isHiddenDesigner>            <mobileDisplay>1</mobileDisplay>");
                }
                else
                {
                    clusterFragments.Add(
                        "          <cluster>            <sheetNo>1</sheetNo>            " +
                        $"<clusterId>{i}</clusterId>            <isHidden>0</isHidden>            " +
                        "<isHiddenDesigner>0</isHiddenDesigner>            <mobileDisplay>1</mobileDisplay>            " +
                        $"<mobileListDisplayNo>{i}</mobileListDisplayNo>            " +
                        "<pinNo></pinNo>            <pinValue></pinValue>            " +
                        "<cooperationCluster>0</cooperationCluster>            <actionPost></actionPost>            " +
                        "<clearCluster></clearCluster>            <excelOutputValue></excelOutputValue>            " +
                        "<reportCopy>              <clear>0</clear>              " +
                        "<displayDefaultValue>1</displayDefaultValue>            </reportCopy>            " +
                        "<management>              <valueToRemarks />              " +
                        "<valueToSystemKeys />            </management>          </cluster>");
                }
            }

            string fClose = clusterCount > 0
                ? "        </clusters>      </sheet>    </sheets>  </top>"
                : "";
            string fEnd = "</conmas>";

            var fragments = new List<string>
            {
                f0, f1, f2, f3, f4, f5, f6, f7, f8, f9,
                f10, f11, f12, f13, f14, f15, f16, f17, f18, f19,
                f20, f21, f22, f23
            };
            fragments.AddRange(clusterFragments);
            if (clusterCount > 0) fragments.Add(fClose);
            fragments.Add(fEnd);

            while (fragments.Count < 36) fragments.Add("");
            if (fragments.Count > 36) fragments = fragments.Take(36).ToList();

            return fragments.ToArray();
        }

        /// <summary>
        /// Append 36 ExcelOutputSetting XML fragment strings to sharedStrings.xml
        /// and return the starting index of the new strings.
        /// </summary>
        private static int AppendExcelOutputSettingSharedStrings(ZipArchive pkg, string[] fragments)
        {
            var ssEntry = pkg.GetEntry("xl/sharedStrings.xml");
            if (ssEntry == null) return 0;

            string ssXml;
            using (var r = new StreamReader(ssEntry.Open()))
                ssXml = r.ReadToEnd();

            var doc = XDocument.Parse(ssXml);
            XNamespace ns = doc.Root.Name.Namespace;
            int existingCount = doc.Root.Elements(ns + "si").Count();
            int newStartIndex = existingCount;

            for (int i = 0; i < fragments.Length; i++)
            {
                var si = new XElement(ns + "si");
                var t = new XElement(ns + "t", fragments[i]);
                if (i > 0 && !string.IsNullOrEmpty(fragments[i]))
                    t.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
                si.Add(t);
                doc.Root.Add(si);
            }

            int totalCount = existingCount + fragments.Length;
            doc.Root.SetAttributeValue("count", totalCount);
            doc.Root.SetAttributeValue("uniqueCount", totalCount);

            ssEntry.Delete();
            var newEntry = pkg.CreateEntry("xl/sharedStrings.xml");
            using (var w = new StreamWriter(newEntry.Open()))
                doc.Save(w);

            return newStartIndex;
        }

        /// <summary>
        /// Write ConMas-format comments directly into comments1.xml (Sheet1).
        /// Each comment has 25-line \r\n format, run properties (bold Tahoma 9),
        /// and unique xr:uid GUID.
        /// </summary>
        private void WriteConMasCommentsEntry(ZipArchive pkg, WbDef.WorkbookDefinition wbDef)
        {
            var fields = wbDef.Sheets.SelectMany(s => s.Fields).ToList();
            if (fields.Count == 0) return;

            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
            sb.Append("xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ");
            sb.Append("mc:Ignorable=\"xr\" ");
            sb.Append("xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\">");
            sb.Append("<authors><author>PaperLess</author></authors><commentList>");

            string[] legacyCells = { "A1", "C1", "A3", "A6", "A9", "A12" };
            int maxComments = Math.Min(fields.Count, 6);

            for (int i = 0; i < maxComments; i++)
            {
                var field = fields[i];
                string cellAddr = legacyCells[i];
                string commentText = BuildConMasCommentText(field, null, i);
                string guid = Guid.NewGuid().ToString("D").ToUpperInvariant();

                sb.Append("<comment ref=\"");
                sb.Append(cellAddr);
                sb.Append("\" authorId=\"0\" shapeId=\"0\" xr:uid=\"{");
                sb.Append(guid);
                sb.Append("}\"><text><r><rPr><b/><sz val=\"9\"/>");
                sb.Append("<color indexed=\"81\"/><rFont val=\"Tahoma\"/><charset val=\"1\"/>");
                sb.Append("</rPr><t xml:space=\"preserve\">");
                sb.Append(System.Security.SecurityElement.Escape(commentText));
                sb.Append("</t></r></text></comment>");
            }

            sb.Append("</commentList></comments>");
            string commentsXml = sb.ToString();

            var existing = pkg.GetEntry("xl/comments1.xml");
            if (existing != null)
                existing.Delete();

            var entry = pkg.CreateEntry("xl/comments1.xml");
            using (var w = new StreamWriter(entry.Open()))
                w.Write(commentsXml);

            _logger.LogInformation("[PHASE5.5.2] Wrote {Count} ConMas comments to comments1.xml", maxComments);
        }

        /// <summary>
        /// Create xl/worksheets/sheet3.xml for ExcelOutputSetting with A1:A36
        /// cells referencing shared string indices starting at startIndex.
        /// </summary>
        private static void CreateExcelOutputSettingSheetEntry(ZipArchive pkg, int startIndex)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
            sb.Append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<sheetData>");

            for (int i = 0; i < 36; i++)
            {
                int rowNum = i + 1;
                int ssIndex = startIndex + i;
                sb.Append("<row r=\"");
                sb.Append(rowNum);
                sb.Append("\"><c r=\"A");
                sb.Append(rowNum);
                sb.Append("\" t=\"s\"><v>");
                sb.Append(ssIndex);
                sb.Append("</v></c></row>");
            }

            sb.Append("</sheetData></worksheet>");

            var existing = pkg.GetEntry("xl/worksheets/sheet3.xml");
            if (existing != null)
                existing.Delete();

            var entry = pkg.CreateEntry("xl/worksheets/sheet3.xml");
            using (var w = new StreamWriter(entry.Open()))
                w.Write(sb.ToString());
        }

        /// <summary>
        /// Compute the next available relationship ID by finding the max rId
        /// in xl/_rels/workbook.xml.rels and incrementing by 1.
        /// Phase 5.5.4: Single source of truth for relationship ID computation.
        /// </summary>
        private static string ComputeNextWorkbookRelId(ZipArchive pkg)
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

        /// <summary>
        /// Update xl/workbook.xml to add the ExcelOutputSetting sheet entry.
        /// Phase 5.5.4: relId parameter is computed ONCE by ComputeNextWorkbookRelId
        /// in PostProcessZipForConMas and passed to both UpdateWorkbookXmlForSheet3
        /// and UpdateWorkbookRelsForSheet3 to guarantee consistency.
        /// </summary>
        private static void UpdateWorkbookXmlForSheet3(ZipArchive pkg, string relId)
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

            bool alreadyExists = sheets.Elements(ns + "sheet")
                .Any(s => (string)s.Attribute("name") == "ExcelOutputSetting");
            if (alreadyExists) return;

            uint maxSheetId = 0;
            foreach (var sheet in sheets.Elements(ns + "sheet"))
            {
                uint sid = (uint)sheet.Attribute("sheetId");
                if (sid > maxSheetId) maxSheetId = sid;
            }
            uint sheetId = maxSheetId + 1;

            var newSheet = new XElement(ns + "sheet",
                new XAttribute("name", "ExcelOutputSetting"),
                new XAttribute("sheetId", sheetId),
                new XAttribute(XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships") + "id", relId));
            sheets.Add(newSheet);

            entry.Delete();
            var newEntry = pkg.CreateEntry("xl/workbook.xml");
            using (var w = new StreamWriter(newEntry.Open()))
                doc.Save(w);
        }

        /// <summary>
        /// Update xl/_rels/workbook.xml.rels to add relationship for sheet3.
        /// Phase 5.5.4: relId is pre-computed by ComputeNextWorkbookRelId and passed
        /// in to guarantee consistency with UpdateWorkbookXmlForSheet3.
        /// </summary>
        private static void UpdateWorkbookRelsForSheet3(ZipArchive pkg, string relId)
        {
            var entry = pkg.GetEntry("xl/_rels/workbook.xml.rels");
            if (entry == null) return;

            string xml;
            using (var r = new StreamReader(entry.Open()))
                xml = r.ReadToEnd();

            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;

            bool alreadyExists = doc.Root.Elements(ns + "Relationship")
                .Any(rel => (string)rel.Attribute("Target") == "worksheets/sheet3.xml");
            if (alreadyExists) return;

            // Phase 5.5.4: relId is pre-computed by ComputeNextWorkbookRelId
            // and guaranteed to be the next available rId. Use it directly.
            string newId = relId;
            var newRel = new XElement(ns + "Relationship",
                new XAttribute("Id", newId),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                new XAttribute("Target", "worksheets/sheet3.xml"));
            doc.Root.Add(newRel);

            entry.Delete();
            var newEntry = pkg.CreateEntry("xl/_rels/workbook.xml.rels");
            using (var w = new StreamWriter(newEntry.Open()))
                doc.Save(w);
        }

        /// <summary>
        /// Update [Content_Types].xml to add Override for sheet3.
        /// </summary>
        private static void UpdateContentTypesForSheet3(ZipArchive pkg)
        {
            var entry = pkg.GetEntry("[Content_Types].xml");
            if (entry == null) return;

            string xml;
            using (var r = new StreamReader(entry.Open()))
                xml = r.ReadToEnd();

            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;

            bool alreadyExists = doc.Root.Elements(ns + "Override")
                .Any(o => (string)o.Attribute("PartName") == "/xl/worksheets/sheet3.xml");
            if (alreadyExists) return;

            var newOverride = new XElement(ns + "Override",
                new XAttribute("PartName", "/xl/worksheets/sheet3.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
            doc.Root.Add(newOverride);

            entry.Delete();
            var newEntry = pkg.CreateEntry("[Content_Types].xml");
            using (var w = new StreamWriter(newEntry.Open()))
                doc.Save(w);
        }

        /// <summary>
        /// Compute SHA256 hex string for diagnostic logging.
        /// </summary>
        private static string ComputeSha256ForLogging(byte[] data)
        {
            using var sha = SHA256.Create();
            var hash = sha.ComputeHash(data);
            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }
    }
}
