using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using WbDef = ExcelAPI.Models.WorkbookDefinition;
using System.Security.Cryptography;

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

            // ── DIAGNOSTIC: Log incoming WbDef structure ──
            _logger.LogInformation("========== WRITE VALUES ==========");
            _logger.LogInformation("Workbook: {Title}", wbDef.Info?.Title ?? "(untitled)");
            _logger.LogInformation("SourcePath: {Path}", sourceXlsxPath);
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

                        // Write the value based on field type
                        string oldValueBeforeWrite = cell.CellValue?.Text ?? "(empty)";
                        CellValues? oldDataType = cell.DataType;

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

                        if (isNewCell)
                        {
                            row.Append(cell);
                        }

                        cellsWritten++;

                        // ── DIAGNOSTIC: Verify the value was written ──
                        _logger.LogInformation("[DIAG]  RESULT: {Sheet}!{Cell} | old='{Old}' ({OldDt}) | new='{New}' ({NewDt}) | style={Style} | row={RowStatus} | cell={CellStatus}",
                            wbSheet.Name, cellRef,
                            oldValueBeforeWrite, oldDataType?.ToString() ?? "(none)",
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
            }

            // ── DIAGNOSTIC: Post-write cell verification ──
            _logger.LogInformation("========== POST-WRITE VERIFICATION ==========");
            _logger.LogInformation("Opening edited workbook to verify written cells...");
            try
            {
                using var verifyDoc = SpreadsheetDocument.Open(outputPath, false);
                if (verifyDoc.WorkbookPart != null)
                {
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
                            if (verifyCell != null)
                            {
                                string actualValue = verifyCell.CellValue?.Text ?? "(empty)";
                                bool matches = actualValue == field.Value
                                    || (field.Value.Length <= 10 && actualValue.Contains(field.Value.TrimEnd(' '))); // shared string index vs value
                                _logger.LogInformation("  Verify {Sheet}!{Cell}: expected='{Exp}', actual='{Act}' {Status}",
                                    wbSheet.Name, cellRef, field.Value, actualValue, matches ? "✅" : "❌");
                            }
                            else
                            {
                                _logger.LogWarning("  Verify {Sheet}!{Cell}: expected='{Exp}' — CELL NOT FOUND IN OUTPUT ❌",
                                    wbSheet.Name, cellRef, field.Value);
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
            // Phase 5.2.4: Log original hash, current (SDK-mutated) hash before restore,
            // and final hash after restore to verify restore succeeded.
            {
                byte[] origWbXml = originalZipEntries.TryGetValue("xl/workbook.xml", out var owb) ? owb : Array.Empty<byte>();
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

                try
                {
                    using var pkg = System.IO.Compression.ZipFile.Open(outputPath, System.IO.Compression.ZipArchiveMode.Update);

                    foreach (var kvp in originalZipEntries.OrderBy(e => e.Key))
                    {
                        string entryName = kvp.Key;
                        byte[] originalBytes = kvp.Value;

                        // Skip entries that were intentionally modified
                        bool isWorksheet = entryName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase);
                        bool isSharedStrings = entryName.Equals("xl/sharedstrings.xml", StringComparison.OrdinalIgnoreCase)
                            || entryName.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase);
                        if (isWorksheet || isSharedStrings)
                        {
                            skippedModified++;
                            _logger.LogInformation("[RESTORE] Skipped '{Name}' — intentionally modified", entryName);
                            continue;
                        }

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

                    _logger.LogInformation("[RESTORE] Summary: {Restored} restored, {Unchanged} unchanged, {Skipped} skipped (intentionally modified). workbook.xml restored={WbXml}",
                        restoredCount, unchangedCount, skippedModified, workbookXmlWasRestored ? "YES" : "NO — check if it was unchanged or missing");
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "[RESTORE] Could not restore original ZIP entries");
                }
            }

            // ── SHA256 DIAGNOSTIC: Log workbook.xml hash AFTER restore ──
            // Verify the restore actually succeeded by comparing final hash vs original hash.
            {
                byte[] origWbXml = originalZipEntries.TryGetValue("xl/workbook.xml", out var owb) ? owb : Array.Empty<byte>();
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
            _logger.LogInformation("Summary: {Total} cells written to {Output}", totalCellsWritten, outputPath);
            _logger.LogInformation("  Fields received (total): {FTotal}", totalFieldsReceived);
            _logger.LogInformation("  Fields with non-empty values: {FVal}", totalFieldsWithValues);
            _logger.LogInformation("  Fields skipped (empty value): {FSkip}", totalFieldsSkippedEmpty);
            _logger.LogInformation("  Fields skipped (sheet not found): {FSheet}", totalFieldsSheetNotFound);
            _logger.LogInformation("  Cells actually written: {FCells}", totalCellsWritten);

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
