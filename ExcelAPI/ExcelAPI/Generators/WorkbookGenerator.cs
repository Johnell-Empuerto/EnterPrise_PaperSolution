using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelAPI.Models;
using Microsoft.Extensions.Logging;

namespace ExcelAPI.Generators
{
    public class WorkbookGenerator
    {
        private readonly ILogger<WorkbookGenerator> _logger;

        public WorkbookGenerator(ILogger<WorkbookGenerator> logger)
        {
            _logger = logger;
        }

        private static long _exportSequence;
        private static readonly object _exportLock = new();

        private string ExportId()
        {
            long seq = Interlocked.Increment(ref _exportSequence);
            return $"EXP_{seq:D4}_{DateTime.UtcNow:HHmmss}";
        }

        private static string ProcessId() => Process.GetCurrentProcess().Id.ToString();

        /// <summary>
        /// Instrument a name-related COM operation with BEFORE/AFTER logging.
        /// Returns false if the operation threw.
        /// </summary>
        private bool NameOp(string exportId, string opLabel, System.Action action)
        {
            _logger.LogInformation("[FORENSIC][{ExportId}] >>> NAME-OP: {Op}", exportId, opLabel);
            try
            {
                action();
                _logger.LogInformation("[FORENSIC][{ExportId}] <<< NAME-OP SUCCESS: {Op}", exportId, opLabel);
                return true;
            }
            catch (COMException ex) when ((uint)ex.ErrorCode == 0x800A03EC)
            {
                _logger.LogError("[FORENSIC][{ExportId}] <<< NAME-OP FAILED (0x800A03EC): {Op}\nMessage: {Msg}\nHResult: 0x{HR:X8}\nStack: {Stack}",
                    exportId, opLabel, ex.Message, (uint)ex.ErrorCode, ex.StackTrace);
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError("[FORENSIC][{ExportId}] <<< NAME-OP FAILED (OTHER): {Op}\nMessage: {Msg}\nType: {Type}\nStack: {Stack}",
                    exportId, opLabel, ex.Message, ex.GetType().FullName, ex.StackTrace);
                return false;
            }
        }

        /// <summary>
        /// Enumerate all defined names in a workbook into a log string.
        /// </summary>
        private string DumpNames(Workbook wb, string exportId)
        {
            var sb = new StringBuilder();
            try
            {
                var names = wb.Names;
                int cnt = names.Count;
                sb.Append($"Names.Count={cnt}");
                if (cnt > 0)
                {
                    foreach (Name n in names)
                    {
                        try
                        {
                            string nName = n.Name ?? "(null)";
                            string nRef = n.RefersTo ?? "(null)";
                            sb.Append($" | [{nName}]->{nRef}");
                        }
                        catch (Exception e)
                        {
                            sb.Append($" | [ERROR reading name: {e.Message}]");
                        }
                        finally
                        {
                            try { Marshal.ReleaseComObject(n); } catch { }
                        }
                    }
                }
                Marshal.ReleaseComObject(names);
            }
            catch (Exception ex)
            {
                sb.Append($"ERROR enumerating names: {ex.Message}");
            }
            string result = sb.ToString();
            _logger.LogInformation("[FORENSIC][{ExportId}] Names dump: {Dump}", exportId, result);
            return result;
        }

        public string Generate(FormDefinition form, string outputPath)
        {
            string exportId = ExportId();
            _logger.LogInformation("[FORENSIC][{ExportId}] ===== GENERATE ENTER =====", exportId);
            _logger.LogInformation("[FORENSIC][{ExportId}] ProcessId={Pid} sheets={SheetCount} clusters={ClusterCount} outputPath={Path}",
                exportId, ProcessId(), form.Sheets?.Count ?? 0, form.Clusters?.Count ?? 0, outputPath);

            Application? excelApp = null;
            Workbook? workbook = null;

            try
            {
                _logger.LogInformation("[FORENSIC][{ExportId}] >>> OP: new Excel.Application", exportId);
                excelApp = new Application { Visible = false, DisplayAlerts = false };
                _logger.LogInformation("[FORENSIC][{ExportId}] <<< OP SUCCESS: new Excel.Application", exportId);

                _logger.LogInformation("[FORENSIC][{ExportId}] >>> OP: Workbooks.Add(xlWBATWorksheet)", exportId);
                workbook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                _logger.LogInformation("[FORENSIC][{ExportId}] <<< OP SUCCESS: Workbooks.Add() — workbook=\"{Wb}\"", exportId, workbook?.Name ?? "null");

                // Dump initial names state (should be empty for a fresh workbook)
                _logger.LogInformation("[FORENSIC][{ExportId}] === INITIAL Names state (fresh workbook) ===", exportId);
                DumpNames(workbook, exportId);

                string printAreaAddress = CalculatePrintArea(form);
                _logger.LogInformation("[FORENSIC][{ExportId}] CalculatePrintArea result: {Addr}", exportId, printAreaAddress);

                _logger.LogInformation("[FORENSIC][{ExportId}] >>> OP: EnsureSheets()", exportId);
                EnsureSheets(workbook, form, out var sheetIndexMap, exportId);
                var mapItems = new List<string>();
                foreach (var kv in sheetIndexMap)
                    mapItems.Add($"{kv.Key}={kv.Value}");
                _logger.LogInformation("[FORENSIC][{ExportId}] <<< OP SUCCESS: EnsureSheets(). Sheet map: {Map}",
                    exportId, string.Join(", ", mapItems));

                _logger.LogInformation("[FORENSIC][{ExportId}] >>> OP: ApplySheetLayout()", exportId);
                ApplySheetLayout(workbook, form, sheetIndexMap, printAreaAddress, exportId);
                _logger.LogInformation("[FORENSIC][{ExportId}] <<< OP SUCCESS: ApplySheetLayout()", exportId);

                _logger.LogInformation("[FORENSIC][{ExportId}] >>> OP: WriteCellValues()", exportId);
                WriteCellValues(workbook, form, sheetIndexMap, exportId);
                _logger.LogInformation("[FORENSIC][{ExportId}] <<< OP SUCCESS: WriteCellValues()", exportId);

                _logger.LogInformation("[FORENSIC][{ExportId}] >>> OP: WriteCellComments()", exportId);
                WriteCellComments(workbook, form, sheetIndexMap, exportId);
                _logger.LogInformation("[FORENSIC][{ExportId}] <<< OP SUCCESS: WriteCellComments()", exportId);

                _logger.LogInformation("[FORENSIC][{ExportId}] >>> OP: PopulateFieldsWorksheet()", exportId);
                PopulateFieldsWorksheet(workbook, form, sheetIndexMap, exportId);
                _logger.LogInformation("[FORENSIC][{ExportId}] <<< OP SUCCESS: PopulateFieldsWorksheet()", exportId);

                _logger.LogInformation("[FORENSIC][{ExportId}] >>> OP: CreateExcelOutputSetting()", exportId);
                CreateExcelOutputSetting(workbook, form, sheetIndexMap, exportId);
                _logger.LogInformation("[FORENSIC][{ExportId}] <<< OP SUCCESS: CreateExcelOutputSetting()", exportId);

                // Dump names BEFORE SetDefinedNames
                _logger.LogInformation("[FORENSIC][{ExportId}] === Names state BEFORE SetDefinedNames ===", exportId);
                DumpNames(workbook, exportId);

                _logger.LogInformation("[FORENSIC][{ExportId}] >>> OP: SetDefinedNames() printAreaAddress=\"{Addr}\"", exportId, printAreaAddress);
                SetDefinedNames(workbook, form, printAreaAddress, exportId);
                _logger.LogInformation("[FORENSIC][{ExportId}] <<< OP: SetDefinedNames() returned (see above for success/failure)", exportId);

                // Dump names AFTER SetDefinedNames
                _logger.LogInformation("[FORENSIC][{ExportId}] === Names state AFTER SetDefinedNames ===", exportId);
                DumpNames(workbook, exportId);

                _logger.LogInformation("[FORENSIC][{ExportId}] >>> OP: workbook.SaveAs({Path})", exportId, outputPath);
                bool saved = NameOp(exportId, $"workbook.SaveAs({outputPath})", () => workbook.SaveAs(outputPath));
                if (saved)
                {
                    _logger.LogInformation("[FORENSIC][{ExportId}] <<< OP SUCCESS: workbook.SaveAs()", exportId);
                    _logger.LogInformation("[FORENSIC][{ExportId}] Workbook saved to {Path}", exportId, outputPath);
                }
                else
                {
                    // Already logged by NameOp
                }
                return outputPath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[FORENSIC][{ExportId}] ===== GENERATE FAILED at: {Msg}\nStack: {Stack}", exportId, ex.Message, ex.StackTrace);
                // Dump names for post-mortem
                try { if (workbook != null) DumpNames(workbook, exportId); } catch { }
                throw;
            }
            finally
            {
                if (workbook != null)
                {
                    try { _logger.LogInformation("[FORENSIC][{ExportId}] Closing workbook", exportId); workbook.Close(SaveChanges: false); }
                    catch (Exception e) { _logger.LogWarning("[FORENSIC][{ExportId}] workbook.Close() exception: {Msg}", exportId, e.Message); }
                    Marshal.ReleaseComObject(workbook);
                    _logger.LogInformation("[FORENSIC][{ExportId}] Workbook COM object released", exportId);
                }
                if (excelApp != null)
                {
                    try { _logger.LogInformation("[FORENSIC][{ExportId}] Quitting Excel", exportId); excelApp.Quit(); }
                    catch (Exception e) { _logger.LogWarning("[FORENSIC][{ExportId}] excelApp.Quit() exception: {Msg}", exportId, e.Message); }
                    Marshal.ReleaseComObject(excelApp);
                    _logger.LogInformation("[FORENSIC][{ExportId}] ExcelApp COM object released", exportId);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                _logger.LogInformation("[FORENSIC][{ExportId}] ===== GENERATE EXIT =====", exportId);
            }
        }

        private string CalculatePrintArea(FormDefinition form)
        {
            if (form.Clusters == null || form.Clusters.Count == 0)
                return "";

            int minRow = int.MaxValue, maxRow = 0;
            int minCol = int.MaxValue, maxCol = 0;

            foreach (var cluster in form.Clusters)
            {
                if (string.IsNullOrWhiteSpace(cluster.CellAddress)) continue;

                string addr = cluster.CellAddress;
                string[] parts = addr.Contains(":") ? addr.Split(':') : [addr, addr];

                foreach (var part in parts)
                {
                    var (col, row) = ParseCellRef(part.Trim().Replace("$", ""));
                    if (row > 0 && col > 0)
                    {
                        minRow = Math.Min(minRow, row);
                        maxRow = Math.Max(maxRow, row);
                        minCol = Math.Min(minCol, col);
                        maxCol = Math.Max(maxCol, col);
                    }
                }
            }

            if (maxRow == 0 || maxCol == 0) return "";

            string startCol = ColumnIndexToLetters(minCol);
            string endCol = ColumnIndexToLetters(maxCol);
            return $"$A$1:${endCol}${maxRow}";
        }

        private static (int col, int row) ParseCellRef(string cellRef)
        {
            if (string.IsNullOrEmpty(cellRef)) return (0, 0);
            int i = 0, col = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i]))
            {
                col = col * 26 + (char.ToUpper(cellRef[i]) - 'A' + 1);
                i++;
            }
            if (i >= cellRef.Length || !int.TryParse(cellRef[i..], out int row)) return (0, 0);
            return (col, row);
        }

        private static string ColumnIndexToLetters(int col)
        {
            if (col <= 0) return "A";
            var result = new StringBuilder();
            while (col > 0) { col--; result.Insert(0, (char)('A' + (col % 26))); col /= 26; }
            return result.ToString();
        }

        // ────────────────────────────────────────────────────────────────────────────
        // EnsureSheets — FORENSIC LOGGING
        // ────────────────────────────────────────────────────────────────────────────
        private void EnsureSheets(Workbook workbook, FormDefinition form,
            out Dictionary<string, int> sheetIndexMap, string exportId)
        {
            sheetIndexMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            int targetSheetCount = 3;
            int currentCount = workbook.Worksheets.Count;
            _logger.LogInformation("[FORENSIC][{ExportId}] EnsureSheets: target={Target} current={Current}", exportId, targetSheetCount, currentCount);

            // Remove excess sheets
            while (workbook.Worksheets.Count > targetSheetCount)
            {
                int idx = workbook.Worksheets.Count;
                _logger.LogInformation("[FORENSIC][{ExportId}] >>> NAME-OP: Deleting sheet at index {Idx}", exportId, idx);
                bool deleted = NameOp(exportId, $"Worksheet[{idx}].Delete()", () =>
                    ((Excel.Worksheet)workbook.Worksheets[idx]).Delete());
                if (!deleted)
                {
                    _logger.LogWarning("[FORENSIC][{ExportId}] Delete sheet {Idx} FAILED — continuing", exportId, idx);
                }
            }

            // Add sheets if needed
            while (workbook.Worksheets.Count < targetSheetCount)
            {
                _logger.LogInformation("[FORENSIC][{ExportId}] >>> NAME-OP: Worksheets.Add() (current count={Count})", exportId, workbook.Worksheets.Count);
                bool added = NameOp(exportId, "Worksheets.Add(After: last)", () =>
                {
                    var newSheet = (Excel.Worksheet)workbook.Worksheets.Add(
                        After: workbook.Worksheets[workbook.Worksheets.Count]);
                    _logger.LogInformation("[FORENSIC][{ExportId}]   New sheet created with default name=\"{Name}\"", exportId, newSheet.Name);
                    Marshal.ReleaseComObject(newSheet);
                });
                if (!added)
                {
                    _logger.LogWarning("[FORENSIC][{ExportId}] Worksheets.Add() FAILED — continuing", exportId);
                }
            }

            string contentSheetName = "Sheet1";
            if (form.Sheets != null && form.Sheets.Count > 0)
                contentSheetName = form.Sheets[0].Name;

            // Rename COM[3] -> ExcelOutputSetting
            _logger.LogInformation("[FORENSIC][{ExportId}] >>> NAME-OP: Rename WS[3] -> \"ExcelOutputSetting\"", exportId);
            NameOp(exportId, "Worksheet[3].Name = \"ExcelOutputSetting\"", () =>
            {
                var ws3 = (Excel.Worksheet)workbook.Worksheets[3];
                _logger.LogInformation("[FORENSIC][{ExportId}]   Before rename WS[3].Name=\"{Name}\"", exportId, ws3.Name);
                ws3.Name = "ExcelOutputSetting";
                Marshal.ReleaseComObject(ws3);
            });

            // Rename COM[1] -> _Fields FIRST to free up the default "Sheet1" name
            _logger.LogInformation("[FORENSIC][{ExportId}] >>> NAME-OP: Rename WS[1] -> \"_Fields\"", exportId);
            NameOp(exportId, "Worksheet[1].Name = \"_Fields\"", () =>
            {
                var ws1 = (Excel.Worksheet)workbook.Worksheets[1];
                _logger.LogInformation("[FORENSIC][{ExportId}]   Before rename WS[1].Name=\"{Name}\"", exportId, ws1.Name);
                ws1.Name = "_Fields";
                Marshal.ReleaseComObject(ws1);
            });

            // Rename COM[2] -> contentSheetName (now "Sheet1" is free since WS[1] was renamed)
            _logger.LogInformation("[FORENSIC][{ExportId}] >>> NAME-OP: Rename WS[2] -> \"{Name}\"", exportId, contentSheetName);
            NameOp(exportId, $"Worksheet[2].Name = \"{contentSheetName}\"", () =>
            {
                var ws2 = (Excel.Worksheet)workbook.Worksheets[2];
                _logger.LogInformation("[FORENSIC][{ExportId}]   Before rename WS[2].Name=\"{Name}\"", exportId, ws2.Name);
                ws2.Name = contentSheetName;
                Marshal.ReleaseComObject(ws2);
            });

            sheetIndexMap["_Fields"] = 1;
            if (form.Sheets != null && form.Sheets.Count > 0)
                sheetIndexMap[form.Sheets[0].Id] = 2;
            sheetIndexMap["ExcelOutputSetting"] = 3;

            _logger.LogInformation("[FORENSIC][{ExportId}] EnsureSheets DONE: _Fields(1), {Content}(2), ExcelOutputSetting(3)", exportId, contentSheetName);
        }

        // ────────────────────────────────────────────────────────────────────────────
        // ApplySheetLayout — FORENSIC LOGGING
        // ────────────────────────────────────────────────────────────────────────────
        private void ApplySheetLayout(Workbook workbook, FormDefinition form,
            Dictionary<string, int> sheetIndexMap, string printAreaAddress, string exportId)
        {
            _logger.LogInformation("[FORENSIC][{ExportId}] ApplySheetLayout ENTER: sheetCount={Count}", exportId, form.Sheets?.Count ?? 0);
            if (form.Sheets == null) return;

            for (int si = 0; si < form.Sheets.Count; si++)
            {
                var sheetDef = form.Sheets[si];
                if (!sheetIndexMap.TryGetValue(sheetDef.Id, out int comIdx))
                {
                    _logger.LogWarning("[FORENSIC][{ExportId}] ApplySheetLayout: sheet {Id} not in map, skipping", exportId, sheetDef.Id);
                    continue;
                }

                _logger.LogInformation("[FORENSIC][{ExportId}] ApplySheetLayout: Processing sheet {Id} at COM index {Idx}", exportId, sheetDef.Id, comIdx);
                var ws = (Excel.Worksheet)workbook.Worksheets[comIdx];
                _logger.LogInformation("[FORENSIC][{ExportId}]   ws.Name={Name}", exportId, ws.Name);

                if (sheetDef.PageSettings != null)
                {
                    _logger.LogInformation("[FORENSIC][{ExportId}]   Calling ApplyPageSettings() orientation={Ori}",
                        exportId, sheetDef.PageSettings.Orientation);
                    ApplyPageSettings(ws, sheetDef.PageSettings);
                    _logger.LogInformation("[FORENSIC][{ExportId}]   ApplyPageSettings() SUCCESS", exportId);
                }

                // PrintArea is handled by SetDefinedNames() only

                if (sheetDef.ColumnWidths != null && sheetDef.ColumnWidths.Count > 0)
                {
                    _logger.LogInformation("[FORENSIC][{ExportId}]   Applying {Count} column widths", exportId, sheetDef.ColumnWidths.Count);
                    ApplyColumnWidths(ws, sheetDef.ColumnWidths);
                }

                if (sheetDef.RowHeights != null && sheetDef.RowHeights.Count > 0)
                {
                    _logger.LogInformation("[FORENSIC][{ExportId}]   Applying {Count} row heights", exportId, sheetDef.RowHeights.Count);
                    ApplyRowHeights(ws, sheetDef.RowHeights);
                }

                _logger.LogInformation("[FORENSIC][{ExportId}]   Calling ApplyMergedCells()", exportId);
                ApplyMergedCells(ws, sheetDef, form);
                _logger.LogInformation("[FORENSIC][{ExportId}]   ApplyMergedCells() SUCCESS", exportId);

                if (!string.IsNullOrEmpty(sheetDef.FreezePane))
                {
                    _logger.LogInformation("[FORENSIC][{ExportId}]   Calling ApplyFreezePane({Pane})", exportId, sheetDef.FreezePane);
                    ApplyFreezePane(ws, sheetDef.FreezePane);
                }

                if (sheetDef.CellStyles != null && sheetDef.CellStyles.Count > 0)
                {
                    _logger.LogInformation("[FORENSIC][{ExportId}]   Calling ApplyCellStyles() with {Count} styles", exportId, sheetDef.CellStyles.Count);
                    ApplyCellStyles(ws, sheetDef.CellStyles);
                }

                Marshal.ReleaseComObject(ws);
                _logger.LogInformation("[FORENSIC][{ExportId}]   Sheet {Id} done", exportId, sheetDef.Id);
            }
            _logger.LogInformation("[FORENSIC][{ExportId}] ApplySheetLayout EXIT", exportId);
        }

        // ────────────────────────────────────────────────────────────────────────────
        // ApplyPageSettings — FORENSIC LOGGING
        // ────────────────────────────────────────────────────────────────────────────
        private void ApplyPageSettings(Excel.Worksheet ws, PageSettings settings)
        {
            try
            {
                _logger.LogInformation("[FORENSIC]     PageSetup.Orientation={Ori}", settings.Orientation);
                ws.PageSetup.Orientation = string.Equals(settings.Orientation, "landscape",
                    StringComparison.OrdinalIgnoreCase)
                    ? XlPageOrientation.xlLandscape
                    : XlPageOrientation.xlPortrait;

                _logger.LogInformation("[FORENSIC]     PageSetup.LeftMargin={Val}", settings.LeftMargin);
                ws.PageSetup.LeftMargin = settings.LeftMargin;
                _logger.LogInformation("[FORENSIC]     PageSetup.TopMargin={Val}", settings.TopMargin);
                ws.PageSetup.TopMargin = settings.TopMargin;
                _logger.LogInformation("[FORENSIC]     PageSetup.RightMargin={Val}", settings.RightMargin);
                ws.PageSetup.RightMargin = settings.RightMargin;
                _logger.LogInformation("[FORENSIC]     PageSetup.BottomMargin={Val}", settings.BottomMargin);
                ws.PageSetup.BottomMargin = settings.BottomMargin;

                _logger.LogInformation("[FORENSIC]     PageSetup.CenterHorizontally={Val}", settings.CenterHorizontally);
                ws.PageSetup.CenterHorizontally = settings.CenterHorizontally;
                _logger.LogInformation("[FORENSIC]     PageSetup.CenterVertically={Val}", settings.CenterVertically);
                ws.PageSetup.CenterVertically = settings.CenterVertically;

                if (settings.Zoom > 0)
                {
                    _logger.LogInformation("[FORENSIC]     PageSetup.Zoom={Val}", settings.Zoom);
                    ws.PageSetup.Zoom = settings.Zoom;
                }
                if (settings.FitToPagesWide > 0)
                {
                    _logger.LogInformation("[FORENSIC]     PageSetup.FitToPagesWide={Val}", settings.FitToPagesWide);
                    ws.PageSetup.FitToPagesWide = settings.FitToPagesWide;
                }
                if (settings.FitToPagesTall > 0)
                {
                    _logger.LogInformation("[FORENSIC]     PageSetup.FitToPagesTall={Val}", settings.FitToPagesTall);
                    ws.PageSetup.FitToPagesTall = settings.FitToPagesTall;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[FORENSIC] ApplyPageSettings FAILED: {Msg}", ex.Message);
                // NOT throwing — original behavior was to log and continue
            }
        }

        private void ApplyColumnWidths(Excel.Worksheet ws, Dictionary<int, double> columnWidths)
        {
            foreach (var kv in columnWidths)
            {
                try
                {
                    _logger.LogInformation("[FORENSIC]   Setting col={Col} width={Width}", kv.Key, kv.Value);
                    var col = (Excel.Range)ws.Columns[kv.Key];
                    col.ColumnWidth = kv.Value;
                    Marshal.ReleaseComObject(col);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "[FORENSIC] ApplyColumnWidths FAILED col={Col}", kv.Key);
                }
            }
        }

        private void ApplyRowHeights(Excel.Worksheet ws, Dictionary<int, double> rowHeights)
        {
            foreach (var kv in rowHeights)
            {
                try
                {
                    _logger.LogInformation("[FORENSIC]   Setting row={Row} height={Height}", kv.Key, kv.Value);
                    var row = (Excel.Range)ws.Rows[kv.Key];
                    row.RowHeight = kv.Value;
                    Marshal.ReleaseComObject(row);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "[FORENSIC] ApplyRowHeights FAILED row={Row}", kv.Key);
                }
            }
        }

        private void ApplyMergedCells(Excel.Worksheet ws, SheetDefinition sheetDef, FormDefinition form)
        {
            var mergedCellAddresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (sheetDef.MergedCells != null)
            {
                foreach (var mc in sheetDef.MergedCells)
                {
                    if (!string.IsNullOrWhiteSpace(mc.Address) && mergedCellAddresses.Add(mc.Address))
                    {
                        _logger.LogInformation("[FORENSIC]   Merging cell {Addr} from sheetDef", mc.Address);
                        MergeRange(ws, mc.Address);
                    }
                }
            }
            foreach (var cluster in form.Clusters)
            {
                if (cluster.CellAddress != null && cluster.CellAddress.Contains(":") &&
                    mergedCellAddresses.Add(cluster.CellAddress))
                {
                    _logger.LogInformation("[FORENSIC]   Merging cell {Addr} from cluster {Cluster}", cluster.CellAddress, cluster.ClusterId);
                    MergeRange(ws, cluster.CellAddress);
                }
            }

            // Apply thin borders to single-cell field addresses (no colon = not merged)
            // These cells otherwise get no border because MergeRange only handles merged ranges.
            // If ApplyCellStyles runs later with captured border data, it will override this.
            foreach (var cluster in form.Clusters)
            {
                if (cluster.CellAddress != null && !cluster.CellAddress.Contains(":"))
                {
                    if (mergedCellAddresses.Add(cluster.CellAddress))
                    {
                        _logger.LogInformation("[FORENSIC]   Applying border to single cell {Addr} from cluster {Cluster}", cluster.CellAddress, cluster.ClusterId);
                        ApplySingleCellBorder(ws, cluster.CellAddress);
                    }
                }
            }
        }

        private void ApplySingleCellBorder(Excel.Worksheet ws, string address)
        {
            try
            {
                var range = ws.Range[address];
                range.Borders.LineStyle = XlLineStyle.xlContinuous;
                range.Borders.Weight = XlBorderWeight.xlThin;
                Marshal.ReleaseComObject(range);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[FORENSIC] ApplySingleCellBorder FAILED {Addr}: {Msg}", address, ex.Message);
            }
        }

        private void MergeRange(Excel.Worksheet ws, string address)
        {
            try
            {
                _logger.LogInformation("[FORENSIC]     ws.Range[{Addr}].Merge()", address);
                var range = ws.Range[address];
                range.Merge();
                _logger.LogInformation("[FORENSIC]     Merge SUCCESS");

                try
                {
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;
                }
                catch { }

                Marshal.ReleaseComObject(range);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[FORENSIC] MergeRange FAILED {Addr}: {Msg}", address, ex.Message);
            }
        }

        private void ApplyFreezePane(Excel.Worksheet ws, string freezePane)
        {
            try
            {
                _logger.LogInformation("[FORENSIC]     Activating worksheet for freeze pane");
                ws.Activate();
                _logger.LogInformation("[FORENSIC]     ws.Range[{Pane}].Activate()", freezePane);
                var range = ws.Range[freezePane];
                range.Activate();
                _logger.LogInformation("[FORENSIC]     ActiveWindow.FreezePanes=true");
                ws.Application.ActiveWindow.FreezePanes = true;
                _logger.LogInformation("[FORENSIC]     FreezePane SUCCESS");
                Marshal.ReleaseComObject(range);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[FORENSIC] ApplyFreezePane FAILED {Pane}: {Msg}", freezePane, ex.Message);
            }
        }

        private void ApplyCellStyles(Excel.Worksheet ws, Dictionary<string, CellStyleInfo> cellStyles)
        {
            foreach (var kv in cellStyles)
            {
                try
                {
                    _logger.LogInformation("[FORENSIC]   Applying style to cell {Cell}", kv.Key);
                    var range = ws.Range[kv.Key];
                    ApplyStyleToRange(range, kv.Value);
                    Marshal.ReleaseComObject(range);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "[FORENSIC] ApplyCellStyles FAILED cell={Cell}", kv.Key);
                }
            }
        }

        private static void ApplyStyleToRange(Excel.Range range, CellStyleInfo style)
        {
            if (style.FontName != null) range.Font.Name = style.FontName;
            if (style.FontSize.HasValue) range.Font.Size = style.FontSize.Value;
            if (style.Bold.HasValue) range.Font.Bold = style.Bold.Value;
            if (style.Italic.HasValue) range.Font.Italic = style.Italic.Value;
            if (style.Underline.HasValue && style.Underline.Value) range.Font.Underline = true;
            if (style.Color != null)
            {
                try { var (r, g, b) = ParseColor(style.Color); range.Font.Color = System.Drawing.Color.FromArgb(r, g, b); } catch { }
            }
            if (style.FillColor != null)
            {
                try { var (r, g, b) = ParseColor(style.FillColor); range.Interior.Color = System.Drawing.Color.FromArgb(r, g, b); } catch { }
            }
            if (style.HorizontalAlignment != null)
            {
                range.HorizontalAlignment = style.HorizontalAlignment switch
                {
                    "left" => XlHAlign.xlHAlignLeft, "center" => XlHAlign.xlHAlignCenter,
                    "right" => XlHAlign.xlHAlignRight, _ => XlHAlign.xlHAlignGeneral
                };
            }
            if (style.VerticalAlignment != null)
            {
                range.VerticalAlignment = style.VerticalAlignment switch
                {
                    "top" => XlVAlign.xlVAlignTop, "center" => XlVAlign.xlVAlignCenter,
                    "bottom" => XlVAlign.xlVAlignBottom, _ => XlVAlign.xlVAlignBottom
                };
            }
            if (style.WrapText.HasValue) range.WrapText = style.WrapText.Value;

            // Borders — apply each edge from CSS-style border strings
            ApplyBorderEdge(range, XlBordersIndex.xlEdgeTop, style.BorderTop);
            ApplyBorderEdge(range, XlBordersIndex.xlEdgeBottom, style.BorderBottom);
            ApplyBorderEdge(range, XlBordersIndex.xlEdgeLeft, style.BorderLeft);
            ApplyBorderEdge(range, XlBordersIndex.xlEdgeRight, style.BorderRight);
        }

        private static void ApplyBorderEdge(Excel.Range range, XlBordersIndex edge, string? borderCss)
        {
            if (string.IsNullOrEmpty(borderCss)) return;

            try
            {
                // Parse format: "{width} {lineStyle} {color}" e.g. "1px solid #000000"
                var parts = borderCss.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 2) return;

                string widthStr = parts[0]; // "1px", "2px", "3px"
                string lineStyleStr = parts[1]; // "solid", "dashed", "dotted", "double"
                string? colorStr = parts.Length >= 3 ? parts[2] : null;

                var border = range.Borders[edge];

                border.LineStyle = lineStyleStr switch
                {
                    "double" => XlLineStyle.xlDouble,
                    "dashed" => XlLineStyle.xlDash,
                    "dotted" => XlLineStyle.xlDot,
                    _ => XlLineStyle.xlContinuous
                };

                border.Weight = widthStr switch
                {
                    "2px" => XlBorderWeight.xlMedium,
                    "3px" => XlBorderWeight.xlThick,
                    _ => XlBorderWeight.xlThin
                };

                if (colorStr != null)
                {
                    var (r, g, b) = ParseColor(colorStr);
                    border.Color = System.Drawing.Color.FromArgb(r, g, b);
                }

                Marshal.ReleaseComObject(border);
            }
            catch { }
        }

        private static (int r, int g, int b) ParseColor(string color)
        {
            if (color.StartsWith("#")) color = color[1..];
            if (color.Length == 6 && int.TryParse(color, System.Globalization.NumberStyles.HexNumber, null, out var rgb))
                return ((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
            return (0, 0, 0);
        }

        // ────────────────────────────────────────────────────────────────────────────
        // WriteCellValues — FORENSIC LOGGING
        // ────────────────────────────────────────────────────────────────────────────
        private void WriteCellValues(Workbook workbook, FormDefinition form,
            Dictionary<string, int> sheetIndexMap, string exportId)
        {
            _logger.LogInformation("[FORENSIC][{ExportId}] WriteCellValues ENTER", exportId);
            if (form.Sheets == null) return;

            foreach (var sheetDef in form.Sheets)
            {
                if (!sheetIndexMap.TryGetValue(sheetDef.Id, out int comIdx))
                {
                    _logger.LogWarning("[FORENSIC][{ExportId}] WriteCellValues: sheet {Id} not in map", exportId, sheetDef.Id);
                    continue;
                }

                var ws = (Excel.Worksheet)workbook.Worksheets[comIdx];
                _logger.LogInformation("[FORENSIC][{ExportId}]   Writing values to sheet {Name} (COM idx {Idx})", exportId, ws.Name, comIdx);

                if (sheetDef.CellValues != null)
                {
                    foreach (var kv in sheetDef.CellValues)
                    {
                        if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                        try
                        {
                            _logger.LogInformation("[FORENSIC][{ExportId}]     ws.Range[{Cell}].Value = {Val}", exportId, kv.Key, kv.Value);
                            var range = ws.Range[kv.Key];
                            range.Value = kv.Value;
                            Marshal.ReleaseComObject(range);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "[FORENSIC][{ExportId}] WriteCellValues FAILED cell={Cell}: {Msg}", exportId, kv.Key, ex.Message);
                        }
                    }
                }

                // Ensure cells exist at all cluster addresses even when no value was written
                // This guarantees ApplyCellStyles (which runs earlier, in ApplySheetLayout) has
                // a physical cell to apply styles to. For cells without a value, accessing the
                // Range via COM creates the cell in the OOXML.
                if (form.Clusters != null)
                {
                    foreach (var cluster in form.Clusters)
                    {
                        if (cluster.CellAddress == null) continue;
                        string cellAddr = cluster.CellAddress.Contains(":")
                            ? cluster.CellAddress.Split(':')[0]
                            : cluster.CellAddress;
                        if (string.IsNullOrWhiteSpace(cellAddr)) continue;
                        if (sheetDef.CellValues != null && sheetDef.CellValues.ContainsKey(cellAddr)) continue;
                        try
                        {
                            _logger.LogInformation("[FORENSIC][{ExportId}]     ws.Range[{Cell}].Value = (empty) — ensuring cell exists", exportId, cellAddr);
                            var range = ws.Range[cellAddr];
                            Marshal.ReleaseComObject(range);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "[FORENSIC][{ExportId}] WriteCellValues (ensure cell) FAILED {Cell}: {Msg}", exportId, cellAddr, ex.Message);
                        }
                    }
                }

                Marshal.ReleaseComObject(ws);
            }
            _logger.LogInformation("[FORENSIC][{ExportId}] WriteCellValues EXIT", exportId);
        }

        // ────────────────────────────────────────────────────────────────────────────
        // WriteCellComments — FORENSIC LOGGING
        // ────────────────────────────────────────────────────────────────────────────
        private void WriteCellComments(Workbook workbook, FormDefinition form,
            Dictionary<string, int> sheetIndexMap, string exportId)
        {
            _logger.LogInformation("[FORENSIC][{ExportId}] WriteCellComments ENTER: {Count} clusters", exportId, form.Clusters?.Count ?? 0);
            if (form.Clusters == null || form.Clusters.Count == 0) return;

            for (int clusterIdx = 0; clusterIdx < form.Clusters.Count; clusterIdx++)
            {
                var cluster = form.Clusters[clusterIdx];
                if (string.IsNullOrWhiteSpace(cluster.CellAddress)) continue;
                if (!sheetIndexMap.TryGetValue(cluster.SheetId, out int comIdx)) continue;

                try
                {
                    var ws = (Excel.Worksheet)workbook.Worksheets[comIdx];
                    string singleCellAddr = cluster.CellAddress.Contains(":")
                        ? cluster.CellAddress.Split(':')[0]
                        : cluster.CellAddress;

                    _logger.LogInformation("[FORENSIC][{ExportId}]   ws.Range[{Addr}].AddComment() for cluster {Idx}",
                        exportId, singleCellAddr, clusterIdx);

                    var commentRange = ws.Range[singleCellAddr];

                    var commentLines = new List<string>
                    {
                        cluster.Name ?? "",
                        cluster.Type ?? "KeyboardText",
                        clusterIdx.ToString(),
                        "0", "0"
                    };

                    string paramStr = "";
                    if (cluster.InputParameters != null && cluster.InputParameters.Count > 0)
                    {
                        var paramParts = new List<string>(cluster.InputParameters.Count);
                        foreach (var kv in cluster.InputParameters)
                            paramParts.Add($"{kv.Key}={kv.Value}");
                        paramStr = string.Join(";", paramParts);
                    }
                    else
                    {
                        paramStr = "Required=0;Lines=1;FontSize=9;FontAutoResizeMode=0";
                    }
                    commentLines.Add(paramStr);
                    for (int i = 0; i < 15; i++) commentLines.Add("");

                    string commentText = string.Join("\n", commentLines);

                    var comment = commentRange.AddComment(commentText);
                    comment.Visible = false;
                    _logger.LogInformation("[FORENSIC][{ExportId}]   AddComment SUCCESS for {Addr}", exportId, singleCellAddr);

                    Marshal.ReleaseComObject(comment);
                    Marshal.ReleaseComObject(commentRange);
                    Marshal.ReleaseComObject(ws);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "[FORENSIC][{ExportId}] WriteCellComments FAILED cluster {Idx} addr={Addr}: {Msg}",
                        exportId, clusterIdx, cluster.CellAddress, ex.Message);
                }
            }
            _logger.LogInformation("[FORENSIC][{ExportId}] WriteCellComments EXIT: {Count} clusters processed", exportId, form.Clusters.Count);
        }
        // ────────────────────────────────────────────────────────────────────────────
        // PopulateFieldsWorksheet — FORENSIC LOGGING
        // ────────────────────────────────────────────────────────────────────────────
        private void PopulateFieldsWorksheet(Workbook workbook, FormDefinition form,
            Dictionary<string, int> sheetIndexMap, string exportId)
        {
            _logger.LogInformation("[FORENSIC][{ExportId}] PopulateFieldsWorksheet ENTER: {Count} clusters", exportId, form.Clusters?.Count ?? 0);
            if (form.Clusters == null || form.Clusters.Count == 0) return;
            if (!sheetIndexMap.TryGetValue("_Fields", out int comIdx)) return;

            try
            {
                var ws = (Excel.Worksheet)workbook.Worksheets[comIdx];
                _logger.LogInformation("[FORENSIC][{ExportId}]   _Fields sheet found at COM idx {Idx}, name={Name}", exportId, comIdx, ws.Name);

                _logger.LogInformation("[FORENSIC][{ExportId}]   Calling ws.Cells.ClearContents()", exportId);
                ws.Cells.ClearContents();
                _logger.LogInformation("[FORENSIC][{ExportId}]   ClearContents SUCCESS", exportId);

                var sheetIdToName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (form.Sheets != null)
                    foreach (var s in form.Sheets) sheetIdToName[s.Id] = s.Name;

                string[] headers = { "Address", "FieldId", "FieldName", "FieldType", "SheetName", "CreatedDate", "Notes" };
                for (int col = 0; col < headers.Length; col++)
                {
                    _logger.LogInformation("[FORENSIC][{ExportId}]   Writing header cell [{Row},{Col}]: {Val}", exportId, 1, col + 1, headers[col]);
                    var headerCell = (Excel.Range)ws.Cells[1, col + 1];
                    headerCell.Value = headers[col];
                    headerCell.Font.Bold = true;
                    Marshal.ReleaseComObject(headerCell);
                }

                string nowStr = DateTime.Now.ToString("o");
                for (int i = 0; i < form.Clusters.Count; i++)
                {
                    var cluster = form.Clusters[i];
                    int row = i + 2;
                    _logger.LogInformation("[FORENSIC][{ExportId}]   Writing field row {Row}: addr={Addr} id={Id} name={Name} type={Type}",
                        exportId, row, cluster.CellAddress, cluster.ClusterId, cluster.Name, cluster.Type);
                    ws.Cells[row, 1] = cluster.CellAddress ?? "";
                    ws.Cells[row, 2] = cluster.ClusterId ?? i.ToString();
                    ws.Cells[row, 3] = cluster.Name ?? "";
                    ws.Cells[row, 4] = cluster.Type ?? "";
                    ws.Cells[row, 5] = sheetIdToName.GetValueOrDefault(cluster.SheetId, "");
                    ws.Cells[row, 6] = nowStr;
                    ws.Cells[row, 7] = cluster.Remarks ?? "";
                }

                _logger.LogInformation("[FORENSIC][{ExportId}]   Setting _Fields sheet visibility to xlSheetHidden", exportId);
                ws.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                _logger.LogInformation("[FORENSIC][{ExportId}]   _Fields hidden SUCCESS", exportId);

                Marshal.ReleaseComObject(ws);
                _logger.LogInformation("[FORENSIC][{ExportId}] PopulateFieldsWorksheet EXIT: {Count} rows", exportId, form.Clusters.Count);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[FORENSIC][{ExportId}] PopulateFieldsWorksheet FAILED: {Msg}", exportId, ex.Message);
                // NOT throwing — original behavior was to log and continue
            }
        }

        // ────────────────────────────────────────────────────────────────────────────
        // CreateExcelOutputSetting — FORENSIC LOGGING
        // ────────────────────────────────────────────────────────────────────────────
        private void CreateExcelOutputSetting(Workbook workbook, FormDefinition form,
            Dictionary<string, int> sheetIndexMap, string exportId)
        {
            _logger.LogInformation("[FORENSIC][{ExportId}] CreateExcelOutputSetting ENTER", exportId);
            if (!sheetIndexMap.TryGetValue("ExcelOutputSetting", out int comIdx)) return;

            try
            {
                var ws = (Excel.Worksheet)workbook.Worksheets[comIdx];
                _logger.LogInformation("[FORENSIC][{ExportId}]   ExcelOutputSetting sheet at COM idx {Idx}, name={Name}", exportId, comIdx, ws.Name);

                _logger.LogInformation("[FORENSIC][{ExportId}]   Calling ws.Cells.ClearContents()", exportId);
                ws.Cells.ClearContents();
                _logger.LogInformation("[FORENSIC][{ExportId}]   ClearContents SUCCESS", exportId);

                string[] xmlFragments = GenerateExcelOutputSettingXmlFragments(form);
                _logger.LogInformation("[FORENSIC][{ExportId}]   Generated {Count} XML fragments", exportId, xmlFragments.Length);

                for (int i = 0; i < xmlFragments.Length; i++)
                {
                    string val = xmlFragments[i];
                    string preview = val.Length > 80 ? val[..80] + "..." : val;
                    int rowNum = i + 1;
                    int valLen = val.Length;
                    _logger.LogInformation("[FORENSIC][{ExportId}]     Writing cell A{Row}: len={Len} preview={Preview}",
                        exportId, rowNum, valLen, preview);
                    ws.Cells[i + 1, 1] = val;
                }
                _logger.LogInformation("[FORENSIC][{ExportId}]   All fragments written SUCCESS", exportId);

                Marshal.ReleaseComObject(ws);
                _logger.LogInformation("[FORENSIC][{ExportId}] CreateExcelOutputSetting EXIT", exportId);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[FORENSIC][{ExportId}] CreateExcelOutputSetting FAILED: {Msg}", exportId, ex.Message);
                // NOT throwing — original behavior was to log and continue
            }
        }

        private string[] GenerateExcelOutputSettingXmlFragments(FormDefinition form)
        {
            int sheetCount = form.Sheets?.Count ?? 1;
            int clusterCount = form.Clusters?.Count ?? 0;

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

        // ────────────────────────────────────────────────────────────────────────────
        // SetDefinedNames — FORENSIC LOGGING (EVERY NAME OPERATION INSTRUMENTED)
        // ────────────────────────────────────────────────────────────────────────────
        private void SetDefinedNames(Workbook workbook, FormDefinition form, string printAreaAddress, string exportId)
        {
            _logger.LogInformation("[FORENSIC][{ExportId}] SetDefinedNames ENTER: printAreaAddress=\"{Addr}\"", exportId, printAreaAddress);
            if (string.IsNullOrEmpty(printAreaAddress))
            {
                _logger.LogInformation("[FORENSIC][{ExportId}] SetDefinedNames: printAreaAddress empty, skipping", exportId);
                return;
            }

            try
            {
                string sheetName = "Sheet1";
                if (form.Sheets != null && form.Sheets.Count > 0)
                    sheetName = form.Sheets[0].Name;

                string escapedSheetName = sheetName.Contains(" ") ? $"'{sheetName}'" : sheetName;
                string printAreaRef = $"{escapedSheetName}!{printAreaAddress}";
                string printAreaFormula = $"={printAreaRef}";

                // OPERATION 1: Access workbook.Names
                Names names = null!;
                bool gotNames = NameOp(exportId, $"workbook.Names getter", () => { names = workbook.Names; });
                if (!gotNames)
                {
                    _logger.LogError("[FORENSIC][{ExportId}] SetDefinedNames: ABORT — cannot access workbook.Names", exportId);
                    return;
                }

                // OPERATION 2: Enumerate all names
                try
                {
                    int nameCount = names.Count;
                    _logger.LogInformation("[FORENSIC][{ExportId}]   workbook.Names.Count = {Count}", exportId, nameCount);
                    if (nameCount > 0)
                    {
                        _logger.LogInformation("[FORENSIC][{ExportId}]   Enumerating all existing workbook names:", exportId);
                        foreach (Name n in names)
                        {
                            try
                            {
                                string nName = n.Name?.ToString() ?? "(null)";
                                string nRefersTo = n.RefersTo?.ToString() ?? "(null)";
                                _logger.LogInformation("[FORENSIC][{ExportId}]     Name: '{Name}' RefersTo: '{Ref}'",
                                    exportId, nName, nRefersTo);
                            }
                            catch (Exception exInner)
                            {
                                _logger.LogWarning("[FORENSIC][{ExportId}]     Failed to read name: {Msg}", exportId, exInner.Message);
                            }
                            finally
                            {
                                try { Marshal.ReleaseComObject(n); } catch { }
                            }
                        }
                    }
                }
                catch (Exception exEnum)
                {
                    _logger.LogWarning("[FORENSIC][{ExportId}]   Failed to enumerate names: {Msg}", exportId, exEnum.Message);
                }

                // OPERATION 3: Names.Item("_xlnm.Print_Area")
                _logger.LogInformation("[FORENSIC][{ExportId}]   >>> NAME-OP: Names.Item(\"_xlnm.Print_Area\")", exportId);
                bool exists = false;
                try
                {
                    var existingName = names.Item("_xlnm.Print_Area") as Name;
                    if (existingName != null)
                    {
                        string itemName = existingName.Name?.ToString() ?? "(null)";
                        string itemRefersTo = existingName.RefersTo?.ToString() ?? "(null)";
                        _logger.LogInformation("[FORENSIC][{ExportId}]     Names.Item() FOUND: Name='{Name}' RefersTo='{Ref}'",
                            exportId, itemName, itemRefersTo);

                        // OPERATION 4: Name.RefersTo assignment
                        bool refSet = NameOp(exportId,
                            $"existingName.RefersTo = \"{printAreaFormula}\"",
                            () => { existingName.RefersTo = printAreaFormula; });

                        if (refSet)
                        {
                            _logger.LogInformation("[FORENSIC][{ExportId}]     RefersTo update SUCCESS", exportId);
                            exists = true;
                        }
                        else
                        {
                            _logger.LogWarning("[FORENSIC][{ExportId}]     RefersTo update FAILED — will try Names.Add() instead", exportId);
                        }

                        Marshal.ReleaseComObject(existingName);
                    }
                    else
                    {
                        _logger.LogInformation("[FORENSIC][{ExportId}]     Names.Item() returned null (name does not exist)", exportId);
                    }
                }
                catch (Exception exItem)
                {
                    _logger.LogWarning(exItem, "[FORENSIC][{ExportId}]     Names.Item() THREW: {Msg}", exportId, exItem.Message);
                    // Name doesn't exist, will create below
                }

                // OPERATION 5: Names.Add("_xlnm.Print_Area", ...) — THE CRITICAL OPERATION
                if (!exists)
                {
                    _logger.LogInformation("[FORENSIC][{ExportId}]   >>> ATTEMPTING: Names.Add(\"_xlnm.Print_Area\", \"{Formula}\", false, ...)",
                        exportId, printAreaFormula);

                    bool addSuccess = NameOp(exportId,
                        $"Names.Add(\"_xlnm.Print_Area\", \"{printAreaFormula}\", false, ...)",
                        () =>
                        {
                            var printAreaName = names.Add(
                                "_xlnm.Print_Area",
                                printAreaFormula,
                                false,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing);
                            try { Marshal.ReleaseComObject(printAreaName); } catch { }
                        });

                    if (addSuccess)
                    {
                        _logger.LogInformation("[FORENSIC][{ExportId}]     Names.Add() SUCCESS — _xlnm.Print_Area created", exportId);
                    }
                    else
                    {
                        _logger.LogError("[FORENSIC][{ExportId}]     Names.Add() FAILED — see NAME-OP log above", exportId);
                        // Dump current state for analysis
                        DumpNames(workbook, exportId);
                    }
                }

                Marshal.ReleaseComObject(names);
                _logger.LogInformation("[FORENSIC][{ExportId}] SetDefinedNames EXIT", exportId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[FORENSIC][{ExportId}] SetDefinedNames OUTER CATCH: {Msg}\nStack: {Stack}",
                    exportId, ex.Message, ex.StackTrace);
                // NOT throwing — original behavior was to log and continue.
            }
        }
    }
}
