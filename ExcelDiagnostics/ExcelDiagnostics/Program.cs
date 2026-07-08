using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDiagnostics;

class Program
{
    static int _step;
    static int _errorCount;

    static void Main(string[] args)
    {
        string workingFile = @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Text by HandWriting.xlsx";
        string problemFile = @"C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest.xlsx";

        Console.WriteLine("====================================================================================================");
        Console.WriteLine("  EXCEL WORKBOOK DIAGNOSTIC - SIDE-BY-SIDE COMPARISON");
        Console.WriteLine("====================================================================================================");
        Console.WriteLine();

        AnalyzeWorkbook("WORKING", workingFile);

        Console.WriteLine();
        Console.WriteLine("----------------------------------------------------------------------------------------------------");
        Console.WriteLine();

        AnalyzeWorkbook("PROBLEM", problemFile);

        Console.WriteLine();
        Console.WriteLine("====================================================================================================");
        Console.WriteLine("  DIAGNOSTIC COMPLETE - Total errors encountered: {0}", _errorCount);
        Console.WriteLine("====================================================================================================");
    }

    static void AnalyzeWorkbook(string label, string filePath)
    {
        _step = 0;

        Console.WriteLine("=====  ANALYZING: {0}  =====", label);
        Console.WriteLine("  File: {0}", filePath);
        Console.WriteLine();

        if (!File.Exists(filePath))
        {
            Step("FILE CHECK", "FAIL - File not found: {0}", filePath);
            _errorCount++;
            return;
        }

        Step("FILE CHECK", "OK - File exists. Size: {0:N0} bytes", new FileInfo(filePath).Length);

        Excel.Application? excelApp = null;
        Excel.Workbook? workbook = null;

        try
        {
            // --- Launch Excel ---
            Step("EXCEL LAUNCH", "Starting Excel application (hidden)...");
            excelApp = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };
            Step("EXCEL LAUNCH", "OK - Excel started.");

            // --- Open workbook ---
            Step("WORKBOOK OPEN", "Opening: {0}", filePath);
            workbook = excelApp.Workbooks.Open(filePath);
            Step("WORKBOOK OPEN", "OK - Opened. Name: \"{0}\"", workbook.Name);

            // Log file format
            string format = workbook.FileFormat switch
            {
                Excel.XlFileFormat.xlOpenXMLWorkbook => ".xlsx (Open XML)",
                Excel.XlFileFormat.xlWorkbookNormal => ".xls (Excel 97-2003)",
                Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled => ".xlsm (Macro-enabled)",
                _ => string.Format("Unknown (code: {0})", (int)workbook.FileFormat)
            };
            Step("WORKBOOK FORMAT", "Format: {0}", format);

            int sheetCount = workbook.Sheets.Count;
            Step("SHEET COUNT", "Sheets: {0}", sheetCount);

            // ===================================================================
            // ENUMERATE ALL WORKSHEETS
            // ===================================================================
            Step("WORKSHEETS", "Enumerating {0} worksheet(s)...", workbook.Worksheets.Count);
            Console.WriteLine();

            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                try
                {
                    Console.WriteLine("  --- Sheet #{0}: \"{1}\" ---", ws.Index, ws.Name);
                    Console.WriteLine();

                    // Visibility
                    string visibility = ws.Visible switch
                    {
                        Excel.XlSheetVisibility.xlSheetVisible => "Visible",
                        Excel.XlSheetVisibility.xlSheetHidden => "Hidden",
                        Excel.XlSheetVisibility.xlSheetVeryHidden => "VeryHidden",
                        _ => string.Format("Unknown ({0})", ws.Visible)
                    };
                    Console.WriteLine("    Visibility:   {0}", visibility);

                    // UsedRange
                    try
                    {
                        var ur = ws.UsedRange;
                        Console.WriteLine("    UsedRange:    {0}", ur.AddressLocal);
                        Marshal.ReleaseComObject(ur);
                    }
                    catch (Exception ex) { Console.WriteLine("    UsedRange:    ERROR: {0}", ex.Message); }

                    Console.WriteLine();
                    Console.WriteLine("    -- Print Settings --");

                    // PrintArea
                    string? printArea = null;
                    try { printArea = ws.PageSetup.PrintArea; } catch (Exception ex) { Console.WriteLine("    PageSetup.PrintArea: ERROR: {0}", ex.Message); }
                    Console.WriteLine("    PrintArea:            \"{0}\"", printArea ?? "(null)");

                    try { Console.WriteLine("    PrintTitleRows:       \"{0}\"", ws.PageSetup.PrintTitleRows ?? "(null)"); } catch { }
                    try { Console.WriteLine("    PrintTitleColumns:    \"{0}\"", ws.PageSetup.PrintTitleColumns ?? "(null)"); } catch { }
                    try { Console.WriteLine("    Zoom:                 {0}", ws.PageSetup.Zoom); } catch { }
                    try { Console.WriteLine("    FitToPagesWide:       {0}", ws.PageSetup.FitToPagesWide); } catch { }
                    try { Console.WriteLine("    FitToPagesTall:       {0}", ws.PageSetup.FitToPagesTall); } catch { }
                    try { Console.WriteLine("    Orientation:          {0}", ws.PageSetup.Orientation); } catch { }
                    try { Console.WriteLine("    PaperSize:            {0}", ws.PageSetup.PaperSize); } catch { }

                    Console.WriteLine();
                    Console.WriteLine("    -- Comments / Fields --");

                    // Comments
                    try
                    {
                        var commentedCells = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeComments);
                        int commentCount = commentedCells.Count;
                        Console.WriteLine("    Cells with comments:  {0}", commentCount);

                        // Extract sample field coordinates from first commented cell
                        try
                        {
                            foreach (Excel.Range c in commentedCells)
                            {
                                try
                                {
                                    string cellAddr = c.AddressLocal[false, false];
                                    string? commentText = null;
                                    try { commentText = c.Comment?.Text(); } catch { }

                                    Console.WriteLine("      Cell: {0}", cellAddr);
                                    Console.WriteLine("        Comment: \"{0}\"", commentText ?? "(null)");
                                    Console.WriteLine("        Left:    {0}", c.Left);
                                    Console.WriteLine("        Top:     {0}", c.Top);
                                    Console.WriteLine("        Width:   {0}", c.Width);
                                    Console.WriteLine("        Height:  {0}", c.Height);

                                    if (commentText != null && commentText.IndexOf("Type=", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        Console.WriteLine("        -> Contains field type definition");
                                    }
                                }
                                finally { Marshal.ReleaseComObject(c); }
                                break; // Only first cell
                            }
                        }
                        catch (Exception ex) { Console.WriteLine("    Sample field extraction error: {0}", ex.Message); }
                        finally { Marshal.ReleaseComObject(commentedCells); }
                    }
                    catch
                    {
                        Console.WriteLine("    Cells with comments:  0 (none found)");
                    }

                    Console.WriteLine();
                    Console.WriteLine("    -- Worksheet-Level Names ({0}) --", ws.Names.Count);
                    if (ws.Names.Count > 0)
                    {
                        foreach (Excel.Name name in ws.Names)
                        {
                            try
                            {
                                string? refersTo = name.RefersTo as string;
                                string? refersToLocal = null;
                                try { refersToLocal = name.RefersToLocal as string; } catch { }
                                string? rangeAddr = null;
                                try
                                {
                                    var range = name.RefersToRange;
                                    if (range != null)
                                    {
                                        rangeAddr = range.AddressLocal;
                                        Marshal.ReleaseComObject(range);
                                    }
                                }
                                catch { }

                                Console.WriteLine("      \"{0}\"", name.Name);
                                Console.WriteLine("        RefersTo:       \"{0}\"", refersTo ?? "(null)");
                                Console.WriteLine("        RefersToLocal:  \"{0}\"", refersToLocal ?? "(null)");
                                Console.WriteLine("        Resolves to:    {0}", rangeAddr ?? "(unavailable)");
                                try { Console.WriteLine("        Visible:        {0}", name.Visible); } catch { }
                            }
                            finally { Marshal.ReleaseComObject(name); }
                        }
                    }
                    else
                    {
                        Console.WriteLine("      (none)");
                    }

                    Console.WriteLine();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  ERROR reading worksheet: {0}", ex.Message);
                    _errorCount++;
                }
                finally { Marshal.ReleaseComObject(ws); }
            }

            // ===================================================================
            // WORKBOOK-LEVEL NAMES
            // ===================================================================
            Console.WriteLine("  --- Workbook-Level Names ({0}) ---", workbook.Names.Count);
            Console.WriteLine();

            if (workbook.Names.Count > 0)
            {
                foreach (Excel.Name name in workbook.Names)
                {
                    try
                    {
                        string nameValue = name.Name ?? "";
                        string? refersTo = name.RefersTo as string;
                        string? refersToLocal = null;
                        try { refersToLocal = name.RefersToLocal as string; } catch { }
                        string? rangeAddr = null;
                        try
                        {
                            var range = name.RefersToRange;
                            if (range != null)
                            {
                                rangeAddr = range.AddressLocal;
                                Marshal.ReleaseComObject(range);
                            }
                        }
                        catch { }

                        Console.WriteLine("    \"{0}\"", nameValue);
                        Console.WriteLine("      RefersTo:       \"{0}\"", refersTo ?? "(null)");
                        Console.WriteLine("      RefersToLocal:  \"{0}\"", refersToLocal ?? "(null)");
                        Console.WriteLine("      Resolves to:    {0}", rangeAddr ?? "(unavailable)");
                        try { Console.WriteLine("      Visible:        {0}", name.Visible); } catch { }

                        // If this is a Print_Area name, flag it
                        if (nameValue.IndexOf("Print_Area", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Console.WriteLine("      *** THIS IS A Print_Area NAME ***");
                            if (!string.IsNullOrEmpty(refersTo) && !string.IsNullOrEmpty(rangeAddr))
                            {
                                Console.WriteLine("      *** Print_Area resolves OK: {0} ***", rangeAddr);
                            }
                            else
                            {
                                Console.WriteLine("      *** Print_Area resolution FAILED! ***");
                                _errorCount++;
                            }
                        }

                        Console.WriteLine();
                    }
                    finally { Marshal.ReleaseComObject(name); }
                }
            }
            else
            {
                Console.WriteLine("    (none)");
                Console.WriteLine();
            }

            // ===================================================================
            // DETECTION METHOD TESTS
            // ===================================================================
            Console.WriteLine("  --- Detection Method Tests ---");
            Console.WriteLine();

            // Method 1: PageSetup.PrintArea on first sheet
            Excel.Worksheet? sheet1 = null;
            try
            {
                sheet1 = (Excel.Worksheet)workbook.Worksheets[1];
                Step("METHOD 1", "worksheet.PageSetup.PrintArea on \"{0}\"", sheet1.Name);
                string? pa = sheet1.PageSetup.PrintArea;
                if (!string.IsNullOrEmpty(pa))
                {
                    Console.WriteLine("    RESULT: SUCCESS - PrintArea = \"{0}\"", pa);
                }
                else
                {
                    Console.WriteLine("    RESULT: FAILED - PrintArea is null/empty");
                }
            }
            finally { if (sheet1 != null) Marshal.ReleaseComObject(sheet1); }

            // Method 2: Search workbook Names for Print_Area
            Step("METHOD 2", "Search workbook-level Names for Print_Area...");
            bool foundMethod2 = false;
            foreach (Excel.Name name in workbook.Names)
            {
                try
                {
                    string nv = name.Name ?? "";
                    if (nv.IndexOf("Print_Area", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        foundMethod2 = true;
                        string? refersTo = name.RefersTo as string;
                        string? rangeAddr = null;
                        try
                        {
                            var r = name.RefersToRange;
                            if (r != null) { rangeAddr = r.AddressLocal; Marshal.ReleaseComObject(r); }
                        }
                        catch { }
                        Console.WriteLine("    Found: \"{0}\" -> RefersTo = \"{1}\"", nv, refersTo ?? "(null)");
                        Console.WriteLine("    Resolves to: {0}", rangeAddr ?? "(unavailable)");
                    }
                }
                finally { Marshal.ReleaseComObject(name); }
            }
            if (!foundMethod2) Console.WriteLine("    (no Print_Area found in workbook Names)");

            // Method 3: Search worksheet Names for Print_Area
            Step("METHOD 3", "Search worksheet-level Names for Print_Area...");
            bool foundMethod3 = false;
            try
            {
                var ws1 = (Excel.Worksheet)workbook.Worksheets[1];
                foreach (Excel.Name name in ws1.Names)
                {
                    try
                    {
                        string nv = name.Name ?? "";
                        if (nv.IndexOf("Print_Area", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            foundMethod3 = true;
                            string? refersTo = name.RefersTo as string;
                            string? rangeAddr = null;
                            try
                            {
                                var r = name.RefersToRange;
                                if (r != null) { rangeAddr = r.AddressLocal; Marshal.ReleaseComObject(r); }
                            }
                            catch { }
                            Console.WriteLine("    Found: \"{0}\" -> RefersTo = \"{1}\"", nv, refersTo ?? "(null)");
                            Console.WriteLine("    Resolves to: {0}", rangeAddr ?? "(unavailable)");
                        }
                    }
                    finally { Marshal.ReleaseComObject(name); }
                }
                Marshal.ReleaseComObject(ws1);
            }
            catch (Exception ex)
            {
                Console.WriteLine("    ERROR: {0}", ex.Message);
            }
            if (!foundMethod3) Console.WriteLine("    (no Print_Area found in worksheet Names)");

            // Method 4: Activate + Calculate + re-read PrintArea
            Step("METHOD 4", "Activate + Calculate + re-read PrintArea...");
            try
            {
                var m4Sheet = (Excel.Worksheet)workbook.Worksheets[1];
                m4Sheet.Activate();
                excelApp.Calculate();
                excelApp.CalculateFull();
                string? pa4 = m4Sheet.PageSetup.PrintArea;
                if (!string.IsNullOrEmpty(pa4))
                {
                    Console.WriteLine("    RESULT: SUCCESS - PrintArea after refresh = \"{0}\"", pa4);
                }
                else
                {
                    Console.WriteLine("    RESULT: Still empty after refresh");
                }
                Marshal.ReleaseComObject(m4Sheet);
            }
            catch (Exception ex)
            {
                Console.WriteLine("    ERROR: {0}", ex.Message);
            }

            // ===================================================================
            // EXPORT TEST
            // ===================================================================
            Console.WriteLine();
            Console.WriteLine("  --- ExportAsFixedFormat Test ---");
            Console.WriteLine();

            Step("EXPORT TEST", "Attempting ExportAsFixedFormat(PDF, IgnorePrintAreas: false)...");
            try
            {
                var exportSheet = (Excel.Worksheet)workbook.Worksheets[1];
                string tempPdf = Path.Combine(Path.GetTempPath(), string.Format("diag_{0:N}.pdf", Guid.NewGuid()));
                try
                {
                    exportSheet.ExportAsFixedFormat(
                        Type: Excel.XlFixedFormatType.xlTypePDF,
                        Filename: tempPdf,
                        Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
                        IncludeDocProperties: false,
                        IgnorePrintAreas: false,
                        OpenAfterPublish: false);
                    Console.WriteLine("    ExportAsFixedFormat: SUCCESS");

                    if (File.Exists(tempPdf))
                    {
                        long size = new FileInfo(tempPdf).Length;
                        Console.WriteLine("    PDF generated: {0}", tempPdf);
                        Console.WriteLine("    PDF size: {0:N0} bytes", size);
                        File.Delete(tempPdf);
                    }
                    else
                    {
                        Console.WriteLine("    PDF NOT created (but no exception)");
                        _errorCount++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("    ExportAsFixedFormat FAILED:");
                    Console.WriteLine("      Type:     {0}", ex.GetType().FullName);
                    Console.WriteLine("      Message:  {0}", ex.Message);
                    if (ex is COMException comEx)
                    {
                        Console.WriteLine("      HRESULT:  0x{0:X8}", comEx.ErrorCode);
                    }
                    Console.WriteLine("      Stack:    {0}", ex.StackTrace);
                    _errorCount++;

                    // Retry with IgnorePrintAreas: true
                    Console.WriteLine();
                    Console.WriteLine("    Retrying with IgnorePrintAreas: true...");
                    try
                    {
                        string tempPdf2 = Path.Combine(Path.GetTempPath(), string.Format("diag_{0:N}.pdf", Guid.NewGuid()));
                        exportSheet.ExportAsFixedFormat(
                            Type: Excel.XlFixedFormatType.xlTypePDF,
                            Filename: tempPdf2,
                            Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
                            IncludeDocProperties: false,
                            IgnorePrintAreas: true,
                            OpenAfterPublish: false);
                        Console.WriteLine("    ExportAsFixedFormat(IgnorePrintAreas: true): SUCCESS");
                        if (File.Exists(tempPdf2))
                        {
                            Console.WriteLine("    PDF size: {0:N0} bytes", new FileInfo(tempPdf2).Length);
                            File.Delete(tempPdf2);
                        }
                    }
                    catch (Exception ex2)
                    {
                        Console.WriteLine("    Also failed: {0}", ex2.Message);
                    }
                }
                finally { Marshal.ReleaseComObject(exportSheet); }
            }
            catch (Exception ex)
            {
                Console.WriteLine("    Export test setup failed: {0}", ex.Message);
                _errorCount++;
            }

            Step("DIAGNOSTIC", "Finished analyzing {0}", label);
            Console.WriteLine();
        }
        catch (Exception ex)
        {
            Step("FATAL", "{0}: {1}", ex.GetType().Name, ex.Message);
            Console.WriteLine(ex.StackTrace);
            _errorCount++;
        }
        finally
        {
            // Cleanup
            if (workbook != null)
            {
                try { workbook.Close(false); } catch { }
                Marshal.ReleaseComObject(workbook);
            }
            if (excelApp != null)
            {
                try { excelApp.Quit(); } catch { }
                Marshal.ReleaseComObject(excelApp);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    static void Step(string label, string message, params object?[] args)
    {
        _step++;
        Console.WriteLine("  [{0,2}] {1,-30} {2}", _step, label, string.Format(message, args));
    }
}
