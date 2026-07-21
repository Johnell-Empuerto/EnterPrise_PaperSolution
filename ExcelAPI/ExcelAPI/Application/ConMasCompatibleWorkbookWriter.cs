using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using System.IO.Compression;
using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using WbDef = ExcelAPI.Models.WorkbookDefinition;

namespace ExcelAPI.Application
{
    public class ConMasCompatibleWorkbookWriter
    {
        private readonly ILogger<ConMasCompatibleWorkbookWriter> _logger;

        public ConMasCompatibleWorkbookWriter(ILogger<ConMasCompatibleWorkbookWriter> logger)
        {
            _logger = logger;
        }

        public int WriteValues(WbDef.WorkbookDefinition wbDef, string sourceXlsxPath, string outputPath)
        {
            if (!System.IO.File.Exists(sourceXlsxPath))
                throw new FileNotFoundException($"Source workbook not found: {sourceXlsxPath}");

            System.IO.File.Copy(sourceXlsxPath, outputPath, overwrite: true);

            int totalCellsWritten = 0;
            int totalFieldsReceived = 0;
            int totalFieldsWithValues = 0;

            foreach (var diagSheet in wbDef.Sheets)
            {
                totalFieldsReceived += diagSheet.Fields.Count;
                totalFieldsWithValues += diagSheet.Fields.Count(f => !string.IsNullOrWhiteSpace(f.Value));
            }

            if (totalFieldsWithValues == 0)
            {
                _logger.LogWarning("No fields with non-empty values — workbook copied unchanged.");
            }

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = new Excel.Application { Visible = false, DisplayAlerts = false };

                _logger.LogInformation("Opening workbook via COM: {Path}", outputPath);
                workbook = excelApp.Workbooks.Open(outputPath);
                _logger.LogInformation("Workbook opened: {Name}", workbook.Name);

                foreach (var wbSheet in wbDef.Sheets)
                {
                    if (wbSheet.Fields.Count == 0) continue;

                    Excel.Worksheet ws = null;
                    try
                    {
                        ws = (Excel.Worksheet)workbook.Worksheets[wbSheet.Name];
                    }
                    catch
                    {
                        _logger.LogWarning("Sheet '{Name}' not found — skipping {Count} fields",
                            wbSheet.Name, wbSheet.Fields.Count);
                        continue;
                    }

                    int cellsWritten = 0;
                    foreach (var field in wbSheet.Fields)
                    {
                        if (string.IsNullOrWhiteSpace(field.Value)) continue;

                        string cellRef = field.Cell.Address;
                        Excel.Range range = null;

                        try
                        {
                            range = ws.Range[cellRef];

                            if (IsNumericField(field) && double.TryParse(field.Value, out double numVal))
                                range.Value = numVal;
                            else
                                range.Value = field.Value;

                            _logger.LogInformation("Wrote cell {Sheet}!{Cell}: '{Value}'",
                                wbSheet.Name, cellRef, field.Value);
                            cellsWritten++;
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to write cell {Sheet}!{Cell}: '{Value}'",
                                wbSheet.Name, cellRef, field.Value);
                        }
                        finally
                        {
                            if (range != null) Marshal.ReleaseComObject(range);
                        }
                    }

                    _logger.LogInformation("Wrote {Count} values to sheet '{Name}'",
                        cellsWritten, wbSheet.Name);
                    totalCellsWritten += cellsWritten;

                    if (ws != null) Marshal.ReleaseComObject(ws);
                }

                _logger.LogInformation("Saving workbook via COM...");
                workbook.Save();
                _logger.LogInformation("Workbook saved successfully");

                workbook.Close(false);
                workbook = null;

                // Idempotent PostProcess — only runs on first export
                PostProcessZipForConMas(outputPath, wbDef);

                _logger.LogInformation("Total cells written: {Count}", totalCellsWritten);
                return totalCellsWritten;
            }
            finally
            {
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
            }
        }

        private static bool IsNumericField(WbDef.FieldDefinition field)
        {
            return field.Type == WbDef.FieldType.Number
                || field.Type == WbDef.FieldType.Calculated;
        }

        // ════════════════════════════════════════════════════════════════════════
        // Post-process ZIP for ExcelOutputSetting (Idempotent)
        // ════════════════════════════════════════════════════════════════════════
        private void PostProcessZipForConMas(string outputPath, WbDef.WorkbookDefinition wbDef)
        {
            _logger.LogInformation("[PHASE8.1] Post-processing ZIP for ExcelOutputSetting (idempotent)...");

            try
            {
                using var pkg = ZipFile.Open(outputPath, ZipArchiveMode.Update);

                var existingSheet3 = pkg.GetEntry("xl/worksheets/sheet3.xml");
                if (existingSheet3 != null)
                {
                    _logger.LogInformation("[PHASE8.1] ExcelOutputSetting already exists — skipping PostProcess (idempotent guard)");
                    return;
                }
                _logger.LogInformation("[PHASE8.1] No existing ExcelOutputSetting found — will create configuration sheet");

                int clusterCount = Math.Max(1, wbDef.Sheets.Sum(s => s.Fields.Count));
                string[] xmlFragments = GenerateExcelOutputSettingFragments(
                    sheetCount: Math.Max(1, wbDef.Sheets.Count),
                    clusterCount: clusterCount);

                int newStartIndex = AppendExcelOutputSettingSharedStrings(pkg, xmlFragments);
                WriteConMasCommentsEntry(pkg, wbDef);
                CreateExcelOutputSettingSheetEntry(pkg, newStartIndex);

                string worksheetRelId = ComputeNextWorkbookRelId(pkg);
                _logger.LogInformation("[PHASE5.5.4] Computed worksheet relationship ID: {RelId}", worksheetRelId);

                UpdateWorkbookXmlForSheet3(pkg, worksheetRelId);
                UpdateWorkbookRelsForSheet3(pkg, worksheetRelId);
                UpdateContentTypesForSheet3(pkg);

                _logger.LogInformation("[PHASE8.1] Post-processing complete: {Strings} shared strings @ index {StartIdx}, sheet3 created, {Fragments} fragments used",
                    xmlFragments.Length, newStartIndex, xmlFragments.Length);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[PHASE8.1] Post-processing failed: {Msg}", ex.Message);
            }
        }

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
                string commentText = BuildConMasCommentText(field, i);
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

        private static string BuildConMasCommentText(WbDef.FieldDefinition field, int fieldIndex)
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

            var newRel = new XElement(ns + "Relationship",
                new XAttribute("Id", relId),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                new XAttribute("Target", "worksheets/sheet3.xml"));
            doc.Root.Add(newRel);

            entry.Delete();
            var newEntry = pkg.CreateEntry("xl/_rels/workbook.xml.rels");
            using (var w = new StreamWriter(newEntry.Open()))
                doc.Save(w);
        }

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
    }
}
