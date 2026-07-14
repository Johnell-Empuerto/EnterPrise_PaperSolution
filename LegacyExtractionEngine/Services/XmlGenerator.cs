using System.Text;
using System.Xml.Linq;
using LegacyExtractionEngine.Models;

namespace LegacyExtractionEngine.Services;

public class XmlGenerator
{
    private readonly IProgress<string> _progress;
    private const string Indent = "  ";

    private const double DefaultPageWidth = 612;
    private const double DefaultPageHeight = 792;
    private const double DefaultRowHeightPt = 14.4;
    private const double DefaultColumnCharWidth = 8.43;
    private const double PerRowExtraPt = 0.36;
    private const double ClusterBaseGapPt = 0.72;

    /// <summary>
    /// Optional COM-rendered coordinate data. When provided, all cluster
    /// positions are computed using Excel's own layout engine via
    /// DB_Ratio = RoundEx((COM_Position + PrintedOrigin) / PageDimension, 7).
    /// When null, falls back to OpenXML column-width estimation.
    /// </summary>
    public ComExtraction? ComData { get; set; }

    public XmlGenerator(IProgress<string> progress)
    {
        _progress = progress;
    }

    public string GenerateConMasXml(ExtractionResult result, DateTime? fixedTimestamp = null, int? defTopId = null, string? defTopName = null)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<conmas>");
        AppendHeader(sb);
        AppendTop(sb, result, fixedTimestamp, defTopId, defTopName);
        sb.Append("</conmas>");
        return sb.ToString();
    }

    private void AppendHeader(StringBuilder sb)
    {
        sb.AppendLine(Indent + "<header>");
        IndentLine(sb, 2, "<version></version>");
        IndentLine(sb, 2, "<command></command>");
        IndentLine(sb, 2, "<resultCode></resultCode>");
        IndentLine(sb, 2, "<message></message>");
        IndentLine(sb, 2, "<requestUser></requestUser>");
        IndentLine(sb, 2, "<createTime></createTime>");
        sb.AppendLine(Indent + "</header>");
    }

    private void AppendTop(StringBuilder sb, ExtractionResult result, DateTime? fixedTimestamp = null, int? defTopId = null, string? defTopName = null)
    {
        sb.AppendLine(Indent + "<top>");

        var wb = result.Workbook;

        int tid = defTopId ?? 546;
        IndentLine(sb, 2, $"<defTopId>{tid}</defTopId>");
        string formName = defTopName ?? wb.Properties.Title ?? Path.GetFileNameWithoutExtension(result.SourceFile) ?? "Form";
        IndentLine(sb, 2, $"<defTopName>{EscapeXml(formName)}</defTopName>");
        IndentLine(sb, 2, "<repTopId></repTopId>");
        IndentLine(sb, 2, "<repTopName></repTopName>");
        IndentLine(sb, 2, "<updateDefinitionApp>ConMasDesigner</updateDefinitionApp>");
        sb.Append(Indent + Indent);
        sb.AppendLine();
        sb.Append(Indent + Indent);
        sb.AppendLine();
        IndentLine(sb, 2, "<constrainedFunction></constrainedFunction>");
        IndentLine(sb, 2, "<constrainedFunctionOther></constrainedFunctionOther>");
        IndentLine(sb, 2, "<isSortReport>0</isSortReport>");
        IndentLine(sb, 2, "<notDisplayRenumberedIndex>0</notDisplayRenumberedIndex>");
        IndentLine(sb, 2, "<reportType>1</reportType>");
        IndentLine(sb, 2, $"<sheetCount>{wb.Sheets.Count}</sheetCount>");
        IndentLine(sb, 2, "<autoGen>0</autoGen>");
        IndentLine(sb, 2, "<tables></tables>");
        IndentLine(sb, 2, "<mobileSave>0</mobileSave>");
        IndentLine(sb, 2, "<mobileReportSave>1</mobileReportSave>");
        IndentLine(sb, 2, "<useBiometrics>0</useBiometrics>");
        IndentLine(sb, 2, "<useIdentification>0</useIdentification>");
        IndentLine(sb, 2, "<safeKeeping>0</safeKeeping>");
        IndentLine(sb, 2, "<autoSelectGen>0</autoSelectGen>");
        IndentLine(sb, 2, "<nameEditable>0</nameEditable>");
        IndentLine(sb, 2, "<nameRegenerate>0</nameRegenerate>");
        IndentLine(sb, 2, "<lifeTime>0</lifeTime>");
        IndentLine(sb, 2, "<creatable></creatable>");
        IndentLine(sb, 2, "<finishOutput>1</finishOutput>");

        AppendFinishOutputFiles(sb);
        IndentLine(sb, 2, "<editOutput>0</editOutput>");
        AppendEditOutputFiles(sb);

        IndentLine(sb, 2, "<excelOutput>1</excelOutput>");
        IndentLine(sb, 2, "<readOnly>0</readOnly>");
        IndentLine(sb, 2, "<lockMode>0</lockMode>");
        IndentLine(sb, 2, "<locked>0</locked>");
        IndentLine(sb, 2, "<editStatus>0</editStatus>");
        IndentLine(sb, 2, "<publicStatus>2</publicStatus>");
        IndentLine(sb, 2, "<picOriginalResolution>0</picOriginalResolution>");
        IndentLine(sb, 2, "<imageSize></imageSize>");
        IndentLine(sb, 2, "<isOriginalWhole>1</isOriginalWhole>");
        IndentLine(sb, 2, "<wholeImageSize></wholeImageSize>");
        IndentLine(sb, 2, "<saveIndividuallyImage>1</saveIndividuallyImage>");

        AppendDefinitionFile(sb, formName);

        IndentLine(sb, 2, "<backgroundImage></backgroundImage>");
        IndentLine(sb, 2, "<thumbnail></thumbnail>");
        IndentLine(sb, 2, "<editMail></editMail>");
        IndentLine(sb, 2, "<completeMail></completeMail>");

        for (int i = 1; i <= 10; i++)
            IndentLine(sb, 2, $"<remarksName{i}>Form remarks {i}</remarksName{i}>");
        for (int i = 1; i <= 10; i++)
            IndentLine(sb, 2, $"<remarksValue{i}></remarksValue{i}>");

        IndentLine(sb, 2, "<remarksEditable>0</remarksEditable>");
        IndentLine(sb, 2, "<remarksClearCooperation>1</remarksClearCooperation>");

        string ts;
        if (fixedTimestamp.HasValue)
            ts = fixedTimestamp.Value.ToString("yyyy/MM/dd HH:mm:ss");
        else
            ts = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
        IndentLine(sb, 2, $"<registTime>{ts}</registTime>");
        IndentLine(sb, 2, "<registUser>conmasadmin</registUser>");
        IndentLine(sb, 2, "<registUserName></registUserName>");
        IndentLine(sb, 2, $"<updateTime>{ts}</updateTime>");
        IndentLine(sb, 2, "<updateUser>conmasadmin</updateUser>");
        IndentLine(sb, 2, "<updateUserName></updateUserName>");
        IndentLine(sb, 2, "<canSendMailAsAttachment>1</canSendMailAsAttachment>");
        IndentLine(sb, 2, "<canOpenAsPdf>0</canOpenAsPdf>");
        IndentLine(sb, 2, "<saveToServerReopen>0</saveToServerReopen>");
        IndentLine(sb, 2, "<saveLocalCameraImage>0</saveLocalCameraImage>");
        IndentLine(sb, 2, "<cooperationTable>0</cooperationTable>");
        IndentLine(sb, 2, "<requiredCheckMode>0</requiredCheckMode>");
        IndentLine(sb, 2, "<requiredSaveMode>0</requiredSaveMode>");
        IndentLine(sb, 2, "<requiredCheckPrint>0</requiredCheckPrint>");
        IndentLine(sb, 2, "<minimumEditSize></minimumEditSize>");
        IndentLine(sb, 2, "<finishCreateSortedReport>0</finishCreateSortedReport>");
        IndentLine(sb, 2, "<isReportCopy>0</isReportCopy>");
        IndentLine(sb, 2, "<reportCopyType>0</reportCopyType>");
        IndentLine(sb, 2, "<mobileEditType>0</mobileEditType>");
        IndentLine(sb, 2, "<useNetworkAutoInputStart>1</useNetworkAutoInputStart>");
        IndentLine(sb, 2, "<existReportMaster></existReportMaster>");
        IndentLine(sb, 2, "<useApplicantLock>0</useApplicantLock>");
        IndentLine(sb, 2, "<useInputHistory>0</useInputHistory>");
        IndentLine(sb, 2, "<useInitInputJudge>0</useInitInputJudge>");
        IndentLine(sb, 2, "<useInitInputJudgeParameters></useInitInputJudgeParameters>");
        IndentLine(sb, 2, "<useChangeReason>0</useChangeReason>");
        IndentLine(sb, 2, "<serverVersion>8.2.26020</serverVersion>");
        IndentLine(sb, 2, "<useExclusiveMode>0</useExclusiveMode>");
        IndentLine(sb, 2, "<useExclusiveModeManager>0</useExclusiveModeManager>");
        IndentLine(sb, 2, "<retinaMode new=\"1\">1</retinaMode>");
        IndentLine(sb, 2, "<calculateMode new=\"1\">1</calculateMode>");
        IndentLine(sb, 2, "<reEditDisable>0</reEditDisable>");
        IndentLine(sb, 2, "<useStartTime></useStartTime>");
        IndentLine(sb, 2, "<useEndTime></useEndTime>");

        for (int i = 1; i <= 5; i++)
            IndentLine(sb, 2, $"<systemKey{i} xml:space=\"preserve\"></systemKey{i}>");

        IndentLine(sb, 2, "<noNeedToFillOut>0</noNeedToFillOut>");
        IndentLine(sb, 2, "<noNeedToFillOutMode>0</noNeedToFillOutMode>");
        IndentLine(sb, 2, "<noNeedToFillOutCluster></noNeedToFillOutCluster>");
        IndentLine(sb, 2, "<noNeedToFillOutType>0</noNeedToFillOutType>");
        IndentLine(sb, 2, "<noNeedToFillOutString></noNeedToFillOutString>");
        IndentLine(sb, 2, "<journalizingDefTopId></journalizingDefTopId>");
        IndentLine(sb, 2, "<pinCooperationSelectCluster></pinCooperationSelectCluster>");
        IndentLine(sb, 2, "<trailOutput>0</trailOutput>");
        IndentLine(sb, 2, "<indexType>0</indexType>");
        IndentLine(sb, 2, "<voiceInputEndTime></voiceInputEndTime>");
        IndentLine(sb, 2, "<networkAnswerbackMode>0</networkAnswerbackMode>");
        IndentLine(sb, 2, "<audioRecordingFileFormat>1</audioRecordingFileFormat>");
        IndentLine(sb, 2, "<fontAutoResizingMode>1</fontAutoResizingMode>");
        IndentLine(sb, 2, "<useReferencedCluster>0</useReferencedCluster>");

        AppendDisplaySaveMenu(sb);
        AppendDividedDeviceCode(sb);

        IndentLine(sb, 2, "<nameParts xml:space=\"preserve\"></nameParts>");
        IndentLine(sb, 2, "<useAutoNumbering>0</useAutoNumbering>");

        AppendAutoNumbering(sb);

        IndentLine(sb, 2, "<workReportType>0</workReportType>");
        IndentLine(sb, 2, "<dailyReportCluster></dailyReportCluster>");
        IndentLine(sb, 2, "<workReportTableNo></workReportTableNo>");
        IndentLine(sb, 2, "<labels></labels>");
        IndentLine(sb, 2, "<networks></networks>");
        AppendMatrixScanSetting(sb);
        AppendEdgeOcrSetting(sb);
        IndentLine(sb, 2, "<pageClusters></pageClusters>");

        AppendSheets(sb, result);

        sb.AppendLine(Indent + "</top>");
    }

    private void AppendFinishOutputFiles(StringBuilder sb)
    {
        sb.AppendLine(Indent + Indent + "<finishOutputFiles>");
        foreach (var tag in new[] { "csv", "csvImageAudio", "csvZip", "dataOutputCsv", "dataOutputCsvImageAudio",
            "xml", "pdf", "pdfLayer", "docuworks", "excel" })
            IndentLine(sb, 3, $"<{tag}></{tag}>");
        sb.AppendLine(Indent + Indent + "</finishOutputFiles>");
    }

    private void AppendEditOutputFiles(StringBuilder sb)
    {
        sb.AppendLine(Indent + Indent + "<editOutputFiles>");
        foreach (var tag in new[] { "csv", "csvImageAudio", "csvZip", "dataOutputCsv", "dataOutputCsvImageAudio",
            "xml", "pdf", "pdfLayer", "docuworks", "excel" })
            IndentLine(sb, 3, $"<{tag}></{tag}>");
        sb.AppendLine(Indent + Indent + "</editOutputFiles>");
    }

    private void AppendDefinitionFile(StringBuilder sb, string formName)
    {
        sb.AppendLine(Indent + Indent + "<definitionFile>");
        IndentLine(sb, 3, "<type>xlsx</type>");
        IndentLine(sb, 3, $"<name>{EscapeXml(formName)}.xlsx</name>");
        IndentLine(sb, 3, "<value></value>");
        sb.AppendLine(Indent + Indent + "</definitionFile>");
    }

    private void AppendDisplaySaveMenu(StringBuilder sb)
    {
        sb.AppendLine(Indent + Indent + "<displaySaveMenu>");
        foreach (var tag in new[] { "localSave", "continuationSave", "serverSave", "finishSave",
            "continuationServerSave", "continuationFinishSave", "mailImage", "mailPdf",
            "openApp", "print", "printReceipt", "savePdf", "saveExcel" })
            IndentLine(sb, 3, $"<{tag}>1</{tag}>");
        sb.AppendLine(Indent + Indent + "</displaySaveMenu>");
    }

    private void AppendDividedDeviceCode(StringBuilder sb)
    {
        sb.AppendLine(Indent + Indent + "<dividedDeviceCode>");
        IndentLine(sb, 3, "<delimiterType></delimiterType>");
        IndentLine(sb, 3, "<encodeType></encodeType>");
        sb.AppendLine(Indent + Indent + "</dividedDeviceCode>");
    }

    private void AppendAutoNumbering(StringBuilder sb)
    {
        sb.AppendLine(Indent + Indent + "<autoNumbering xml:space=\"preserve\">");
        IndentLine(sb, 3, "<startValue></startValue>");
        IndentLine(sb, 3, "<increment></increment>");
        IndentLine(sb, 3, "<digit></digit>");
        IndentLine(sb, 3, "<zeroPadding></zeroPadding>");
        IndentLine(sb, 3, "<numberingTiming></numberingTiming>");
        IndentLine(sb, 3, "<numberingForReeditFromComplete></numberingForReeditFromComplete>");
        IndentLine(sb, 3, "<numberingCyclicEnabled></numberingCyclicEnabled>");
        IndentLine(sb, 3, "<useReset></useReset>");

        sb.AppendLine(Indent + Indent + Indent + "<reset>");
        IndentLine(sb, 4, "<termType></termType>");
        IndentLine(sb, 4, "<startDate></startDate>");
        IndentLine(sb, 4, "<executeTime></executeTime>");
        IndentLine(sb, 4, "<interval></interval>");
        IndentLine(sb, 4, "<week></week>");
        IndentLine(sb, 4, "<month></month>");
        IndentLine(sb, 4, "<day></day>");
        sb.AppendLine(Indent + Indent + Indent + "</reset>");

        IndentLine(sb, 3, "<useSaveCount></useSaveCount>");
        sb.AppendLine(Indent + Indent + Indent + "<saveCount>");
        IndentLine(sb, 4, "<digit></digit>");
        IndentLine(sb, 4, "<zeroPadding></zeroPadding>");
        sb.AppendLine(Indent + Indent + Indent + "</saveCount>");

        sb.AppendLine(Indent + Indent + "</autoNumbering>");
    }

    private void AppendMatrixScanSetting(StringBuilder sb)
    {
        sb.AppendLine(Indent + Indent + "<matrixScanSetting>");
        IndentLine(sb, 3, "<useScanditCluster>0</useScanditCluster>");
        IndentLine(sb, 3, "<scanditClusters></scanditClusters>");
        sb.AppendLine(Indent + Indent + "</matrixScanSetting>");
    }

    private void AppendEdgeOcrSetting(StringBuilder sb)
    {
        sb.AppendLine(Indent + Indent + "<edgeOcrSetting>");
        IndentLine(sb, 3, "<edgeOcrClusters />");
        sb.AppendLine(Indent + Indent + "</edgeOcrSetting>");
    }

    private void AppendSheets(StringBuilder sb, ExtractionResult result)
    {
        sb.AppendLine(Indent + Indent + "<sheets>");

        int sheetNo = 0;
        foreach (var sheet in result.Workbook.Sheets)
        {
            sheetNo++;
            AppendSheet(sb, sheet, result.Styles, sheetNo);
        }

        sb.AppendLine(Indent + Indent + "</sheets>");
    }

    private void AppendSheet(StringBuilder sb, SheetData sheet, List<StyleData> styles, int seqSheetNo)
    {
        sb.AppendLine(Indent + Indent + Indent + "<sheet>");

        IndentLine(sb, 4, $"<defSheetId>1706</defSheetId>");
        IndentLine(sb, 4, $"<defSheetName>{EscapeXml(sheet.Name)}</defSheetName>");
        IndentLine(sb, 4, "<repSheetId></repSheetId>");
        IndentLine(sb, 4, "<repSheetName></repSheetName>");
        IndentLine(sb, 4, $"<sheetNo>{seqSheetNo}</sheetNo>");
        IndentLine(sb, 4, "<autoSelectGen>0</autoSelectGen>");
        IndentLine(sb, 4, "<backgroundImage></backgroundImage>");
        IndentLine(sb, 4, "<inputImage></inputImage>");
        IndentLine(sb, 4, "<importImage>0</importImage>");
        IndentLine(sb, 4, "<thumbnail></thumbnail>");
        IndentLine(sb, 4, "<width>612</width>");
        IndentLine(sb, 4, "<height>792</height>");
        IndentLine(sb, 4, "<copyDisable>0</copyDisable>");
        IndentLine(sb, 4, "<sheetCopyCheck>0</sheetCopyCheck>");
        IndentLine(sb, 4, "<focusClusterIndex></focusClusterIndex>");

        for (int i = 1; i <= 10; i++)
            IndentLine(sb, 4, $"<remarksName{i}>Sheet remarks {i}</remarksName{i}>");
        for (int i = 1; i <= 10; i++)
            IndentLine(sb, 4, $"<remarksValue{i} xml:space=\"preserve\"></remarksValue{i}>");

        IndentLine(sb, 4, "<references></references>");

        var clusters = BuildClusters(sheet, styles, seqSheetNo);
        AppendClustersXml(sb, clusters);

        sb.AppendLine(Indent + Indent + Indent + "</sheet>");
    }

    private List<ConMasCluster> BuildClusters(SheetData sheet, List<StyleData> styles, int seqSheetNo)
    {
        var clusters = new List<ConMasCluster>();
        int clusterId = 0;

        double pageWidth = DefaultPageWidth;
        double pageHeight = DefaultPageHeight;

        bool horizCenter = false;
        if (sheet.PageSetup.AllAttributes.TryGetValue("horizontalCentered", out var hc) && hc is string hcs)
            horizCenter = hcs == "1" || hcs == "true";
        bool vertCenter = false;
        if (sheet.PageSetup.AllAttributes.TryGetValue("verticalCentered", out var vc) && vc is string vcs)
            vertCenter = vcs == "1" || vcs == "true";

        double defaultRowHeightPt = sheet.SheetFormatProperties?.DefaultRowHeight ?? DefaultRowHeightPt;
        double effectiveRowHeightPt = defaultRowHeightPt + PerRowExtraPt;
        int maxCol = sheet.Cells.Count > 0 ? sheet.Cells.Max(c => c.Column) : 4;
        if (sheet.MergedCells.Count > 0)
        {
            int mcMax = sheet.MergedCells.Max(m => m.EndColumn);
            if (mcMax > maxCol) maxCol = mcMax;
        }

        // Legacy COM coordinate system: use COM positions if available
        double printedOriginX = 0, printedOriginY = 0;
        if (ComData != null)
        {
            printedOriginX = ComData.PrintedOriginX;
            printedOriginY = ComData.PrintedOriginY;
            _progress.Report($"  [COM] PrintedOrigin: X={printedOriginX:F2}pt, Y={printedOriginY:F2}pt, " +
                $"Page: {ComData.PageWidth:F1}x{ComData.PageHeight:F1}, " +
                $"Content: {ComData.ContentWidth:F1}x{ComData.ContentHeight:F1}");
        }
        else
        {
            // Fallback: OpenXML estimation (remains for backward compat)
            double contentWidth = ComputeContentWidth(sheet, maxCol);
            double contentHeight = ComputeContentHeight(sheet, defaultRowHeightPt, effectiveRowHeightPt);
            double horizOffset = horizCenter ? (pageWidth - contentWidth) / 2.0 : sheet.PageMargins.Left * 72;
            double vertOffset = vertCenter ? (pageHeight - contentHeight) / 2.0 : sheet.PageMargins.Top * 72;
            printedOriginX = horizCenter ? (pageWidth - contentWidth) / 2.0 : sheet.PageMargins.Left * 72;
            printedOriginY = vertCenter ? (pageHeight - contentHeight) / 2.0 : sheet.PageMargins.Top * 72;
            _progress.Report($"  [FALLBACK] maxCol={maxCol} contentWidth={contentWidth:F2} contentHeight={contentHeight:F2} printedOrigin=({printedOriginX:F2},{printedOriginY:F2}) pageW={pageWidth} pageH={pageHeight}");
        }

        var allBoundaries = new List<(int StartRow, int EndRow)>();
        allBoundaries.AddRange(sheet.MergedCells.Select(m => (m.StartRow, m.EndRow)));
        var standaloneCells = sheet.Cells
            .Where(c => !sheet.MergedCells.Any(m =>
                c.Row >= m.StartRow && c.Row <= m.EndRow &&
                c.Column >= m.StartColumn && c.Column <= m.EndColumn))
            .GroupBy(c => new { c.Row, c.Column })
            .Select(g => g.First())
            .OrderBy(c => c.Row)
            .ThenBy(c => c.Column)
            .ToList();
        allBoundaries.AddRange(standaloneCells.Select(c => (c.Row, c.Row)));
        allBoundaries = allBoundaries
            .GroupBy(b => b.StartRow)
            .Select(g => g.OrderByDescending(b => b.EndRow).First())
            .OrderBy(b => b.StartRow)
            .ThenBy(b => b.EndRow)
            .ToList();

        int prevEndRow = 0;
        bool first = true;
        var clusterGaps = new Dictionary<int, double>();
        double cumulativeGap = 0;
        foreach (var b in allBoundaries)
        {
            if (first)
            {
                first = false;
                prevEndRow = b.EndRow;
                continue;
            }
            double gap = ClusterBaseGapPt + PerRowExtraPt * (b.StartRow - prevEndRow);
            cumulativeGap += gap;
            clusterGaps[b.StartRow] = cumulativeGap;
            prevEndRow = b.EndRow;
        }

        // Compute cluster positions using COM-rendered positions (preferred) or OpenXML fallback
        foreach (var mc in sheet.MergedCells)
        {
            var addr = ToAbsoluteRef(mc.Reference);
            ComPosition? comPos = null;
            ComData?.ComPositions.TryGetValue(addr, out comPos);

            double left, right, top, bottom;

            if (comPos != null)
            {
                // Use COM-rendered positions (exactly as legacy VB6 importer did)
                double pageW = ComData!.PageWidth;
                double pageH = ComData.PageHeight;
                double ox = ComData.PrintedOriginX;
                double oy = ComData.PrintedOriginY;

                left = ComCoordinateService.ComputeRatio(comPos.Left, ox, pageW);
                top = ComCoordinateService.ComputeRatio(comPos.Top, oy, pageH);
                // Right/Bottom include cluster gap derived from first cluster
                // The gap may vary per row due to cumulative row spacing
                double gapX = ComData.ClusterGapX;
                double gapY = ComData.ClusterGapY;
                right = ComCoordinateService.ComputeRatio(comPos.Left + comPos.Width + gapX, ox, pageW);
                bottom = ComCoordinateService.ComputeRatio(comPos.Top + comPos.Height + gapY, oy, pageH);
            }
            else
            {
                // Fallback: legacy OpenXML estimation
                double gapBefore = clusterGaps.GetValueOrDefault(mc.StartRow, 0);
                double topPos = printedOriginY + (mc.StartRow - 1) * effectiveRowHeightPt + gapBefore;
                double bottomPos = printedOriginY + mc.EndRow * effectiveRowHeightPt + gapBefore;
                left = ComputeColumnLeft(mc.StartColumn, sheet.Columns, pageWidth, printedOriginX);
                right = ComputeColumnRight(mc.EndColumn, sheet.Columns, pageWidth, printedOriginX);
                if (mc.EndColumn < maxCol)
                    right = RoundRatio(right - (ClusterBaseGapPt + PerRowExtraPt) / pageWidth);
                top = RoundRatio(topPos / pageHeight);
                bottom = RoundRatio(bottomPos / pageHeight);
            }

            var topLeftCell = sheet.Cells.FirstOrDefault(c =>
                c.Row == mc.StartRow && c.Column == mc.StartColumn);
            var style = topLeftCell?.StyleIndex != null && topLeftCell.StyleIndex.Value < styles.Count
                ? styles[topLeftCell.StyleIndex.Value] : null;

            clusters.Add(new ConMasCluster
            {
                ClusterId = clusterId,
                SheetNo = seqSheetNo,
                CellAddress = addr,
                Left = left,
                Right = right,
                Top = top,
                Bottom = bottom,
                Value = topLeftCell?.Value?.ToString() ?? "",
                FontName = GetFontName(style),
                FontSize = GetFontSize(style),
                FontBold = IsBold(style),
                Align = GetHorizontalAlign(style),
                VerticalAlign = GetVerticalAlign(style),
                FillColor = GetFontColor(style)
            });
            clusterId++;
        }

        foreach (var cell in standaloneCells)
        {
            var addr = ToAbsoluteRef(cell.Reference);
            ComPosition? comPos = null;
            ComData?.ComPositions.TryGetValue(addr, out comPos);

            double left, right, top, bottom;

            if (comPos != null)
            {
                // Use COM-rendered positions
                double pageW = ComData!.PageWidth;
                double pageH = ComData.PageHeight;
                double ox = ComData.PrintedOriginX;
                double oy = ComData.PrintedOriginY;

                left = ComCoordinateService.ComputeRatio(comPos.Left, ox, pageW);
                top = ComCoordinateService.ComputeRatio(comPos.Top, oy, pageH);
                // Right/Bottom include cluster gap derived from first cluster
                // The gap may vary per row due to cumulative row spacing
                double gapX = ComData.ClusterGapX;
                double gapY = ComData.ClusterGapY;
                right = ComCoordinateService.ComputeRatio(comPos.Left + comPos.Width + gapX, ox, pageW);
                bottom = ComCoordinateService.ComputeRatio(comPos.Top + comPos.Height + gapY, oy, pageH);
            }
            else
            {
                // Fallback: legacy OpenXML estimation
                double gapBefore = clusterGaps.GetValueOrDefault(cell.Row, 0);
                double topPos = printedOriginY + (cell.Row - 1) * effectiveRowHeightPt + gapBefore;
                double bottomPos = topPos + effectiveRowHeightPt;
                left = ComputeColumnLeft(cell.Column, sheet.Columns, pageWidth, printedOriginX);
                right = ComputeColumnRight(cell.Column, sheet.Columns, pageWidth, printedOriginX);
                if (cell.Column < maxCol)
                    right = RoundRatio(right - (ClusterBaseGapPt + PerRowExtraPt) / pageWidth);
                top = RoundRatio(topPos / pageHeight);
                bottom = RoundRatio(bottomPos / pageHeight);
            }

            var style = cell.StyleIndex != null && cell.StyleIndex.Value < styles.Count
                ? styles[cell.StyleIndex.Value] : null;

            clusters.Add(new ConMasCluster
            {
                ClusterId = clusterId,
                SheetNo = seqSheetNo,
                CellAddress = addr,
                Left = left,
                Right = right,
                Top = top,
                Bottom = bottom,
                Value = cell.Value?.ToString() ?? "",
                FontName = GetFontName(style),
                FontSize = GetFontSize(style),
                FontBold = IsBold(style),
                Align = GetHorizontalAlign(style),
                VerticalAlign = GetVerticalAlign(style),
                FillColor = GetFontColor(style)
            });
            clusterId++;
        }

        return clusters.OrderBy(c => c.Top).ThenBy(c => c.Left).ToList();
    }

    private double ComputeContentWidth(SheetData sheet, int maxUsedCol)
    {
        if (sheet.Columns.Count > 0)
        {
            double total = 0;
            // Sum each column position individually to avoid double-counting
            // overlapping column definitions. For each column position 1..maxUsedCol,
            // find the FIRST matching column definition (ordered by Min ascending).
            var sortedCols = sheet.Columns.OrderBy(c => c.Min).ThenBy(c => c.Max).ToList();
            for (int i = 1; i <= maxUsedCol; i++)
            {
                var col = sortedCols.FirstOrDefault(c => i >= c.Min && i <= c.Max);
                double chars = col?.Width ?? DefaultColumnCharWidth;
                total += ColumnCharsToPoints(chars);
            }
            return total;
        }
        return maxUsedCol * ColumnCharsToPoints(DefaultColumnCharWidth);
    }

    private double ComputeContentHeight(SheetData sheet, double defaultRowHeightPt, double effectiveRowHeightPt)
    {
        if (sheet.Rows.Count > 0)
        {
            int maxRow = sheet.Rows.Max(r => r.RowIndex);
            double marginForCentering = 0.75;
            double extra = Math.Round(marginForCentering, 1) * defaultRowHeightPt / 2.0;
            return maxRow * effectiveRowHeightPt + extra;
        }
        return 0;
    }

    private static double ColumnCharsToPoints(double charWidth)
    {
        return charWidth * 50.04 / 8.43;
    }

    private double ComputeColumnLeft(int column, List<ColumnData> columns, double pageWidth, double horizOffset)
    {
        double cumulative = 0;
        // Sort by Min to ensure correct definition priority (narrowest match first)
        var sortedCols = columns.OrderBy(c => c.Min).ThenBy(c => c.Max).ToList();
        for (int i = 1; i < column; i++)
        {
            var col = sortedCols.FirstOrDefault(c => i >= c.Min && i <= c.Max);
            double chars = col?.Width ?? DefaultColumnCharWidth;
            cumulative += ColumnCharsToPoints(chars);
        }
        return Math.Round((horizOffset + cumulative) / pageWidth, 7);
    }

    private double ComputeColumnRight(int column, List<ColumnData> columns, double pageWidth, double horizOffset)
    {
        double cumulative = 0;
        // Sort by Min to ensure correct definition priority (narrowest match first)
        var sortedCols = columns.OrderBy(c => c.Min).ThenBy(c => c.Max).ToList();
        for (int i = 1; i <= column; i++)
        {
            var col = sortedCols.FirstOrDefault(c => i >= c.Min && i <= c.Max);
            double chars = col?.Width ?? DefaultColumnCharWidth;
            cumulative += ColumnCharsToPoints(chars);
        }
        return Math.Round((horizOffset + cumulative) / pageWidth, 7);
    }

    private static double RoundRatio(double value)
    {
        return Math.Round((float)value, 7, MidpointRounding.AwayFromZero);
    }

    private void AppendClustersXml(StringBuilder sb, List<ConMasCluster> clusters)
    {
        sb.AppendLine(Indent + Indent + Indent + Indent + "<clusters>");

        foreach (var c in clusters)
        {
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + "<cluster>");

            IndentLine(sb, 6, $"<sheetNo>{c.SheetNo}</sheetNo>");
            IndentLine(sb, 6, $"<clusterId>{c.ClusterId}</clusterId>");
            IndentLine(sb, 6, "<isHidden>0</isHidden>");
            IndentLine(sb, 6, "<isHiddenDesigner>0</isHiddenDesigner>");
            IndentLine(sb, 6, "<mobileDisplay>1</mobileDisplay>");
            IndentLine(sb, 6, "<mobileListDisplayNo></mobileListDisplayNo>");
            IndentLine(sb, 6, "<pinNo></pinNo>");
            IndentLine(sb, 6, "<pinValue></pinValue>");
            IndentLine(sb, 6, "<name>samples</name>");
            IndentLine(sb, 6, "<type>KeyboardText</type>");
            IndentLine(sb, 6, $"<top>{FormatRatio(c.Top)}</top>");
            IndentLine(sb, 6, $"<bottom>{FormatRatio(c.Bottom)}</bottom>");
            IndentLine(sb, 6, $"<right>{FormatRatio(c.Right)}</right>");
            IndentLine(sb, 6, $"<left>{FormatRatio(c.Left)}</left>");
            IndentLine(sb, 6, $"<value>{EscapeXml(c.Value)}</value>");
            IndentLine(sb, 6, "<external>0</external>");
            IndentLine(sb, 6, "<displayValue></displayValue>");
            IndentLine(sb, 6, "<cooperationCluster>0</cooperationCluster>");
            IndentLine(sb, 6, "<readOnly>0</readOnly>");
            IndentLine(sb, 6, "<function></function>");
            IndentLine(sb, 6, "<actionPost></actionPost>");
            IndentLine(sb, 6, "<clearCluster></clearCluster>");
            IndentLine(sb, 6, "<originalFunction></originalFunction>");
            IndentLine(sb, 6, "<excelOutputValue></excelOutputValue>");
            IndentLine(sb, 6, $"<inputParameters>{EscapeXml(c.GetInputParameters())}</inputParameters>");
            IndentLine(sb, 6, "<carbonCopy></carbonCopy>");
            AppendUserCustomMaster(sb);
            AppendReportCopy(sb);
            AppendDividedCopy(sb);
            IndentLine(sb, 6, "<buttonImage></buttonImage>");
            IndentLine(sb, 6, "<buttonImageName></buttonImageName>");

            for (int i = 1; i <= 10; i++)
                IndentLine(sb, 6, $"<remarksValue{i}></remarksValue{i}>");

            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent + "<management>");
            IndentLine(sb, 7, "<valueToRemarks />");
            IndentLine(sb, 7, "<valueToSystemKeys />");
            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent + "</management>");

            IndentLine(sb, 6, $"<cellAddress>{EscapeXml(c.CellAddress)}</cellAddress>");

            sb.AppendLine(Indent + Indent + Indent + Indent + Indent + "</cluster>");
        }

        sb.AppendLine(Indent + Indent + Indent + Indent + "</clusters>");
    }

    private void AppendUserCustomMaster(StringBuilder sb)
    {
        sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent + "<userCustomMaster>");
        IndentLine(sb, 7, "<masterTableId></masterTableId>");
        IndentLine(sb, 7, "<masterKey></masterKey>");
        sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent + "</userCustomMaster>");
    }

    private void AppendReportCopy(StringBuilder sb)
    {
        sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent + "<reportCopy>");
        IndentLine(sb, 7, "<clear>0</clear>");
        IndentLine(sb, 7, "<displayDefaultValue>1</displayDefaultValue>");
        sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent + "</reportCopy>");
    }

    private void AppendDividedCopy(StringBuilder sb)
    {
        sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent + "<dividedCopy>");
        IndentLine(sb, 7, "<delimiterType></delimiterType>");
        IndentLine(sb, 7, "<encodeType></encodeType>");
        sb.AppendLine(Indent + Indent + Indent + Indent + Indent + Indent + "</dividedCopy>");
    }

    private static string GetFontName(StyleData? style)
    {
        return "Arial";
    }

    private static string GetFontSize(StyleData? style)
    {
        return "11";
    }

    private static string IsBold(StyleData? style)
    {
        return style?.Font?.Bold == true ? "Bold" : "Normal";
    }

    private static string GetHorizontalAlign(StyleData? style)
    {
        if (style?.Alignment?.Horizontal != null)
            return style.Alignment.Horizontal switch
            {
                "left" => "Left",
                "center" => "Center",
                "right" => "Right",
                _ => "Center"
            };
        return "Center";
    }

    private static string GetVerticalAlign(StyleData? style)
    {
        if (style?.Alignment?.Vertical != null)
            return style.Alignment.Vertical switch
            {
                "top" => "0",
                "center" => "1",
                "bottom" => "2",
                _ => "2"
            };
        return "2";
    }

    private static string GetFontColor(StyleData? style)
    {
        if (style?.Font?.Color != null)
        {
            if (style.Font.Color.StartsWith("#"))
            {
                var hex = style.Font.Color.TrimStart('#');
                if (hex.Length >= 6)
                {
                    var r = Convert.ToInt32(hex.Substring(0, 2), 16);
                    var g = Convert.ToInt32(hex.Substring(2, 2), 16);
                    var b = Convert.ToInt32(hex.Substring(4, 2), 16);
                    return $"{r},{g},{b}";
                }
            }
        }
        return "0,0,0";
    }



    private static string FormatRatio(double ratio)
    {
        return ratio.ToString("0.0#######");
    }

    private static void IndentLine(StringBuilder sb, int level, string content)
    {
        for (int i = 0; i < level; i++)
            sb.Append(Indent);
        sb.AppendLine(content);
    }

    private static string ToAbsoluteRef(string reference)
    {
        if (string.IsNullOrEmpty(reference)) return reference;
        if (reference.StartsWith("$") && reference.Contains("$", StringComparison.Ordinal) && reference.IndexOf('$') != reference.LastIndexOf('$'))
            return reference;
        var parts = reference.Split(':');
        for (int i = 0; i < parts.Length; i++)
        {
            if (string.IsNullOrEmpty(parts[i])) continue;
            var letters = new string(parts[i].TakeWhile(char.IsLetter).ToArray());
            var digits = new string(parts[i].SkipWhile(char.IsLetter).ToArray());
            parts[i] = "$" + letters + "$" + digits;
        }
        return string.Join(":", parts);
    }

    private static string EscapeXml(string s)
    {
        return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
            .Replace("\"", "&quot;").Replace("'", "&apos;");
    }
}

public class ConMasCluster
{
    public int ClusterId { get; set; }
    public int SheetNo { get; set; }
    public string CellAddress { get; set; } = "";
    public double Left { get; set; }
    public double Right { get; set; }
    public double Top { get; set; }
    public double Bottom { get; set; }
    public string Value { get; set; } = "";
    public string FontName { get; set; } = "Arial";
    public string FontSize { get; set; } = "11";
    public string FontBold { get; set; } = "Normal";
    public string Align { get; set; } = "Center";
    public string VerticalAlign { get; set; } = "2";
    public string FillColor { get; set; } = "0,0,0";

    public string GetInputParameters()
    {
        return $"Required=0;Lines=1;InputRestriction=None;MaxLength=0;Align={Align};Font={FontName};FontSize={FontSize};Weight={FontBold};Color={FillColor};VerticalAlignment={VerticalAlign};DefaultFontSize={FontSize}";
    }
}
