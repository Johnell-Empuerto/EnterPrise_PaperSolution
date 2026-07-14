using ExcelAPI.LegacyEngine.Models;
using ExcelAPI.LegacyEngine.PublishEngine;

namespace ExcelAPI.LegacyEngine.VerificationEngine;

public class VerificationEngine : IVerificationEngine
{
    private readonly IPublishEngine _publishEngine;
    private readonly IVerificationDatabaseReader _dbReader;
    private readonly IXmlComparer _xmlComparer;
    private readonly IImageComparer _imageComparer;
    private readonly IVerificationReportGenerator _reportGenerator;
    private readonly ILogger<VerificationEngine> _logger;
    private const double CoordinateTolerance = 0.0001;

    public VerificationEngine(
        IPublishEngine publishEngine,
        IVerificationDatabaseReader dbReader,
        IXmlComparer xmlComparer,
        IImageComparer imageComparer,
        IVerificationReportGenerator reportGenerator,
        ILogger<VerificationEngine> logger)
    {
        _publishEngine = publishEngine;
        _dbReader = dbReader;
        _xmlComparer = xmlComparer;
        _imageComparer = imageComparer;
        _reportGenerator = reportGenerator;
        _logger = logger;
    }

    public async Task<VerificationReport> VerifyAsync(int defTopId, string excelFilePath, string? outputDir = null)
    {
        outputDir ??= Path.Combine(Path.GetTempPath(), "PaperLessVerify",
            Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(outputDir);

        var report = new VerificationReport
        {
            DefTopId = defTopId,
            GeneratedAt = DateTime.UtcNow
        };

        try
        {
            var defTop = await _dbReader.ReadDefTopAsync(defTopId);
            var dbClusters = await _dbReader.ReadDefClustersAsync(defTopId);

            if (defTop == null)
            {
                _logger.LogWarning("DefTop {Id} not found in database", defTopId);
                report.Overall = "FAIL";
                report.Summary.Failed++;
                return report;
            }

            var publishResult = await _publishEngine.PublishAsync(excelFilePath, outputDir);

            if (!publishResult.Success)
            {
                _logger.LogError("Publish failed for DefTop {Id}: {Error}", defTopId, publishResult.ErrorMessage);
                report.Overall = "FAIL";
                report.Summary.Failed++;
                return report;
            }

            return await CompleteVerification(report, publishResult, defTop, dbClusters, outputDir);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Verification failed for DefTop {Id}", defTopId);
            report.Overall = "FAIL";
            return report;
        }
    }

    public async Task<VerificationReport> VerifyAsync(int defTopId, PublishResult publishResult, string? outputDir = null)
    {
        outputDir ??= Path.Combine(Path.GetTempPath(), "PaperLessVerify",
            Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(outputDir);

        var report = new VerificationReport
        {
            DefTopId = defTopId,
            GeneratedAt = DateTime.UtcNow
        };

        try
        {
            var defTop = await _dbReader.ReadDefTopAsync(defTopId);
            var dbClusters = await _dbReader.ReadDefClustersAsync(defTopId);

            if (defTop == null)
            {
                _logger.LogWarning("DefTop {Id} not found in database", defTopId);
                report.Overall = "FAIL";
                report.Summary.Failed++;
                return report;
            }

            return await CompleteVerification(report, publishResult, defTop, dbClusters, outputDir);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Verification failed for DefTop {Id}", defTopId);
            report.Overall = "FAIL";
            return report;
        }
    }

    private async Task<VerificationReport> CompleteVerification(
        VerificationReport report,
        PublishResult publishResult,
        DefTopData defTop,
        List<DefClusterData> dbClusters,
        string outputDir)
    {
        report.XmlComparison = _xmlComparer.Compare(
            publishResult.XmlData ?? "", defTop.XmlData ?? "");
        report.Summary.TotalChecks++;
        if (report.XmlComparison.Status == "PASS") report.Summary.Passed++;
        else report.Summary.Failed++;

        double pageW = publishResult.Sheets.FirstOrDefault()?.WidthPoints ?? 612;
        double pageH = publishResult.Sheets.FirstOrDefault()?.HeightPoints ?? 792;

        report.CoordinateComparison = CompareCoordinates(publishResult, dbClusters, pageW, pageH);
        report.Summary.TotalChecks++;
        if (report.CoordinateComparison.Status == "PASS") report.Summary.Passed++;
        else report.Summary.Failed++;

        var bgPath = publishResult.BackgroundImagePath;
        report.ImageComparison = _imageComparer.Compare(bgPath ?? "", defTop.BackgroundImageData);
        report.Summary.TotalChecks++;
        if (report.ImageComparison.Status == "PASS") report.Summary.Passed++;
        else if (report.ImageComparison.Status == "FAIL") report.Summary.Failed++;

        int totalChecks = report.Summary.Passed + report.Summary.Failed;
        report.Summary.OverallMatchPercentage = totalChecks > 0
            ? (double)report.Summary.Passed / totalChecks * 100.0 : 0;
        report.Overall = report.Summary.Failed == 0 ? "PASS" : "FAIL";

        var verDir = Path.Combine(outputDir, "Verification");
        await _reportGenerator.GenerateJsonAsync(report, verDir);
        await _reportGenerator.GenerateHtmlAsync(report, verDir);
        await _reportGenerator.GenerateCsvAsync(report, verDir);

        return report;
    }

    private CoordinateComparisonResult CompareCoordinates(
        PublishResult publishResult,
        List<DefClusterData> dbClusters,
        double pageWidth,
        double pageHeight)
    {
        var result = new CoordinateComparisonResult();

        foreach (var dbCluster in dbClusters)
        {
            var genCluster = publishResult.Clusters
                .FirstOrDefault(c => c.ClusterId == dbCluster.ClusterId);

            if (genCluster == null)
            {
                result.Failed++;
                result.Differences.Add(new CoordinateDiff
                {
                    ClusterId = dbCluster.ClusterId,
                    CellAddress = dbCluster.CellAddr ?? "",
                    Status = "MISSING"
                });
                continue;
            }

            double dbLeft = ParseRatio(dbCluster.LeftPosition);
            double dbTop = ParseRatio(dbCluster.TopPosition);
            double dbRight = ParseRatio(dbCluster.RightPosition);
            double dbBottom = ParseRatio(dbCluster.BottomPosition);

            double diffL = Math.Abs(genCluster.Left - dbLeft);
            double diffT = Math.Abs(genCluster.Top - dbTop);
            double diffR = Math.Abs(genCluster.Right - dbRight);
            double diffB = Math.Abs(genCluster.Bottom - dbBottom);

            bool pass = diffL < CoordinateTolerance && diffT < CoordinateTolerance
                     && diffR < CoordinateTolerance && diffB < CoordinateTolerance;

            result.Differences.Add(new CoordinateDiff
            {
                ClusterId = genCluster.ClusterId,
                SheetNo = genCluster.SheetNo,
                CellAddress = genCluster.CellAddress,
                Status = pass ? "PASS" : "FAIL",
                ExpectedLeft = dbLeft,
                ExpectedTop = dbTop,
                ExpectedRight = dbRight,
                ExpectedBottom = dbBottom,
                ActualLeft = genCluster.Left,
                ActualTop = genCluster.Top,
                ActualRight = genCluster.Right,
                ActualBottom = genCluster.Bottom,
                DiffLeftPoints = diffL * pageWidth,
                DiffTopPoints = diffT * pageHeight,
                DiffRightPoints = diffR * pageWidth,
                DiffBottomPoints = diffB * pageHeight,
                DiffLeftRatio = diffL,
                DiffTopRatio = diffT,
                DiffRightRatio = diffR,
                DiffBottomRatio = diffB
            });

            if (pass) result.Passed++;
            else result.Failed++;
        }

        result.TotalClusters = result.Passed + result.Failed;
        result.MatchPercentage = result.TotalClusters > 0
            ? (double)result.Passed / result.TotalClusters * 100.0 : 0;
        result.Status = result.MatchPercentage >= 100 ? "PASS"
            : result.MatchPercentage > 0 ? "PARTIAL" : "FAIL";

        return result;
    }

    private static double ParseRatio(string? value)
    {
        if (string.IsNullOrEmpty(value)) return 0;
        return double.TryParse(value,
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture,
            out var result) ? result : 0;
    }
}
