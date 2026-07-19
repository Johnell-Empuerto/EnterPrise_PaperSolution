using ExcelAPI.Designer.Legacy.Models;

namespace ExcelAPI.Designer.Legacy.VerificationEngine;

public interface IVerificationEngine
{
    Task<VerificationReport> VerifyAsync(int defTopId, string excelFilePath, string? outputDir = null);
    Task<VerificationReport> VerifyAsync(int defTopId, PublishResult publishResult, string? outputDir = null);
}

public interface IVerificationDatabaseReader
{
    Task<DefTopData?> ReadDefTopAsync(int defTopId);
    Task<List<DefClusterData>> ReadDefClustersAsync(int defTopId);
    Task<byte[]?> ReadBackgroundImageAsync(int defTopId);
}

public interface IXmlComparer
{
    XmlComparisonResult Compare(string generatedXml, string databaseXml);
}

public interface IImageComparer
{
    ImageComparisonResult Compare(string generatedPath, byte[]? databaseImageBytes);
}

public interface IVerificationReportGenerator
{
    Task<string> GenerateJsonAsync(VerificationReport report, string outputDir);
    Task<string> GenerateHtmlAsync(VerificationReport report, string outputDir);
    Task<string> GenerateCsvAsync(VerificationReport report, string outputDir);
}
