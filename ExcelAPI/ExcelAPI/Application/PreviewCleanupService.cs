using Microsoft.Extensions.Options;
using ExcelAPI.Models;

namespace ExcelAPI.Application
{
    /// <summary>
    /// Background service that periodically removes old preview images and temporary upload files.
    /// Runs on a configurable interval and respects the configured retention period.
    /// Never deletes files currently in use (checks if file is locked).
    /// </summary>
    public class PreviewCleanupService : BackgroundService
    {
        private readonly IWebHostEnvironment _env;
        private readonly IOptions<ExcelCaptureOptions> _options;
        private readonly ILogger<PreviewCleanupService> _logger;

        public PreviewCleanupService(
            IWebHostEnvironment env,
            IOptions<ExcelCaptureOptions> options,
            ILogger<PreviewCleanupService> logger)
        {
            _env = env;
            _options = options;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation(
                "Preview cleanup service started. Interval: {Interval}min, Retention: {Retention}h",
                _options.Value.CleanupIntervalMinutes,
                _options.Value.DeletePreviewAfterHours);

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    await Task.Delay(
                        TimeSpan.FromMinutes(_options.Value.CleanupIntervalMinutes),
                        stoppingToken);

                    if (!stoppingToken.IsCancellationRequested)
                    {
                        CleanupOldFiles();
                    }
                }
                catch (OperationCanceledException)
                {
                    // Shutdown requested, exit gracefully
                    break;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error during preview cleanup cycle.");
                }
            }

            _logger.LogInformation("Preview cleanup service stopped.");
        }

        /// <summary>
        /// Deletes old preview images and upload files that exceed the retention period.
        /// Files that are currently locked (in use) are skipped.
        /// </summary>
        private void CleanupOldFiles()
        {
            var retention = TimeSpan.FromHours(_options.Value.DeletePreviewAfterHours);
            var cutoff = DateTime.UtcNow - retention;
            int deletedCount = 0;

            // Clean up Preview directory
            var previewPath = Path.Combine(_env.ContentRootPath, _options.Value.PreviewDirectory);
            if (Directory.Exists(previewPath))
            {
                deletedCount += CleanupDirectory(previewPath, "*.png", cutoff);
            }

            // Clean up Uploads directory (any leftover files)
            var uploadsPath = Path.Combine(_env.ContentRootPath, _options.Value.UploadDirectory);
            if (Directory.Exists(uploadsPath))
            {
                deletedCount += CleanupDirectory(uploadsPath, "*.*", cutoff);
            }

            if (deletedCount > 0)
            {
                _logger.LogInformation(
                    "Cleanup completed: {Count} old file(s) deleted (retention: {Retention}h).",
                    deletedCount,
                    _options.Value.DeletePreviewAfterHours);
            }
        }

        /// <summary>
        /// Deletes files in a directory older than the cutoff date.
        /// Skips files that are currently locked by another process.
        /// </summary>
        /// <returns>Number of files successfully deleted.</returns>
        private static int CleanupDirectory(string directoryPath, string searchPattern, DateTime cutoff)
        {
            int deleted = 0;

            try
            {
                var files = Directory.GetFiles(directoryPath, searchPattern);

                foreach (var file in files)
                {
                    try
                    {
                        var lastWrite = System.IO.File.GetLastWriteTimeUtc(file);

                        if (lastWrite < cutoff)
                        {
                            // Check if file is in use by trying to open exclusively
                            using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.None))
                            {
                                // File is not locked, safe to delete
                            }

                            System.IO.File.Delete(file);
                            deleted++;
                        }
                    }
                    catch (IOException)
                    {
                        // File is in use or locked; skip it
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // No permission to delete; skip it
                    }
                }
            }
            catch (DirectoryNotFoundException)
            {
                // Directory doesn't exist yet; nothing to clean
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"Error cleaning directory '{directoryPath}': {ex.Message}");
            }

            return deleted;
        }
    }
}
