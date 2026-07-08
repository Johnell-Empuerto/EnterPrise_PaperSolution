using Microsoft.AspNetCore.Mvc;
using ExcelAPI.Models;
using Microsoft.Extensions.Options;

namespace ExcelAPI.Controllers
{
    /// <summary>
    /// Health check endpoint for monitoring and diagnostics.
    /// Provides a quick status overview of the API and its dependencies.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class HealthController : ControllerBase
    {
        private readonly IWebHostEnvironment _env;
        private readonly IOptions<ExcelCaptureOptions> _options;

        public HealthController(IWebHostEnvironment env, IOptions<ExcelCaptureOptions> options)
        {
            _env = env;
            _options = options;
        }

        /// <summary>
        /// Returns the current health status of the API.
        /// Checks Excel availability, directory existence, and version info.
        /// </summary>
        [HttpGet]
        [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
        public IActionResult Get()
        {
            var uploadsPath = Path.Combine(_env.ContentRootPath, _options.Value.UploadDirectory);
            var previewPath = Path.Combine(_env.ContentRootPath, _options.Value.PreviewDirectory);

            // Check if Excel is installed by attempting to detect the registered type library
            bool excelInstalled = IsExcelInstalled();

            return Ok(new
            {
                status = "Healthy",
                excelInstalled,
                uploadsDirectory = Directory.Exists(uploadsPath),
                previewDirectory = Directory.Exists(previewPath),
                version = "1.3",
                timestamp = DateTime.UtcNow
            });
        }

        /// <summary>
        /// Checks if Microsoft Excel is likely installed on this machine
        /// by looking for the Excel type library in the registry.
        /// </summary>
        private static bool IsExcelInstalled()
        {
            try
            {
                // Attempt to create an Excel Application instance to verify it's available
                // Using a simple registry check is safer than launching Excel
                var excelKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot");

                if (excelKey != null)
                {
                    var path = excelKey.GetValue("Path") as string;
                    return !string.IsNullOrEmpty(path) && System.IO.File.Exists(
                        System.IO.Path.Combine(path, "EXCEL.EXE"));
                }

                // Fall back to checking for Excel type library registration
                var typeLibKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(
                    @"TypeLib\{00020813-0000-0000-C000-000000000046}\1.9");
                return typeLibKey != null;
            }
            catch
            {
                return false;
            }
        }
    }
}
