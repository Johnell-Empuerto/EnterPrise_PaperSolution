using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using SkiaSharp;

namespace ExcelAPI.LegacyEngine.PublishEngine;

public class BackgroundExporter : IBackgroundExporter
{
    public async Task<string?> ExportAsync(string excelFilePath, int sheetIndex, string outputDir)
    {
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, $"background_{sheetIndex}.pdf");
        string pngPath = Path.Combine(outputDir, $"background_{sheetIndex}.png");

        Application excel = new();
        Workbooks? workbooks = null;
        _Workbook? workbook = null;

        try
        {
            excel.DisplayAlerts = false;
            workbooks = excel.Workbooks;
            workbook = workbooks.Open(excelFilePath);

            if (sheetIndex > 1)
            {
                Worksheet? target = workbook.Sheets[sheetIndex] as Worksheet;
                target?.Activate();
            }

            await Task.Run(() =>
            {
                workbook.ExportAsFixedFormat(
                    Type: XlFixedFormatType.xlTypePDF,
                    Filename: pdfPath,
                    Quality: XlFixedFormatQuality.xlQualityStandard,
                    IncludeDocProperties: true,
                    IgnorePrintAreas: false,
                    OpenAfterPublish: false);
            });

            await ConvertPdfToPng(pdfPath, pngPath);

            return File.Exists(pngPath) ? pngPath : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Background export failed: {ex.Message}", ex);
        }
        finally
        {
            workbook?.Close(false);
            if (workbook != null) Marshal.ReleaseComObject(workbook);
            workbooks?.Close();
            if (workbooks != null) Marshal.ReleaseComObject(workbooks);
            excel.Quit();
            Marshal.ReleaseComObject(excel);
        }
    }

    private static Task ConvertPdfToPng(string pdfPath, string pngPath)
    {
        return Task.Run(() =>
        {
            byte[] pdfBytes = File.ReadAllBytes(pdfPath);
            using var bitmap = PDFtoImage.Conversion.ToImage(
                pdfBytes,
                page: 0,
                options: new PDFtoImage.RenderOptions
                {
                    Dpi = 200,
                    WithAnnotations = false
                });
            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            File.WriteAllBytes(pngPath, data.ToArray());
        });
    }
}
