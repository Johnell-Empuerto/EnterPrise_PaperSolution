using SkiaSharp;

namespace ExcelAPI.Designer.Legacy.VerificationEngine;

public class ImageComparer : IImageComparer
{
    private readonly ILogger<ImageComparer> _logger;

    public ImageComparer(ILogger<ImageComparer> logger)
    {
        _logger = logger;
    }

    public ImageComparisonResult Compare(string generatedPath, byte[]? databaseImageBytes)
    {
        var result = new ImageComparisonResult
        {
            GeneratedImageExists = File.Exists(generatedPath),
            DatabaseImageExists = databaseImageBytes != null && databaseImageBytes.Length > 0
        };

        if (!result.GeneratedImageExists && !result.DatabaseImageExists)
        {
            result.Status = "SKIP";
            return result;
        }

        if (!result.GeneratedImageExists)
        {
            result.Status = "FAIL";
            return result;
        }

        if (!result.DatabaseImageExists)
        {
            result.Status = "PARTIAL";
            return result;
        }

        try
        {
            using var genBitmap = SKBitmap.Decode(generatedPath);
            using var dbBitmap = SKBitmap.Decode(databaseImageBytes);

            if (genBitmap == null || dbBitmap == null)
            {
                result.Status = "FAIL";
                return result;
            }

            result.GeneratedWidth = genBitmap.Width;
            result.GeneratedHeight = genBitmap.Height;
            result.DatabaseWidth = dbBitmap.Width;
            result.DatabaseHeight = dbBitmap.Height;

            using var resizedGen = ResizeToMatch(genBitmap, dbBitmap.Width, dbBitmap.Height);

            var mse = ComputeMse(resizedGen, dbBitmap);
            var pixelDiff = ComputePixelDiffPercent(resizedGen, dbBitmap);
            var ssim = ComputeApproximateSsim(resizedGen, dbBitmap);

            result.Mse = mse;
            result.PixelDiffPercent = pixelDiff;
            result.Ssim = ssim;

            double mseScore = mse > 0 ? Math.Max(0, 1.0 - mse / 10000.0) : 1.0;
            double pixelScore = pixelDiff.HasValue ? (1.0 - pixelDiff.Value / 100.0) : 1.0;
            double ssimScore = ssim;
            result.SimilarityScore = mseScore * 0.2 + pixelScore * 0.3 + ssimScore * 0.5;

            result.Status = result.SimilarityScore >= 0.95 ? "PASS"
                : result.SimilarityScore >= 0.7 ? "PARTIAL" : "FAIL";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Image comparison failed");
            result.Status = "FAIL";
        }

        return result;
    }

    private static SKBitmap ResizeToMatch(SKBitmap source, int targetWidth, int targetHeight)
    {
        if (source.Width == targetWidth && source.Height == targetHeight)
            return source;

        var resized = new SKBitmap(targetWidth, targetHeight);
        using var canvas = new SKCanvas(resized);
        canvas.DrawBitmap(source, new SKRect(0, 0, targetWidth, targetHeight));
        return resized;
    }

    private static double ComputeMse(SKBitmap a, SKBitmap b)
    {
        if (a.Width != b.Width || a.Height != b.Height) return double.MaxValue;

        double sum = 0;
        int count = a.Width * a.Height * 3;

        for (int y = 0; y < a.Height; y++)
        {
            for (int x = 0; x < a.Width; x++)
            {
                var pa = a.GetPixel(x, y);
                var pb = b.GetPixel(x, y);
                sum += Math.Pow(pa.Red - pb.Red, 2);
                sum += Math.Pow(pa.Green - pb.Green, 2);
                sum += Math.Pow(pa.Blue - pb.Blue, 2);
            }
        }

        return count > 0 ? sum / count : 0;
    }

    private static double? ComputePixelDiffPercent(SKBitmap a, SKBitmap b)
    {
        if (a.Width != b.Width || a.Height != b.Height) return null;

        int diffPixels = 0;
        int totalPixels = a.Width * a.Height;

        for (int y = 0; y < a.Height; y++)
        {
            for (int x = 0; x < a.Width; x++)
            {
                if (a.GetPixel(x, y) != b.GetPixel(x, y))
                    diffPixels++;
            }
        }

        return totalPixels > 0 ? (double)diffPixels / totalPixels * 100.0 : 0;
    }

    private static double ComputeApproximateSsim(SKBitmap a, SKBitmap b)
    {
        int width = Math.Min(a.Width, b.Width);
        int height = Math.Min(a.Height, b.Height);

        const double C1 = 6.5025;
        const double C2 = 58.5225;
        const int BlockSize = 8;

        double totalSsim = 0;
        int blockCount = 0;

        for (int y = 0; y + BlockSize <= height; y += BlockSize)
        {
            for (int x = 0; x + BlockSize <= width; x += BlockSize)
            {
                double muX = 0, muY = 0;
                for (int dy = 0; dy < BlockSize; dy++)
                {
                    for (int dx = 0; dx < BlockSize; dx++)
                    {
                        var pa = a.GetPixel(x + dx, y + dy);
                        var pb = b.GetPixel(x + dx, y + dy);
                        double grayA = 0.299 * pa.Red + 0.587 * pa.Green + 0.114 * pa.Blue;
                        double grayB = 0.299 * pb.Red + 0.587 * pb.Green + 0.114 * pb.Blue;
                        muX += grayA;
                        muY += grayB;
                    }
                }

                double n = BlockSize * BlockSize;
                muX /= n;
                muY /= n;

                double sigmaX2 = 0, sigmaY2 = 0, sigmaXY = 0;
                for (int dy = 0; dy < BlockSize; dy++)
                {
                    for (int dx = 0; dx < BlockSize; dx++)
                    {
                        var pa = a.GetPixel(x + dx, y + dy);
                        var pb = b.GetPixel(x + dx, y + dy);
                        double grayA = 0.299 * pa.Red + 0.587 * pa.Green + 0.114 * pa.Blue;
                        double grayB = 0.299 * pb.Red + 0.587 * pb.Green + 0.114 * pb.Blue;
                        sigmaX2 += Math.Pow(grayA - muX, 2);
                        sigmaY2 += Math.Pow(grayB - muY, 2);
                        sigmaXY += (grayA - muX) * (grayB - muY);
                    }
                }

                sigmaX2 /= n;
                sigmaY2 /= n;
                sigmaXY /= n;

                double ssim = (2 * muX * muY + C1) * (2 * sigmaXY + C2)
                            / ((muX * muX + muY * muY + C1) * (sigmaX2 + sigmaY2 + C2));
                totalSsim += ssim;
                blockCount++;
            }
        }

        return blockCount > 0 ? totalSsim / blockCount : 0;
    }
}
