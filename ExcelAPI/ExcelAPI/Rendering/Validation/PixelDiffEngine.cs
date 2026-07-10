using SkiaSharp;

namespace ExcelAPI.Rendering.Validation
{
    /// <summary>
    /// Per-pixel image comparison engine.
    /// Compares two SKBitmaps pixel-by-pixel and produces:
    ///   - Difference statistics (total pixels, different pixels, percentage, max/average error)
    ///   - Bounding rectangles of changed regions
    ///   - A diff image highlighting differences (accessible via PixelDiffResult.DiffBitmap)
    ///
    /// Uses safe GetPixel/SetPixel API — no unsafe pointer access required.
    /// Supports configurable tolerance per channel and optional alpha channel skipping.
    /// </summary>
    public class PixelDiffEngine
    {
        /// <summary>
        /// Result of a pixel-diff comparison between two images.
        /// </summary>
        public class PixelDiffResult
        {
            /// <summary>Whether the images are considered identical within tolerance.</summary>
            public bool Passed { get; set; }

            /// <summary>Total pixel count.</summary>
            public int TotalPixels { get; set; }

            /// <summary>Number of pixels that differ beyond tolerance.</summary>
            public int DifferentPixels { get; set; }

            /// <summary>Percentage of different pixels (0.0-100.0).</summary>
            public double DifferencePercent { get; set; }

            /// <summary>Maximum per-channel error across all pixels (0-255).</summary>
            public int MaxError { get; set; }

            /// <summary>Average per-channel error across all different pixels.</summary>
            public double AverageError { get; set; }

            /// <summary>Width of the compared images.</summary>
            public int Width { get; set; }

            /// <summary>Height of the compared images.</summary>
            public int Height { get; set; }

            /// <summary>Bounding rectangle of all changed regions (or empty if none).</summary>
            public SKRectI ChangedBounds { get; set; }

            /// <summary>The diff bitmap (null if generateDiffImage was false).</summary>
            public SKBitmap? DiffBitmap { get; set; }

            /// <summary>Path to the saved diff image (set by caller).</summary>
            public string? DiffImagePath { get; set; }
        }

        /// <summary>
        /// Compare two images pixel-by-pixel within the configured tolerance.
        /// </summary>
        /// <param name="expected">The expected (baseline) image.</param>
        /// <param name="actual">The actual (FormLess-rendered) image.</param>
        /// <param name="tolerance">Per-channel tolerance (0 = exact).</param>
        /// <param name="skipAlpha">Whether to skip alpha channel comparison.</param>
        /// <param name="generateDiffImage">Whether to generate a diff image bitmap.</param>
        /// <returns>PixelDiffResult with comparison statistics and optionally a diff bitmap.</returns>
        public PixelDiffResult Compare(
            SKBitmap expected,
            SKBitmap actual,
            int tolerance = 1,
            bool skipAlpha = false,
            bool generateDiffImage = true)
        {
            if (expected == null || actual == null)
                throw new ArgumentNullException(expected == null ? nameof(expected) : nameof(actual));

            int width = Math.Min(expected.Width, actual.Width);
            int height = Math.Min(expected.Height, actual.Height);

            int totalPixels = width * height;
            int diffPixels = 0;
            int maxError = 0;
            double totalErrorSum = 0;
            int minChangedX = width, minChangedY = height;
            int maxChangedX = 0, maxChangedY = 0;

            // Create diff bitmap if requested
            SKBitmap? diffBitmap = generateDiffImage ? new SKBitmap(width, height) : null;

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    // Get pixel colors (safe API — no unsafe code needed)
                    SKColor expColor = expected.GetPixel(x, y);
                    SKColor actColor = actual.GetPixel(x, y);

                    // Compute per-channel difference
                    int diffR = Math.Abs(expColor.Red - actColor.Red);
                    int diffG = Math.Abs(expColor.Green - actColor.Green);
                    int diffB = Math.Abs(expColor.Blue - actColor.Blue);
                    int diffA = skipAlpha ? 0 : Math.Abs(expColor.Alpha - actColor.Alpha);

                    int maxChannelError = Math.Max(Math.Max(diffR, diffG), Math.Max(diffB, diffA));
                    bool isDifferent = maxChannelError > tolerance;

                    if (isDifferent)
                    {
                        diffPixels++;
                        if (maxChannelError > maxError) maxError = maxChannelError;
                        totalErrorSum += maxChannelError;

                        if (x < minChangedX) minChangedX = x;
                        if (y < minChangedY) minChangedY = y;
                        if (x > maxChangedX) maxChangedX = x;
                        if (y > maxChangedY) maxChangedY = y;

                        // Color the diff pixel
                        if (diffBitmap != null)
                        {
                            SKColor diffColor;
                            if (maxChannelError > 10)
                                diffColor = new SKColor(255, 0, 255, 255); // magenta
                            else if (maxChannelError > 3)
                                diffColor = new SKColor(255, 0, 0, 255);   // red
                            else
                                diffColor = new SKColor(255, 128, 0, 255); // orange
                            diffBitmap.SetPixel(x, y, diffColor);
                        }
                    }
                    else if (diffBitmap != null)
                    {
                        // Identical pixels → transparent
                        diffBitmap.SetPixel(x, y, SKColors.Transparent);
                    }
                }
            }

            double diffPercent = totalPixels > 0
                ? (double)diffPixels / totalPixels * 100.0
                : 0.0;

            double avgError = diffPixels > 0
                ? totalErrorSum / diffPixels
                : 0.0;

            bool hasChanges = diffPixels > 0;
            var changedBounds = hasChanges
                ? new SKRectI(minChangedX, minChangedY, maxChangedX + 1, maxChangedY + 1)
                : SKRectI.Empty;

            return new PixelDiffResult
            {
                Passed = diffPercent <= 0.5,
                TotalPixels = totalPixels,
                DifferentPixels = diffPixels,
                DifferencePercent = Math.Round(diffPercent, 4),
                MaxError = maxError,
                AverageError = Math.Round(avgError, 2),
                Width = width,
                Height = height,
                ChangedBounds = changedBounds,
                DiffBitmap = diffBitmap,
                DiffImagePath = null
            };
        }

        /// <summary>
        /// Save a diff bitmap to disk.
        /// </summary>
        public static void SaveDiffImage(SKBitmap diffBitmap, string outputPath)
        {
            if (diffBitmap == null) return;

            var dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using var image = SKImage.FromBitmap(diffBitmap);
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            System.IO.File.WriteAllBytes(outputPath, data.ToArray());
        }

        /// <summary>
        /// Create a side-by-side comparison image (expected | actual | diff).
        /// </summary>
        public static SKBitmap CreateSideBySide(SKBitmap expected, SKBitmap actual, SKBitmap? diff)
        {
            int w = expected.Width;
            int h = expected.Height;
            int totalW = w * (diff != null ? 3 : 2);

            var result = new SKBitmap(totalW, h);
            using var canvas = new SKCanvas(result);

            canvas.DrawBitmap(expected, 0, 0);
            canvas.DrawBitmap(actual, w, 0);
            if (diff != null)
                canvas.DrawBitmap(diff, w * 2, 0);

            return result;
        }
    }
}
