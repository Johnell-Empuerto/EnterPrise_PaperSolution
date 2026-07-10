using DocumentFormat.OpenXml.Packaging;
using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Extracts and caches images from xl/media/ via worksheet drawing relationships.
    /// Images are decoded once and cached by relationship ID.
    /// Supports PNG, JPEG, GIF, BMP, TIFF via SkiaSharp's built-in decoders.
    /// </summary>
    public class ImageResolver
    {
        private readonly Dictionary<string, RenderImage> _cache = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Resolve an image by relationship ID from the worksheet's drawing part.
        /// Returns null if the image part cannot be found or decoded.
        /// </summary>
        public RenderImage? Resolve(WorksheetPart wsPart, string relId)
        {
            // Check cache first
            if (_cache.TryGetValue(relId, out var cached))
                return cached;

            // WorksheetPart.DrawingsPart returns a single DrawingsPart (or null)
            var dp = wsPart.DrawingsPart;
            if (dp == null) return null;

            // Find the image part by relationship ID
            var imgPart = dp.GetPartById(relId) as ImagePart;
            if (imgPart == null) return null;

            try
            {
                using var stream = imgPart.GetStream();
                if (stream == null || stream.Length == 0) return null;

                var bitmap = SKBitmap.Decode(stream);
                if (bitmap == null) return null;

                // Use Segments instead of UriSegments (CS1061)
                var segments = imgPart.Uri?.Segments;
                var fileName = (segments != null && segments.Length > 0)
                    ? segments[^1]
                    : relId;

                var renderImage = new RenderImage
                {
                    Bitmap = bitmap,
                    RelationshipId = relId,
                    ContentType = imgPart.ContentType ?? "image/png",
                    FileName = fileName
                };

                _cache[relId] = renderImage;
                return renderImage;
            }
            catch (Exception ex)
            {
                // Log but don't crash — image may be corrupt or unsupported
                System.Diagnostics.Debug.WriteLine(
                    $"[ImageResolver] Failed to decode image part {relId}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Clear the image cache (e.g., when a new workbook is loaded).
        /// </summary>
        public void ClearCache()
        {
            _cache.Clear();
        }
    }
}
