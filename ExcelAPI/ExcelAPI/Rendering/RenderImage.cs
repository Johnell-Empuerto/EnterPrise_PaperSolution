using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// A decoded image ready for rendering.
    /// Produced by ImageResolver from xl/media/ files.
    /// </summary>
    public class RenderImage
    {
        /// <summary>The decoded SkiaSharp bitmap.</summary>
        public SKBitmap Bitmap { get; set; } = null!;

        /// <summary>Relationship ID for cache lookup.</summary>
        public string RelationshipId { get; set; } = "";

        /// <summary>Content type (image/png, image/jpeg, etc.).</summary>
        public string ContentType { get; set; } = "";

        /// <summary>Original file name in xl/media/. </summary>
        public string FileName { get; set; } = "";

        /// <summary>Pixel width of the decoded bitmap.</summary>
        public int PixelWidth => Bitmap.Width;

        /// <summary>Pixel height of the decoded bitmap.</summary>
        public int PixelHeight => Bitmap.Height;
    }
}
