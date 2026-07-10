using SkiaSharp;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// A single rendering layer in the FormLess rendering pipeline.
    /// Each layer draws one visual element (fills, borders, text, images, etc.)
    /// and receives the full RenderingContext for coordinate lookups.
    ///
    /// Layers are executed in order by RendererCoordinator:
    ///   1. FillLayer
    ///   2. GridlineLayer
    ///   3. BorderLayer
    ///   4. TextLayer (future)
    ///   5. ImageLayer (future)
    ///   6. ShapeLayer (future)
    ///   7. AnnotationLayer (future)
    /// </summary>
    public interface IRenderLayer
    {
        /// <summary>
        /// Render this layer onto the SkiaSharp canvas.
        /// </summary>
        /// <param name="canvas">The SkiaSharp canvas to draw on.</param>
        /// <param name="context">Shared rendering context (workbook, sheet, origin, DPI).</param>
        void Render(SKCanvas canvas, RenderingContext context);

        /// <summary>
        /// Display name for logging and diagnostics.
        /// </summary>
        string Name { get; }
    }
}
