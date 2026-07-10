namespace ExcelAPI.Runtime
{
    /// <summary>
    /// Represents a single worksheet in the runtime form.
    /// Contains all editable fields, images, shapes, and print area info for this sheet.
    /// </summary>
    public class RuntimeSheet
    {
        /// <summary>Sheet name (e.g., "Sheet1", "Invoice").</summary>
        public string Name { get; set; } = "";

        /// <summary>Sheet index (0-based).</summary>
        public int Index { get; set; }

        /// <summary>Editable fields detected on this sheet.</summary>
        public List<RuntimeField> Fields { get; set; } = new();

        /// <summary>Images on this sheet (from drawing objects).</summary>
        public List<RuntimeImage> Images { get; set; } = new();

        /// <summary>Shapes on this sheet (from drawing objects).</summary>
        public List<RuntimeShape> Shapes { get; set; } = new();

        /// <summary>Print area info, if configured.</summary>
        public RuntimePrintArea? PrintArea { get; set; }

        /// <summary>Page width in pixels (at rendering DPI).</summary>
        public int PageWidthPx { get; set; }

        /// <summary>Page height in pixels (at rendering DPI).</summary>
        public int PageHeightPx { get; set; }
    }

    /// <summary>
    /// Lightweight image reference for runtime consumption.
    /// </summary>
    public class RuntimeImage
    {
        public string Name { get; set; } = "";
        public double LeftPx { get; set; }
        public double TopPx { get; set; }
        public double WidthPx { get; set; }
        public double HeightPx { get; set; }
        public string ContentType { get; set; } = "image/png";
    }

    /// <summary>
    /// Lightweight shape reference for runtime consumption.
    /// </summary>
    public class RuntimeShape
    {
        public string Name { get; set; } = "";
        public string ShapeType { get; set; } = "rectangle";
        public double LeftPx { get; set; }
        public double TopPx { get; set; }
        public double WidthPx { get; set; }
        public double HeightPx { get; set; }
    }

    /// <summary>
    /// Print area bounds for the runtime.
    /// </summary>
    public class RuntimePrintArea
    {
        public string Address { get; set; } = "";
        public double LeftPx { get; set; }
        public double TopPx { get; set; }
        public double WidthPx { get; set; }
        public double HeightPx { get; set; }
    }
}
