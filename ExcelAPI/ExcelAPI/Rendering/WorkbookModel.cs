namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Internal workbook model parsed from OpenXML. Used by all rendering engines
    /// (FillEngine, BorderEngine, CellGeometryEngine, TextEngine, etc.).
    /// This is the single source of truth for rendering — not the legacy DB.
    /// </summary>
    public class RenderWorkbook
    {
        public string FilePath { get; set; } = "";
        public List<RenderSheet> Sheets { get; set; } = new();
    }

    public class RenderSheet
    {
        public string Name { get; set; } = "";
        public int Index { get; set; }
        public string? SheetDimension { get; set; }
        public List<RenderColumn> Columns { get; set; } = new();
        public List<RenderRow> Rows { get; set; } = new();
        public List<RenderCell> Cells { get; set; } = new();
        public List<RenderMerge> Merges { get; set; } = new();
        public Dictionary<uint, uint> ColumnOutlineLevels { get; set; } = new();
        public Dictionary<uint, uint> RowOutlineLevels { get; set; } = new();
        public double DefaultColumnWidth { get; set; } = 8.43;
        public double DefaultRowHeight { get; set; } = 15.0;

        // Computed geometry
        public Dictionary<uint, double> ComputedColWidthsPt { get; set; } = new();
        public Dictionary<uint, double> ComputedColCumLeftPt { get; set; } = new();
        public Dictionary<uint, double> ComputedRowHeightsPt { get; set; } = new();
        public Dictionary<uint, double> ComputedRowCumTopPt { get; set; } = new();
        public double TotalWidthPt { get; set; }
        public double TotalHeightPt { get; set; }

        // Drawing objects (images and shapes from drawing.xml)
        public List<DrawingParser.DrawingObject>? DrawingObjects { get; set; }

        // Style references
        public uint? DefaultFontId { get; set; }
        public double MaxDigitWidth { get; set; } = 7.33;
    }

    public class RenderColumn
    {
        public uint Min { get; set; }
        public uint Max { get; set; }
        public double Width { get; set; }          // Character width
        public double PointWidth { get; set; }     // Converted to points
        public bool Hidden { get; set; }
        public bool CustomWidth { get; set; }
        public bool BestFit { get; set; }
        public uint OutlineLevel { get; set; }
    }

    public class RenderRow
    {
        public uint RowIndex { get; set; }
        public double? Height { get; set; }         // Points
        public bool Hidden { get; set; }
        public bool CustomHeight { get; set; }
        public uint OutlineLevel { get; set; }
    }

    public class RenderCell
    {
        public string? Reference { get; set; }
        public uint? StyleIndex { get; set; }
        public string? DataType { get; set; }
        public string? Value { get; set; }
        public string? SharedString { get; set; }
        public uint RowIndex { get; set; }
        public uint ColumnIndex { get; set; }

        // Resolved style properties
        public string? FontName { get; set; }
        public double FontSize { get; set; } = 11;
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikeout { get; set; }
        public string? FontColorArgb { get; set; }

        // Fill
        public string? FillColorArgb { get; set; }
        public string? PatternType { get; set; }  // none, solid, gray125, etc.

        // Border
        public RenderBorder? Border { get; set; }

        // Alignment (stored as string to avoid OpenXml enum conflicts)
        public string? HorizontalAlignment { get; set; }  // "general", "left", "center", "right", etc.
        public string? VerticalAlignment { get; set; }    // "top", "center", "bottom", "justify", etc.
        public bool WrapText { get; set; }
        public uint Indent { get; set; }
        public double TextRotation { get; set; }

        // Is this cell part of a merge
        public int? MergeIndex { get; set; } = null; // Index into sheet.Merges

        // Phase 11E: Resolved style from StyleResolver (cached, single source of truth)
        public ResolvedCellStyle? ResolvedStyle { get; set; }
    }

    public class RenderBorder
    {
        public RenderBorderItem? Top { get; set; }
        public RenderBorderItem? Bottom { get; set; }
        public RenderBorderItem? Left { get; set; }
        public RenderBorderItem? Right { get; set; }
        public RenderBorderItem? DiagonalUp { get; set; }
        public RenderBorderItem? DiagonalDown { get; set; }
    }

    public class RenderBorderItem
    {
        /// <summary>Border style name: "thin", "medium", "thick", "double", "dotted", "dashed", "dashDot", etc.</summary>
        public string Style { get; set; } = "thin";
        public string? ColorArgb { get; set; }
        public bool AutoColor { get; set; }

        /// <summary>Render weight in points (for line width).</summary>
        public double WeightPt
        {
            get
            {
                switch (Style)
                {
                    case "hair": return 0.25;
                    case "dotted": return 0.5;
                    case "dashed": return 0.5;
                    case "dashDot": return 0.5;
                    case "dashDotDot": return 0.5;
                    case "thin": return 0.5;
                    case "medium": return 1.0;
                    case "mediumDashed": return 1.0;
                    case "mediumDashDot": return 1.0;
                    case "mediumDashDotDot": return 1.0;
                    case "thick": return 2.0;
                    case "double": return 2.0;
                    default: return 0.5;
                }
            }
        }

        /// <summary>Whether this is a single solid line (for stroke cap optimization).</summary>
        public bool IsSingleLine
        {
            get
            {
                switch (Style)
                {
                    case "hair":
                    case "thin":
                    case "medium":
                    case "thick":
                        return true;
                    default:
                        return false;
                }
            }
        }
    }

    public class RenderMerge
    {
        public string? Reference { get; set; }
        public uint FirstCol { get; set; }
        public uint FirstRow { get; set; }
        public uint LastCol { get; set; }
        public uint LastRow { get; set; }

        // Computed geometry (in points, uses cumulative column/row positions)
        public double LeftPt { get; set; }
        public double TopPt { get; set; }
        public double WidthPt { get; set; }
        public double HeightPt { get; set; }
    }
}
