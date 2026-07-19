// ────────────────────────────────────────────────────────────────────────────
// WbDefConverter — WorkbookDefinition → Rendering Adapter
//
// Bridges the canonical WorkbookDefinition into the Rendering layer's internal
// types (ResolvedCellStyle, PrintLayoutResult). Rendering now consumes
// WorkbookDefinition directly — no intermediate RenderWorkbook conversion.
//
// Phase 4.1: Removed ToRenderWorkbook() — WbDef is now consumed directly by
//            Rendering. Only style and print layout bridges remain.
//
// Ownership: Rendering (adapter layer)
// ────────────────────────────────────────────────────────────────────────────

using WbDef = ExcelAPI.Models.WorkbookDefinition;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Bridges WorkbookDefinition domain types into Rendering internal models.
    /// Remaining adapters: ToResolvedCellStyle (style bridge),
    /// ToPrintLayoutResult (print layout bridge).
    /// Phase 4.1: ToRenderWorkbook removed — no longer needed.
    /// </summary>
    public static class WbDefConverter
    {
        // ── Style Bridge ─────────────────────────────────────────────────

        /// <summary>
        /// Convert a canonical WbDef CellStyle into the Rendering layer's ResolvedCellStyle.
        /// Enables the rendering pipeline to consume WbDef styles without re-parsing OOXML.
        /// </summary>
        public static ResolvedCellStyle ToResolvedCellStyle(WbDef.CellStyle? style)
        {
            var resolved = new ResolvedCellStyle();
            if (style == null) return resolved;

            // Font
            if (style.Font != null)
            {
                resolved.FontName = style.Font.Name ?? "Calibri";
                resolved.FontSize = style.Font.SizePt > 0 ? style.Font.SizePt : 11;
                resolved.Bold = style.Font.Bold;
                resolved.Italic = style.Font.Italic;
                resolved.Underline = style.Font.Underline;
                resolved.Strikeout = style.Font.Strikeout;
                resolved.FontColorArgb = style.Font.ColorArgb;
            }

            // Fill
            if (style.Fill != null)
            {
                resolved.PatternType = style.Fill.PatternType;
                resolved.FillColorArgb = style.Fill.ColorArgb;
            }

            // Border
            if (style.Border != null)
            {
                resolved.Border = new ResolvedBorder
                {
                    Left = ToBorderItem(style.Border.Left),
                    Right = ToBorderItem(style.Border.Right),
                    Top = ToBorderItem(style.Border.Top),
                    Bottom = ToBorderItem(style.Border.Bottom),
                    DiagonalUp = ToBorderItem(style.Border.DiagonalUp),
                    DiagonalDown = ToBorderItem(style.Border.DiagonalDown)
                };
            }

            // Alignment
            if (style.Alignment != null)
            {
                resolved.HorizontalAlignment = style.Alignment.Horizontal;
                resolved.VerticalAlignment = style.Alignment.Vertical;
            }
            resolved.WrapText = style.WrapText;
            resolved.Indent = style.Indent;
            resolved.TextRotation = style.TextRotation;

            return resolved;
        }

        private static ResolvedBorderItem? ToBorderItem(WbDef.BorderEdge? edge)
        {
            if (edge == null) return null;
            return new ResolvedBorderItem
            {
                Style = string.IsNullOrEmpty(edge.Style) ? "thin" : edge.Style,
                ColorArgb = edge.ColorArgb
            };
        }

        // ── Print Layout Bridge ──────────────────────────────────────────

        /// <summary>
        /// Convert a canonical WbDef PrintLayout into the Rendering layer's PrintLayoutResult.
        /// Delegates to PrintLayoutEngine.Compute() — the single authoritative implementation.
        /// This adapter only maps properties; it does NOT duplicate layout computation.
        /// </summary>
        /// <param name="layout">WbDef PrintLayout to convert (null → default Letter portrait).</param>
        /// <param name="engine">PrintLayoutEngine instance (injected by caller).</param>
        /// <param name="totalContentWidthPt">Total content width in points (required for centering/scaling).</param>
        /// <param name="totalContentHeightPt">Total content height in points (required for centering/scaling).</param>
        public static PrintLayoutResult ToPrintLayoutResult(
            WbDef.PrintLayout? layout,
            PrintLayoutEngine engine,
            double totalContentWidthPt = 0,
            double totalContentHeightPt = 0)
        {
            if (layout == null)
            {
                // Default to Letter portrait with default margins
                return engine.Compute(
                    paperSize: "Letter",
                    orientation: "portrait",
                    leftMargin: 50.4, rightMargin: 50.4,
                    topMargin: 54.0, bottomMargin: 54.0,
                    centerHorizontally: false, centerVertically: false,
                    zoom: 100, fitToPagesWide: 0, fitToPagesTall: 0,
                    totalContentWidthPt: totalContentWidthPt,
                    totalContentHeightPt: totalContentHeightPt
                );
            }

            string orientation = layout.Orientation == WbDef.PageOrientation.Landscape
                ? "landscape" : "portrait";

            return engine.Compute(
                paperSize: layout.PaperSize?.Name ?? "Letter",
                orientation: orientation,
                leftMargin: layout.Margins?.LeftPt ?? 50.4,
                rightMargin: layout.Margins?.RightPt ?? 50.4,
                topMargin: layout.Margins?.TopPt ?? 54.0,
                bottomMargin: layout.Margins?.BottomPt ?? 54.0,
                centerHorizontally: layout.Scaling?.CenterHorizontally ?? false,
                centerVertically: layout.Scaling?.CenterVertically ?? false,
                zoom: layout.Scaling?.Zoom ?? 100,
                fitToPagesWide: layout.Scaling?.FitToPagesWide ?? 0,
                fitToPagesTall: layout.Scaling?.FitToPagesTall ?? 0,
                totalContentWidthPt: totalContentWidthPt,
                totalContentHeightPt: totalContentHeightPt
            );
        }

        // Phase 4.1: ToRenderWorkbook removed — Rendering now consumes WorkbookDefinition
        // directly through the Runtime and Rendering pipelines. No intermediate
        // RenderWorkbook conversion is required.
    }
}
