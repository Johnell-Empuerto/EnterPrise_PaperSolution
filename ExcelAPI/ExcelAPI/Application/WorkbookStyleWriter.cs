using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using WbDef = ExcelAPI.Models.WorkbookDefinition;

namespace ExcelAPI.Application
{
    /// <summary>
    /// Phase 22 — Browser Style Persistence.
    ///
    /// Applies user-edited cell styles (font, fill, alignment) to cells in
    /// the Excel workbook after WorkbookValueWriter has written cell values.
    ///
    /// Works with the existing OpenXML Stylesheet — creates new CellFormat
    /// records (xf) for each unique style and sets the cell's StyleIndex
    /// to reference the appropriate format.
    ///
    /// Rules:
    ///   1. Never modify existing styles in the stylesheet (no side effects)
    ///   2. Create new CellFormat records for each unique style combination
    ///   3. Deduplicate identical style requests (reuse existing xf indices)
    ///   4. Preserve existing cell values and formulas
    ///   5. Only modify cells that have explicit style overrides
    /// </summary>
    public class WorkbookStyleWriter
    {
        private readonly ILogger<WorkbookStyleWriter> _logger;

        public WorkbookStyleWriter(ILogger<WorkbookStyleWriter> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Apply style changes from WorkbookDefinition fields to cells in the workbook.
        /// Called after WorkbookValueWriter.WriteValues().
        /// </summary>
        /// <param name="wbDef">WorkbookDefinition with field style properties.</param>
        /// <param name="workbookPath">Path to the already-saved workbook.</param>
        /// <returns>Number of cells with styles applied.</returns>
        public int ApplyStyles(WbDef.WorkbookDefinition wbDef, string workbookPath)
        {
            if (!System.IO.File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook not found: {workbookPath}");

            int totalStyled = 0;

            _logger.LogInformation("=========================================================");
            _logger.LogInformation("WORKBOOK STYLE WRITER — Phase 22");
            _logger.LogInformation("=========================================================");
            _logger.LogInformation("Workbook: {Path}", workbookPath);
            _logger.LogInformation("Sheets: {Count}", wbDef.Sheets.Count);

            using (var doc = SpreadsheetDocument.Open(workbookPath, true))
            {
                if (doc.WorkbookPart == null)
                    throw new InvalidOperationException("WorkbookPart is null — file may be corrupt.");

                var wbPart = doc.WorkbookPart;

                // Get or create WorkbookStylesPart
                var stylesPart = wbPart.WorkbookStylesPart;
                if (stylesPart == null)
                {
                    stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
                    stylesPart.Stylesheet = new Stylesheet();
                }
                var stylesheet = stylesPart.Stylesheet;

                // Ensure required collections exist
                if (stylesheet.Fonts == null)
                    stylesheet.Fonts = new Fonts();
                if (stylesheet.Fills == null)
                    stylesheet.Fills = new Fills();
                if (stylesheet.Borders == null)
                    stylesheet.Borders = new Borders();
                if (stylesheet.CellFormats == null)
                    stylesheet.CellFormats = new CellFormats();

                // Ensure StyleSheet has minimum defaults (fonts, fills)
                EnsureMinimalStylesheet(stylesheet);

                // Build sheet name → WorksheetPart mapping
                var sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                var sheetParts = new Dictionary<string, WorksheetPart>(StringComparer.OrdinalIgnoreCase);
                foreach (var sheet in sheets)
                {
                    if (sheet.Id == null || sheet.Name == null) continue;
                    var wsPart = wbPart.GetPartById(sheet.Id) as WorksheetPart;
                    if (wsPart != null)
                        sheetParts[sheet.Name] = wsPart;
                }

                // Process each field that has style properties
                foreach (var wbSheet in wbDef.Sheets)
                {
                    if (!sheetParts.TryGetValue(wbSheet.Name, out var wsPart)) continue;

                    var worksheet = wsPart.Worksheet;
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData == null) continue;

                    foreach (var field in wbSheet.Fields)
                    {
                        if (field.Style == null) continue;
                        bool hasStyle = false;

                        // Check if there's anything to apply
                        var s = field.Style;
                        bool hasFont = s.Font != null &&
                            (!string.IsNullOrEmpty(s.Font.Name) ||
                             s.Font.SizePt > 0 ||
                             s.Font.Bold ||
                             s.Font.Italic ||
                             !string.IsNullOrEmpty(s.Font.ColorArgb));
                        bool hasFill = s.Fill != null &&
                            !string.IsNullOrEmpty(s.Fill.ColorArgb) &&
                            s.Fill.PatternType != "none";
                        bool hasAlign = s.Alignment != null &&
                            (!string.IsNullOrEmpty(s.Alignment.Horizontal) ||
                             !string.IsNullOrEmpty(s.Alignment.Vertical));
                        bool hasWrap = s.WrapText;
                        hasStyle = hasFont || hasFill || hasAlign || hasWrap;

                        if (!hasStyle) continue;

                        string cellRef = field.Cell?.Address ?? "";
                        if (string.IsNullOrEmpty(cellRef)) continue;

                        uint rowIndex = (uint)field.Cell.RowIndex;
                        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
                        if (row == null) continue;

                        var cell = row.Elements<Cell>().FirstOrDefault(c =>
                            string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
                        if (cell == null) continue;

                        // Find or create a CellFormat (xf) matching this style
                        uint styleIndex = FindOrCreateCellFormat(stylesheet, field.Style);

                        // Apply the style index to the cell
                        cell.StyleIndex = styleIndex;

                        // ═════════════════════════════════════════════════════════
                        // PHASE 22.1 — STAGE 24: STYLE WRITE (Applied to Cell)
                        // ═════════════════════════════════════════════════════════
                        _logger.LogInformation("========================================================");
                        _logger.LogInformation("  STAGE 24 — Style Write");
                        _logger.LogInformation("========================================================");
                        _logger.LogInformation("  Applying");
                        _logger.LogInformation($"  Cell:        {cellRef}");
                        _logger.LogInformation($"  ✓ Font:       {s.Font?.Name ?? "(default)"}");
                        _logger.LogInformation($"  ✓ Size:       {s.Font?.SizePt ?? 0}");
                        _logger.LogInformation($"  ✓ Bold:       {s.Font?.Bold ?? false}");
                        _logger.LogInformation($"  ✓ Italic:     {s.Font?.Italic ?? false}");
                        _logger.LogInformation($"  ✓ Color:      {s.Font?.ColorArgb ?? "(default)"}");
                        _logger.LogInformation($"  ✓ Fill:       {s.Fill?.ColorArgb ?? "(none)"}");
                        _logger.LogInformation($"  ✓ H-Align:    {s.Alignment?.Horizontal ?? "(default)"}");
                        _logger.LogInformation($"  ✓ V-Align:    {s.Alignment?.Vertical ?? "(default)"}");
                        _logger.LogInformation($"  ✓ Wrap:       {s.WrapText}");
                        _logger.LogInformation($"  ✓ xfIdx:      {styleIndex}");
                        _logger.LogInformation("========================================================");

                        totalStyled++;
                    }

                    wsPart.Worksheet.Save();
                }

                // Save stylesheet
                stylesPart.Stylesheet.Save();
            }

            _logger.LogInformation("=========================================================");
            _logger.LogInformation("PHASE 22 COMPLETE: {Count} cells styled", totalStyled);
            _logger.LogInformation("=========================================================");

            return totalStyled;
        }

        /// <summary>
        /// Ensure the stylesheet has at least the minimum required records:
        /// - 1 Font (default)
        /// - 2 Fills (none + gray125)
        /// - 1 Border (default)
        /// - 1 CellFormat (xf) referencing them
        /// These are required by the OPC specification.
        /// </summary>
        private static void EnsureMinimalStylesheet(Stylesheet stylesheet)
        {
            // Fonts: need at least 1
            if (stylesheet.Fonts.Count() == 0)
            {
                stylesheet.Fonts.AppendChild(new Font());
            }

            // Fills: need at least 2 (none + gray125)
            if (stylesheet.Fills.Count() == 0)
            {
                stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } });
                stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } });
            }
            else if (stylesheet.Fills.Count() == 1)
            {
                // Only one fill exists — add gray125 as second
                stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } });
            }

            // Borders: need at least 1
            if (stylesheet.Borders.Count() == 0)
            {
                stylesheet.Borders.AppendChild(new Border());
            }

            // CellFormats: need at least 1 (default format)
            if (stylesheet.CellFormats.Count() == 0)
            {
                stylesheet.CellFormats.AppendChild(new CellFormat { FontId = 0, FillId = 0, BorderId = 0 });
            }

            // Update count attributes
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
        }

        /// <summary>
        /// Find an existing CellFormat (xf) that matches the requested style,
        /// or create a new one. Returns the index into the CellFormats collection.
        /// </summary>
        private static uint FindOrCreateCellFormat(Stylesheet stylesheet, WbDef.CellStyle style)
        {
            var cellFormats = stylesheet.CellFormats;
            var fonts = stylesheet.Fonts;
            var fills = stylesheet.Fills;
            var borders = stylesheet.Borders;

            // Find or create a Font matching the requested font properties
            uint fontId = FindOrCreateFont(fonts, style.Font);
            // Find or create a Fill matching the requested fill/background
            uint fillId = FindOrCreateFill(fills, style.Fill);
            // Use default border (index 0) — we don't modify borders
            uint borderId = 0;

            // Build the alignment we need
            HorizontalAlignmentValues? hAlign = null;
            VerticalAlignmentValues? vAlign = null;
            bool wrapText = style.WrapText;

            if (style.Alignment != null)
            {
                if (!string.IsNullOrEmpty(style.Alignment.Horizontal))
                    hAlign = ParseHorizontalAlignment(style.Alignment.Horizontal);
                if (!string.IsNullOrEmpty(style.Alignment.Vertical))
                    vAlign = ParseVerticalAlignment(style.Alignment.Vertical);
            }

            // Look for an existing CellFormat with the same font, fill, border, alignment
            for (uint i = 0; i < cellFormats.Count(); i++)
            {
                var xf = cellFormats.ElementAt((int)i) as CellFormat;
                if (xf == null) continue;

                if (xf.FontId?.Value == fontId &&
                    xf.FillId?.Value == fillId &&
                    xf.BorderId?.Value == borderId)
                {
                    // Check alignment match
                    bool alignMatch = true;
                    if (hAlign.HasValue || vAlign.HasValue || wrapText)
                    {
                        var xfAlign = xf.Alignment;
                        if (xfAlign == null)
                        {
                            // We need alignment but existing xf has none
                            if (hAlign.HasValue || vAlign.HasValue || wrapText)
                                alignMatch = false;
                        }
                        else
                        {
                            if (hAlign.HasValue && xfAlign.Horizontal?.Value != hAlign.Value)
                                alignMatch = false;
                            if (vAlign.HasValue && xfAlign.Vertical?.Value != vAlign.Value)
                                alignMatch = false;
                            if (wrapText && (xfAlign.WrapText?.Value ?? false) != wrapText)
                                alignMatch = false;
                        }
                    }
                    else if (xf.Alignment != null)
                    {
                        // We have no alignment request but xf has non-default alignment
                        if (xf.Alignment.Horizontal != null || xf.Alignment.Vertical != null || (xf.Alignment.WrapText?.Value ?? false))
                            alignMatch = false;
                    }

                    if (alignMatch)
                        return i;
                }
            }

            // Create a new CellFormat
            var newXf = new CellFormat
            {
                FontId = fontId,
                FillId = fillId,
                BorderId = borderId,
            };

            // Apply alignment if specified
            if (hAlign.HasValue || vAlign.HasValue || wrapText)
            {
                newXf.Alignment = new Alignment();
                if (hAlign.HasValue)
                    newXf.Alignment.Horizontal = hAlign.Value;
                if (vAlign.HasValue)
                    newXf.Alignment.Vertical = vAlign.Value;
                if (wrapText)
                    newXf.Alignment.WrapText = wrapText;
            }

            cellFormats.AppendChild(newXf);
            cellFormats.Count = (uint)cellFormats.Count();

            return (uint)(cellFormats.Count() - 1);
        }

        /// <summary>
        /// Find an existing Font matching the requested properties, or create a new one.
        /// </summary>
        private static uint FindOrCreateFont(Fonts fonts, WbDef.FontDefinition? fontDef)
        {
            if (fontDef == null) return 0;

            string targetName = fontDef.Name ?? "";
            double targetSize = fontDef.SizePt;
            bool targetBold = fontDef.Bold;
            bool targetItalic = fontDef.Italic;
            bool targetUnderline = fontDef.Underline;
            string targetColor = fontDef.ColorArgb ?? "";

            // Look for an existing matching font
            for (uint i = 0; i < fonts.Count(); i++)
            {
                var f = fonts.ElementAt((int)i) as Font;
                if (f == null) continue;

                string fName = f.FontName?.Val?.Value ?? "";
                double fSize = f.FontSize?.Val?.Value ?? 0;
                bool fBold = f.Bold?.Val?.Value ?? false;
                bool fItalic = f.Italic?.Val?.Value ?? false;
                bool fUnder = f.Underline?.Val?.Value != null;

                if (fName == targetName &&
                    Math.Abs(fSize - targetSize) < 0.01 &&
                    fBold == targetBold &&
                    fItalic == targetItalic &&
                    fUnder == targetUnderline)
                {
                    // Check color
                    string fColor = "";
                    var colorEl = f.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>();
                    if (colorEl?.Rgb?.Value != null)
                        fColor = "#" + colorEl.Rgb.Value;
                    else if (colorEl?.Indexed?.Value != null)
                        fColor = $"indexed={colorEl.Indexed.Value}";

                    if (fColor == targetColor)
                        return i;
                }
            }

            // Create a new Font
            var newFont = new Font();

            if (!string.IsNullOrEmpty(targetName))
                newFont.AppendChild(new FontName { Val = targetName });

            if (targetSize > 0)
                newFont.AppendChild(new FontSize { Val = targetSize });

            if (targetBold)
                newFont.AppendChild(new Bold());

            if (targetItalic)
                newFont.AppendChild(new Italic());

            if (targetUnderline)
                newFont.AppendChild(new Underline());

            if (!string.IsNullOrEmpty(targetColor) && targetColor.StartsWith("#"))
            {
                // Strip "#" prefix — Rgb value should be RRGGBB or AARRGGBB
                string rgb = targetColor[1..]; // Remove '#'
                newFont.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = new HexBinaryValue(rgb) });
            }

            fonts.AppendChild(newFont);
            fonts.Count = (uint)fonts.Count();

            return (uint)(fonts.Count() - 1);
        }

        /// <summary>
        /// Find an existing Fill matching the requested background, or create a new one.
        /// </summary>
        private static uint FindOrCreateFill(Fills fills, WbDef.FillDefinition? fillDef)
        {
            if (fillDef == null || string.IsNullOrEmpty(fillDef.ColorArgb) || fillDef.PatternType == "none")
                return 0; // Return index 0 (none/no-fill)

            string targetColor = fillDef.ColorArgb;
            string targetPattern = fillDef.PatternType;

            // Start searching from index 2 (skip "none" at 0 and "gray125" at 1)
            uint startIdx = Math.Min(2, (uint)fills.Count());
            for (uint i = startIdx; i < fills.Count(); i++)
            {
                var f = fills.ElementAt((int)i) as Fill;
                if (f == null) continue;

                var pf = f.PatternFill;
                if (pf?.PatternType?.Value == null) continue;

                // Check pattern type
                string patternStr = pf.PatternType.Value.ToString();
                if (!patternStr.Equals(targetPattern, StringComparison.OrdinalIgnoreCase))
                    continue;

                // Check foreground color
                string fColor = "";
                if (pf.ForegroundColor?.Rgb?.Value != null)
                    fColor = "#" + pf.ForegroundColor.Rgb.Value;

                if (fColor == targetColor)
                    return i;
            }

            // Create a new Fill
            var newFill = new Fill
            {
                PatternFill = new PatternFill
                {
                    PatternType = targetPattern.ToLowerInvariant() switch
                    {
                        "solid" => PatternValues.Solid,
                        "gray125" => PatternValues.Gray125,
                        "gray0625" => PatternValues.Gray0625,
                        _ => PatternValues.Solid
                    },
                }
            };

            if (!string.IsNullOrEmpty(targetColor) && targetColor.StartsWith("#"))
            {
                string rgb = targetColor[1..]; // Remove '#'
                newFill.PatternFill.ForegroundColor = new ForegroundColor { Rgb = new HexBinaryValue(rgb) };
            }

            fills.AppendChild(newFill);
            fills.Count = (uint)fills.Count();

            return (uint)(fills.Count() - 1);
        }

        /// <summary>
        /// Parse a horizontal alignment string from the frontend into OpenXML enum.
        /// </summary>
        private static HorizontalAlignmentValues? ParseHorizontalAlignment(string align)
        {
            return align.ToLowerInvariant() switch
            {
                "left" => HorizontalAlignmentValues.Left,
                "center" => HorizontalAlignmentValues.Center,
                "right" => HorizontalAlignmentValues.Right,
                "general" => HorizontalAlignmentValues.General,
                "fill" => HorizontalAlignmentValues.Fill,
                "justify" => HorizontalAlignmentValues.Justify,
                "distributed" => HorizontalAlignmentValues.Distributed,
                _ => null
            };
        }

        /// <summary>
        /// Parse a vertical alignment string from the frontend into OpenXML enum.
        /// </summary>
        private static VerticalAlignmentValues? ParseVerticalAlignment(string align)
        {
            return align.ToLowerInvariant() switch
            {
                "top" => VerticalAlignmentValues.Top,
                "center" => VerticalAlignmentValues.Center,
                "bottom" => VerticalAlignmentValues.Bottom,
                "justify" => VerticalAlignmentValues.Justify,
                "distributed" => VerticalAlignmentValues.Distributed,
                _ => null
            };
        }
    }
}
