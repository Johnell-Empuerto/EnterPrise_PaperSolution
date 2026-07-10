using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Parses .xlsx files using DocumentFormat.OpenXml and populates the internal
    /// RenderWorkbook model. This is the single data source for the rendering engines.
    /// No COM interop. No legacy database dependency.
    /// </summary>
    public class OpenXmlParser
    {
        private readonly GeometryBuilder _geometry;
        private readonly StyleResolver _styleResolver;
        private readonly DrawingParser _drawingParser;
        private readonly ImageResolver _imageResolver;

        public OpenXmlParser()
        {
            _geometry = new GeometryBuilder();
            _styleResolver = new StyleResolver(new ColorResolver(new ThemeResolver()), new StyleCache());
            _drawingParser = new DrawingParser();
            _imageResolver = new ImageResolver();
        }

        public OpenXmlParser(GeometryBuilder geometry)
        {
            _geometry = geometry;
            _styleResolver = new StyleResolver(new ColorResolver(new ThemeResolver()), new StyleCache());
            _drawingParser = new DrawingParser();
            _imageResolver = new ImageResolver();
        }

        public OpenXmlParser(GeometryBuilder geometry, StyleResolver styleResolver)
        {
            _geometry = geometry;
            _styleResolver = styleResolver;
            _drawingParser = new DrawingParser();
            _imageResolver = new ImageResolver();
        }

        public OpenXmlParser(GeometryBuilder geometry, StyleResolver styleResolver, DrawingParser drawingParser)
        {
            _geometry = geometry;
            _styleResolver = styleResolver;
            _drawingParser = drawingParser;
            _imageResolver = new ImageResolver();
        }

        public OpenXmlParser(GeometryBuilder geometry, StyleResolver styleResolver,
            DrawingParser drawingParser, ImageResolver imageResolver)
        {
            _geometry = geometry;
            _styleResolver = styleResolver;
            _drawingParser = drawingParser;
            _imageResolver = imageResolver;
        }

        /// <summary>
        /// Parse an .xlsx file into a RenderWorkbook.
        /// </summary>
        public RenderWorkbook Parse(string filePath)
        {
            var workbook = new RenderWorkbook { FilePath = filePath };

            using var doc = SpreadsheetDocument.Open(filePath, false);

            if (doc.WorkbookPart == null) return workbook;

            // Load shared strings for text cell values
            var sharedStrings = LoadSharedStrings(doc.WorkbookPart);

            // Parse styles (fonts, fills, borders, cellFormatXfs)
            var styleSheet = LoadStyleSheet(doc.WorkbookPart);

            // Determine default font and maxDigitWidth
            double maxDigitWidth = ComputeMaxDigitWidth(styleSheet);

            // Phase 11E: Load theme colors once per workbook
            _styleResolver.LoadTheme(doc.WorkbookPart);

            // Parse each worksheet, skipping hidden/veryHidden sheets
            int sheetIndex = 0;
            foreach (var sheetEntry in doc.WorkbookPart.Workbook.Descendants<Sheet>())
            {
                // Skip hidden sheets (e.g., legacy "_Fields" metadata sheet from VSTO Add-in)
                var sheetState = sheetEntry.State?.Value;
                if (sheetState == SheetStateValues.Hidden || sheetState == SheetStateValues.VeryHidden)
                    continue;

                var wsPart = doc.WorkbookPart.GetPartById(sheetEntry.Id!) as WorksheetPart;
                if (wsPart == null) continue;

                var renderSheet = ParseSheet(
                    wsPart, sheetEntry, sheetIndex, styleSheet, sharedStrings, maxDigitWidth);

                workbook.Sheets.Add(renderSheet);
                sheetIndex++;
            }

            return workbook;
        }

        private RenderSheet ParseSheet(
            WorksheetPart wsPart,
            Sheet sheetEntry,
            int sheetIndex,
            Stylesheet? styleSheet,
            List<string> sharedStrings,
            double maxDigitWidth)
        {
            var sheet = new RenderSheet
            {
                Name = sheetEntry.Name!,
                Index = sheetIndex,
                MaxDigitWidth = maxDigitWidth,
                DefaultColumnWidth = 8.43,
                DefaultRowHeight = 15.0
            };

            var ws = wsPart.Worksheet;

            // Sheet dimension string
            sheet.SheetDimension = ws.SheetDimension?.Reference?.Value;

            // Sheet format properties (default column width, row height)
            var sheetProps = ws.SheetProperties;
            if (sheetProps != null)
            {
                var sfp = sheetProps.GetFirstChild<SheetFormatProperties>();
                if (sfp != null)
                {
                    if (sfp.DefaultColumnWidth?.Value > 0)
                        sheet.DefaultColumnWidth = (double)sfp.DefaultColumnWidth.Value;
                    if (sfp.DefaultRowHeight?.Value > 0)
                        sheet.DefaultRowHeight = (double)sfp.DefaultRowHeight.Value;
                }
            }

            // Phase 11E: Pre-resolve all styles via StyleResolver
            if (styleSheet != null)
            {
                // Read actual default font from stylesheet (not hardcoded)
                string defaultFontName = "Calibri";
                double defaultFontSize = 11;
                if (styleSheet.Fonts?.Count > 0)
                {
                    var defaultFont = styleSheet.Fonts.ChildElements[0] as Font;
                    if (defaultFont != null)
                    {
                        defaultFontName = defaultFont.FontName?.Val?.Value ?? defaultFontName;
                        defaultFontSize = (double)(defaultFont.FontSize?.Val?.Value ?? defaultFontSize);
                    }
                }
                _styleResolver.PreResolveAll(styleSheet, defaultFontName, defaultFontSize);
            }

            // Parse columns
            ParseColumns(ws, sheet);

            // Parse rows and cells
            ParseRowsAndCells(ws, sheet, styleSheet, sharedStrings);

            // Parse merges
            ParseMerges(ws, sheet);

            // Parse drawing objects (images and shapes)
            var drawObjects = _drawingParser.ParseDrawings(wsPart);

            // Resolve image data for each image drawing object
            foreach (var drawObj in drawObjects)
            {
                if (drawObj.IsImage && !string.IsNullOrEmpty(drawObj.ImageRelId))
                {
                    drawObj.ImageData = _imageResolver.Resolve(wsPart, drawObj.ImageRelId);
                }
            }

            sheet.DrawingObjects = drawObjects;

            return sheet;
        }

        private void ParseColumns(Worksheet ws, RenderSheet sheet)
        {
            var cols = ws.Descendants<Column>().ToList();
            foreach (var col in cols)
            {
                uint min = col.Min?.Value ?? 1;
                uint max = col.Max?.Value ?? min;

                double charWidth = col.Width?.Value ?? sheet.DefaultColumnWidth;
                double pointWidth = CharWidthToPoints(charWidth, sheet.MaxDigitWidth);

                // Map outline level across the column range
                uint outlineLevel = col.OutlineLevel?.Value ?? 0;
                for (uint i = min; i <= max; i++)
                {
                    sheet.ColumnOutlineLevels[i] = outlineLevel;
                }

                sheet.Columns.Add(new RenderColumn
                {
                    Min = min,
                    Max = max,
                    Width = charWidth,
                    PointWidth = pointWidth,
                    Hidden = col.Hidden?.Value ?? false,
                    CustomWidth = col.CustomWidth?.Value ?? false,
                    BestFit = col.BestFit?.Value ?? false,
                    OutlineLevel = outlineLevel
                });
            }
        }

        private void ParseRowsAndCells(Worksheet ws, RenderSheet sheet,
            Stylesheet? styleSheet, List<string> sharedStrings)
        {
            // Default font info
            string? defaultFontName = "Calibri";
            double defaultFontSize = 11;
            if (styleSheet?.Fonts?.Count > 0)
            {
                var defaultFont = styleSheet.Fonts.ChildElements[0] as Font;
                if (defaultFont != null)
                {
                    defaultFontName = defaultFont.FontName?.Val?.Value ?? "Calibri";
                    defaultFontSize = (double)(defaultFont.FontSize?.Val?.Value ?? 11);
                }
            }

            var rows = ws.Descendants<Row>().ToList();
            foreach (var row in rows)
            {
                uint rowIndex = row.RowIndex?.Value ?? 0;
                bool hidden = row.Hidden?.Value ?? false;
                uint outlineLevel = row.OutlineLevel?.Value ?? 0;

                sheet.RowOutlineLevels[rowIndex] = outlineLevel;

                sheet.Rows.Add(new RenderRow
                {
                    RowIndex = rowIndex,
                    Height = row.Height?.Value,
                    Hidden = hidden,
                    CustomHeight = row.CustomHeight?.Value ?? false,
                    OutlineLevel = outlineLevel
                });

                // Parse cells within this row
                foreach (var cell in row.Descendants<Cell>())
                {
                    var rc = ParseCell(cell, styleSheet, sharedStrings,
                        defaultFontName, defaultFontSize);
                    if (rc != null)
                    {
                        rc.RowIndex = rowIndex;
                        sheet.Cells.Add(rc);
                    }
                }
            }
        }

        private RenderCell? ParseCell(Cell cell, Stylesheet? styleSheet,
            List<string> sharedStrings, string defaultFontName, double defaultFontSize)
        {
            var rc = new RenderCell
            {
                Reference = cell.CellReference?.Value,
                StyleIndex = cell.StyleIndex?.Value,
                FontName = defaultFontName,
                FontSize = defaultFontSize
            };

            // Parse cell value
            if (cell.DataType?.Value == CellValues.SharedString && cell.CellValue?.Text != null)
            {
                int ssIndex = int.Parse(cell.CellValue.Text);
                rc.SharedString = ssIndex >= 0 && ssIndex < sharedStrings.Count
                    ? sharedStrings[ssIndex]
                    : cell.CellValue.Text;
                rc.Value = rc.SharedString;
                rc.DataType = "sharedString";
            }
            else if (cell.CellValue?.Text != null)
            {
                rc.Value = cell.CellValue.Text;
                rc.DataType = (cell.DataType?.Value ?? CellValues.Number).ToString();
            }

            // Parse column index from reference
            var refStr = cell.CellReference?.Value ?? "";
            uint colIndex = 0;
            for (int i = 0; i < refStr.Length && char.IsLetter(refStr[i]); i++)
            {
                colIndex = colIndex * 26 + (uint)(refStr[i] - 'A' + 1);
            }
            rc.ColumnIndex = colIndex;

            // Apply style via StyleResolver (Phase 11E)
            if (styleSheet != null && rc.StyleIndex.HasValue)
            {
                var resolved = _styleResolver.Resolve(
                    styleSheet, rc.StyleIndex.Value,
                    defaultFontName, defaultFontSize);
                ApplyResolvedStyle(rc, resolved);
            }

            return rc;
        }

        /// <summary>
        /// Copies properties from ResolvedCellStyle (produced by StyleResolver) to RenderCell.
        /// This is the replacement for the old ApplyCellStyle method.
        /// Render layers continue reading individual properties for backward compatibility.
        /// </summary>
        private static void ApplyResolvedStyle(RenderCell rc, ResolvedCellStyle style)
        {
            rc.ResolvedStyle = style;

            rc.FontName = style.FontName;
            rc.FontSize = style.FontSize;
            rc.Bold = style.Bold;
            rc.Italic = style.Italic;
            rc.Underline = style.Underline;
            rc.Strikeout = style.Strikeout;
            rc.FontColorArgb = style.FontColorArgb;

            rc.FillColorArgb = style.FillColorArgb;
            rc.PatternType = style.PatternType;

            if (style.Border != null)
            {
                rc.Border = new RenderBorder
                {
                    Left = MapBorderItem(style.Border.Left),
                    Right = MapBorderItem(style.Border.Right),
                    Top = MapBorderItem(style.Border.Top),
                    Bottom = MapBorderItem(style.Border.Bottom)
                };
            }

            rc.HorizontalAlignment = style.HorizontalAlignment;
            rc.VerticalAlignment = style.VerticalAlignment;
            rc.WrapText = style.WrapText;
            rc.Indent = style.Indent;
            rc.TextRotation = style.TextRotation;
        }

        /// <summary>Map ResolvedBorderItem to RenderBorderItem for backward compat.</summary>
        private static RenderBorderItem? MapBorderItem(ResolvedBorderItem? item)
        {
            if (item == null) return null;
            return new RenderBorderItem
            {
                Style = item.Style,
                ColorArgb = item.ColorArgb
            };
        }

        private void ParseMerges(Worksheet ws, RenderSheet sheet)
        {
            var mergeCells = ws.Descendants<MergeCell>().ToList();

            // Build a merge lookup from cells
            foreach (var mc in mergeCells)
            {
                var range = mc.Reference?.Value ?? "";
                if (string.IsNullOrEmpty(range)) continue;

                var parts = range.Split(':');
                if (parts.Length != 2) continue;

                // Parse first and last cell references
                var (firstCol, firstRow) = ParseCellRef(parts[0]);
                var (lastCol, lastRow) = ParseCellRef(parts[1]);

                if (firstCol == 0) continue;

                var merge = new RenderMerge
                {
                    Reference = range,
                    FirstCol = firstCol,
                    FirstRow = firstRow,
                    LastCol = lastCol,
                    LastRow = lastRow
                };

                sheet.Merges.Add(merge);
            }

            // Mark cells that are part of merges
            for (int mi = 0; mi < sheet.Merges.Count; mi++)
            {
                var m = sheet.Merges[mi];
                foreach (var cell in sheet.Cells)
                {
                    if (cell.ColumnIndex >= m.FirstCol && cell.ColumnIndex <= m.LastCol
                        && cell.RowIndex >= m.FirstRow && cell.RowIndex <= m.LastRow)
                    {
                        cell.MergeIndex = mi;
                    }
                }
            }
        }

        // ── Helpers ─────────────────────────────────────────────────────

        private List<string> LoadSharedStrings(WorkbookPart wbPart)
        {
            var result = new List<string>();
            var ssp = wbPart.SharedStringTablePart;
            if (ssp == null) return result;

            foreach (var si in ssp.SharedStringTable.Descendants<SharedStringItem>())
            {
                result.Add(si.Text?.Text ?? si.InnerText);
            }
            return result;
        }

        private Stylesheet? LoadStyleSheet(WorkbookPart wbPart)
        {
            var sp = wbPart.WorkbookStylesPart;
            return sp?.Stylesheet;
        }

        private double ComputeMaxDigitWidth(Stylesheet? styleSheet)
        {
            // Read default font from styles.xml and compute maxDigitWidth
            if (styleSheet?.Fonts?.Count > 0)
            {
                var defaultFont = styleSheet.Fonts.ChildElements[0] as Font;
                if (defaultFont?.FontSize?.Val != null)
                {
                    double fontSize = (double)defaultFont.FontSize.Val;

                    // Approximate maxDigitWidth based on font size
                    // Calibri 11pt → 7.33, Arial 10pt → 7.0, etc.
                    // This can be refined per font family
                    if (fontSize > 0)
                    {
                        return 7.33 * (fontSize / 11.0);
                    }
                }
            }
            return 7.33;
        }

        private double CharWidthToPoints(double charWidth, double maxDigitWidth)
        {
            // Delegated to shared GeometryBuilder
            return GeometryBuilder.CharWidthToPoints(charWidth, maxDigitWidth);
        }

        private (uint col, uint row) ParseCellRef(string reference)
        {
            uint col = 0;
            int i = 0;
            for (; i < reference.Length && char.IsLetter(reference[i]); i++)
            {
                col = col * 26 + (uint)(char.ToUpper(reference[i]) - 'A' + 1);
            }
            uint row = i < reference.Length ? uint.Parse(reference[i..]) : 0;
            return (col, row);
        }

        /// <summary>
        /// Save the parsed workbooks to a JSON file for debugging.
        /// </summary>
        public string? SaveAsJson(RenderWorkbook workbook, string outputDir)
        {
            try
            {
                var json = System.Text.Json.JsonSerializer.Serialize(workbook,
                    new System.Text.Json.JsonSerializerOptions
                    {
                        WriteIndented = true,
                        DefaultIgnoreCondition =
                            System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
                    });
                string path = Path.Combine(outputDir, $"renderModel_{Guid.NewGuid():N}.json");
                System.IO.File.WriteAllText(path, json);
                return path;
            }
            catch { return null; }
        }
    }
}
