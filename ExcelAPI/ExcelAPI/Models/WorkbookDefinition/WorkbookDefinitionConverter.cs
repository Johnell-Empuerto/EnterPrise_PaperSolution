// ────────────────────────────────────────────────────────────────────────────
// WorkbookDefinitionConverter — Adapter from existing models to canonical form
//
// This adapter converts the existing ExcelAPI.Models types (FormDefinition,
// CaptureResult, SheetDefinition, etc.) into the canonical WorkbookDefinition.
//
// It is a pure function that does NOT modify existing business logic.
// It provides a migration path: existing code can produce a WorkbookDefinition
// without being rewritten.
//
// Usage in future phases:
//   ExcelCaptureService → WorkbookDefinition → CaptureResult (embedded)
//   WorkbookDefinition → FormRuntimeBuilder → RuntimeForm
//   WorkbookDefinition → Renderer (replacing scattered DTOs)
//
// Ownership: Application
// ────────────────────────────────────────────────────────────────────────────

namespace ExcelAPI.Models.WorkbookDefinition
{
    /// <summary>
    /// Converts existing ExcelAPI.Models types to the canonical WorkbookDefinition.
    /// Each method is an independent conversion that can be used incrementally.
    /// </summary>
    public static class WorkbookDefinitionConverter
    {
        // ── From CaptureResult ─────────────────────────────────────────

        /// <summary>
        /// Convert a CaptureResult (from COM capture) into a basic WorkbookDefinition.
        /// The CaptureResult's Fields become FieldDefinitions on the first sheet.
        /// </summary>
        public static WorkbookDefinition FromCaptureResult(
            CaptureResult capture,
            string fileName,
            PrintLayout? printLayout = null)
        {
            var wb = new WorkbookDefinition
            {
                SourceFileName = fileName,
                CapturedAt = DateTime.UtcNow,
                Info = new WorkbookInfo
                {
                    Title = Path.GetFileNameWithoutExtension(fileName),
                    Created = DateTime.UtcNow,
                    Modified = DateTime.UtcNow,
                    Version = "1.0"
                }
            };

            var sheet = new SheetDefinition
            {
                Id = "sheet_0",
                Name = "Sheet1",
                Index = 0,
                PrintLayout = printLayout ?? CreateDefaultPrintLayout(capture)
            };

            // Convert fields
            foreach (var f in capture.Fields)
            {
                var field = new FieldDefinition
                {
                    Id = f.Id,
                    Name = f.Name,
                    Cell = new CellReference(f.Cell, 0, 0), // row/col from address
                    BoundsPt = new Rectangle(f.Left / 300.0 * 72.0, f.Top / 300.0 * 72.0,
                                              f.Width / 300.0 * 72.0, f.Height / 300.0 * 72.0),
                    Type = ParseFieldType(f.Type),
                    MergeInfo = new MergedFieldInfo
                    {
                        IsMerged = f.IsMerged,
                        MergeAddress = f.MergeAddress
                    },
                    Metadata = new Dictionary<string, string>
                    {
                        ["comment"] = f.Comment,
                        ["excelLeft"] = f.ExcelLeft.ToString("F2"),
                        ["excelTop"] = f.ExcelTop.ToString("F2"),
                        ["printAreaLeft"] = f.PrintAreaLeft.ToString("F2"),
                        ["printAreaTop"] = f.PrintAreaTop.ToString("F2"),
                        ["excelWidthPt"] = f.ExcelWidthPt.ToString("F2"),
                        ["excelHeightPt"] = f.ExcelHeightPt.ToString("F2")
                    }
                };
                sheet.Fields.Add(field);
            }

            wb.Sheets.Add(sheet);
            return wb;
        }

        // ── From FormDefinition ────────────────────────────────────────

        /// <summary>
        /// Convert a FormDefinition (from Designer generation or workbook upload)
        /// into a complete WorkbookDefinition with all sheet, field, and image data.
        /// </summary>
        public static WorkbookDefinition FromFormDefinition(FormDefinition form)
        {
            var wb = new WorkbookDefinition
            {
                SourceFileName = form.Workbook?.Title ?? "Untitled",
                CapturedAt = DateTime.UtcNow,
                Info = new WorkbookInfo
                {
                    Title = form.Workbook?.Title ?? "Untitled",
                    Author = form.Workbook?.Author ?? string.Empty,
                    Description = form.Workbook?.Description ?? string.Empty,
                    Version = form.Workbook?.Version ?? "1.0",
                    Created = TryParseDateTime(form.Workbook?.Created),
                    Modified = TryParseDateTime(form.Workbook?.Modified)
                }
            };

            // Convert sheets
            foreach (var s in form.Sheets)
            {
                var sheet = new SheetDefinition
                {
                    Id = s.Id,
                    Name = s.Name,
                    Index = s.Index,
                    PrintLayout = ConvertPageSettings(s.PageSettings, s.PrintArea),
                    FreezePane = s.FreezePane,
                    DefaultColumnWidthChars = 8.43,
                    DefaultRowHeightPt = 15.0
                };

                // Convert merged cells
                foreach (var m in s.MergedCells)
                {
                    sheet.MergedRanges.Add(new MergedRangeDefinition
                    {
                        Address = m.Address,
                        BoundsPt = new Rectangle(m.LeftPt, m.TopPt, m.WidthPt, m.HeightPt)
                    });
                }

                // Convert row heights
                foreach (var (rowIdx, height) in s.RowHeights)
                {
                    sheet.Rows.Add(new RowDefinition
                    {
                        Index = (uint)rowIdx,
                        HeightPt = height,
                        CustomHeight = true
                    });
                }

                // Convert column widths
                foreach (var (colIdx, width) in s.ColumnWidths)
                {
                    sheet.Columns.Add(new ColumnDefinition
                    {
                        Index = (uint)colIdx,
                        WidthChars = width,
                        WidthPt = width * 7.0, // approximate conversion
                        CustomWidth = true
                    });
                }

                wb.Sheets.Add(sheet);
            }

            // Convert clusters (interactive fields)
            foreach (var c in form.Clusters)
            {
                var targetSheet = wb.Sheets.FirstOrDefault(s => s.Id == c.SheetId);
                if (targetSheet == null) continue;

                var field = new FieldDefinition
                {
                    Id = c.ClusterId,
                    Name = c.Name,
                    Cell = new CellReference(c.CellAddress, 0, 0),
                    BoundsPt = new Rectangle(c.LeftPt, c.TopPt, c.WidthPt, c.HeightPt),
                    Type = ParseFieldType(c.Type),
                    Required = false,
                    Locked = c.Readonly,
                    Visible = c.Visibility == "visible",
                    TabIndex = form.Clusters.IndexOf(c),
                    MergeInfo = ParseMergeMetadata(c.Metadata),
                    Metadata = new Dictionary<string, string>(c.Metadata ?? new Dictionary<string, string>())
                };

                // Copy input parameters
                foreach (var kvp in c.InputParameters)
                {
                    field.Metadata[$"param_{kvp.Key}"] = kvp.Value;
                }

                targetSheet.Fields.Add(field);
            }

            // Populate field styles from source sheet CellStyles
            foreach (var s in form.Sheets)
            {
                var targetSheet = wb.Sheets.FirstOrDefault(ws => ws.Id == s.Id);
                if (targetSheet == null) continue;

                foreach (var field in targetSheet.Fields)
                {
                    if (s.CellStyles.TryGetValue(field.Cell.Address, out var cellStyle))
                    {
                        field.Style = ConvertCellStyleInfo(cellStyle);
                    }
                }
            }

            // Convert images
            foreach (var img in form.Images)
            {
                var targetSheet = wb.Sheets.FirstOrDefault(s => s.Id == img.SheetId);
                if (targetSheet == null) continue;

                targetSheet.Images.Add(new ImageDefinition
                {
                    Id = img.Id,
                    Name = img.Name,
                    BoundsPt = new Rectangle(img.LeftPt, img.TopPt, img.WidthPt, img.HeightPt),
                    Data = img.Data,
                    ContentType = img.Format switch
                    {
                        "png" => "image/png",
                        "jpg" or "jpeg" => "image/jpeg",
                        "gif" => "image/gif",
                        _ => "image/png"
                    }
                });
            }

            return wb;
        }

        // ── Private Helpers ────────────────────────────────────────────

        private static PrintLayout CreateDefaultPrintLayout(CaptureResult capture)
        {
            var layout = new PrintLayout
            {
                PaperSize = new PaperSize { Name = "Letter", WidthPt = 612, HeightPt = 792 },
                Margins = new Margins(),
                Scaling = new ScalingDefinition()
            };

            if (capture.PageSetup != null)
            {
                layout.PaperSize = new PaperSize
                {
                    WidthPt = capture.PageSetup.PageWidthPt > 0
                        ? capture.PageSetup.PageWidthPt : 612,
                    HeightPt = capture.PageSetup.PageHeightPt > 0
                        ? capture.PageSetup.PageHeightPt : 792
                };
                layout.Margins.LeftPt = capture.PageSetup.LeftMargin > 0
                    ? capture.PageSetup.LeftMargin : 50.4;
                layout.Margins.TopPt = capture.PageSetup.TopMargin > 0
                    ? capture.PageSetup.TopMargin : 54.0;
                layout.Scaling.CenterHorizontally = capture.PageSetup.CenterHorizontally;
                layout.Scaling.CenterVertically = capture.PageSetup.CenterVertically;
                layout.Scaling.Zoom = capture.PageSetup.Zoom > 0
                    ? capture.PageSetup.Zoom : 100;
                layout.Scaling.FitToPagesWide = capture.PageSetup.FitToPagesWide;
                layout.Scaling.FitToPagesTall = capture.PageSetup.FitToPagesTall;
            }

            if (capture.Page != null)
            {
                double ptsToPx = Rectangle.PtToPx(1, 300);
                layout.PrintArea = new PrintAreaDefinition
                {
                    Address = string.Empty,
                    BoundsPt = new Rectangle(
                        0, 0,
                        capture.Page.Width / ptsToPx,
                        capture.Page.Height / ptsToPx)
                };
            }

            return layout;
        }

        private static PrintLayout ConvertPageSettings(PageSettings ps, PrintAreaInfo? printArea)
        {
            var layout = new PrintLayout
            {
                PaperSize = new PaperSize
                {
                    Name = ps.PaperSize,
                    WidthPt = ps.WidthPt,
                    HeightPt = ps.HeightPt
                },
                Orientation = string.Equals(ps.Orientation, "landscape", StringComparison.OrdinalIgnoreCase)
                    ? PageOrientation.Landscape
                    : PageOrientation.Portrait,
                Margins = new Margins
                {
                    LeftPt = ps.LeftMargin,
                    RightPt = ps.RightMargin,
                    TopPt = ps.TopMargin,
                    BottomPt = ps.BottomMargin
                },
                Scaling = new ScalingDefinition
                {
                    Zoom = ps.Zoom,
                    FitToPagesWide = ps.FitToPagesWide,
                    FitToPagesTall = ps.FitToPagesTall,
                    CenterHorizontally = ps.CenterHorizontally,
                    CenterVertically = ps.CenterVertically
                }
            };

            if (printArea != null && !string.IsNullOrEmpty(printArea.Address))
            {
                layout.PrintArea = new PrintAreaDefinition
                {
                    Address = printArea.Address,
                    BoundsPt = new Rectangle(
                        printArea.LeftPt, printArea.TopPt,
                        printArea.WidthPt, printArea.HeightPt)
                };
            }

            return layout;
        }

        private static FieldType ParseFieldType(string type)
        {
            return type.ToLowerInvariant() switch
            {
                "text" or "textbox" or "input" => FieldType.Text,
                "number" or "amount" or "price" => FieldType.Number,
                "date" => FieldType.Date,
                "checkbox" or "check box" or "tick" => FieldType.Checkbox,
                "signature" => FieldType.Signature,
                "dropdown" or "combo" or "list" => FieldType.Dropdown,
                "calculated" or "formula" or "auto" => FieldType.Calculated,
                _ => FieldType.Text
            };
        }

        private static MergedFieldInfo? ParseMergeMetadata(Dictionary<string, string>? metadata)
        {
            if (metadata == null) return null;

            if (metadata.TryGetValue("isMerged", out var isMerged)
                && bool.TryParse(isMerged, out var merged) && merged)
            {
                metadata.TryGetValue("mergeAddress", out var mergeAddress);
                return new MergedFieldInfo
                {
                    IsMerged = true,
                    MergeAddress = mergeAddress
                };
            }

            return null;
        }

        // ── Reverse: WorkbookDefinition → FormDefinition ─────────────

        /// <summary>
        /// Convert a canonical WorkbookDefinition back into a FormDefinition.
        /// This enables existing service contracts (FormSaveService, PublishEngine)
        /// to consume WorkbookDefinition without being rewritten.
        /// </summary>
        public static FormDefinition ToFormDefinition(WorkbookDefinition wbDef)
        {
            var form = new FormDefinition
            {
                Workbook = new WorkbookMetadata
                {
                    Title = wbDef.Info?.Title ?? "Untitled",
                    Author = wbDef.Info?.Author ?? string.Empty,
                    Created = (wbDef.Info?.Created ?? DateTime.UtcNow).ToString("o"),
                    Modified = (wbDef.Info?.Modified ?? DateTime.UtcNow).ToString("o"),
                    Version = wbDef.Info?.Version ?? "1.0",
                    Description = wbDef.Info?.Description ?? string.Empty
                },
                Metadata = new Dictionary<string, string>
                {
                    ["sourceFile"] = wbDef.SourceFileName,
                    ["capturedAt"] = wbDef.CapturedAt.ToString("o"),
                    ["schemaVersion"] = WorkbookDefinition.SchemaVersion
                }
            };

            foreach (var sheet in wbDef.Sheets)
            {
                // MUST use fully qualified OLD types here — the current namespace
                // has new SheetDefinition/ImageDefinition that conflict with
                // ExcelAPI.Models.SheetDefinition/ImageDefinition expected by FormDefinition.
                var sheetDef = new Models.SheetDefinition
                {
                    Id = sheet.Id,
                    Name = sheet.Name,
                    Index = sheet.Index,
                    PageSettings = ConvertPrintLayout(sheet.PrintLayout),
                    FreezePane = sheet.FreezePane,
                    RowHeights = sheet.Rows.ToDictionary(r => (int)r.Index, r => r.HeightPt ?? 15.0),
                    ColumnWidths = sheet.Columns.ToDictionary(c => (int)c.Index, c => c.WidthChars),
                    MergedCells = sheet.MergedRanges.Select(m => new Models.MergedCellInfo
                    {
                        Address = m.Address,
                        CellAddress = m.Address?.Split(':')?.FirstOrDefault() ?? "",
                        LeftPt = m.BoundsPt?.Left ?? 0,
                        TopPt = m.BoundsPt?.Top ?? 0,
                        WidthPt = m.BoundsPt?.Width ?? 0,
                        HeightPt = m.BoundsPt?.Height ?? 0
                    }).ToList(),
                    CellValues = new Dictionary<string, string>(),
                    CellStyles = new Dictionary<string, Models.CellStyleInfo>()
                };

                // Convert fields to clusters
                foreach (var field in sheet.Fields)
                {
                    form.Clusters.Add(new Models.ClusterDefinition
                    {
                        ClusterId = field.Id,
                        Name = field.Name,
                        Type = field.Type.ToString().ToLower(),
                        SheetId = sheet.Id,
                        CellAddress = field.Cell.Address,
                        LeftPt = field.BoundsPt.Left,
                        TopPt = field.BoundsPt.Top,
                        WidthPt = field.BoundsPt.Width,
                        HeightPt = field.BoundsPt.Height,
                        Visibility = field.Visible ? "visible" : "hidden",
                        Readonly = field.Locked,
                        Remarks = field.Placeholder ?? string.Empty,
                        Metadata = field.Metadata ?? new Dictionary<string, string>()
                    });
                }

                // Convert images
                foreach (var img in sheet.Images)
                {
                    form.Images.Add(new Models.ImageDefinition
                    {
                        Id = img.Id,
                        Name = img.Name,
                        SheetId = sheet.Id,
                        LeftPt = img.BoundsPt.Left,
                        TopPt = img.BoundsPt.Top,
                        WidthPt = img.BoundsPt.Width,
                        HeightPt = img.BoundsPt.Height,
                        Data = img.Data,
                        Format = img.ContentType switch
                        {
                            "image/png" => "png",
                            "image/jpeg" => "jpg",
                            "image/gif" => "gif",
                            _ => "png"
                        }
                    });
                }

                form.Sheets.Add(sheetDef);
            }

            return form;
        }

        /// <summary>
        /// Convert a canonical PrintLayout back to PageSettings + PrintAreaInfo.
        /// </summary>
        private static PageSettings ConvertPrintLayout(PrintLayout layout)
        {
            return new PageSettings
            {
                PaperSize = layout.PaperSize?.Name ?? "Letter",
                Orientation = layout.Orientation == PageOrientation.Landscape ? "landscape" : "portrait",
                WidthPt = layout.PageWidthPt,
                HeightPt = layout.PageHeightPt,
                LeftMargin = layout.Margins?.LeftPt ?? 50.4,
                RightMargin = layout.Margins?.RightPt ?? 50.4,
                TopMargin = layout.Margins?.TopPt ?? 54.0,
                BottomMargin = layout.Margins?.BottomPt ?? 54.0,
                CenterHorizontally = layout.Scaling?.CenterHorizontally ?? false,
                CenterVertically = layout.Scaling?.CenterVertically ?? false,
                Zoom = layout.Scaling?.Zoom ?? 100,
                FitToPagesWide = layout.Scaling?.FitToPagesWide ?? 0,
                FitToPagesTall = layout.Scaling?.FitToPagesTall ?? 0
            };
        }

        private static DateTime TryParseDateTime(string? value)
        {
            if (string.IsNullOrEmpty(value)) return DateTime.UtcNow;
            if (DateTime.TryParse(value, out var result)) return result;
            return DateTime.UtcNow;
        }

        /// <summary>
        /// Convert an existing CellStyleInfo (from ExcelAPI.Models) into the canonical CellStyle.
        /// </summary>
        private static CellStyle ConvertCellStyleInfo(CellStyleInfo oldStyle)
        {
            var style = new CellStyle
            {
                Font = new FontDefinition
                {
                    Name = oldStyle.FontName ?? "Calibri",
                    SizePt = oldStyle.FontSize ?? 11,
                    Bold = oldStyle.Bold ?? false,
                    Italic = oldStyle.Italic ?? false,
                    Underline = oldStyle.Underline ?? false,
                    ColorArgb = oldStyle.Color != null ? $"#{oldStyle.Color.TrimStart('#')}" : null
                },
                WrapText = oldStyle.WrapText ?? false
            };

            // Fill
            if (!string.IsNullOrEmpty(oldStyle.FillColor))
            {
                style.Fill = new FillDefinition
                {
                    PatternType = "solid",
                    ColorArgb = $"#{oldStyle.FillColor.TrimStart('#')}"
                };
            }

            // Alignment
            if (oldStyle.HorizontalAlignment != null)
            {
                style.Alignment.Horizontal = oldStyle.HorizontalAlignment.ToLower();
            }
            if (oldStyle.VerticalAlignment != null)
            {
                style.Alignment.Vertical = oldStyle.VerticalAlignment.ToLower();
            }

            // Borders
            var hasBorder = oldStyle.BorderTop != null || oldStyle.BorderBottom != null
                         || oldStyle.BorderLeft != null || oldStyle.BorderRight != null;
            if (hasBorder)
            {
                style.Border = new BorderDefinition();
                style.Border.Top = ParseBorderEdge(oldStyle.BorderTop);
                style.Border.Bottom = ParseBorderEdge(oldStyle.BorderBottom);
                style.Border.Left = ParseBorderEdge(oldStyle.BorderLeft);
                style.Border.Right = ParseBorderEdge(oldStyle.BorderRight);
            }

            return style;
        }

        /// <summary>
        /// Parse a CSS-style border string (e.g., "1px solid #000000") into a BorderEdge.
        /// </summary>
        private static BorderEdge? ParseBorderEdge(string? borderCss)
        {
            if (string.IsNullOrEmpty(borderCss)) return null;

            var parts = borderCss.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2) return null;

            // parts: [width] [style] [color?]
            string style = parts.Length >= 2 ? parts[1] : "solid";
            string? color = parts.Length >= 3 ? parts[2] : null;

            return new BorderEdge
            {
                Style = style == "solid" || style == "dashed" || style == "dotted" || style == "double"
                    ? style : "thin",
                ColorArgb = color
            };
        }
    }
}
