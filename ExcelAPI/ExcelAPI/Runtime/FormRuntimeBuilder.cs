using ExcelAPI.Rendering;
using WbDef = ExcelAPI.Models.WorkbookDefinition;

namespace ExcelAPI.Runtime
{
    /// <summary>
    /// Orchestrates the construction of a RuntimeForm from a WorkbookDefinition.
    ///
    /// Phase 4.2: Build(RenderWorkbook) path retained as [Obsolete] for LegacyRuntimeController
    ///            debug endpoint. BuildFromDefinitionDirect is the canonical execution path.
    ///
    /// Workflow:
    ///   1. Accept a WorkbookDefinition
    ///   2. For each sheet, map WbDef fields directly to RuntimeFields
    ///   3. Build RuntimeForm with all metadata
    ///   4. Serialize to JSON for the frontend API
    /// </summary>
    public class FormRuntimeBuilder
    {
        private readonly GeometryBuilder _geometry;
        private readonly CoordinateEngine _coords;
        private readonly FieldDetector _fieldDetector;

        public FormRuntimeBuilder()
        {
            // Default constructor for canonical BuildFromDefinitionDirect path.
            // No GeometryBuilder/CoordinateEngine/FieldDetector required.
            _geometry = null!;
            _coords = null!;
            _fieldDetector = null!;
        }

        /// <summary>
        /// Full constructor — used by DI and LegacyRuntimeController.
        /// </summary>
        public FormRuntimeBuilder(
            GeometryBuilder geometry,
            CoordinateEngine coords,
            FieldDetector fieldDetector)
        {
            _geometry = geometry;
            _coords = coords;
            _fieldDetector = fieldDetector;
        }

        /// <summary>
        /// Build a RuntimeForm from a RenderWorkbook (legacy OpenXml path).
        /// Retained for LegacyRuntimeController debug endpoint only.
        /// </summary>        [Obsolete("Use BuildFromDefinitionDirect with a WorkbookDefinition instead.")]
        public RuntimeForm Build(
            RenderWorkbook workbook,
            int dpi = 300,
            double originXPt = 0,
            double originYPt = 0)
        {
            return BuildInternal(workbook, null, dpi, originXPt, originYPt);
        }

        private RuntimeForm BuildInternal(
            RenderWorkbook workbook,
            WbDef.WorkbookDefinition? wbDef,
            int dpi = 300,
            double originXPt = 0,
            double originYPt = 0)
        {
            foreach (var sheet in workbook.Sheets)
            {
                if (sheet.TotalWidthPt <= 0)
                    _geometry.ComputeGeometry(sheet);
            }

            string workbookName = wbDef?.Info?.Title
                ?? Path.GetFileNameWithoutExtension(workbook.FilePath);
            string title = wbDef?.Info?.Title
                ?? Path.GetFileNameWithoutExtension(workbook.FilePath);

            var runtimeForm = new RuntimeForm
            {
                WorkbookName = workbookName,
                Title = title,
                Dpi = dpi,
                Scale = 1.0,
                Version = "1.0"
            };

            foreach (var sheet in workbook.Sheets)
            {
                var runtimeSheet = new RuntimeSheet
                {
                    Name = sheet.Name,
                    Index = sheet.Index,
                    PageWidthPx = (int)Math.Round(WbDef.Rectangle.PtToPx(sheet.TotalWidthPt, dpi)),
                    PageHeightPx = (int)Math.Round(WbDef.Rectangle.PtToPx(sheet.TotalHeightPt, dpi))
                };

                runtimeSheet.Fields = _fieldDetector.DetectFields(
                    sheet, originXPt, originYPt, dpi, sheet.Index + 1);

                if (sheet.DrawingObjects != null)
                {
                    foreach (var drawObj in sheet.DrawingObjects)
                    {
                        if (drawObj.IsImage && drawObj.ImageData != null)
                        {
                            runtimeSheet.Images.Add(new RuntimeImage
                            {
                                Name = drawObj.ImageData.FileName,
                                LeftPx = CoordinateEngine.PtToPx(
                                    DrawingParser.EmuToPt(drawObj.LeftEmu)),
                                TopPx = CoordinateEngine.PtToPx(
                                    DrawingParser.EmuToPt(drawObj.TopEmu)),
                                WidthPx = CoordinateEngine.PtToPx(
                                    DrawingParser.EmuToPt(drawObj.RightEmu - drawObj.LeftEmu)),
                                HeightPx = CoordinateEngine.PtToPx(
                                    DrawingParser.EmuToPt(drawObj.BottomEmu - drawObj.TopEmu)),
                                ContentType = drawObj.ImageData.ContentType
                            });
                        }
                    }
                }

                if (runtimeForm.Sheets.Count == 0)
                {
                    runtimeForm.PageWidth = runtimeSheet.PageWidthPx;
                    runtimeForm.PageHeight = runtimeSheet.PageHeightPx;
                }

                runtimeForm.Sheets.Add(runtimeSheet);
            }

            return runtimeForm;
        }

        /// <summary>
        /// Build a RuntimeForm directly from a WorkbookDefinition.
        /// This is the only remaining execution path:
        ///
        ///     WbDef → RuntimeForm (no adapter, no geometry/FieldDetector overhead)
        ///
        /// Sheet layout comes from WbDef.PrintLayout.
        /// Fields come from WbDef.Sheets[].Fields directly.
        /// </summary>
        public RuntimeForm BuildFromDefinitionDirect(
            WbDef.WorkbookDefinition wbDef,
            int dpi = 300)
        {
            var runtimeForm = new RuntimeForm
            {
                WorkbookName = wbDef.Info?.Title ?? "Untitled",
                Title = wbDef.Info?.Title ?? "Untitled",
                Dpi = dpi,
                Scale = 1.0,
                Version = "1.0"
            };

            foreach (var wbSheet in wbDef.Sheets)
            {
                var layout = wbSheet.PrintLayout;
                double pageWidthPt = layout.PageWidthPt;
                double pageHeightPt = layout.PageHeightPt;
                double pageWidthPx = WbDef.Rectangle.PtToPx(pageWidthPt, dpi);
                double pageHeightPx = WbDef.Rectangle.PtToPx(pageHeightPt, dpi);

                var rtSheet = new RuntimeSheet
                {
                    Name = wbSheet.Name,
                    Index = wbSheet.Index,
                    PageWidthPx = (int)Math.Round(pageWidthPx),
                    PageHeightPx = (int)Math.Round(pageHeightPx)
                };

                int tabIndex = 0;
                foreach (var field in wbSheet.Fields)
                {
                    var pxBounds = field.BoundsPt.ToPixels(dpi);

                    rtSheet.Fields.Add(new RuntimeField
                    {
                        Id = field.Id,
                        Name = field.Name,
                        CellReference = field.Cell.Address,
                        Row = field.Cell.RowIndex,
                        Column = field.Cell.ColumnIndex,
                        LeftPx = Math.Round(pxBounds.Left, 1),
                        TopPx = Math.Round(pxBounds.Top, 1),
                        WidthPx = Math.Round(pxBounds.Width, 1),
                        HeightPx = Math.Round(pxBounds.Height, 1),
                        LeftRatio = pageWidthPx > 0 ? Math.Round(pxBounds.Left / pageWidthPx, 7) : 0,
                        TopRatio = pageHeightPx > 0 ? Math.Round(pxBounds.Top / pageHeightPx, 7) : 0,
                        WidthRatio = pageWidthPx > 0 ? Math.Round(pxBounds.Width / pageWidthPx, 7) : 0,
                        HeightRatio = pageHeightPx > 0 ? Math.Round(pxBounds.Height / pageHeightPx, 7) : 0,
                        DataType = field.Type.ToString().ToLowerInvariant(),
                        ReadOnly = field.Locked,
                        Required = field.Required,
                        Alignment = field.Style?.Alignment?.Horizontal,
                        Font = field.Style?.Font?.Name,
                        FontSize = field.Style?.Font?.SizePt ?? 11,
                        Bold = field.Style?.Font?.Bold ?? false,
                        FontColor = field.Style?.Font?.ColorArgb,
                        BackgroundColor = field.Style?.Fill?.ColorArgb,
                        Border = field.Style?.Border?.HasBorder == true ? "thin" : null,
                        Placeholder = field.Placeholder,
                        DefaultValue = field.DefaultValue,
                        MaxLength = field.MaxLength,
                        TabIndex = tabIndex,
                        IsMerged = field.MergeInfo?.IsMerged ?? false,
                        MergeRange = field.MergeInfo?.MergeAddress
                    });
                    tabIndex++;
                }

                if (runtimeForm.Sheets.Count == 0)
                {
                    runtimeForm.PageWidth = (int)Math.Round(pageWidthPx);
                    runtimeForm.PageHeight = (int)Math.Round(pageHeightPx);
                }

                runtimeForm.Sheets.Add(rtSheet);
            }

            return runtimeForm;
        }
    }
}
