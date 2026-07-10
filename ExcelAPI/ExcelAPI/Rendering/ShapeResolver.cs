using D = DocumentFormat.OpenXml.Drawing;
using DSS = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Parses DrawingML shape elements (from drawing.xml) into RenderShape models.
    ///
    /// Supports:
    ///   - Rectangles (rect)
    ///   - Rounded rectangles (roundRect)
    ///   - Ellipses (ellipse)
    ///   - Lines (line)
    ///   - Arrows (various arrow presets)
    ///   - Text boxes (textBox)
    ///   - Polygons (various presets)
    ///
    /// Consumes DrawingParser.DrawingObject (which holds DSS types = xdr namespace)
    /// and produces RenderShape.
    /// </summary>
    public class ShapeResolver
    {
        /// <summary>
        /// Resolve a drawing object into a RenderShape ready for rendering.
        /// Returns null if the object is not a shape (e.g., it's an image).
        /// </summary>
        public RenderShape? Resolve(DrawingParser.DrawingObject drawObj)
        {
            if (!drawObj.IsShape) return null;

            var shape = new RenderShape
            {
                LeftPt = DrawingParser.EmuToPt(drawObj.LeftEmu),
                TopPt = DrawingParser.EmuToPt(drawObj.TopEmu),
                WidthPt = DrawingParser.EmuToPt(drawObj.RightEmu - drawObj.LeftEmu),
                HeightPt = DrawingParser.EmuToPt(drawObj.BottomEmu - drawObj.TopEmu)
            };

            shape.ShapeType = MapPresetGeometry(drawObj.PresetGeometry);

            // DSS.ShapeProperties IS the xdr:spPr element; its children use D namespace
            if (drawObj.ShapeProperties != null)
                ResolveShapeProperties(drawObj.ShapeProperties, shape);

            // DSS.TextBody IS the xdr:txBody element; its children use D namespace
            if (drawObj.TextBody != null)
                ResolveTextBody(drawObj.TextBody, shape);

            if (drawObj.NonVisualShapeProps != null)
                ResolveNonVisualProperties(drawObj.NonVisualShapeProps, shape);

            return shape;
        }

        private static string MapPresetGeometry(string? presetName)
        {
            if (string.IsNullOrEmpty(presetName)) return "rectangle";

            return presetName.ToLowerInvariant() switch
            {
                "rect" => "rectangle",
                "roundrect" => "roundedRect",
                "ellipse" => "ellipse",
                "line" => "line",
                "straightconnector1" => "line",
                "bentconnector1" => "line",
                "bentconnector2" => "line",
                "bentconnector3" => "line",
                "bentconnector4" => "line",
                "curvedconnector1" => "line",
                "curvedconnector2" => "line",
                "curvedconnector3" => "line",
                "curvedconnector4" => "line",
                "rightarrow" => "arrow",
                "leftarrow" => "arrow",
                "uparrow" => "arrow",
                "downarrow" => "arrow",
                "notchedrightarrow" => "arrow",
                "pentagon" => "polygon",
                "hexagon" => "polygon",
                "octagon" => "polygon",
                "parallelogram" => "polygon",
                "trapezoid" => "polygon",
                "flowchartprocess" => "roundedRect",
                _ => "rectangle"
            };
        }

        private static void ResolveShapeProperties(DSS.ShapeProperties spPr, RenderShape shape)
        {
            // Solid Fill — D.SolidFill is a child of xdr:spPr
            var solidFill = spPr.GetFirstChild<D.SolidFill>();
            if (solidFill != null)
            {
                var color = ResolveColor(solidFill);
                if (color != null)
                    shape.FillColorArgb = color;
            }

            // Outline / Border
            var outline = spPr.GetFirstChild<D.Outline>();
            if (outline != null)
            {
                if (outline.Width?.Value > 0)
                    shape.BorderWidthPt = outline.Width.Value / 12700.0;

                var outlineSolidFill = outline.GetFirstChild<D.SolidFill>();
                if (outlineSolidFill != null)
                {
                    var color = ResolveColor(outlineSolidFill);
                    if (color != null)
                        shape.BorderColorArgb = color;
                }

                var prstDash = outline.GetFirstChild<D.PresetDash>();
                if (prstDash?.Val?.Value != null)
                    shape.BorderDashStyle = prstDash.Val.Value.ToString().ToLowerInvariant();
            }

            // Transform (rotation)
            var xfrm = spPr.GetFirstChild<D.Transform2D>();
            if (xfrm?.Rotation?.Value > 0)
                shape.RotationDegrees = xfrm.Rotation.Value / 60000.0;
        }

        private static void ResolveTextBody(DSS.TextBody textBody, RenderShape shape)
        {
            var paragraphs = textBody.GetFirstChild<D.Paragraph>();
            if (paragraphs != null)
            {
                var text = new System.Text.StringBuilder();
                // Iterate ALL paragraphs in the text body
                foreach (var para in textBody.Descendants<D.Paragraph>())
                {
                    foreach (var run in para.Descendants<D.Run>())
                    {
                        var runText = run.GetFirstChild<D.Text>();
                        if (runText?.Text != null)
                            text.Append(runText.Text);
                    }
                    text.Append('\n'); // Paragraph separator
                }
                shape.Text = text.ToString().TrimEnd('\n');
            }

            // Text body properties
            var bodyPr = textBody.GetFirstChild<D.BodyProperties>();
            if (bodyPr != null)
            {
                if (bodyPr.Anchor?.Value != null)
                    shape.VerticalAlignment = bodyPr.Anchor.Value.ToString().ToLowerInvariant();
                if (bodyPr.Wrap?.Value != null)
                    shape.WrapText = bodyPr.Wrap.Value == D.TextWrappingValues.Square;
            }

            // Paragraph properties (horizontal alignment) — use first paragraph
            var firstPara = textBody.GetFirstChild<D.Paragraph>();
            var pPr = firstPara?.GetFirstChild<D.ParagraphProperties>();
            if (pPr?.Alignment?.Value != null)
                shape.HorizontalAlignment = pPr.Alignment.Value.ToString().ToLowerInvariant();

            // Font from default end paragraph run properties
            var endParaRPr = pPr?.GetFirstChild<D.EndParagraphRunProperties>();
            if (endParaRPr != null)
                ResolveRunProperties(endParaRPr, shape);

            // Font from first run
            var firstRun = firstPara?.GetFirstChild<D.Run>();
            if (firstRun != null)
            {
                var rPr = firstRun.GetFirstChild<D.RunProperties>();
                if (rPr != null)
                    ResolveRunProperties(rPr, shape);
            }
        }

        private static void ResolveRunProperties(D.TextCharacterPropertiesType rPr, RenderShape shape)
        {
            // LatinFont is a child element, not a direct property
            var latinFont = rPr.GetFirstChild<D.LatinFont>();
            if (latinFont?.Typeface != null)
                shape.FontName = latinFont.Typeface;

            if (rPr.FontSize?.Value > 0)
                shape.FontSizePt = rPr.FontSize.Value / 100.0;
            if (rPr.Bold?.Value == true)
                shape.Bold = rPr.Bold.Value;
            if (rPr.Italic?.Value == true)
                shape.Italic = rPr.Italic.Value;

            var solidFill = rPr.GetFirstChild<D.SolidFill>();
            if (solidFill != null)
            {
                var color = ResolveColor(solidFill);
                if (color != null)
                    shape.FontColorArgb = color;
            }
        }

        private static void ResolveNonVisualProperties(DSS.NonVisualShapeProperties nvSpPr, RenderShape shape)
        {
            var cNvPr = nvSpPr.GetFirstChild<D.NonVisualDrawingProperties>();
            if (cNvPr?.Name?.Value != null)
                shape.Name = cNvPr.Name.Value;
        }

        /// <summary>
        /// Resolve a solid fill color to #AARRGGBB.
        /// Supports srgbClr and schemeClr.
        /// </summary>
        private static string? ResolveColor(D.SolidFill solidFill)
        {
            var rgbColor = solidFill.GetFirstChild<D.RgbColorModelHex>();
            if (rgbColor?.Val?.Value != null)
                return "#FF" + rgbColor.Val.Value;

            var schemeColor = solidFill.GetFirstChild<D.SchemeColor>();
            if (schemeColor?.Val?.Value != null)
            {
                return schemeColor.Val.Value.ToString().ToLowerInvariant() switch
                {
                    "dk1" => "#FF000000",
                    "dk2" => "#FF44546A",
                    "lt1" => "#FFFFFFFF",
                    "lt2" => "#FFF2F2F2",
                    "accent1" => "#FF4472C4",
                    "accent2" => "#FFED7D31",
                    "accent3" => "#FFA5A5A5",
                    "accent4" => "#FFFFC000",
                    "accent5" => "#FF5B9BD5",
                    "accent6" => "#FF70AD47",
                    _ => "#FF000000"
                };
            }

            return null;
        }
    }
}
