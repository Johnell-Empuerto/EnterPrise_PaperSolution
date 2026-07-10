using DocumentFormat.OpenXml.Packaging;
using DSS = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using D = DocumentFormat.OpenXml.Drawing;

namespace ExcelAPI.Rendering
{
    /// <summary>
    /// Parses worksheet drawing relationships (drawing.xml, drawing#.xml)
    /// and produces lists of anchored drawing objects.
    ///
    /// Supports:
    ///   - oneCellAnchor (position relative to a cell)
    ///   - twoCellAnchor (position from/to cells)
    ///   - absoluteAnchor (fixed page position)
    ///
    /// Produces intermediate DrawingObject records consumed by
    /// ImageResolver and ShapeResolver.
    /// </summary>
    public class DrawingParser
    {
        /// <summary>
        /// A raw drawing object extracted from the worksheet's drawing part.
        /// </summary>
        public class DrawingObject
        {
            public string AnchorType { get; set; } = "oneCell";

            /// <summary>Left position in EMU.</summary>
            public long LeftEmu { get; set; }
            /// <summary>Top position in EMU.</summary>
            public long TopEmu { get; set; }
            /// <summary>Width/right position in EMU.</summary>
            public long RightEmu { get; set; }
            /// <summary>Height/bottom position in EMU.</summary>
            public long BottomEmu { get; set; }

            /// <summary>Column offset for oneCellAnchor (in EMU).</summary>
            public long ColOffsetEmu { get; set; }
            /// <summary>Row offset for oneCellAnchor (in EMU).</summary>
            public long RowOffsetEmu { get; set; }

            /// <summary>Anchor column for oneCellAnchor.</summary>
            public uint AnchorCol { get; set; }
            /// <summary>Anchor row for oneCellAnchor.</summary>
            public uint AnchorRow { get; set; }
            /// <summary>End column for twoCellAnchor.</summary>
            public uint EndCol { get; set; }
            /// <summary>End row for twoCellAnchor.</summary>
            public uint EndRow { get; set; }

            /// <summary>Whether this object is an image (has blipFill).</summary>
            public bool IsImage { get; set; }
            /// <summary>Relationship ID for the image part (if IsImage).</summary>
            public string? ImageRelId { get; set; }
            /// <summary>Pre-resolved image data (set by OpenXmlParser).</summary>
            public RenderImage? ImageData { get; set; }

            /// <summary>Whether this object is a shape.</summary>
            public bool IsShape { get; set; }

            // These store the direct OpenXml elements for lazy resolution by ShapeResolver.
            // DSS.ShapeProperties is the xdr:spPr element that contains a-namespace children.
            public DSS.ShapeProperties? ShapeProperties { get; set; }
            public DSS.TextBody? TextBody { get; set; }
            public DSS.NonVisualShapeProperties? NonVisualShapeProps { get; set; }
            public string? PresetGeometry { get; set; }
        }

        public List<DrawingObject> ParseDrawings(WorksheetPart wsPart)
        {
            var results = new List<DrawingObject>();

            // WorksheetPart.DrawingsPart returns a single DrawingsPart (or null)
            var dp = wsPart.DrawingsPart;
            if (dp != null)
            {
                ParseDrawingPart(dp, results);
            }

            return results;
        }

        private void ParseDrawingPart(DrawingsPart dp, List<DrawingObject> results)
        {
            var wsDr = dp.WorksheetDrawing;
            if (wsDr == null) return;

            foreach (var anchor in wsDr.ChildElements)
            {
                var obj = ParseAnchor(anchor);
                if (obj != null) results.Add(obj);
            }
        }

        private DrawingObject? ParseAnchor(DocumentFormat.OpenXml.OpenXmlElement anchor)
        {
            var obj = new DrawingObject();

            switch (anchor.LocalName)
            {
                case "oneCellAnchor":
                    obj.AnchorType = "oneCell";
                    ParseOneCellAnchor(anchor, obj);
                    break;
                case "twoCellAnchor":
                    obj.AnchorType = "twoCell";
                    ParseTwoCellAnchor(anchor, obj);
                    break;
                case "absoluteAnchor":
                    obj.AnchorType = "absolute";
                    ParseAbsoluteAnchor(anchor, obj);
                    break;
                default:
                    return null;
            }

            var picture = anchor.GetFirstChild<DSS.Picture>();
            if (picture != null)
            {
                obj.IsImage = true;
                ParsePicture(picture, obj);
                return obj;
            }

            var shape = anchor.GetFirstChild<DSS.Shape>();
            if (shape != null)
            {
                obj.IsShape = true;
                ParseShape(shape, obj);
                return obj;
            }

            // Connector shapes (straight line connectors, arrows) are not supported
            // in this version of DocumentFormat.OpenXml (type not available in API).
            return null;
        }

        private void ParseOneCellAnchor(DocumentFormat.OpenXml.OpenXmlElement anchor, DrawingObject obj)
        {
            var from = anchor.GetFirstChild<DSS.FromMarker>();
            if (from != null)
            {
                TryParseCellMarker(from, out uint ac, out uint ar,
                    out long co, out long ro);
                obj.AnchorCol = ac; obj.AnchorRow = ar;
                obj.ColOffsetEmu = co; obj.RowOffsetEmu = ro;
            }

            var ext = anchor.GetFirstChild<DSS.Extent>();
            if (ext != null)
            {
                // Cx/Cy are Int64Value with .Value returning long?
                obj.RightEmu = ext.Cx?.Value ?? 0;
                obj.BottomEmu = ext.Cy?.Value ?? 0;
            }
        }

        private void ParseTwoCellAnchor(DocumentFormat.OpenXml.OpenXmlElement anchor, DrawingObject obj)
        {
            var from = anchor.GetFirstChild<DSS.FromMarker>();
            if (from != null)
            {
                TryParseCellMarker(from, out uint ac, out uint ar,
                    out long co, out long ro);
                obj.AnchorCol = ac; obj.AnchorRow = ar;
                obj.ColOffsetEmu = co; obj.RowOffsetEmu = ro;
            }

            var to = anchor.GetFirstChild<DSS.ToMarker>();
            if (to != null)
            {
                TryParseCellMarker(to, out uint ec, out uint er, out _, out _);
                obj.EndCol = ec; obj.EndRow = er;
            }
        }

        private void ParseAbsoluteAnchor(DocumentFormat.OpenXml.OpenXmlElement anchor, DrawingObject obj)
        {
            var pos = anchor.GetFirstChild<DSS.Position>();
            if (pos != null)
            {
                // X/Y are Int64Value with .Value returning long?
                obj.LeftEmu = pos.X?.Value ?? 0;
                obj.TopEmu = pos.Y?.Value ?? 0;
            }

            var ext = anchor.GetFirstChild<DSS.Extent>();
            if (ext != null)
            {
                obj.RightEmu = ext.Cx?.Value ?? 0;
                obj.BottomEmu = ext.Cy?.Value ?? 0;
            }
        }

        private void ParsePicture(DSS.Picture picture, DrawingObject obj)
        {
            var blipFill = picture.GetFirstChild<DSS.BlipFill>();
            if (blipFill != null)
            {
                var blip = blipFill.GetFirstChild<D.Blip>();
                if (blip != null)
                    obj.ImageRelId = blip.Embed?.Value;
            }

            // DSS.ShapeProperties IS the xdr:spPr element with D-namespace children
            obj.ShapeProperties = picture.GetFirstChild<DSS.ShapeProperties>();
            if (obj.ShapeProperties != null)
                ParsePresetGeometry(obj.ShapeProperties, obj);
        }

        private void ParseShape(DSS.Shape shape, DrawingObject obj)
        {
            obj.ShapeProperties = shape.GetFirstChild<DSS.ShapeProperties>();
            if (obj.ShapeProperties != null)
                ParsePresetGeometry(obj.ShapeProperties, obj);

            obj.NonVisualShapeProps = shape.GetFirstChild<DSS.NonVisualShapeProperties>();
            obj.TextBody = shape.GetFirstChild<DSS.TextBody>();
        }

        private void ParsePresetGeometry(DSS.ShapeProperties spPr, DrawingObject obj)
        {
            var prstGeom = spPr.GetFirstChild<D.PresetGeometry>();
            if (prstGeom?.Preset?.Value != null)
                obj.PresetGeometry = prstGeom.Preset.Value.ToString();
        }

        private static void TryParseCellMarker(DSS.FromMarker marker, out uint col, out uint row,
            out long colOffsetEmu, out long rowOffsetEmu)
        {
            col = 0; row = 0; colOffsetEmu = 0; rowOffsetEmu = 0;
            try
            {
                uint.TryParse(marker.ColumnId?.Text, out uint c);
                uint.TryParse(marker.RowId?.Text, out uint r);
                col = c + 1; row = r + 1;
                long.TryParse(marker.ColumnOffset?.Text, out long co);
                long.TryParse(marker.RowOffset?.Text, out long ro);
                colOffsetEmu = co; rowOffsetEmu = ro;
            }
            catch { }
        }

        private static void TryParseCellMarker(DSS.ToMarker marker, out uint col, out uint row,
            out long colOffsetEmu, out long rowOffsetEmu)
        {
            col = 0; row = 0; colOffsetEmu = 0; rowOffsetEmu = 0;
            try
            {
                uint.TryParse(marker.ColumnId?.Text, out uint c);
                uint.TryParse(marker.RowId?.Text, out uint r);
                col = c + 1; row = r + 1;
                long.TryParse(marker.ColumnOffset?.Text, out long co);
                long.TryParse(marker.RowOffset?.Text, out long ro);
                colOffsetEmu = co; rowOffsetEmu = ro;
            }
            catch { }
        }

        /// <summary>Convert EMU to points. 1 pt = 12700 EMU.</summary>
        public static double EmuToPt(long emu) => emu / 12700.0;
    }
}
