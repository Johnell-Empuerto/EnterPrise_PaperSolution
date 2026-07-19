using System.Xml.Linq;
using ExcelAPI.Models;

namespace ExcelAPI.Application
{
    public class XmlGenerator
    {
        public string Generate(FormDefinition form)
        {
            var doc = new XDocument(
                new XElement("top",
                    new XElement("metadata",
                        new XElement("title", form.Workbook.Title),
                        new XElement("author", form.Workbook.Author),
                        new XElement("created", form.Workbook.Created),
                        new XElement("modified", form.Workbook.Modified),
                        new XElement("version", form.Workbook.Version),
                        new XElement("description", form.Workbook.Description)
                    ),
                    GenerateSheetsXml(form),
                    GenerateClustersXml(form),
                    GenerateImagesXml(form),
                    new XElement("metadata_items",
                        form.Metadata.Select(kv =>
                            new XElement("item",
                                new XAttribute("key", kv.Key),
                                new XAttribute("value", kv.Value)
                            )
                        )
                    )
                )
            );

            return doc.Declaration + Environment.NewLine + doc.ToString();
        }

        private XElement GenerateSheetsXml(FormDefinition form)
        {
            return new XElement("sheets",
                form.Sheets.Select(s => new XElement("sheet",
                    new XAttribute("id", s.Id),
                    new XAttribute("name", s.Name),
                    new XAttribute("index", s.Index),
                    new XElement("page_settings",
                        new XElement("paper_size", s.PageSettings.PaperSize),
                        new XElement("orientation", s.PageSettings.Orientation),
                        new XElement("width_pt", s.PageSettings.WidthPt),
                        new XElement("height_pt", s.PageSettings.HeightPt),
                        new XElement("margins",
                            new XElement("left", s.PageSettings.LeftMargin),
                            new XElement("top", s.PageSettings.TopMargin),
                            new XElement("right", s.PageSettings.RightMargin),
                            new XElement("bottom", s.PageSettings.BottomMargin)
                        ),
                        new XElement("center_horizontally", s.PageSettings.CenterHorizontally),
                        new XElement("center_vertically", s.PageSettings.CenterVertically),
                        new XElement("zoom", s.PageSettings.Zoom),
                        new XElement("fit_to_pages_wide", s.PageSettings.FitToPagesWide),
                        new XElement("fit_to_pages_tall", s.PageSettings.FitToPagesTall)
                    ),
                    s.PrintArea != null
                        ? new XElement("print_area",
                            new XAttribute("address", s.PrintArea.Address),
                            new XAttribute("left_pt", s.PrintArea.LeftPt),
                            new XAttribute("top_pt", s.PrintArea.TopPt),
                            new XAttribute("width_pt", s.PrintArea.WidthPt),
                            new XAttribute("height_pt", s.PrintArea.HeightPt),
                            new XAttribute("cols", s.PrintArea.Cols),
                            new XAttribute("rows", s.PrintArea.Rows)
                        )
                        : null,
                    s.BackgroundImage != null
                        ? new XElement("background_image", s.BackgroundImage)
                        : null,
                    new XElement("row_heights",
                        s.RowHeights.Select(rh =>
                            new XElement("row",
                                new XAttribute("index", rh.Key),
                                new XAttribute("height", rh.Value)
                            )
                        )
                    ),
                    new XElement("column_widths",
                        s.ColumnWidths.Select(cw =>
                            new XElement("column",
                                new XAttribute("index", cw.Key),
                                new XAttribute("width", cw.Value)
                            )
                        )
                    ),
                    new XElement("merged_cells",
                        s.MergedCells.Select(mc => new XElement("merged_cell",
                            new XAttribute("address", mc.Address),
                            new XAttribute("cell_address", mc.CellAddress),
                            new XAttribute("left_pt", mc.LeftPt),
                            new XAttribute("top_pt", mc.TopPt),
                            new XAttribute("width_pt", mc.WidthPt),
                            new XAttribute("height_pt", mc.HeightPt)
                        ))
                    ),
                    s.FreezePane != null
                        ? new XElement("freeze_pane", s.FreezePane)
                        : null
                ))
            );
        }

        private XElement GenerateClustersXml(FormDefinition form)
        {
            return new XElement("clusters",
                form.Clusters.Select(c => new XElement("cluster",
                    new XAttribute("id", c.ClusterId),
                    new XAttribute("name", c.Name),
                    new XAttribute("type", c.Type),
                    new XAttribute("sheet_id", c.SheetId),
                    new XElement("cell_address", c.CellAddress),
                    new XElement("left", c.Left),
                    new XElement("right", c.Right),
                    new XElement("top", c.Top),
                    new XElement("bottom", c.Bottom),
                    new XElement("left_pt", c.LeftPt),
                    new XElement("top_pt", c.TopPt),
                    new XElement("width_pt", c.WidthPt),
                    new XElement("height_pt", c.HeightPt),
                    new XElement("input_parameters",
                        c.InputParameters.Select(kv =>
                            new XElement("parameter",
                                new XAttribute("key", kv.Key),
                                new XAttribute("value", kv.Value)
                            )
                        )
                    ),
                    new XElement("visibility", c.Visibility),
                    new XElement("readonly", c.Readonly),
                    new XElement("remarks", c.Remarks),
                    new XElement("functions",
                        c.Functions.Select(f => new XElement("function", f))
                    ),
                    new XElement("metadata",
                        c.Metadata.Select(kv =>
                            new XElement("item",
                                new XAttribute("key", kv.Key),
                                new XAttribute("value", kv.Value)
                            )
                        )
                    )
                ))
            );
        }

        private XElement GenerateImagesXml(FormDefinition form)
        {
            return new XElement("images",
                form.Images.Select(img => new XElement("image",
                    new XAttribute("id", img.Id),
                    new XAttribute("sheet_id", img.SheetId),
                    new XAttribute("name", img.Name),
                    new XElement("left_pt", img.LeftPt),
                    new XElement("top_pt", img.TopPt),
                    new XElement("width_pt", img.WidthPt),
                    new XElement("height_pt", img.HeightPt),
                    new XElement("format", img.Format),
                    img.Data != null
                        ? new XElement("data", Convert.ToBase64String(img.Data))
                        : null
                ))
            );
        }
    }
}
