using Microsoft.AspNetCore.Mvc;
using ExcelAPI.LegacyEngine.OpenXml;
using ExcelAPI.Models;

namespace ExcelAPI.Controllers;

[ApiController]
[Route("api/[controller]")]
public class TemplateController : ControllerBase
{
    private readonly IOpenXmlWorkbookLoader _loader;
    private readonly IWebHostEnvironment _env;
    private readonly ILogger<TemplateController> _logger;

    // Known template IDs mapped to xlsx file paths (relative to ContentRootPath)
    private static readonly Dictionary<int, string> TemplatePaths = new()
    {
        [546] = @"..\..\Investigation_546\original.xlsx",
        [547] = @"..\..\[V3.1_Sample]アンケート用紙.xlsx",
        [548] = @"Forms\0f8e09082b5c4861abe5d92985ace548.xlsx"
    };

    public TemplateController(
        IOpenXmlWorkbookLoader loader,
        IWebHostEnvironment env,
        ILogger<TemplateController> logger)
    {
        _loader = loader;
        _env = env;
        _logger = logger;
    }

    [HttpGet("{id}")]
    [ProducesResponseType(typeof(TemplateModelDto), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public IActionResult GetTemplate(int id)
    {
        if (!TemplatePaths.TryGetValue(id, out var relativePath))
            return NotFound(new { error = $"Template {id} not found" });

        var fullPath = Path.Combine(_env.ContentRootPath, relativePath);
        if (!System.IO.File.Exists(fullPath))
            return NotFound(new { error = $"Template file not found: {fullPath}" });

        try
        {
            var workbook = _loader.Load(fullPath);
            var sheet = workbook.Sheets.FirstOrDefault();
            if (sheet is null)
                return NotFound(new { error = "No sheets found in workbook" });

            var dto = MapToDto(sheet);
            return Ok(dto);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to load template {TemplateId}", id);
            return StatusCode(500, new { error = $"Failed to load template: {ex.Message}" });
        }
    }

    private static TemplateModelDto MapToDto(OpenXmlSheetModel sheet)
    {
        var dto = new TemplateModelDto
        {
            SheetName = sheet.Name,
            PageSetup = new TemplatePageSetupDto
            {
                PaperWidthPt = sheet.PageSetup.PaperWidthPoints,
                PaperHeightPt = sheet.PageSetup.PaperHeightPoints,
                MarginLeftIn = sheet.PageSetup.MarginLeft,
                MarginRightIn = sheet.PageSetup.MarginRight,
                MarginTopIn = sheet.PageSetup.MarginTop,
                MarginBottomIn = sheet.PageSetup.MarginBottom,
                MarginHeaderIn = sheet.PageSetup.MarginHeader,
                MarginFooterIn = sheet.PageSetup.MarginFooter,
                CenterHorizontally = sheet.PageSetup.CenterHorizontally,
                CenterVertically = sheet.PageSetup.CenterVertically,
                Orientation = sheet.PageSetup.Orientation,
                PaperSize = 1
            },
            PrintArea = new TemplatePrintAreaDto
            {
                StartCol = sheet.PrintArea.StartColumn,
                StartRow = sheet.PrintArea.StartRow,
                EndCol = sheet.PrintArea.EndColumn,
                EndRow = sheet.PrintArea.EndRow
            },
            ColumnWidths = sheet.ColumnWidths.ToList(),
            RowHeights = sheet.RowHeights.ToList(),
            MergedCells = sheet.MergeAreas.Select(m => new TemplateMergedCellDto
            {
                StartCol = m.StartColumn,
                StartRow = m.StartRow,
                EndCol = m.EndColumn,
                EndRow = m.EndRow
            }).ToList(),
            Comments = sheet.Comments.Select(c => new TemplateCommentDto
            {
                Address = $"${GetColumnLetter(c.Column)}${c.Row}",
                Text = c.Text
            }).ToList(),
            Images = sheet.Images.Select(i => new TemplateImageDto
            {
                AnchorCol = i.Column,
                AnchorRow = i.Row,
                WidthPt = 0,
                HeightPt = 0
            }).ToList()
        };

        // Cell values and styles keyed by A1 address
        foreach (var cell in sheet.Cells)
        {
            string address = GetColumnLetter(cell.Column) + cell.Row;
            dto.CellValues[address] = cell.Value;
        }

        foreach (var kvp in sheet.CellStyles)
        {
            var style = kvp.Value;
            var styleDto = new TemplateCellStyleDto();

            if (!string.IsNullOrEmpty(style.FontName))
                styleDto.FontName = style.FontName;
            if (style.FontSize > 0)
                styleDto.FontSize = style.FontSize;
            if (style.Bold)
                styleDto.Bold = true;
            if (style.Italic)
                styleDto.Italic = true;
            if (style.Underline)
                styleDto.Underline = true;
            if (!string.IsNullOrEmpty(style.FontColor))
                styleDto.FontColor = style.FontColor;
            if (!string.IsNullOrEmpty(style.FillColor))
                styleDto.FillColor = style.FillColor;
            if (!string.IsNullOrEmpty(style.BorderTop))
                styleDto.BorderTop = style.BorderTop;
            if (!string.IsNullOrEmpty(style.BorderBottom))
                styleDto.BorderBottom = style.BorderBottom;
            if (!string.IsNullOrEmpty(style.BorderLeft))
                styleDto.BorderLeft = style.BorderLeft;
            if (!string.IsNullOrEmpty(style.BorderRight))
                styleDto.BorderRight = style.BorderRight;
            if (!string.IsNullOrEmpty(style.HorizontalAlignment))
                styleDto.HorizontalAlignment = style.HorizontalAlignment;
            if (!string.IsNullOrEmpty(style.VerticalAlignment))
                styleDto.VerticalAlignment = style.VerticalAlignment;
            if (style.WrapText)
                styleDto.WrapText = true;
            if (style.Indent > 0)
                styleDto.Indent = style.Indent;

            dto.CellStyles[kvp.Key] = styleDto;
        }

        return dto;
    }

    private static string GetColumnLetter(int col)
    {
        if (col < 1) return "";
        col--;
        string result = "";
        while (col >= 0)
        {
            result = (char)('A' + col % 26) + result;
            col = col / 26 - 1;
        }
        return result;
    }
}
