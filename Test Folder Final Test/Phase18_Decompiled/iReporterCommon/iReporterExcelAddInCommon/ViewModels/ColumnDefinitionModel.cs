using System.Windows;

namespace iReporterExcelAddInCommon.ViewModels;

public class ColumnDefinitionModel
{
	public string Header { get; set; }

	public string PropertyName { get; set; }

	public ColumnType ColumnType { get; set; }

	public double? Width { get; set; }

	public double? MinWidth { get; set; }

	public bool IsReadOnly { get; set; } = true;

	public HorizontalAlignment Alignment { get; set; }

	public bool IsVisible { get; set; } = true;

	public int? MaxLength { get; set; }

	public bool IsTextWrapping { get; set; }

	public bool IsSortable { get; set; }
}
