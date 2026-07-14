using System.Collections.Generic;

namespace iReporterExcelAddInCommon.Tables;

public class TableLine
{
	protected readonly List<TableItem> items = new List<TableItem>();

	public int Index { get; set; }

	public IEnumerable<TableItem> Items => items;

	public void Add(TableItem item)
	{
		items.Add(item);
	}
}
