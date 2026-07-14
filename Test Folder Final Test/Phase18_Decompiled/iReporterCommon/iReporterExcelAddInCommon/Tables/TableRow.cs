using System.Linq;

namespace iReporterExcelAddInCommon.Tables;

public class TableRow : TableLine
{
	public int Y { get; }

	public int B { get; }

	public string Name
	{
		get
		{
			return base.Items.First().RowName;
		}
		set
		{
			foreach (TableItem item in base.Items)
			{
				item.RowName = value;
			}
		}
	}

	public TableRow(TableItem item)
	{
		Y = item.Y;
		B = item.B;
		items.Add(item);
	}
}
