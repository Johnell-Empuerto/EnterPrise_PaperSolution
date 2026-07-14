using System;
using System.Linq;

namespace iReporterExcelAddInCommon.Tables;

public class TableCol : TableLine
{
	public int X { get; }

	public int R { get; }

	public string Name
	{
		get
		{
			return base.Items.First().ColName;
		}
		set
		{
			foreach (TableItem item in base.Items)
			{
				item.ColName = value;
			}
		}
	}

	public string Key
	{
		get
		{
			return base.Items.First().ColKey;
		}
		set
		{
			foreach (TableItem item in base.Items)
			{
				item.ColKey = value;
			}
		}
	}

	public ColType Type
	{
		get
		{
			if (!Enum.TryParse<ColType>(base.Items.First().ColType, out var result))
			{
				return ColType.NONE;
			}
			return result;
		}
		set
		{
			foreach (TableItem item in base.Items)
			{
				item.ColType = value.ToString();
			}
		}
	}

	public string Kind { get; set; }

	public bool IsCalc { get; set; }

	public bool IsNotOutput => Type == ColType.NONE;

	public TableCol(TableItem item)
	{
		X = item.X;
		R = item.R;
		Add(item);
	}
}
