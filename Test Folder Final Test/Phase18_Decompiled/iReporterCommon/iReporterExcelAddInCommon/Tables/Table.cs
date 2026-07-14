using System;
using System.Collections.Generic;
using System.Linq;
using iReporterExcelAddInCommon.Properties;

namespace iReporterExcelAddInCommon.Tables;

public class Table
{
	private readonly List<TableRow> rows = new List<TableRow>();

	private readonly List<TableCol> cols = new List<TableCol>();

	public static Table Empty { get; } = new Table();

	public bool IsReal { get; set; }

	public string No
	{
		get
		{
			return Item?.TableNo ?? "";
		}
		set
		{
			foreach (TableItem item in Items)
			{
				item.TableNo = value;
			}
		}
	}

	public string Name
	{
		get
		{
			return Item?.TableName ?? "";
		}
		set
		{
			foreach (TableItem item in Items)
			{
				item.TableName = value;
			}
		}
	}

	public bool OutputCollabo
	{
		get
		{
			return Item?.OutputCollabo ?? false;
		}
		set
		{
			foreach (TableItem item in Items)
			{
				item.OutputCollabo = value;
			}
		}
	}

	public string SheetName { get; }

	public IEnumerable<TableRow> Rows => rows;

	public IEnumerable<TableCol> Cols => cols;

	private IEnumerable<TableItem> Items => Rows.SelectMany((TableRow r) => r.Items);

	private TableItem Item => Items.FirstOrDefault();

	public int RowCount => rows.Count;

	public int ColCount => cols.Count;

	private Table()
	{
		SheetName = "";
	}

	public Table(string sheetName, TableItem item, bool isReal)
	{
		SheetName = sheetName;
		IsReal = isReal;
		Add(item);
	}

	public void Add(TableItem item)
	{
		TableRow tableRow = Rows.FirstOrDefault((TableRow r) => r.Y == item.Y && r.B == item.B);
		if (tableRow == null)
		{
			rows.Add(tableRow = new TableRow(item));
		}
		else
		{
			tableRow.Add(item);
		}
		item.Row = tableRow;
		TableCol tableCol = Cols.FirstOrDefault((TableCol r) => r.X == item.X && r.R == item.R);
		if (tableCol == null)
		{
			cols.Add(tableCol = new TableCol(item));
		}
		else
		{
			tableCol.Add(item);
		}
		item.Col = tableCol;
	}

	public void Adjust()
	{
		rows.Sort((TableRow lhs, TableRow rhs) => lhs.Y - rhs.Y);
		cols.Sort((TableCol lhs, TableCol rhs) => lhs.X - rhs.X);
		for (int num = 0; num < rows.Count; num++)
		{
			TableRow tableRow = rows[num];
			tableRow.Index = num + 1;
			foreach (TableItem item in tableRow.Items)
			{
				item.RowIndex = tableRow.Index;
			}
		}
		for (int num2 = 0; num2 < cols.Count; num2++)
		{
			TableCol tableCol = cols[num2];
			tableCol.Index = num2 + 1;
			foreach (TableItem item2 in tableCol.Items)
			{
				item2.ColIndex = tableCol.Index;
			}
		}
	}

	public string Check()
	{
		foreach (TableCol col in Cols)
		{
			string[] ary = col.Items.Select((TableItem i) => i.ClusterKind).Distinct().ToArray();
			if (ary.Length != 1)
			{
				return Resources.NOT_SAME_CLUSTER_COL;
			}
			ClusterTypeInfo clusterTypeInfo = InputFormController2.GetInstance().ApplicationConfigData.ListClusters().FirstOrDefault((ClusterTypeInfo c) => c.TypeKey == ary[0]);
			if (clusterTypeInfo == null)
			{
				return Resources.NOT_FOUND_CLUSTER;
			}
			for (int num = 1; num < cols.Count; num++)
			{
				if (cols[num].X <= cols[num - 1].R)
				{
					return Resources.TABLE_COL_OVERLAP;
				}
			}
			for (int num2 = 1; num2 < rows.Count; num2++)
			{
				if (rows[num2].Y <= rows[num2 - 1].B)
				{
					return Resources.TABLE_ROW_OVERLAP;
				}
			}
			col.Kind = clusterTypeInfo.CultureName;
			Enum.TryParse<ColType>(clusterTypeInfo.TableColType, ignoreCase: true, out var result);
			col.IsCalc = result == ColType.calculate;
			if (col.IsCalc)
			{
				ColType type = col.Type;
				if ((uint)(type - 1) <= 3u)
				{
					col.Type = col.Type;
				}
				else
				{
					bool[] array = col.Items.Select((TableItem i) => i.Fmt.Contains('h') || i.Fmt.Contains('s')).Distinct().ToArray();
					if (1 < array.Length)
					{
						return Resources.MULTI_CALC_FORMATS;
					}
					col.Type = ((!array[0]) ? ColType.numeric : ColType.interval);
				}
			}
			else
			{
				col.Type = result;
				if (result == ColType.INVALID)
				{
					return Resources.INVALID_COL_TYPE;
				}
			}
			foreach (TableItem item in col.Items)
			{
				item.ColType = col.Type.ToString();
			}
		}
		return null;
	}

	public override string ToString()
	{
		return No.PadLeft(3) + ".  " + Name;
	}
}
