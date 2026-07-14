using System.Collections.Generic;
using System.Linq;

namespace iReporterExcelAddInCommon.Tables;

public class TableItem
{
	private const int FIRST = 16;

	private readonly List<string> comment;

	public int X { get; }

	public int Y { get; }

	public int R { get; }

	public int B { get; }

	public string Fmt { get; }

	public TableRow Row { get; internal set; }

	public TableCol Col { get; internal set; }

	public string TableNo
	{
		get
		{
			return Get(0);
		}
		set
		{
			Set(0, value);
		}
	}

	public string TableName
	{
		get
		{
			return Get(1);
		}
		set
		{
			Set(1, value);
		}
	}

	public bool OutputCollabo
	{
		get
		{
			return Get(2) == "1";
		}
		set
		{
			Set(2, value ? "1" : "0");
		}
	}

	public int ColIndex
	{
		set
		{
			Set(3, value.ToString());
		}
	}

	public string ColName
	{
		get
		{
			return Get(4);
		}
		set
		{
			Set(4, value);
		}
	}

	public string ColKey
	{
		get
		{
			return Get(5);
		}
		set
		{
			Set(5, value);
		}
	}

	public string ColType
	{
		get
		{
			return Get(6);
		}
		set
		{
			Set(6, value);
		}
	}

	public int RowIndex
	{
		set
		{
			Set(7, value.ToString());
		}
	}

	public string RowName
	{
		get
		{
			return Get(8);
		}
		set
		{
			Set(8, value);
		}
	}

	public string ClusterName => Get(-16);

	public string ClusterKind => Get(-15);

	public string ClusterIndex => Get(-14);

	public IEnumerable<string> Comments => comment;

	public void RemoveTableInfos()
	{
		for (int i = 16; i < comment.Count; i++)
		{
			comment[i] = "";
		}
	}

	public TableItem(int x, int y, int r, int b, string comment, string fmt)
	{
		X = x;
		Y = y;
		R = r;
		B = b;
		this.comment = comment.Replace("\r", "").Split('\n').ToList();
		Fmt = fmt;
	}

	private string Get(int index)
	{
		index += 16;
		if (0 <= index && index < comment.Count)
		{
			return comment[index];
		}
		return "";
	}

	private void Set(int index, string value)
	{
		index += 16;
		while (comment.Count <= index)
		{
			comment.Add("");
		}
		comment[index] = value.Replace("\r", "").Replace("\n", "");
	}
}
