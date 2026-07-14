using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Cimtops.Excel;
using Cimtops.R2Cluster;

namespace iReporterExcelAddInCommon;

public class ClusterList : IComparer<ClusterInfo>, IEnumerable<ClusterInfo>, IEnumerable
{
	private List<ClusterInfo> _list = new List<ClusterInfo>();

	public int Count => _list.Count;

	public string ClusterTypeKey => GetLine((ClusterInfo c) => c.ClusterTypeKey, "-");

	public string ExcelFilePath { get; private set; }

	public string SheetName { get; private set; }

	public TitleInfo Title { get; private set; }

	public bool CountChanged { get; set; }

	public bool HasTable => _list.Any((ClusterInfo c) => c.TableNo.HasValue);

	public ClusterList(string filePath, string sheetName, TitleInfo title)
	{
		ExcelFilePath = filePath;
		SheetName = sheetName;
		Title = title;
		CountChanged = true;
	}

	public void Add(int row, int col, int rowCount, int colCount, string comment, bool selected)
	{
		Decoder.AIInfo aIInfo = Decoder.IsNoConfidence(ExcelFilePath, SheetName, new Point(col, row));
		_list.Add(new ClusterInfo(row, col, rowCount, colCount, comment, selected, aIInfo.UseAI, aIInfo.IsNoConfi));
	}

	public void Sort()
	{
		_list.Sort(this);
	}

	public void Copy(ClusterList src)
	{
		if (Count == src.Count)
		{
			for (int i = 0; i < _list.Count; i++)
			{
				_list[i].Copy(src._list[i]);
			}
		}
		else
		{
			_list = src._list;
			CountChanged = true;
		}
		ExcelFilePath = src.ExcelFilePath;
		SheetName = src.SheetName;
		Title = src.Title;
	}

	public string GetLine(int index, string multi = "", string none = "")
	{
		return GetLine((ClusterInfo c) => c.GetLine(index), multi, none);
	}

	public ClusterList GetSelected()
	{
		return new ClusterList(ExcelFilePath, SheetName, Title)
		{
			_list = _list.FindAll((ClusterInfo c) => c.IsSelected)
		};
	}

	public int Compare(ClusterInfo x, ClusterInfo y)
	{
		return Comp(x.ClusterIndex, y.ClusterIndex) ?? Comp(x.Row, y.Row) ?? Comp(x.Col, y.Col) ?? 0;
	}

	private int? Comp(int x, int y)
	{
		if (x < y)
		{
			return -1;
		}
		if (y < x)
		{
			return 1;
		}
		return null;
	}

	internal ClusterInfo[] SetupArray(ClusterTypeInfo[] clusterTypes, DeviceType device)
	{
		foreach (ClusterInfo info in _list)
		{
			info.Setup(clusterTypes.FirstOrDefault((ClusterTypeInfo t) => t.TypeKey == info.ClusterTypeKey), device);
		}
		return _list.ToArray();
	}

	private string GetLine(Func<ClusterInfo, string> func, string multi = "", string none = "")
	{
		if (_list.Count == 0)
		{
			return none;
		}
		string text = func(_list[0]);
		for (int i = 1; i < _list.Count; i++)
		{
			string text2 = func(_list[i]);
			if (text != text2)
			{
				return multi;
			}
		}
		return text;
	}

	public IEnumerator<ClusterInfo> GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _list.GetEnumerator();
	}
}
