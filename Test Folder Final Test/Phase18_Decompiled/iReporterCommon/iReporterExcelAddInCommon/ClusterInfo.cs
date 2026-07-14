using Cimtops.Excel;

namespace iReporterExcelAddInCommon;

public class ClusterInfo
{
	private string[] _texts;

	public int Row { get; private set; }

	public int Col { get; private set; }

	private int Bottom { get; set; }

	private int Right { get; set; }

	public bool IsSelected { get; private set; }

	public bool IsUnknown { get; private set; }

	public bool UseAI { get; private set; }

	public bool IsNoConfi { get; private set; }

	public string TypeName { get; private set; }

	public string ClusterName => GetLine(0);

	public string ClusterTypeKey => GetLine(1);

	public int ClusterIndex
	{
		get
		{
			if (!int.TryParse(GetLine(2), out var result))
			{
				return -1;
			}
			return result;
		}
	}

	public int? TableNo
	{
		get
		{
			if (!int.TryParse(GetLine(16), out var result))
			{
				return null;
			}
			return result;
		}
	}

	internal void Copy(ClusterInfo src)
	{
		_texts = src._texts;
		Row = src.Row;
		Col = src.Col;
		Bottom = src.Bottom;
		Right = src.Right;
		IsSelected = src.IsSelected;
		UseAI = src.UseAI;
		IsNoConfi = src.IsNoConfi;
	}

	public CellRect GetRect()
	{
		return new CellRect(Row, Col, Bottom, Right);
	}

	public string GetLine(int index)
	{
		if ((uint)index < _texts.Length)
		{
			return _texts[index] ?? "";
		}
		return "";
	}

	public ClusterInfo(int row, int col, int rowCount, int colCount, string comment, bool selected, bool useAI, bool isNo)
	{
		Row = row;
		Col = col;
		Bottom = row + rowCount - 1;
		Right = col + colCount - 1;
		IsSelected = selected;
		UseAI = useAI;
		IsNoConfi = isNo;
		_texts = comment.Replace("\r", "").Split('\n');
	}

	public void Setup(ClusterTypeInfo type, DeviceType device)
	{
		IsUnknown = type == null || !type.Available[device];
		TypeName = (IsUnknown ? ClusterTypeKey : type.CultureName);
	}
}
