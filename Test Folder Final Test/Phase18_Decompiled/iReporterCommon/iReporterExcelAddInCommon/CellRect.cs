using Cimtops.Excel;

namespace iReporterExcelAddInCommon;

public class CellRect
{
	public int Top { get; private set; }

	public int Left { get; private set; }

	public int Bottom { get; private set; }

	public int Right { get; private set; }

	internal CellRect(int top, int left, int bottom, int right)
	{
		Top = top;
		Left = left;
		Bottom = bottom;
		Right = right;
	}

	public bool Union(CellRect other)
	{
		if (this == other)
		{
			return false;
		}
		if (Top == other.Top && Bottom == other.Bottom)
		{
			if (Right + 1 == other.Left)
			{
				Right = other.Right;
				return true;
			}
			if (other.Right + 1 == Left)
			{
				Left = other.Left;
				return true;
			}
		}
		if (Left == other.Left && Right == other.Right)
		{
			if (Bottom + 1 == other.Top)
			{
				Bottom = other.Bottom;
				return true;
			}
			if (other.Bottom + 1 == Top)
			{
				Top = other.Top;
				return true;
			}
		}
		return false;
	}

	public override string ToString()
	{
		string text = XLRangeUtil.ToAddressName(Top, Left, 1);
		if (Top != Bottom || Left != Right)
		{
			text = text + ":" + XLRangeUtil.ToAddressName(Bottom, Right, 1);
		}
		return text;
	}
}
