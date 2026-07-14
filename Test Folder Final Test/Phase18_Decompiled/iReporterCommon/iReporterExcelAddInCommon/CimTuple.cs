using System;

namespace iReporterExcelAddInCommon;

public class CimTuple
{
	public string Item1;

	public string Item2;

	public string Item3;

	public Tuple<string, string, string> ToTuple()
	{
		return Tuple.Create(Item1, Item2, Item3);
	}

	public static CimTuple Create(string item1, string item2, string item3)
	{
		return new CimTuple
		{
			Item1 = item1,
			Item2 = item2,
			Item3 = item3
		};
	}
}
