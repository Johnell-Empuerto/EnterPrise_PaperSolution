using iReporterExcelAddInCommon.Properties;

namespace iReporterExcelAddInCommon.Tables;

public static class A
{
	public static string ToText(this ColType type)
	{
		return type switch
		{
			ColType.numeric => Resources.TYPE_NAME_NUMERIC, 
			ColType.date => Resources.TYPE_NAME_DATE, 
			ColType.interval => Resources.TYPE_NAME_INTERVAL, 
			ColType.text => Resources.TYPE_NAME_TEXT, 
			ColType.NONE => Resources.NOT_OUTPUT_TYPE, 
			_ => "", 
		};
	}
}
