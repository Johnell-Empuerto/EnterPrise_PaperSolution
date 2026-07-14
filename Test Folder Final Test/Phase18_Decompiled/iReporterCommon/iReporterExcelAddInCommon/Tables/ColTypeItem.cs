namespace iReporterExcelAddInCommon.Tables;

public class ColTypeItem
{
	public ColType Type { get; set; }

	public override string ToString()
	{
		return Type.ToText();
	}
}
