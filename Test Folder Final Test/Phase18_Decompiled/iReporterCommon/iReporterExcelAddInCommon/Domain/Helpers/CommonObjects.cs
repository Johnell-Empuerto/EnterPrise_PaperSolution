namespace iReporterExcelAddInCommon.Domain.Helpers;

public static class CommonObjects
{
	public static string Language => InputFormController2.GetInstance().GetSelectedCulture();
}
