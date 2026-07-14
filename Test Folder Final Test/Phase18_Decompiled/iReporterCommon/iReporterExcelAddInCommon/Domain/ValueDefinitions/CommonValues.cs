using System.Collections.Generic;

namespace iReporterExcelAddInCommon.Domain.ValueDefinitions;

public static class CommonValues
{
	public static Dictionary<string, string> LangNameCultureName = new Dictionary<string, string>
	{
		{ "japanese", "ja" },
		{ "english", "en" },
		{ "chinese", "zh-CN" },
		{ "traditionalChinese", "zh-TW" }
	};
}
