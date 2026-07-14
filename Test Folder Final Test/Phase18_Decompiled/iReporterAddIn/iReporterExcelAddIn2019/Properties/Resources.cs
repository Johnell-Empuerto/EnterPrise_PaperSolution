using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Resources;
using System.Runtime.CompilerServices;

namespace iReporterExcelAddIn2019.Properties;

[GeneratedCode("System.Resources.Tools.StronglyTypedResourceBuilder", "15.0.0.0")]
[DebuggerNonUserCode]
[CompilerGenerated]
internal class Resources
{
	private static ResourceManager resourceMan;

	private static CultureInfo resourceCulture;

	[EditorBrowsable(EditorBrowsableState.Advanced)]
	internal static ResourceManager ResourceManager
	{
		get
		{
			if (resourceMan == null)
			{
				resourceMan = new ResourceManager("iReporterExcelAddIn2019.Properties.Resources", typeof(Resources).Assembly);
			}
			return resourceMan;
		}
	}

	[EditorBrowsable(EditorBrowsableState.Advanced)]
	internal static CultureInfo Culture
	{
		get
		{
			return resourceCulture;
		}
		set
		{
			resourceCulture = value;
		}
	}

	internal static string API_ERROR => ResourceManager.GetString("API_ERROR", resourceCulture);

	internal static string CHECK_SHEET_NAME => ResourceManager.GetString("CHECK_SHEET_NAME", resourceCulture);

	internal static string CheckSheetNameComplete => ResourceManager.GetString("CheckSheetNameComplete", resourceCulture);

	internal static string CheckSheetNameSuccess => ResourceManager.GetString("CheckSheetNameSuccess", resourceCulture);

	internal static string EMPTY_SHEET => ResourceManager.GetString("EMPTY_SHEET", resourceCulture);

	internal static string EXCEL_ERROR => ResourceManager.GetString("EXCEL_ERROR", resourceCulture);

	internal static string ExcelMessage64 => ResourceManager.GetString("ExcelMessage64", resourceCulture);

	internal static string ExcelMessage65 => ResourceManager.GetString("ExcelMessage65", resourceCulture);

	internal static string ExcelProcessorMessage5 => ResourceManager.GetString("ExcelProcessorMessage5", resourceCulture);

	internal static string MENU_EXTRACTION => ResourceManager.GetString("MENU_EXTRACTION", resourceCulture);

	internal static string MENU_KIND => ResourceManager.GetString("MENU_KIND", resourceCulture);

	internal static string MENU_SETTING => ResourceManager.GetString("MENU_SETTING", resourceCulture);

	internal static string MENU_TABLE => ResourceManager.GetString("MENU_TABLE", resourceCulture);

	internal static string MULTI_PRINT => ResourceManager.GetString("MULTI_PRINT", resourceCulture);

	internal static string SHEET_NAME => ResourceManager.GetString("SHEET_NAME", resourceCulture);

	internal static string SheetNameNoticeGridColumnNameHeader => ResourceManager.GetString("SheetNameNoticeGridColumnNameHeader", resourceCulture);

	internal static string SheetNameNoticeGridColumnSymbolHeader => ResourceManager.GetString("SheetNameNoticeGridColumnSymbolHeader", resourceCulture);

	internal static string SheetNameNoticeGridDesc => ResourceManager.GetString("SheetNameNoticeGridDesc", resourceCulture);

	internal static string SheetNameNoticeMessageDesc => ResourceManager.GetString("SheetNameNoticeMessageDesc", resourceCulture);

	internal static string SheetNameNoticeWindow => ResourceManager.GetString("SheetNameNoticeWindow", resourceCulture);

	internal static string TOO_MANY_SELECT => ResourceManager.GetString("TOO_MANY_SELECT", resourceCulture);

	internal Resources()
	{
	}
}
