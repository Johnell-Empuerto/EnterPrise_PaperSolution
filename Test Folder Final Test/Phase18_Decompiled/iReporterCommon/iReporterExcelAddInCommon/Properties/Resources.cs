using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Resources;
using System.Runtime.CompilerServices;

namespace iReporterExcelAddInCommon.Properties;

[GeneratedCode("System.Resources.Tools.StronglyTypedResourceBuilder", "15.0.0.0")]
[DebuggerNonUserCode]
[CompilerGenerated]
public class Resources
{
	private static ResourceManager resourceMan;

	private static CultureInfo resourceCulture;

	[EditorBrowsable(EditorBrowsableState.Advanced)]
	public static ResourceManager ResourceManager
	{
		get
		{
			if (resourceMan == null)
			{
				resourceMan = new ResourceManager("iReporterExcelAddInCommon.Properties.Resources", typeof(Resources).Assembly);
			}
			return resourceMan;
		}
	}

	[EditorBrowsable(EditorBrowsableState.Advanced)]
	public static CultureInfo Culture
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

	public static string ACTION_KIND => ResourceManager.GetString("ACTION_KIND", resourceCulture);

	public static string ACTION_KIND_AUTO_INPUT_CLUSTER => ResourceManager.GetString("ACTION_KIND_AUTO_INPUT_CLUSTER", resourceCulture);

	public static string ACTION_KIND_BATCHCLEAR => ResourceManager.GetString("ACTION_KIND_BATCHCLEAR", resourceCulture);

	public static string ACTION_KIND_BIOMETRICS => ResourceManager.GetString("ACTION_KIND_BIOMETRICS", resourceCulture);

	public static string ACTION_KIND_BTL => ResourceManager.GetString("ACTION_KIND_BTL", resourceCulture);

	public static string ACTION_KIND_CIOTG => ResourceManager.GetString("ACTION_KIND_CIOTG", resourceCulture);

	public static string ACTION_KIND_DOCUMENT => ResourceManager.GetString("ACTION_KIND_DOCUMENT", resourceCulture);

	public static string ACTION_KIND_MENU => ResourceManager.GetString("ACTION_KIND_MENU", resourceCulture);

	public static string ACTION_KIND_NO_NTFO => ResourceManager.GetString("ACTION_KIND_NO_NTFO", resourceCulture);

	public static string ACTION_KIND_OPEN_URL => ResourceManager.GetString("ACTION_KIND_OPEN_URL", resourceCulture);

	public static string ACTION_KIND_OUTPUT_TEXT => ResourceManager.GetString("ACTION_KIND_OUTPUT_TEXT", resourceCulture);

	public static string ACTION_KIND_QR_CODE => ResourceManager.GetString("ACTION_KIND_QR_CODE", resourceCulture);

	public static string ACTION_KIND_RUN_COMMAND => ResourceManager.GetString("ACTION_KIND_RUN_COMMAND", resourceCulture);

	public static string ACTION_KIND_SHEET_COPY => ResourceManager.GetString("ACTION_KIND_SHEET_COPY", resourceCulture);

	public static string ACTION_KIND_SHEET_JUMP => ResourceManager.GetString("ACTION_KIND_SHEET_JUMP", resourceCulture);

	public static string ACTION_KIND_START_TIMER => ResourceManager.GetString("ACTION_KIND_START_TIMER", resourceCulture);

	public static string CLUSTER_CAPTION => ResourceManager.GetString("CLUSTER_CAPTION", resourceCulture);

	public static string CLUSTER_HAS_TABLE => ResourceManager.GetString("CLUSTER_HAS_TABLE", resourceCulture);

	public static string CLUSTER_KIND => ResourceManager.GetString("CLUSTER_KIND", resourceCulture);

	public static string CLUSTER_TYPE_ACTION => ResourceManager.GetString("CLUSTER_TYPE_ACTION", resourceCulture);

	public static string CLUSTER_TYPE_DATE => ResourceManager.GetString("CLUSTER_TYPE_DATE", resourceCulture);

	public static string CLUSTER_TYPE_NUM => ResourceManager.GetString("CLUSTER_TYPE_NUM", resourceCulture);

	public static string CLUSTER_TYPE_TEXT => ResourceManager.GetString("CLUSTER_TYPE_TEXT", resourceCulture);

	public static string CLUSTER_TYPE_TIME => ResourceManager.GetString("CLUSTER_TYPE_TIME", resourceCulture);

	public static string DELETE => ResourceManager.GetString("DELETE", resourceCulture);

	public static string FAILED_FILE_LOAD => ResourceManager.GetString("FAILED_FILE_LOAD", resourceCulture);

	public static string FAILED_FILE_SAVE => ResourceManager.GetString("FAILED_FILE_SAVE", resourceCulture);

	public static string INVALID_COL_TYPE => ResourceManager.GetString("INVALID_COL_TYPE", resourceCulture);

	public static string MULTI_CALC_FORMATS => ResourceManager.GetString("MULTI_CALC_FORMATS", resourceCulture);

	public static string NO_TABLE => ResourceManager.GetString("NO_TABLE", resourceCulture);

	public static string NOT_FOUND_CLUSTER => ResourceManager.GetString("NOT_FOUND_CLUSTER", resourceCulture);

	public static string NOT_FOUND_CLUSTERS => ResourceManager.GetString("NOT_FOUND_CLUSTERS", resourceCulture);

	public static string NOT_FOUND_SHEET => ResourceManager.GetString("NOT_FOUND_SHEET", resourceCulture);

	public static string NOT_OUTPUT_TYPE => ResourceManager.GetString("NOT_OUTPUT_TYPE", resourceCulture);

	public static string NOT_SAME_CLUSTER_COL => ResourceManager.GetString("NOT_SAME_CLUSTER_COL", resourceCulture);

	public static string SHEET_NAME_EMPTY => ResourceManager.GetString("SHEET_NAME_EMPTY", resourceCulture);

	public static string SHEET_NAME_INVALID => ResourceManager.GetString("SHEET_NAME_INVALID", resourceCulture);

	public static string TABLE_COL_KEY_EMPTY => ResourceManager.GetString("TABLE_COL_KEY_EMPTY", resourceCulture);

	public static string TABLE_COL_KEY_INVALID => ResourceManager.GetString("TABLE_COL_KEY_INVALID", resourceCulture);

	public static string TABLE_COL_KEY_OVERLAP => ResourceManager.GetString("TABLE_COL_KEY_OVERLAP", resourceCulture);

	public static string TABLE_COL_NAME_EMPTY => ResourceManager.GetString("TABLE_COL_NAME_EMPTY", resourceCulture);

	public static string TABLE_COL_OVERLAP => ResourceManager.GetString("TABLE_COL_OVERLAP", resourceCulture);

	public static string TABLE_INVALID => ResourceManager.GetString("TABLE_INVALID", resourceCulture);

	public static string TABLE_NAME_EMPTY => ResourceManager.GetString("TABLE_NAME_EMPTY", resourceCulture);

	public static string TABLE_NO_CLUSTERS => ResourceManager.GetString("TABLE_NO_CLUSTERS", resourceCulture);

	public static string TABLE_NO_INVALID => ResourceManager.GetString("TABLE_NO_INVALID", resourceCulture);

	public static string TABLE_NO_OVERLAP => ResourceManager.GetString("TABLE_NO_OVERLAP", resourceCulture);

	public static string TABLE_ROW_OVERLAP => ResourceManager.GetString("TABLE_ROW_OVERLAP", resourceCulture);

	public static string TABLE_SAVED => ResourceManager.GetString("TABLE_SAVED", resourceCulture);

	public static string TYPE_NAME_DATE => ResourceManager.GetString("TYPE_NAME_DATE", resourceCulture);

	public static string TYPE_NAME_INTERVAL => ResourceManager.GetString("TYPE_NAME_INTERVAL", resourceCulture);

	public static string TYPE_NAME_NUMERIC => ResourceManager.GetString("TYPE_NAME_NUMERIC", resourceCulture);

	public static string TYPE_NAME_TEXT => ResourceManager.GetString("TYPE_NAME_TEXT", resourceCulture);

	internal Resources()
	{
	}
}
