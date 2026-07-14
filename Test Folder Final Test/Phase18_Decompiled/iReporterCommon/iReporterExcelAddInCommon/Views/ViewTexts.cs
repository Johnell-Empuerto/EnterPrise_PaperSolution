using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Resources;
using System.Runtime.CompilerServices;

namespace iReporterExcelAddInCommon.Views;

[GeneratedCode("System.Resources.Tools.StronglyTypedResourceBuilder", "15.0.0.0")]
[DebuggerNonUserCode]
[CompilerGenerated]
public class ViewTexts
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
				resourceMan = new ResourceManager("iReporterExcelAddInCommon.Views.resources.ViewTexts", typeof(ViewTexts).Assembly);
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

	public static string Cancel => ResourceManager.GetString("Cancel", resourceCulture);

	public static string MessageGrid => ResourceManager.GetString("MessageGrid", resourceCulture);

	public static string MessageGridGridCopy => ResourceManager.GetString("MessageGridGridCopy", resourceCulture);

	public static string MessageGridTextCopy => ResourceManager.GetString("MessageGridTextCopy", resourceCulture);

	public static string OK => ResourceManager.GetString("OK", resourceCulture);

	internal ViewTexts()
	{
	}
}
