using System.Collections.Generic;
using Cimtops.Excel;

namespace iReporterExcelAddInCommon;

public class ClusterTypeInfo
{
	public string TypeKey { get; }

	public Dictionary<string, string> TypeNames { get; }

	public Dictionary<DeviceType, bool> Available { get; }

	public Dictionary<DeviceType, string> AutoConvert { get; }

	public string TableColType { get; }

	public static string Culture { get; set; }

	public string CultureName
	{
		get
		{
			if (Culture == null)
			{
				Culture = InputFormController2.GetInstance().GetSelectedCulture();
			}
			if (!TypeNames.TryGetValue(Culture, out var value))
			{
				return TypeKey;
			}
			return value;
		}
	}

	public double Likelihood { get; set; }

	public ClusterTypeInfo(string typeKey, Dictionary<string, string> typeNames, Dictionary<DeviceType, bool> available, Dictionary<DeviceType, string> autoConvert, string tableColType)
	{
		TypeKey = typeKey;
		TypeNames = typeNames;
		Available = available;
		AutoConvert = autoConvert;
		TableColType = tableColType;
	}

	public override string ToString()
	{
		return CultureName;
	}
}
