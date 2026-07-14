using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using Cimtops.R2Cluster;

namespace iReporterExcelAddInCommon;

public class InputFormController2
{
	private static InputFormController2 _instance;

	public ApplicationConfig ApplicationConfigData { get; private set; }

	public UserConfig UserConfigData => UserConfig.GetInstance();

	public static InputFormController2 GetInstance()
	{
		return _instance;
	}

	static InputFormController2()
	{
		_instance = new InputFormController2();
		_instance.InitializeApplicationConfig();
	}

	private void InitializeApplicationConfig()
	{
		using Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("iReporterExcelAddInCommon.ApplicationConfig.xml");
		int num = (int)stream.Length;
		BinaryReader binaryReader = new BinaryReader(stream);
		MemoryStream memoryStream = new MemoryStream(num);
		memoryStream.Write(binaryReader.ReadBytes(num), 0, num);
		memoryStream.Seek(0L, SeekOrigin.Begin);
		ApplicationConfigData = ApplicationConfig.LoadFromStream(memoryStream);
	}

	public ClusterTypeInfo[] GetClusterList(ClusterList selected)
	{
		List<ClusterTypeInfo> list = ApplicationConfigData.ListClusters(UserConfigData.SelectedDeviceType);
		if (0 < selected.Count)
		{
			IEnumerable<Point> points = selected.Select((ClusterInfo c) => new Point(c.Col, c.Row));
			foreach (ClusterTypeInfo item in list)
			{
				item.Likelihood = Decoder.GetLikelihood(selected.ExcelFilePath, selected.SheetName, points, item.TypeKey);
			}
			return list.OrderByDescending((ClusterTypeInfo cl) => cl.Likelihood).ToArray();
		}
		foreach (ClusterTypeInfo item2 in list)
		{
			item2.Likelihood = 0.5;
		}
		return list.ToArray();
	}

	public string GetSelectedCulture()
	{
		return UserConfigData.Language.Culture;
	}
}
