using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using Cimtops.Excel;

namespace iReporterExcelAddInCommon;

public class ApplicationConfig
{
	private List<ClusterTypeInfo> clusterTypes = new List<ClusterTypeInfo>();

	public List<ClusterTypeInfo> ListClusters(DeviceType? os = null)
	{
		return clusterTypes.FindAll((ClusterTypeInfo c) => !os.HasValue || c.Available[os.Value]);
	}

	public static ApplicationConfig LoadFromStream(Stream stream)
	{
		ApplicationConfig applicationConfig = new ApplicationConfig();
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.Load(stream);
		foreach (XmlNode item2 in xmlDocument.SelectNodes("/conmas/clusters/cluster"))
		{
			string innerText = item2["key"].InnerText;
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			foreach (XmlNode item3 in item2.SelectNodes("text"))
			{
				string value = item3.Attributes["culture"].Value;
				string innerText2 = item3.InnerText;
				dictionary.Add(value, innerText2);
			}
			Dictionary<DeviceType, bool> dictionary2 = new Dictionary<DeviceType, bool>();
			Dictionary<DeviceType, string> dictionary3 = new Dictionary<DeviceType, string>();
			dictionary2[DeviceType.Windows] = false;
			dictionary2[DeviceType.IOS] = false;
			dictionary2[DeviceType.Android] = false;
			XmlNode xmlNode2 = item2["available"]["windows"];
			if (xmlNode2 != null)
			{
				if (xmlNode2.InnerText.Equals("1"))
				{
					dictionary2[DeviceType.Windows] = true;
				}
				else if (xmlNode2.Attributes["autoConvert"] != null)
				{
					dictionary3[DeviceType.Windows] = xmlNode2.Attributes["autoConvert"].Value;
				}
			}
			XmlNode xmlNode3 = item2["available"]["ios"];
			if (xmlNode3 != null)
			{
				if (xmlNode3.InnerText.Equals("1"))
				{
					dictionary2[DeviceType.IOS] = true;
				}
				else if (xmlNode3.Attributes["autoConvert"] != null)
				{
					dictionary3[DeviceType.IOS] = xmlNode3.Attributes["autoConvert"].Value;
				}
			}
			XmlNode xmlNode4 = item2["available"]["android"];
			if (xmlNode4 != null)
			{
				if (xmlNode4.InnerText.Equals("1"))
				{
					dictionary2[DeviceType.Android] = true;
				}
				else if (xmlNode4.Attributes["autoConvert"] != null)
				{
					dictionary3[DeviceType.Android] = xmlNode4.Attributes["autoConvert"].Value;
				}
			}
			ClusterTypeInfo item = new ClusterTypeInfo(innerText, dictionary, dictionary2, dictionary3, item2["tableColType"]?.InnerText ?? "text");
			applicationConfig.clusterTypes.Add(item);
		}
		return applicationConfig;
	}

	public string ClusterKeyForDevice(string clusterKey, DeviceType deviceType)
	{
		foreach (ClusterTypeInfo clusterType in clusterTypes)
		{
			if (clusterType.TypeKey.Equals(clusterKey))
			{
				if (clusterType.Available[deviceType])
				{
					return clusterKey;
				}
				if (clusterType.AutoConvert.Keys.Contains(deviceType))
				{
					return clusterType.AutoConvert[deviceType];
				}
				return clusterKey;
			}
		}
		return clusterKey;
	}
}
