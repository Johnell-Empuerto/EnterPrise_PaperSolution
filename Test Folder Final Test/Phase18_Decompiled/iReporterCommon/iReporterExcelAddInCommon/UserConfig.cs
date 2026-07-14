using System;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using Cimtops.Excel;
using Cimtops.R2Cluster;

namespace iReporterExcelAddInCommon;

public class UserConfig
{
	public DeviceType SelectedDeviceType = DeviceType.IOS;

	public readonly PairList ClusterColors = new PairList("clusterColor", "color");

	public readonly PairList NotClusterColors = new PairList("notClusterColor", "color");

	public readonly PairList CapAndKinds = new PairList("capAndKind")
	{
		TupleCount = 3
	};

	private CaptionPriority captionPriority = Global.CaptionPriority;

	private bool useAI = Global.UseAI;

	private static UserConfig _instance = null;

	private static UserConfig _default = null;

	private static readonly object _locker = new object();

	public Language Language { get; private set; }

	public CaptionPriority CaptionPriority
	{
		get
		{
			if (!IsRoot)
			{
				return captionPriority;
			}
			return Global.CaptionPriority;
		}
		set
		{
			if (IsRoot)
			{
				Global.CaptionPriority = value;
			}
			else
			{
				captionPriority = value;
			}
		}
	}

	public string APIRoot => "http://52.246.177.117/";

	public bool HintUp { get; set; } = true;

	public bool HintLeft { get; set; } = true;

	public bool HintRight { get; set; } = true;

	public bool HintDown { get; set; } = true;

	public bool UseAI
	{
		get
		{
			if (!IsRoot)
			{
				return useAI;
			}
			return Global.UseAI;
		}
		set
		{
			if (IsRoot)
			{
				Global.UseAI = value;
			}
			else
			{
				useAI = value;
			}
		}
	}

	public bool IsRoot { get; }

	public void SetLang(string lang)
	{
		Language language = Language;
		Language = Language.FromName(lang);
		if (language != Language)
		{
			CultureInfo cultureInfo = new CultureInfo(Language.Culture);
			try
			{
				Thread.CurrentThread.CurrentUICulture = cultureInfo;
				Thread.CurrentThread.CurrentCulture = cultureInfo;
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
	}

	private UserConfig(bool isRoot)
	{
		IsRoot = isRoot;
		Language = Language.FromCulture(CultureInfo.CurrentCulture.Name);
	}

	public static UserConfig GetInstance(ConfigKind kind = ConfigKind.Normal)
	{
		lock (_locker)
		{
			if (_default == null)
			{
				_default = GetInstance(GetUserConfigPath(), "default");
			}
			if (_instance == null)
			{
				_instance = GetInstance(GetUserConfigPath(), "");
			}
		}
		return kind switch
		{
			ConfigKind.Export => new UserConfig(isRoot: false).CopyFrom(_instance), 
			ConfigKind.Default => _default, 
			_ => _instance, 
		};
	}

	public static UserConfig GetInstance(string filePath, string child)
	{
		XmlNode xmlNode = null;
		if (File.Exists(filePath))
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.Load(filePath);
			xmlNode = xmlDocument.SelectSingleNode("/conmas/userConfig");
			if (child != "")
			{
				xmlNode = xmlNode?.SelectSingleNode(child);
			}
		}
		UserConfig userConfig = new UserConfig(child == "");
		if (xmlNode == null)
		{
			return userConfig;
		}
		if (xmlNode["deviceType"].InnerText.Equals("windows"))
		{
			userConfig.SelectedDeviceType = DeviceType.Windows;
		}
		userConfig.SetLang(xmlNode["language"].InnerText);
		userConfig.ClusterColors.Setup(xmlNode, _default?.ClusterColors);
		userConfig.NotClusterColors.Setup(xmlNode, _default?.NotClusterColors);
		userConfig.CapAndKinds.Setup(xmlNode, _default?.CapAndKinds);
		userConfig.CaptionPriority = (CaptionPriority)Text2Enum<CaptionPriority>(xmlNode["captionPriority"]?.InnerText);
		if (!bool.TryParse(xmlNode["useAI"]?.InnerText, out var result))
		{
			result = true;
		}
		userConfig.UseAI = result;
		userConfig.HintUp = xmlNode["hintUp"]?.InnerText != "0";
		userConfig.HintDown = xmlNode["HintDown"]?.InnerText != "0";
		userConfig.HintLeft = xmlNode["HintLeft"]?.InnerText != "0";
		userConfig.HintRight = xmlNode["HintRight"]?.InnerText != "0";
		return userConfig;
	}

	private static int Text2Enum<T>(string name)
	{
		return Math.Max(Array.IndexOf(Enum.GetNames(typeof(T)), name), 0);
	}

	private static int Elm2Int(XmlElement elm, int def)
	{
		if (!int.TryParse(elm?.InnerText, out var result))
		{
			return def;
		}
		return result;
	}

	public void SaveToFile(string filePath = null)
	{
		XmlDocument xmlDocument = new XmlDocument();
		XmlElement xmlElement = xmlDocument.CreateElement("conmas");
		XmlElement xmlElement2 = xmlDocument.CreateElement("userConfig");
		SaveToFile(xmlDocument, xmlElement2);
		if (filePath == null)
		{
			XmlElement xmlElement3 = xmlDocument.CreateElement("default");
			xmlElement2.AppendChild(xmlElement3);
			_default?.SaveToFile(xmlDocument, xmlElement3);
		}
		xmlElement.AppendChild(xmlElement2);
		xmlDocument.AppendChild(xmlElement);
		xmlDocument.Save(filePath ?? GetUserConfigPath());
	}

	private void SaveToFile(XmlDocument document, XmlElement userConfigElement)
	{
		XmlElement xmlElement = document.CreateElement("deviceType");
		XmlElement xmlElement2 = document.CreateElement("language");
		if (SelectedDeviceType == DeviceType.Windows)
		{
			xmlElement.InnerText = "windows";
		}
		else if (SelectedDeviceType == DeviceType.Android)
		{
			xmlElement.InnerText = "android";
		}
		else
		{
			xmlElement.InnerText = "ios";
		}
		xmlElement2.InnerText = Language.Name;
		userConfigElement.AppendChild(xmlElement);
		userConfigElement.AppendChild(xmlElement2);
		ClusterColors.WriteXML(userConfigElement);
		NotClusterColors.WriteXML(userConfigElement);
		CapAndKinds.WriteXML(userConfigElement);
		AddChild(userConfigElement, "captionPriority", CaptionPriority.ToString());
		AddChild(userConfigElement, "HintUp", HintUp ? "1" : "0");
		AddChild(userConfigElement, "HintDown", HintDown ? "1" : "0");
		AddChild(userConfigElement, "HintLeft", HintLeft ? "1" : "0");
		AddChild(userConfigElement, "HintRight", HintRight ? "1" : "0");
		AddChild(userConfigElement, "useAI", UseAI.ToString());
	}

	private void AddChild(XmlElement parent, string name, string value)
	{
		XmlElement xmlElement = parent.OwnerDocument.CreateElement(name);
		xmlElement.InnerText = value;
		parent.AppendChild(xmlElement);
	}

	private UserConfig CopyFrom(UserConfig src)
	{
		if (this == src)
		{
			return this;
		}
		ClusterColors.Setup(null, src.ClusterColors);
		NotClusterColors.Setup(null, src.NotClusterColors);
		CapAndKinds.Setup(null, src.CapAndKinds);
		CaptionPriority = src.CaptionPriority;
		UseAI = src.UseAI;
		return this;
	}

	private static string GetUserConfigPath()
	{
		string text = string.Format("{0}\\{1}\\{2}", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CIMTOPS", "ConMas_i-Reporter_ExcelAddIn");
		if (!Directory.Exists(text))
		{
			Directory.CreateDirectory(text);
		}
		return string.Format("{0}\\{1}", text, "UserSetting.xml");
	}
}
