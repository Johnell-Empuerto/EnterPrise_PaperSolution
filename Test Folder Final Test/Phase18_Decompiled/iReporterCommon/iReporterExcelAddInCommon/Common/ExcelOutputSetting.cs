using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace iReporterExcelAddInCommon.Common;

[XmlRoot("conmas")]
public class ExcelOutputSetting
{
	public class TopSection
	{
		[XmlArray("originalSheetNames")]
		[XmlArrayItem("originalSheetName")]
		public List<OriginalSheetName> OriginalSheetNames { get; set; } = new List<OriginalSheetName>();
	}

	public class OriginalSheetName
	{
		[XmlElement("sheetNo")]
		public int SheetNo { get; set; }

		[XmlElement("sheetName")]
		public string SheetName { get; set; }
	}

	[XmlElement("top")]
	public TopSection Top { get; set; }

	public static ExcelOutputSetting FromXml(string xml)
	{
		if (string.IsNullOrEmpty(xml))
		{
			return null;
		}
		XmlSerializer xmlSerializer = new XmlSerializer(typeof(ExcelOutputSetting));
		using StringReader textReader = new StringReader(xml);
		return (ExcelOutputSetting)xmlSerializer.Deserialize(textReader);
	}
}
