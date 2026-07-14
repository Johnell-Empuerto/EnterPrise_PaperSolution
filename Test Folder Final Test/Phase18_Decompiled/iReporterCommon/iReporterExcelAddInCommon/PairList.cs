using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace iReporterExcelAddInCommon;

public class PairList : IEnumerable<CimTuple>, IEnumerable
{
	private readonly List<CimTuple> list = new List<CimTuple>();

	private string Tag1Name { get; }

	private string Tag2Name { get; }

	public int TupleCount { get; set; }

	internal PairList(string tag1Name, string tag2Name = null)
	{
		Tag1Name = tag1Name;
		Tag2Name = tag2Name ?? "item";
		TupleCount = 2;
	}

	internal void Setup(XmlNode root, PairList def)
	{
		Clear();
		XmlElement xmlElement = root?[Tag1Name];
		if (xmlElement != null)
		{
			foreach (XmlNode item in xmlElement.GetElementsByTagName(Tag2Name))
			{
				Add(item["before"]?.InnerText, item["after"]?.InnerText, item["third"]?.InnerText);
			}
			return;
		}
		if (def == null)
		{
			return;
		}
		foreach (CimTuple item2 in def.list)
		{
			Add(item2.Item1, item2.Item2, item2.Item3);
		}
	}

	internal void WriteXML(XmlElement parent)
	{
		XmlDocument ownerDocument = parent.OwnerDocument;
		XmlNode xmlNode = parent.AppendChild(ownerDocument.CreateElement(Tag1Name));
		foreach (CimTuple item in list)
		{
			XmlNode xmlNode2 = xmlNode.AppendChild(ownerDocument.CreateElement(Tag2Name));
			xmlNode2.AppendChild(ownerDocument.CreateElement("before")).InnerText = item.Item1;
			xmlNode2.AppendChild(ownerDocument.CreateElement("after")).InnerText = item.Item2;
			if (3 <= TupleCount)
			{
				xmlNode2.AppendChild(ownerDocument.CreateElement("third")).InnerText = item.Item3;
			}
		}
	}

	internal void Add(string key, string value, string third)
	{
		list.Add(CimTuple.Create(key, value, third));
	}

	internal void Clear()
	{
		list.Clear();
	}

	public IEnumerator<CimTuple> GetEnumerator()
	{
		return list.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return GetEnumerator();
	}
}
