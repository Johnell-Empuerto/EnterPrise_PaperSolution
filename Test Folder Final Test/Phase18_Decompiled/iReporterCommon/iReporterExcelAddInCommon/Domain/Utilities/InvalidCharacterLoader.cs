using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using iReporterExcelAddInCommon.Domain.Models;

namespace iReporterExcelAddInCommon.Domain.Utilities;

public static class InvalidCharacterLoader
{
	private static string _currentLanguage;

	private static readonly Dictionary<string, List<SheetNameInvalidCharacter>> _cache = new Dictionary<string, List<SheetNameInvalidCharacter>>();

	private const string SheetNameInvalidCharacterResource = "iReporterExcelAddInCommon.Domain.Resources.SheetNameInvalidCharacters";

	private static readonly Dictionary<string, string> ResourceMap = new Dictionary<string, string>
	{
		{ "ja", "iReporterExcelAddInCommon.Domain.Resources.SheetNameInvalidCharacters_ja.xml" },
		{ "en", "iReporterExcelAddInCommon.Domain.Resources.SheetNameInvalidCharacters.xml" },
		{ "zh-Hans", "iReporterExcelAddInCommon.Domain.Resources.SheetNameInvalidCharacters_zh-CN.xml" },
		{ "zh-Hant", "iReporterExcelAddInCommon.Domain.Resources.SheetNameInvalidCharacters_zh-TW.xml" }
	};

	public static List<SheetNameInvalidCharacter> LoadFromStream(Stream stream)
	{
		return ParseDocument(XDocument.Load(stream));
	}

	private static List<SheetNameInvalidCharacter> ParseDocument(XDocument doc)
	{
		return (from c in doc.Descendants("Character")
			select new SheetNameInvalidCharacter
			{
				Symbol = c.Element("Symbol")?.Value,
				Name = c.Element("Name")?.Value,
				Method = c.Element("Method")?.Value
			} into c
			where !string.IsNullOrEmpty(c.Symbol)
			select c).ToList();
	}

	public static List<SheetNameInvalidCharacter> Load(string lang)
	{
		if (_currentLanguage == lang && _cache.TryGetValue(lang, out var value))
		{
			return value;
		}
		List<SheetNameInvalidCharacter> list = LoadFromEmbeddedXml(GetLocalizedResourceName(lang));
		_cache[lang] = list;
		_currentLanguage = lang;
		return list;
	}

	public static List<SheetNameInvalidCharacter> LoadFromEmbeddedXml(string resourceName)
	{
		using Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);
		if (stream == null)
		{
			throw new FileNotFoundException("Resource not found: " + resourceName);
		}
		return LoadFromStream(stream);
	}

	private static string GetLocalizedResourceName(string lang)
	{
		if (!ResourceMap.TryGetValue(lang, out var value))
		{
			return "iReporterExcelAddInCommon.Domain.Resources.SheetNameInvalidCharacters.xml";
		}
		return value;
	}
}
