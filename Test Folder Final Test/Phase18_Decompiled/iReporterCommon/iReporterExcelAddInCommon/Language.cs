using System;
using System.Collections.Generic;
using System.Linq;

namespace iReporterExcelAddInCommon;

public class Language
{
	public string Name { get; }

	public string Culture { get; }

	public bool CanUseAI => true;

	public int Index => GetAll().ToList().IndexOf(this);

	private static Language Ja { get; } = new Language("Japanese", "ja");

	private static Language En { get; } = new Language("English", "en");

	private static Language Ch { get; } = new Language("Chinese(Simplified)", "zh-Hans");

	private static Language ChT { get; } = new Language("Chinese(Traditional)", "zh-Hant");

	private Language(string name, string culture)
	{
		Name = name;
		Culture = culture;
	}

	public static IEnumerable<Language> GetAll()
	{
		yield return En;
		yield return Ja;
		yield return Ch;
		yield return ChT;
	}

	public static Language FromCulture(string calName)
	{
		Language language = GetAll().FirstOrDefault((Language la) => calName.Equals(la.Name, StringComparison.CurrentCultureIgnoreCase));
		if (language != null)
		{
			return language;
		}
		if (calName.Contains('-'))
		{
			calName = calName.Substring(0, calName.IndexOf('-'));
		}
		return GetAll().FirstOrDefault((Language la) => string.Compare(calName, la.Culture, ignoreCase: true) == 0) ?? GetAll().First();
	}

	public static Language FromName(string name)
	{
		if (name == "chinese")
		{
			return Ch;
		}
		return GetAll().FirstOrDefault((Language la) => string.Compare(la.Name, name, ignoreCase: true) == 0) ?? GetAll().First();
	}
}
