using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using iReporterExcelAddInCommon.Domain.Models;
using iReporterExcelAddInCommon.Domain.Utilities;

namespace iReporterExcelAddInCommon.Common;

public class WorksheetChecker
{
	private Workbook workbook;

	public WorksheetChecker(Workbook workbook)
	{
		this.workbook = workbook;
	}

	public static bool HasCircularReference(Sheets worksheets)
	{
		if (worksheets == null)
		{
			return false;
		}
		if (worksheets.Count == 0)
		{
			return false;
		}
		bool result = false;
		for (int i = 1; i <= worksheets.Count; i++)
		{
			Worksheet worksheet = (dynamic)worksheets[i];
			Range range = null;
			try
			{
				if (worksheet != null)
				{
					range = worksheet.CircularReference;
					if (range != null)
					{
						result = true;
						break;
					}
				}
			}
			finally
			{
				if (range != null)
				{
					Marshal.FinalReleaseComObject(range);
					range = null;
				}
				if (worksheet != null)
				{
					Marshal.FinalReleaseComObject(worksheet);
					worksheet = null;
				}
			}
		}
		return result;
	}

	public static bool IsBlankName(string sheetName, bool treatFullWidthSpaceAsBlank = false)
	{
		if (string.IsNullOrEmpty(sheetName))
		{
			return true;
		}
		if (treatFullWidthSpaceAsBlank)
		{
			return string.IsNullOrWhiteSpace(sheetName);
		}
		return sheetName.Trim(' ').Length == 0;
	}

	public static bool ValidateSheetNamesNotBlank(IEnumerable<string> sheetNames, out List<int> invalidIndexes)
	{
		invalidIndexes = new List<int>();
		if (sheetNames == null)
		{
			return false;
		}
		List<string> list = sheetNames.ToList();
		for (int i = 0; i < list.Count; i++)
		{
			if (IsBlankName(list[i]))
			{
				invalidIndexes.Add(i + 1);
			}
		}
		return !invalidIndexes.Any();
	}

	public static List<string> ValidateSheetNames(IEnumerable<string> sheetNames, List<SheetNameInvalidCharacter> invalidCharacters, string messageTemplate)
	{
		HashSet<char> invalidCharSet = SheetNameValidator.BuildInvalidCharSet(invalidCharacters);
		List<string> list = new List<string>();
		foreach (string sheetName in sheetNames)
		{
			if (SheetNameValidator.IsValid(sheetName, invalidCharSet))
			{
				continue;
			}
			foreach (SheetNameInvalidCharacter invalidCharacter in SheetNameValidator.GetInvalidCharacters(sheetName, invalidCharacters))
			{
				list.Add(string.Format(messageTemplate, sheetName, invalidCharacter.Name, invalidCharacter.Symbol));
			}
		}
		return list;
	}

	public static string ReplaceSheetNameInFormula(string formula, string oldName, string newName)
	{
		string pattern = "'(?<quoted>(?:[^']|'')+)'!|(?<plain>[^\\s'!\\+\\-\\*/\\^\\&=\\(\\),:<>]+)!";
		return Regex.Replace(formula, pattern, delegate(Match match)
		{
			string obj = (match.Groups["quoted"].Success ? match.Groups["quoted"].Value : null);
			string text = (match.Groups["plain"].Success ? match.Groups["plain"].Value : null);
			string text2 = obj ?? text;
			return (((obj != null) ? text2.Replace("''", "'").Replace("’’", "’") : text2) == oldName) ? (EscapeSheetNameForFormula(newName) + "!") : match.Value;
		});
	}

	public static string EscapeSheetNameForFormula(string sheetName)
	{
		string text = sheetName.Replace("'", "''");
		text = text.Replace("’", "’’");
		return "'" + text + "'";
	}

	public static bool AreSheetNamesEqual(string name1, string name2)
	{
		if (name1 == null || name2 == null)
		{
			return false;
		}
		string a = NormalizeSheetName(name1);
		string b = NormalizeSheetName(name2);
		return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
	}

	private static string NormalizeSheetName(string input)
	{
		string text = Strings.StrConv(input.Normalize(NormalizationForm.FormC), VbStrConv.Narrow);
		StringBuilder stringBuilder = new StringBuilder(text.Length);
		string text2 = text;
		foreach (char c in text2)
		{
			if (c >= '①' && c <= '⑨')
			{
				stringBuilder.Append((char)(49 + (c - 9312)));
			}
			else if (c >= '⓵' && c <= '⓽')
			{
				stringBuilder.Append((char)(49 + (c - 9461)));
			}
			else if (c >= 'Ⓐ' && c <= 'Ⓩ')
			{
				stringBuilder.Append((char)(c - 9398 + 9424));
			}
			else
			{
				stringBuilder.Append(c);
			}
		}
		return stringBuilder.ToString();
	}

	public static List<SheetDiffResult> CompareSheetConfiguration(string xmlSetting, List<string> currentSheetNames)
	{
		ExcelOutputSetting excelOutputSetting = ExcelOutputSetting.FromXml(xmlSetting);
		if (excelOutputSetting == null || excelOutputSetting.Top == null)
		{
			return new List<SheetDiffResult>();
		}
		List<ExcelOutputSetting.OriginalSheetName> originalSheetNames = excelOutputSetting.Top.OriginalSheetNames;
		List<SheetDiffResult> list = new List<SheetDiffResult>();
		HashSet<ExcelOutputSetting.OriginalSheetName> usedOriginals = new HashSet<ExcelOutputSetting.OriginalSheetName>();
		for (int i = 0; i < currentSheetNames.Count; i++)
		{
			string currentName = currentSheetNames[i];
			int currentNo = i + 1;
			ExcelOutputSetting.OriginalSheetName originalSheetName = originalSheetNames.FirstOrDefault((ExcelOutputSetting.OriginalSheetName s) => s.SheetNo == currentNo && s.SheetName == currentName);
			if (originalSheetName != null)
			{
				usedOriginals.Add(originalSheetName);
				continue;
			}
			ExcelOutputSetting.OriginalSheetName originalSheetName2 = originalSheetNames.FirstOrDefault((ExcelOutputSetting.OriginalSheetName s) => s.SheetName == currentName);
			if (originalSheetName2 != null)
			{
				list.Add(new SheetDiffResult
				{
					Status = SheetStatus.Moved,
					Name = currentName,
					OldNo = originalSheetName2.SheetNo,
					NewNo = currentNo
				});
				usedOriginals.Add(originalSheetName2);
			}
			else
			{
				list.Add(new SheetDiffResult
				{
					Status = SheetStatus.Added,
					Name = currentName,
					NewNo = currentNo
				});
			}
		}
		foreach (ExcelOutputSetting.OriginalSheetName item in originalSheetNames.Where((ExcelOutputSetting.OriginalSheetName s) => !usedOriginals.Contains(s)))
		{
			list.Add(new SheetDiffResult
			{
				Status = SheetStatus.Deleted,
				Name = item.SheetName,
				OldNo = item.SheetNo
			});
		}
		return list;
	}
}
