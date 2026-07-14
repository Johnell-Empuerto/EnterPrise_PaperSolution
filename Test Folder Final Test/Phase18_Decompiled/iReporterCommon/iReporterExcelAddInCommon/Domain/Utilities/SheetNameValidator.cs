using System.Collections.Generic;
using System.Linq;
using iReporterExcelAddInCommon.Domain.Models;

namespace iReporterExcelAddInCommon.Domain.Utilities;

public static class SheetNameValidator
{
	public static bool IsValid(string sheetName, HashSet<char> invalidCharSet)
	{
		if (string.IsNullOrEmpty(sheetName))
		{
			return false;
		}
		return !sheetName.Any((char c) => invalidCharSet.Contains(c));
	}

	public static List<SheetNameInvalidCharacter> GetInvalidCharacters(string sheetName, List<SheetNameInvalidCharacter> invalidCharacters)
	{
		return invalidCharacters.Where((SheetNameInvalidCharacter c) => !string.IsNullOrEmpty(c.Symbol) && sheetName.Contains(c.Symbol)).ToList();
	}

	public static HashSet<char> BuildInvalidCharSet(List<SheetNameInvalidCharacter> invalidCharacters)
	{
		return new HashSet<char>(invalidCharacters.Where((SheetNameInvalidCharacter c) => !string.IsNullOrEmpty(c.Symbol)).SelectMany((SheetNameInvalidCharacter c) => c.Symbol));
	}
}
