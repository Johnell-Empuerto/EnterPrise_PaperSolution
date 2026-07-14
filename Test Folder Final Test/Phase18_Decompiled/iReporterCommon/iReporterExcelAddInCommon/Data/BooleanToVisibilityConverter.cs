using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace iReporterExcelAddInCommon.Data;

[ValueConversion(typeof(bool), typeof(Visibility))]
public class BooleanToVisibilityConverter : IValueConverter
{
	public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
	{
		Visibility visibility = Visibility.Collapsed;
		if (bool.Parse(value.ToString()))
		{
			visibility = Visibility.Visible;
		}
		return visibility;
	}

	public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
	{
		if ((Visibility)value == Visibility.Visible)
		{
			return true;
		}
		return false;
	}
}
