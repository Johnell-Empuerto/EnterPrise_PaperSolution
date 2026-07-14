using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using iReporterExcelAddInCommon.Domain.ValueDefinitions;

namespace iReporterExcelAddInCommon.Domain.Service;

public class CommonResourceService : INotifyPropertyChanged
{
	public static CommonResourceService Current { get; } = new CommonResourceService();

	public CommonResource Res { get; } = new CommonResource();

	public event PropertyChangedEventHandler PropertyChanged;

	protected void OnPropertyChanged([CallerMemberName] string name = null)
	{
		this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
	}

	public void ChangeCulture(string name)
	{
		CommonResource.Culture = CultureInfo.GetCultureInfo(name);
		OnPropertyChanged("Res");
	}

	public static void ChangeLanguage()
	{
	}
}
