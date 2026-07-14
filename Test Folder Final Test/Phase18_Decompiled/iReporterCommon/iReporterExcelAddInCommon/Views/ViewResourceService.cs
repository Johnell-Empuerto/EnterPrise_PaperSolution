using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace iReporterExcelAddInCommon.Views;

public class ViewResourceService : INotifyPropertyChanged
{
	public static ViewResourceService Current { get; } = new ViewResourceService();

	public ViewTexts Texts { get; } = new ViewTexts();

	public event PropertyChangedEventHandler PropertyChanged;

	protected void OnPropertyChanged([CallerMemberName] string name = null)
	{
		this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
	}
}
