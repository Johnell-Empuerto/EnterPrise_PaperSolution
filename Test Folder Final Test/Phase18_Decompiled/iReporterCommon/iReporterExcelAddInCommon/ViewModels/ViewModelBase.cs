using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace iReporterExcelAddInCommon.ViewModels;

public abstract class ViewModelBase : INotifyPropertyChanged
{
	public event PropertyChangedEventHandler PropertyChanged;

	protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
	{
		this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
	}

	protected bool SetValue<T>(ref T field, T newValue, [CallerMemberName] string propertyName = null)
	{
		if (EqualityComparer<T>.Default.Equals(field, newValue))
		{
			return false;
		}
		field = newValue;
		OnPropertyChanged(propertyName);
		return true;
	}

	protected void UpdatePropertySilent<T>(ref T field, T value, string propertyName)
	{
		field = value;
		OnPropertyChanged(propertyName);
	}

	protected void ForceRaisePropertyChanged(string propertyName)
	{
		this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
	}
}
