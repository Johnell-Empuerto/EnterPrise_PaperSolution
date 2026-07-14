using System;
using System.Windows.Input;

namespace iReporterExcelAddInCommon.Domain.Helpers;

public class DelegateCommand : ICommand
{
	public static DelegateCommand InvalidCommand = new DelegateCommand
	{
		CanExecuteHandler = (object obj) => false
	};

	public Action<object> ExecuteHandler { get; set; }

	public Func<object, bool> CanExecuteHandler { get; set; }

	public event EventHandler CanExecuteChanged;

	public bool CanExecute(object parameter)
	{
		if (CanExecuteHandler == null)
		{
			return true;
		}
		return CanExecuteHandler(parameter);
	}

	public void Execute(object parameter)
	{
		ExecuteHandler?.Invoke(parameter);
	}

	public void RaiseCanExecuteChanged()
	{
		this.CanExecuteChanged?.Invoke(this, null);
	}
}
