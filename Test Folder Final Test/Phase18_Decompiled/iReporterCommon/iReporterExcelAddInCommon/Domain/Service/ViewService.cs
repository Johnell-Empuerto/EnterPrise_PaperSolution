using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using iReporterExcelAddInCommon.Domain.Helpers;

namespace iReporterExcelAddInCommon.Domain.Service;

public class ViewService
{
	public static ICommand CloseCommand;

	public static ViewService Instance { get; private set; }

	private Dictionary<string, Type> _windowTypes { get; set; } = new Dictionary<string, Type>();

	static ViewService()
	{
		CloseCommand = new DelegateCommand
		{
			ExecuteHandler = CloseWindow
		};
		Instance = new ViewService();
	}

	private ViewService()
	{
	}

	public void AddType(Type viewType)
	{
		if (!_windowTypes.ContainsKey(viewType.Name))
		{
			_windowTypes.Add(viewType.Name, viewType);
		}
	}

	public bool? ShowDialog(string typeName, object context = null, Window owner = null)
	{
		if (!_windowTypes.ContainsKey(typeName))
		{
			return false;
		}
		if (!(_windowTypes[typeName].GetConstructor(new Type[0]).Invoke(new object[0]) is Window window))
		{
			return false;
		}
		if (context != null)
		{
			window.DataContext = context;
		}
		if (owner != null)
		{
			window.Owner = owner;
			window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
		}
		return window.ShowDialog();
	}

	public void Show(string typeName, object context = null)
	{
		if (_windowTypes.ContainsKey(typeName) && _windowTypes[typeName].GetConstructor(new Type[0]).Invoke(new object[0]) is Window window)
		{
			if (context != null)
			{
				window.DataContext = context;
			}
			window.Show();
		}
	}

	public static void CloseWindow(object args)
	{
		if (args is Window window)
		{
			window.Close();
		}
	}

	public static void OpenWindow(string windowTypeKey, IWindowViewModel viewModel, Action actionWhenOk, object owner = null)
	{
		if (Instance != null)
		{
			Instance.ShowDialog(windowTypeKey, viewModel, owner as Window);
			if (viewModel != null && viewModel.WindowResult)
			{
				actionWhenOk();
			}
		}
	}

	public static IEnumerable<DependencyObject> GetAncestors(DependencyObject obj)
	{
		while (obj != null)
		{
			obj = VisualTreeHelper.GetParent(obj);
			yield return obj;
		}
	}

	public static T GetAncestorByType<T>(DependencyObject obj) where T : DependencyObject
	{
		return GetAncestors(obj).FirstOrDefault((DependencyObject a) => a is T) as T;
	}
}
