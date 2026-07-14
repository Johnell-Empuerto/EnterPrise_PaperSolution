using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Dynamic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using iReporterExcelAddInCommon.Domain.Helpers;
using iReporterExcelAddInCommon.Domain.Service;
using iReporterExcelAddInCommon.Domain.ValueDefinitions;
using iReporterExcelAddInCommon.Services;
using iReporterExcelAddInCommon.Views;

namespace iReporterExcelAddInCommon.ViewModels;

public class MessageGridViewModel : ViewModelBase
{
	private DisplayMode _mode = DisplayMode.Both;

	private bool _isMessageSectionVisible = true;

	private bool _isGridSectionVisible = true;

	private bool _isOkButtonEnabled = true;

	private bool _isCopyButtonVisible = true;

	private bool _isCloseButtonVisible = true;

	private readonly Action _onOkAction;

	private string _windowTitle;

	private string _okButtonText;

	private string _closeButtonText;

	private string _messageGridTextDesc;

	private string _messageGridGridDesc;

	private string _message;

	private readonly INotificationService _notificationService;

	public DisplayMode Mode
	{
		get
		{
			return _mode;
		}
		set
		{
			if (_mode != value)
			{
				_mode = value;
				UpdateVisibility();
				OnPropertyChanged("Mode");
			}
		}
	}

	public bool IsMessageSectionVisible
	{
		get
		{
			return _isMessageSectionVisible;
		}
		set
		{
			if (_isMessageSectionVisible != value)
			{
				_isMessageSectionVisible = value;
				OnPropertyChanged("IsMessageSectionVisible");
			}
		}
	}

	public bool IsGridSectionVisible
	{
		get
		{
			return _isGridSectionVisible;
		}
		set
		{
			if (_isGridSectionVisible != value)
			{
				_isGridSectionVisible = value;
				OnPropertyChanged("IsGridSectionVisible");
			}
		}
	}

	public bool IsOkButtonEnabled
	{
		get
		{
			return _isOkButtonEnabled;
		}
		set
		{
			if (SetValue(ref _isOkButtonEnabled, value, "IsOkButtonEnabled"))
			{
				(OKCommand as DelegateCommand)?.RaiseCanExecuteChanged();
			}
		}
	}

	public bool IsCopyButtonVisible
	{
		get
		{
			return _isCopyButtonVisible;
		}
		private set
		{
			SetValue(ref _isCopyButtonVisible, value, "IsCopyButtonVisible");
		}
	}

	public bool IsCloseButtonVisible
	{
		get
		{
			return _isCloseButtonVisible;
		}
		private set
		{
			SetValue(ref _isCloseButtonVisible, value, "IsCloseButtonVisible");
		}
	}

	public string WindowTitle
	{
		get
		{
			return _windowTitle;
		}
		set
		{
			SetValue(ref _windowTitle, value, "WindowTitle");
		}
	}

	public string OkButtonText
	{
		get
		{
			return _okButtonText;
		}
		set
		{
			SetValue(ref _okButtonText, value, "OkButtonText");
		}
	}

	public string CloseButtonText
	{
		get
		{
			return _closeButtonText;
		}
		set
		{
			SetValue(ref _closeButtonText, value, "CloseButtonText");
		}
	}

	public string MessageGridTextDesc
	{
		get
		{
			return _messageGridTextDesc;
		}
		set
		{
			_messageGridTextDesc = value;
			OnPropertyChanged("MessageGridTextDesc");
		}
	}

	public string MessageGridGridDesc
	{
		get
		{
			return _messageGridGridDesc;
		}
		set
		{
			_messageGridGridDesc = value;
			OnPropertyChanged("MessageGridGridDesc");
		}
	}

	public string Message
	{
		get
		{
			return _message;
		}
		set
		{
			SetValue(ref _message, value, "Message");
		}
	}

	public ObservableCollection<ColumnDefinitionModel> ColumnDefinitions { get; }

	public ObservableCollection<ExpandoObject> GridRows { get; }

	public ICommand OKCommand { get; }

	public ICommand CloseCommand { get; }

	public ICommand CopyMessageCommand { get; }

	public ICommand CopyGridCommand { get; }

	public Action<MessageGridViewModel, object> OnCellValueChanged { get; set; }

	public bool WindowResult { get; private set; }

	public Func<MessageGridViewModel, bool> ErrorCheckPredicate { get; set; }

	private void UpdateVisibility()
	{
		IsMessageSectionVisible = Mode == DisplayMode.MessageOnly || Mode == DisplayMode.Both;
		IsGridSectionVisible = Mode == DisplayMode.GridOnly || Mode == DisplayMode.Both;
	}

	public MessageGridViewModel(string messageTextDesc, string messageGridDesc, string message, IEnumerable<ColumnDefinitionModel> columns, IEnumerable<IEnumerable<string>> rowData, DisplayMode displayMode = DisplayMode.Both, INotificationService notificationService = null, Action onOkAction = null, bool isCopyButtonVisible = true, bool isCloseButtonVisible = true, string windowTitle = null, string okButtonText = null, string closeButtonText = null)
	{
		MessageGridTextDesc = messageTextDesc;
		MessageGridGridDesc = messageGridDesc;
		Message = message;
		ColumnDefinitions = new ObservableCollection<ColumnDefinitionModel>(columns);
		GridRows = new ObservableCollection<ExpandoObject>();
		Mode = displayMode;
		_notificationService = notificationService;
		_onOkAction = onOkAction;
		IsCopyButtonVisible = isCopyButtonVisible;
		IsCloseButtonVisible = isCloseButtonVisible;
		WindowTitle = windowTitle ?? ViewTexts.MessageGrid;
		OkButtonText = okButtonText ?? ViewTexts.OK;
		CloseButtonText = closeButtonText ?? ViewTexts.Cancel;
		foreach (IEnumerable<string> rowDatum in rowData)
		{
			dynamic val = new ExpandoObject();
			IDictionary<string, object> dictionary = (IDictionary<string, object>)val;
			dictionary["BackgroundColor"] = Brushes.White;
			int num = 0;
			foreach (string item in rowDatum)
			{
				dictionary[columns.ElementAt(num++).PropertyName] = item;
			}
			((INotifyPropertyChanged)val).PropertyChanged += delegate(object s, PropertyChangedEventArgs e)
			{
				if (e.PropertyName != "BackgroundColor")
				{
					OnCellValueChanged?.Invoke(this, s);
				}
			};
			GridRows.Add(val);
		}
		OKCommand = new DelegateCommand
		{
			ExecuteHandler = delegate(object args)
			{
				_onOkAction?.Invoke();
				WindowResult = true;
				if (args is Window window)
				{
					window.DialogResult = true;
				}
				ViewService.CloseWindow(args);
			},
			CanExecuteHandler = (object _) => IsOkButtonEnabled
		};
		CloseCommand = new DelegateCommand
		{
			ExecuteHandler = delegate(object args)
			{
				WindowResult = false;
				ViewService.CloseWindow(args);
			}
		};
		CopyMessageCommand = new DelegateCommand
		{
			ExecuteHandler = delegate
			{
				Clipboard.SetText(Message);
				_notificationService?.ShowMessage(CommonResource.MessageGridCopied);
			}
		};
		CopyGridCommand = new DelegateCommand
		{
			ExecuteHandler = delegate
			{
				string[] excludeKeywords = new string[3] { "Color", "Background", "IsSelected" };
				Clipboard.SetText(string.Join(Environment.NewLine, GridRows.Select(delegate(ExpandoObject row)
				{
					IEnumerable<string> values = from kvp in row
						where !excludeKeywords.Any((string k) => kvp.Key.Contains(k))
						select kvp.Value?.ToString() ?? "";
					return string.Join("\t", values);
				})));
				_notificationService?.ShowMessage(CommonResource.MessageGridCopied);
			}
		};
	}

	public void SetRowBackgroundColor(int rowIndex, Brush brush)
	{
		if (rowIndex >= 0 && rowIndex < GridRows.Count)
		{
			IDictionary<string, object> dictionary = GridRows[rowIndex];
			if (!dictionary.ContainsKey("BackgroundColor") || dictionary["BackgroundColor"] != brush)
			{
				dictionary["BackgroundColor"] = brush;
			}
		}
	}

	public void RefreshOkButtonState()
	{
		if (ErrorCheckPredicate != null)
		{
			IsOkButtonEnabled = ErrorCheckPredicate(this);
		}
		else
		{
			IsOkButtonEnabled = true;
		}
	}
}
