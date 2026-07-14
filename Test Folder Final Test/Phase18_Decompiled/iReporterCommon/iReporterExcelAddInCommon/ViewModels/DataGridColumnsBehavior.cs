using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace iReporterExcelAddInCommon.ViewModels;

public static class DataGridColumnsBehavior
{
	public static readonly DependencyProperty BindableColumnsProperty = DependencyProperty.RegisterAttached("BindableColumns", typeof(ObservableCollection<ColumnDefinitionModel>), typeof(DataGridColumnsBehavior), new PropertyMetadata(null, OnBindableColumnsChanged));

	public static ObservableCollection<ColumnDefinitionModel> GetBindableColumns(DependencyObject element)
	{
		return (ObservableCollection<ColumnDefinitionModel>)element.GetValue(BindableColumnsProperty);
	}

	public static void SetBindableColumns(DependencyObject element, ObservableCollection<ColumnDefinitionModel> value)
	{
		element.SetValue(BindableColumnsProperty, value);
	}

	private static void OnBindableColumnsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
	{
		if (!(d is DataGrid dataGrid) || !(e.NewValue is ObservableCollection<ColumnDefinitionModel> observableCollection))
		{
			return;
		}
		dataGrid.Columns.Clear();
		foreach (ColumnDefinitionModel item in observableCollection)
		{
			DataGridColumn dataGridColumn;
			switch (item.ColumnType)
			{
			case ColumnType.CheckBox:
				dataGridColumn = new DataGridCheckBoxColumn
				{
					Header = item.Header,
					Binding = new Binding(item.PropertyName)
				};
				break;
			case ColumnType.ComboBox:
				dataGridColumn = new DataGridComboBoxColumn
				{
					Header = item.Header,
					SelectedValueBinding = new Binding(item.PropertyName)
				};
				break;
			default:
			{
				DataGridTextColumn dataGridTextColumn = new DataGridTextColumn
				{
					Header = item.Header,
					Binding = new Binding(item.PropertyName)
					{
						Mode = (item.IsReadOnly ? BindingMode.OneWay : BindingMode.TwoWay),
						UpdateSourceTrigger = UpdateSourceTrigger.Default
					}
				};
				if (item.MaxLength.HasValue && !item.IsReadOnly)
				{
					Style style = new Style(typeof(TextBox));
					style.Setters.Add(new Setter(TextBox.MaxLengthProperty, item.MaxLength.Value));
					dataGridTextColumn.EditingElementStyle = style;
				}
				Style style2 = new Style(typeof(TextBlock));
				if (item.IsTextWrapping)
				{
					style2.Setters.Add(new Setter(TextBlock.TextWrappingProperty, TextWrapping.Wrap));
				}
				else
				{
					style2.Setters.Add(new Setter(TextBlock.TextWrappingProperty, TextWrapping.NoWrap));
				}
				style2.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, (item.Alignment == HorizontalAlignment.Center) ? TextAlignment.Center : ((item.Alignment == HorizontalAlignment.Right) ? TextAlignment.Right : TextAlignment.Left)));
				style2.Setters.Add(new Setter(FrameworkElement.HorizontalAlignmentProperty, item.Alignment));
				dataGridTextColumn.ElementStyle = style2;
				dataGridColumn = dataGridTextColumn;
				break;
			}
			}
			if (item.Width.HasValue)
			{
				dataGridColumn.Width = item.Width.Value;
			}
			if (item.MinWidth.HasValue)
			{
				dataGridColumn.MinWidth = item.MinWidth.Value;
			}
			dataGridColumn.IsReadOnly = item.IsReadOnly;
			dataGridColumn.Visibility = ((!item.IsVisible) ? Visibility.Collapsed : Visibility.Visible);
			dataGridColumn.CanUserSort = item.IsSortable;
			dataGrid.Columns.Add(dataGridColumn);
		}
	}
}
