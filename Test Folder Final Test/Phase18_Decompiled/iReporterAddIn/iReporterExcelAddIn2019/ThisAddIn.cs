using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using Cimtops.Excel;
using Cimtops.R2Cluster;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using iReporterExcelAddIn2019.Properties;
using iReporterExcelAddInCommon;
using iReporterExcelAddInCommon.Common;
using iReporterExcelAddInCommon.Domain.Helpers;
using iReporterExcelAddInCommon.Domain.Models;
using iReporterExcelAddInCommon.Domain.Utilities;
using iReporterExcelAddInCommon.Properties;
using iReporterExcelAddInCommon.Services;
using iReporterExcelAddInCommon.Tables;
using iReporterExcelAddInCommon.ViewModels;
using iReporterExcelAddInCommon.Views;

namespace iReporterExcelAddIn2019;

[StartupObject(0)]
[PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
public sealed class ThisAddIn : AddInBase
{
	private class Win32Window : System.Windows.Forms.IWin32Window
	{
		public IntPtr _handle;

		public IntPtr Handle => _handle;
	}

	private class Judged
	{
		public Exception ex;

		public List<string> msgs = new List<string>();
	}

	private const uint GA_ROOT = 2u;

	private Microsoft.Office.Interop.Excel.Application excelApplication;

	private List<CommandBarButton> MenuTexts = new List<CommandBarButton>();

	private InputForm2 _input2;

	private TableForm tableForm;

	private object _inputLocker = new object();

	private bool _selectionChangedSet;

	private ColorForm _color;

	private object judgeLock = new object();

	internal CustomTaskPaneCollection CustomTaskPanes;

	internal SmartTagCollection VstoSmartTags;

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	private object missing = Type.Missing;

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	internal Microsoft.Office.Interop.Excel.Application Application;

	private bool IsJudging { get; set; }

	[DllImport("user32.dll")]
	private static extern int GetWindowRect(IntPtr hWnd, out RECT lpRect);

	[DllImport("user32.dll")]
	[return: MarshalAs(UnmanagedType.Bool)]
	private static extern bool SetForegroundWindow(IntPtr hWnd);

	private System.Windows.Forms.IWin32Window CreateWinrowHandle()
	{
		return new Win32Window
		{
			_handle = new IntPtr(Application.Hwnd)
		};
	}

	[DllImport("user32.dll")]
	private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

	private void SetUpMenu()
	{
		int num = 10;
		foreach (CommandBar commandBar in Application.CommandBars)
		{
			if (commandBar.Name.Equals("Cell"))
			{
				CommandBarControls controls = commandBar.Controls;
				CommandBarButton commandBarButton = (CommandBarButton)controls.Add(MsoControlType.msoControlButton, missing, missing, missing, true);
				commandBarButton.Style = MsoButtonStyle.msoButtonCaption;
				commandBarButton.Tag = num++.ToString();
				new ComAwareEventInfo(typeof(_CommandBarButtonEvents_Event), "Click").AddEventHandler(commandBarButton, (_CommandBarButtonEvents_ClickEventHandler)delegate
				{
					AutoJudge(openDlg: true);
				});
				MenuTexts.Add(commandBarButton);
				commandBarButton = (CommandBarButton)controls.Add(MsoControlType.msoControlButton, missing, missing, missing, true);
				commandBarButton.Style = MsoButtonStyle.msoButtonCaption;
				commandBarButton.Tag = num++.ToString();
				new ComAwareEventInfo(typeof(_CommandBarButtonEvents_Event), "Click").AddEventHandler(commandBarButton, (_CommandBarButtonEvents_ClickEventHandler)delegate
				{
					ManualCluster();
				});
				MenuTexts.Add(commandBarButton);
				commandBarButton = (CommandBarButton)controls.Add(MsoControlType.msoControlButton, missing, missing, missing, true);
				commandBarButton.Style = MsoButtonStyle.msoButtonCaption;
				commandBarButton.Tag = num++.ToString();
				new ComAwareEventInfo(typeof(_CommandBarButtonEvents_Event), "Click").AddEventHandler(commandBarButton, (_CommandBarButtonEvents_ClickEventHandler)delegate
				{
					SetClusterColors();
				});
				MenuTexts.Add(commandBarButton);
				commandBarButton = (CommandBarButton)controls.Add(MsoControlType.msoControlButton, missing, missing, missing, true);
				commandBarButton.Style = MsoButtonStyle.msoButtonCaption;
				commandBarButton.Tag = num++.ToString();
				new ComAwareEventInfo(typeof(_CommandBarButtonEvents_Event), "Click").AddEventHandler(commandBarButton, (_CommandBarButtonEvents_ClickEventHandler)delegate
				{
					ShowTableForm();
				});
				MenuTexts.Add(commandBarButton);
			}
		}
		UserConfig.GetInstance();
		SetupMenuCaptions();
	}

	public void SetupMenuCaptions()
	{
		Globals.Ribbons.Ribbon1.autoJudgeButton.Label = iReporterExcelAddIn2019.Properties.Resources.MENU_EXTRACTION;
		Globals.Ribbons.Ribbon1.clusterSettingButton.Label = iReporterExcelAddIn2019.Properties.Resources.MENU_KIND;
		Globals.Ribbons.Ribbon1.colorSettingButton.Label = iReporterExcelAddIn2019.Properties.Resources.MENU_SETTING;
		Globals.Ribbons.Ribbon1.tableSettingButton.Label = iReporterExcelAddIn2019.Properties.Resources.MENU_TABLE;
		Globals.Ribbons.Ribbon1.checkSheetNameButton.Label = iReporterExcelAddIn2019.Properties.Resources.CHECK_SHEET_NAME;
		foreach (CommandBarButton item in MenuTexts.ToList())
		{
			try
			{
				if (int.TryParse(item.Tag, out var result))
				{
					switch (result)
					{
					case 10:
					case 15:
						item.Caption = Globals.Ribbons.Ribbon1.autoJudgeButton.Label.Replace("\n", "");
						break;
					case 11:
					case 16:
						item.Caption = Globals.Ribbons.Ribbon1.clusterSettingButton.Label.Replace("\n", "");
						break;
					case 12:
					case 17:
						item.Caption = Globals.Ribbons.Ribbon1.colorSettingButton.Label.Replace("\n", "");
						break;
					case 13:
					case 18:
						item.Caption = Globals.Ribbons.Ribbon1.tableSettingButton.Label.Replace("\n", "");
						break;
					case 14:
					case 19:
						item.Caption = Globals.Ribbons.Ribbon1.checkSheetNameButton.Label.Replace("\n", "");
						break;
					}
				}
			}
			catch (COMException)
			{
			}
		}
	}

	private void ThisAddIn_Startup(object sender, EventArgs e)
	{
		excelApplication = Globals.ThisAddIn.Application;
		SetUpMenu();
	}

	private void ThisAddIn_Shutdown(object sender, EventArgs e)
	{
		MenuTexts.Clear();
	}

	public void ManualCluster()
	{
		lock (_inputLocker)
		{
			System.Windows.Forms.IWin32Window window = CreateWinrowHandle();
			System.Drawing.Point? location = _input2?.DesktopLocation;
			Microsoft.Office.Interop.Excel.Worksheet obj = Application.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
			_input2?.Close();
			obj?.Activate();
			_input2 = new InputForm2();
			_input2.AllClusters = InitClusterList();
			_input2.WriteCommentEvent += WriteComment;
			_input2.ClearCommentEvent += ClearComments;
			_input2.CancelAllEvent += CancelAll;
			_input2.LanguageChangeEvent += ManualCluster;
			_input2.LanguageChangeEvent += SetupMenuCaptions;
			_input2.SelectCellEvent += SelectCell;
			_input2.ExtractEvent += delegate
			{
				AutoJudge(openDlg: false);
			};
			_input2.GetCellTextEvent += GetCellText;
			_input2.SetSheetNameEvent += SetSheetName;
			ShowForm(_input2, location, window);
		}
	}

	public void AutoJudge(bool openDlg)
	{
		lock (judgeLock)
		{
			if (IsJudging)
			{
				return;
			}
			IsJudging = true;
		}
		CancellationTokenSource cancellationTokenSource;
		WaitDialog wait;
		try
		{
			cancellationTokenSource = new CancellationTokenSource();
			wait = new WaitDialog(cancellationTokenSource);
		}
		catch
		{
			IsJudging = false;
			throw;
		}
		AutoJudgeAsync(cancellationTokenSource.Token, wait, openDlg);
	}

	private async Task AutoJudgeAsync(CancellationToken token, Form wait, bool openDlg)
	{
		try
		{
			using (wait)
			{
				wait.Show(CreateWinrowHandle());
				Judged judged = new Judged();
				Microsoft.Office.Interop.Excel.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
				Book book = BookFactory.Create(activeWorkbook, ref judged.ex);
				if (book != null)
				{
					UserConfig instance = UserConfig.GetInstance();
					Tuple<Color, Color>[] clusterColors = instance.ClusterColors.Select((CimTuple t) => Tuple.Create(t.Item1.ToColor(), t.Item2.ToColor())).ToArray();
					Tuple<Color, Color>[] notClusterColors = instance.NotClusterColors.Select((CimTuple t) => Tuple.Create(t.Item1.ToColor(), t.Item2.ToColor())).ToArray();
					Tuple<string, string, string>[] clusterCaptions = instance.CapAndKinds.Select((CimTuple t) => t.ToTuple()).ToArray();
					Decoder.CreateColorDic(clusterColors, notClusterColors, clusterCaptions);
					Decoder.AutoJudgeResult autoJudgeResult = Decoder.DoAutomaticJudgement(book, token, instance.UseAI && instance.Language.CanUseAI, instance.SelectedDeviceType);
					if (token.IsCancellationRequested)
					{
						return;
					}
					if (autoJudgeResult.Empties != null)
					{
						judged.msgs.Add(iReporterExcelAddIn2019.Properties.Resources.EMPTY_SHEET + Environment.NewLine + iReporterExcelAddIn2019.Properties.Resources.SHEET_NAME + autoJudgeResult.Empties);
					}
					if (autoJudgeResult.Multis != null)
					{
						judged.msgs.Add(iReporterExcelAddIn2019.Properties.Resources.MULTI_PRINT + Environment.NewLine + iReporterExcelAddIn2019.Properties.Resources.SHEET_NAME + autoJudgeResult.Multis);
					}
					if (autoJudgeResult.IsAPIError)
					{
						judged.msgs.Add(iReporterExcelAddIn2019.Properties.Resources.API_ERROR);
					}
					try
					{
						foreach (Microsoft.Office.Interop.Excel.Window item in Application.Windows?.OfType<Microsoft.Office.Interop.Excel.Window>() ?? Enumerable.Empty<Microsoft.Office.Interop.Excel.Window>())
						{
							Sheets selectedSheets = item.SelectedSheets;
							if (selectedSheets != null && selectedSheets.Count > 1 && selectedSheets[1] is Microsoft.Office.Interop.Excel.Worksheet worksheet)
							{
								worksheet.Select(Type.Missing);
							}
						}
						foreach (object sheet in activeWorkbook.Sheets)
						{
							if (!(sheet is Microsoft.Office.Interop.Excel.Worksheet worksheet2) || !autoJudgeResult.Modify.TryGetValue(worksheet2.Name, out var value))
							{
								continue;
							}
							foreach (KeyValuePair<System.Drawing.Point, CellModify> item2 in value)
							{
								Range range = (dynamic)worksheet2.Cells[item2.Key.Y, item2.Key.X];
								if (token.IsCancellationRequested)
								{
									return;
								}
								if (item2.Value.Comment != null)
								{
									range.ClearComments();
									range.AddComment(Type.Missing);
									range.Comment.Text(item2.Value.Comment, Type.Missing, Type.Missing);
								}
								if (token.IsCancellationRequested)
								{
									return;
								}
								if (item2.Value.Color.HasValue)
								{
									Color value2 = item2.Value.Color.Value;
									range.Interior.Color = value2.R | (value2.G << 8) | (value2.B << 16);
								}
							}
						}
					}
					catch (Exception ex)
					{
						judged.ex = ex;
					}
				}
				if (token.IsCancellationRequested)
				{
					return;
				}
				wait.Close();
				ResetInput();
				if (token.IsCancellationRequested)
				{
					return;
				}
				if (judged.ex != null)
				{
					System.Windows.Forms.MessageBox.Show(iReporterExcelAddIn2019.Properties.Resources.EXCEL_ERROR);
					return;
				}
				foreach (string msg in judged.msgs)
				{
					System.Windows.Forms.MessageBox.Show(msg);
				}
				if (!token.IsCancellationRequested && openDlg)
				{
					ManualCluster();
				}
			}
		}
		finally
		{
			IsJudging = false;
		}
	}

	public void SetClusterColors()
	{
		lock (_inputLocker)
		{
			ColorForm color = _color;
			if (color != null && !color.IsDisposed)
			{
				_color.Show();
				return;
			}
			_color = new ColorForm();
			_color.SetColorEvent += SetSelectedColor;
			_color.GetColorEvent += GetSelectedColor;
			ShowForm(_color, null, null);
		}
	}

	private void ShowForm(Form form, System.Drawing.Point? location, System.Windows.Forms.IWin32Window window)
	{
		window = window ?? CreateWinrowHandle();
		if (!_selectionChangedSet)
		{
			_selectionChangedSet = true;
			new ComAwareEventInfo(typeof(AppEvents_Event), "SheetSelectionChange").AddEventHandler(Application, new AppEvents_SheetSelectionChangeEventHandler(OnSelectionChanged));
			new ComAwareEventInfo(typeof(AppEvents_Event), "SheetChange").AddEventHandler(Application, new AppEvents_SheetChangeEventHandler(OnSelectionChanged));
			new ComAwareEventInfo(typeof(AppEvents_Event), "SheetActivate").AddEventHandler(Application, (AppEvents_SheetActivateEventHandler)delegate
			{
				ResetInput();
			});
		}
		if (!location.HasValue)
		{
			GetWindowRect(window.Handle, out var lpRect);
			location = new System.Drawing.Point(lpRect.Right - form.Width, Math.Max(lpRect.Top, 0));
		}
		form.DesktopLocation = location.Value;
		form.Show(window);
		form.FormClosed += delegate
		{
			SetForegroundWindow(window.Handle);
		};
	}

	private void ResetInput()
	{
		InputForm2 input = _input2;
		if (input != null && input.Visible)
		{
			_input2.AllClusters.Copy(InitClusterList());
			_input2.Setup();
		}
	}

	private Color ToColor(object color)
	{
		int num2;
		if (color is double num)
		{
			num2 = (int)num;
		}
		else if (color is long num3)
		{
			num2 = (int)num3;
		}
		else if (color is int num4)
		{
			num2 = num4;
		}
		else if (color is short num5)
		{
			num2 = num5;
		}
		else
		{
			if (!(color is byte b))
			{
				return Color.FromArgb(255, 255, 255);
			}
			num2 = b;
		}
		return Color.FromArgb(num2 & 0xFF, (num2 >> 8) & 0xFF, (num2 >> 16) & 0xFF);
	}

	private void OnSelectionChanged(object Sh, Range Target)
	{
		ResetInput();
	}

	private bool GetSelected(out Range sel)
	{
		return (sel = Application.Selection as Range) != null;
	}

	private bool GetSelectedOneCell(out Range cell)
	{
		cell = null;
		if (GetSelected(out var sel))
		{
			cell = sel.OfType<Range>().FirstOrDefault();
		}
		return cell != null;
	}

	private void SetSelectedColor(Color color)
	{
		if (GetSelected(out var sel))
		{
			sel.Interior.Color = color.R | (color.G << 8) | (color.B << 16);
		}
	}

	private Color GetSelectedColor()
	{
		try
		{
			if (GetSelectedOneCell(out var cell))
			{
				return ToColor((dynamic)cell.Interior.Color);
			}
		}
		catch
		{
		}
		return Color.Empty;
	}

	private IEnumerable<string> SplitLine(string baseText)
	{
		return (baseText ?? "").Replace("\r", "").Split('\n');
	}

	private void WriteComment(Tuple<int, string>[] keyAndTexts, bool focusExcel)
	{
		bool flag = keyAndTexts.Select((Tuple<int, string> t) => t.Item1).Any((int n) => n == 0 || n == 1 || n == 2);
		int num = 0;
		if (GetSelected(out var sel))
		{
			IEnumerable<string> enumerable = sel.get_Address((object)false, (object)false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing)?.Split(',');
			foreach (string item in enumerable ?? Enumerable.Empty<string>())
			{
				int num2 = item.IndexOf(':');
				Tuple<int, int> tuple;
				Tuple<int, int> tuple2;
				if (num2 < 0)
				{
					tuple = (tuple2 = XLRangeUtil.ToRowCol(item));
				}
				else
				{
					tuple = XLRangeUtil.ToRowCol(item.Substring(0, num2));
					tuple2 = XLRangeUtil.ToRowCol(item.Substring(num2 + 1));
				}
				if (tuple.Item1 <= 0 || tuple.Item2 <= 0)
				{
					continue;
				}
				List<Rectangle> list = new List<Rectangle>();
				int row = tuple.Item1;
				while (row <= tuple2.Item1 && 1000 >= num)
				{
					int col;
					int num4;
					for (col = tuple.Item2; col <= tuple2.Item2; col = num4)
					{
						int num3 = list.FindIndex((Rectangle r) => r.Contains(col, row));
						if (num3 < 0)
						{
							if (Application.Cells[row, col] is Range range)
							{
								if ((dynamic)range.MergeCells)
								{
									Range mergeArea = range.MergeArea;
									list.Add(new Rectangle(mergeArea.Column, mergeArea.Row, mergeArea.Columns.Count, mergeArea.Rows.Count));
								}
								if (1000 < ++num)
								{
									System.Windows.Forms.MessageBox.Show(string.Format(iReporterExcelAddIn2019.Properties.Resources.TOO_MANY_SELECT, 1000));
									break;
								}
								List<string> list2 = SplitLine(range.Comment?.Text(Type.Missing, Type.Missing, Type.Missing)).ToList();
								while (list2.Count < 16)
								{
									list2.Add((list2.Count == 3 || list2.Count == 4) ? "0" : "");
								}
								foreach (Tuple<int, string> tuple3 in keyAndTexts)
								{
									if (tuple3.Item1 >= 0 && 16 > tuple3.Item1)
									{
										list2[tuple3.Item1] = tuple3.Item2;
									}
								}
								string text = string.Join(Environment.NewLine, list2);
								if (range.Comment == null)
								{
									flag = true;
									range.AddComment(Type.Missing);
									range.Comment.Text(text, Type.Missing, Type.Missing);
								}
								else
								{
									range.Comment.Text(text, Type.Missing, Type.Missing);
								}
							}
						}
						else
						{
							col = list[num3].Right - 1;
						}
						num4 = col + 1;
					}
					num4 = row + 1;
					row = num4;
				}
			}
		}
		if (flag && 0 < num)
		{
			ResetInput();
		}
		if (focusExcel)
		{
			SetForegroundWindow(CreateWinrowHandle().Handle);
		}
	}

	private void WriteTable(Table table)
	{
		try
		{
			Microsoft.Office.Interop.Excel.Worksheet worksheet = Application.ActiveWorkbook?.Sheets?.OfType<Microsoft.Office.Interop.Excel.Worksheet>().FirstOrDefault((Microsoft.Office.Interop.Excel.Worksheet s) => s.Name == table.SheetName);
			if (worksheet == null)
			{
				System.Windows.Forms.MessageBox.Show(string.Format(iReporterExcelAddInCommon.Properties.Resources.NOT_FOUND_SHEET, table.SheetName));
				return;
			}
			foreach (TableRow row in table.Rows)
			{
				foreach (TableItem item in row.Items)
				{
					if (worksheet.Cells[item.Y, item.X] is Range range)
					{
						string text = string.Join(Environment.NewLine, item.Comments);
						if (range.Comment == null)
						{
							range.AddComment(Type.Missing);
							range.Comment.Text(text, Type.Missing, Type.Missing);
						}
						else
						{
							range.Comment.Text(text, Type.Missing, Type.Missing);
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			System.Windows.Forms.MessageBox.Show(ex.ToString(), "error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
		}
	}

	private void ClearComments()
	{
		if (GetSelected(out var sel))
		{
			sel.ClearComments();
			ResetInput();
		}
	}

	private void CancelAll()
	{
		if (Application.ActiveSheet is Microsoft.Office.Interop.Excel.Worksheet worksheet)
		{
			worksheet.Cells?.ClearComments();
			Decoder.ClearTitle(GetFilePath(), worksheet.Name);
			ResetInput();
		}
	}

	private void SelectCell(string rangeName)
	{
		if (Application.ActiveWorkbook?.ActiveSheet is Microsoft.Office.Interop.Excel.Worksheet worksheet)
		{
			((_Worksheet)worksheet).get_Range((object)rangeName, Type.Missing)?.Select();
		}
	}

	private void SetSheetName(string sheetName)
	{
		if (Application.ActiveWorkbook?.ActiveSheet is Microsoft.Office.Interop.Excel.Worksheet worksheet)
		{
			worksheet.Name = sheetName;
		}
	}

	private string[] GetCellText(bool left, bool right, bool up, bool down)
	{
		List<string> list = new List<string>();
		HashSet<string> set;
		if (GetSelectedOneCell(out var cell) && Application.ActiveWindow?.ActiveSheet is Microsoft.Office.Interop.Excel.Worksheet s)
		{
			set = new HashSet<string>();
			int row = cell.Row;
			int column = cell.Column;
			Add(s, row, column);
			for (int i = 0; i < 5; i++)
			{
				for (int j = 0; j <= i; j++)
				{
					if (left)
					{
						Add(s, row - j, column - i);
						Add(s, row + j, column - i);
					}
					if (right)
					{
						Add(s, row - j, column + i);
						Add(s, row + j, column + i);
					}
					if (up)
					{
						Add(s, row - i, column - j);
						Add(s, row - i, column + j);
					}
					if (down)
					{
						Add(s, row + i, column - j);
						Add(s, row + i, column + j);
					}
				}
			}
		}
		return list.ToArray();
		void Add(Microsoft.Office.Interop.Excel.Worksheet worksheet, int num, int col)
		{
			if (num < 1 || col < 1)
			{
				return;
			}
			try
			{
				dynamic val = ((dynamic)worksheet.Cells[num, col]).MergeArea.Cells[1, 1].Value2;
				if (!((val == null) ? true : false))
				{
					dynamic val2 = Convert.ToString(val).Replace("\r\n", " ").Replace("\r", " ")
						.Replace("\n", " ")
						.Replace("\t", " ");
					if (!(string.IsNullOrWhiteSpace(val2) ? true : false) && !((!set.Add(val2)) ? true : false))
					{
						list.Add(val2);
					}
				}
			}
			catch
			{
			}
		}
	}

	private string GetFilePath()
	{
		return Path.Combine(Application.ActiveWorkbook?.Path ?? "", Application.ActiveWorkbook?.Name ?? "");
	}

	private ClusterList InitClusterList()
	{
		string filePath = GetFilePath();
		if (Application.ActiveWorkbook?.ActiveSheet is Microsoft.Office.Interop.Excel.Worksheet worksheet)
		{
			Exception ex = default(Exception);
			TitleInfo title = Decoder.FindTitle(filePath, worksheet.Name) ?? Decoder.GetSheetTitleS(BookFactory.Create(Application.ActiveWorkbook, worksheet, ref ex));
			ClusterList clusterList = new ClusterList(filePath, worksheet.Name, title);
			Comments comments = worksheet.Comments;
			int? num = comments?.Count;
			if (0 < num)
			{
				GetSelected(out var sel);
				for (int i = 0; i < num; i++)
				{
					Comment comment = comments[i + 1];
					if (comment.Parent is Range range)
					{
						bool selected = sel != null && Application.Intersect(sel, range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null;
						Range range2 = range.MergeArea ?? range;
						clusterList.Add(range2.Row, range2.Column, range2.Rows.Count, range2.Columns.Count, comment.Text(Type.Missing, Type.Missing, Type.Missing) ?? "", selected);
					}
				}
			}
			clusterList.Sort();
			return clusterList;
		}
		return new ClusterList(filePath, "", null);
	}

	public void ShowTableForm()
	{
		try
		{
			if (Application.ActiveWorkbook?.Sheets == null)
			{
				return;
			}
			List<Table> list = new List<Table>();
			foreach (object sheet in Application.ActiveWorkbook.Sheets)
			{
				Microsoft.Office.Interop.Excel.Worksheet worksheet = sheet as Microsoft.Office.Interop.Excel.Worksheet;
				Comments comments = worksheet?.Comments;
				int? num = comments?.Count;
				if (num.GetValueOrDefault() == 0)
				{
					continue;
				}
				for (int i = 0; i < num; i++)
				{
					Comment comment = comments[i + 1];
					if (!(comment.Parent is Range cell))
					{
						continue;
					}
					TableItem item = CreateTableItem(cell, comment);
					int.TryParse(item.TableNo, out var result);
					if (result >= 1)
					{
						Table table = list.Find((Table t) => t.No == item.TableNo);
						if (table == null)
						{
							table = new Table(worksheet.Name, item, isReal: true);
							list.Add(table);
						}
						else if (!(table.SheetName != worksheet.Name))
						{
							table.Add(item);
						}
					}
				}
			}
			foreach (Table item2 in list)
			{
				item2.Adjust();
			}
			list.Sort((Table lhs, Table rhs) => (lhs.No.TryParseInt() ?? 0) - (rhs.No.TryParseInt() ?? 0));
			lock (_inputLocker)
			{
				TableForm obj = tableForm;
				if (obj != null && !obj.IsDisposed)
				{
					tableForm.SetupTableList(list);
					tableForm.Show();
					return;
				}
				tableForm = new TableForm();
				tableForm.SetupTableList(list);
				tableForm.ReadTable += ReadTables;
				tableForm.WriteTable += WriteTable;
				tableForm.Reload += delegate
				{
					ShowTableForm();
				};
				ShowForm(tableForm, null, null);
			}
		}
		catch (Exception ex)
		{
			Util.Catched(ex);
		}
	}

	private TableItem CreateTableItem(Range cell, Comment cm)
	{
		cell = cell.MergeArea ?? cell;
		int column = cell.Column;
		int row = cell.Row;
		int r = column + cell.Columns.Count - 1;
		int b = row + cell.Rows.Count - 1;
		string fmt = ((dynamic)cell.NumberFormatLocal)?.ToString() ?? "";
		return new TableItem(column, row, r, b, cm.Text(Type.Missing, Type.Missing, Type.Missing) ?? "", fmt);
	}

	private Table ReadTables()
	{
		Table table = null;
		Microsoft.Office.Interop.Excel.Worksheet worksheet = Application.ActiveWorkbook?.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
		Comments comments = worksheet?.Comments;
		int num = comments?.Count ?? 0;
		if (num == 0 || !GetSelected(out var sel))
		{
			return table;
		}
		for (int i = 0; i < num; i++)
		{
			Comment comment = comments[i + 1];
			if (comment.Parent is Range range && Application.Intersect(sel, range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null)
			{
				TableItem item = CreateTableItem(range, comment);
				if (table == null)
				{
					table = new Table(worksheet.Name, item, isReal: false);
				}
				else
				{
					table.Add(item);
				}
			}
		}
		if (table != null)
		{
			table.Adjust();
			foreach (TableCol col in table.Cols)
			{
				if (string.IsNullOrEmpty(col.Key))
				{
					col.Key = "F_" + col.Index.ToString("000");
				}
				TableItem tableItem = col.Items.First();
				if (1 < tableItem.Y && worksheet.Cells[tableItem.Y - 1, tableItem.X] is Range range2)
				{
					Range mergeArea;
					if ((mergeArea = range2.MergeArea) != null)
					{
						range2 = (dynamic)worksheet.Cells[mergeArea.Row, mergeArea.Column];
					}
					col.Name = ((dynamic)range2.Text)?.ToString() ?? "";
				}
			}
			foreach (TableRow row in table.Rows)
			{
				TableItem tableItem2 = row.Items.First();
				if (1 < tableItem2.X && worksheet.Cells[tableItem2.Y, tableItem2.X - 1] is Range range3)
				{
					Range mergeArea2;
					if ((mergeArea2 = range3.MergeArea) != null)
					{
						range3 = (dynamic)worksheet.Cells[mergeArea2.Row, mergeArea2.Column];
					}
					row.Name = ((dynamic)range3.Text)?.ToString() ?? "";
				}
			}
		}
		return table;
	}

	public void CheckAllSheetName()
	{
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = Application.ActiveWorkbook;
		if (activeWorkbook == null)
		{
			return;
		}
		List<string> sheetNames = GetSheetNames(activeWorkbook).ToList();
		if (!WorksheetChecker.ValidateSheetNamesNotBlank(sheetNames, out var invalidIndexes))
		{
			string arg = string.Join(", ", invalidIndexes);
			string excelMessage = iReporterExcelAddIn2019.Properties.Resources.ExcelMessage65;
			string excelMessage2 = iReporterExcelAddIn2019.Properties.Resources.ExcelMessage64;
			System.Windows.Forms.MessageBox.Show(string.Format(excelMessage, arg, excelMessage2));
			return;
		}
		List<SheetNameInvalidCharacter> list = InvalidCharacterLoader.Load(CommonObjects.Language);
		string excelProcessorMessage = iReporterExcelAddIn2019.Properties.Resources.ExcelProcessorMessage5;
		List<string> list2 = WorksheetChecker.ValidateSheetNames(sheetNames, list, excelProcessorMessage);
		if (list2.Count > 0)
		{
			List<ColumnDefinitionModel> columns = new List<ColumnDefinitionModel>
			{
				new ColumnDefinitionModel
				{
					Header = iReporterExcelAddIn2019.Properties.Resources.SheetNameNoticeGridColumnSymbolHeader,
					PropertyName = "Symbol",
					Width = 60.0
				},
				new ColumnDefinitionModel
				{
					Header = iReporterExcelAddIn2019.Properties.Resources.SheetNameNoticeGridColumnNameHeader,
					PropertyName = "SymbolName",
					Width = 420.0
				}
			};
			List<List<string>> rowData = list.Select((SheetNameInvalidCharacter x) => new List<string> { x.Symbol, x.Name }).ToList();
			string message = string.Join(Environment.NewLine, list2);
			NotificationService notificationService = new NotificationService();
			MessageGridViewModel dataContext = new MessageGridViewModel(iReporterExcelAddIn2019.Properties.Resources.SheetNameNoticeMessageDesc, iReporterExcelAddIn2019.Properties.Resources.SheetNameNoticeGridDesc, message, columns, rowData, DisplayMode.Both, notificationService, null, isCopyButtonVisible: true, isCloseButtonVisible: false, iReporterExcelAddIn2019.Properties.Resources.SheetNameNoticeWindow);
			MessageGrid messageGrid = new MessageGrid
			{
				DataContext = dataContext
			};
			WindowInteropHelper windowInteropHelper = new WindowInteropHelper(messageGrid);
			IntPtr excelMainWindowHandle = GetExcelMainWindowHandle();
			if (excelMainWindowHandle != IntPtr.Zero)
			{
				windowInteropHelper.Owner = excelMainWindowHandle;
				messageGrid.Height = 620.0;
				messageGrid.WindowStartupLocation = WindowStartupLocation.CenterOwner;
			}
			else
			{
				messageGrid.WindowStartupLocation = WindowStartupLocation.CenterScreen;
			}
			messageGrid.ShowDialog();
		}
		else
		{
			System.Windows.Forms.MessageBox.Show(iReporterExcelAddIn2019.Properties.Resources.CheckSheetNameSuccess, iReporterExcelAddIn2019.Properties.Resources.CheckSheetNameComplete, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
		}
	}

	public IntPtr GetExcelMainWindowHandle()
	{
		return GetAncestor(new IntPtr(Application.Hwnd), 2u);
	}

	public List<int> GetReportIndex(Microsoft.Office.Interop.Excel.Workbook workbook)
	{
		List<int> list = new List<int>();
		int num = 1;
		foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in workbook.Sheets)
		{
			try
			{
				bool num2 = sheet.Visible == XlSheetVisibility.xlSheetVisible;
				bool flag = !string.IsNullOrWhiteSpace(sheet.PageSetup.PrintArea);
				if (num2 && flag)
				{
					list.Add(num);
				}
			}
			finally
			{
				num++;
			}
		}
		return list;
	}

	public IEnumerable<int> GetSheetIndexes(Microsoft.Office.Interop.Excel.Workbook workbook)
	{
		return from Microsoft.Office.Interop.Excel.Worksheet ws in workbook.Sheets
			select ws.Index;
	}

	public IEnumerable<string> GetSheetNames(Microsoft.Office.Interop.Excel.Workbook workbook)
	{
		return from Microsoft.Office.Interop.Excel.Worksheet ws in workbook.Sheets
			select ws.Name;
	}

	private void InternalStartup()
	{
		((AddInBase)this).Startup += ThisAddIn_Startup;
		((AddInBase)this).Shutdown += ThisAddIn_Shutdown;
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public ThisAddIn(ApplicationFactory factory, IServiceProvider serviceProvider)
		: base((Microsoft.Office.Tools.Factory)factory, serviceProvider, "AddIn", "ThisAddIn")
	{
		Globals.Factory = factory;
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	protected override void Initialize()
	{
		((AddInBase)this).Initialize();
		Application = ((AddInBase)this).GetHostItem<Microsoft.Office.Interop.Excel.Application>(typeof(Microsoft.Office.Interop.Excel.Application), "Application");
		Globals.ThisAddIn = this;
		System.Windows.Forms.Application.EnableVisualStyles();
		InitializeCachedData();
		InitializeControls();
		InitializeComponents();
		InitializeData();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	protected override void FinishInitialization()
	{
		InternalStartup();
		((AddInBase)this).OnStartup();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	protected override void InitializeDataBindings()
	{
		BeginInitialization();
		BindToData();
		EndInitialization();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void InitializeCachedData()
	{
		if (((AddInBase)this).DataHost != null && ((AddInBase)this).DataHost.IsCacheInitialized)
		{
			((AddInBase)this).DataHost.FillCachedData(this);
		}
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void InitializeData()
	{
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void BindToData()
	{
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Advanced)]
	private void StartCaching(string MemberName)
	{
		((AddInBase)this).DataHost.StartCaching(this, MemberName);
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Advanced)]
	private void StopCaching(string MemberName)
	{
		((AddInBase)this).DataHost.StopCaching(this, MemberName);
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Advanced)]
	private bool IsCached(string MemberName)
	{
		return ((AddInBase)this).DataHost.IsCached(this, MemberName);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void BeginInitialization()
	{
		((AddInBase)this).BeginInit();
		CustomTaskPanes.BeginInit();
		VstoSmartTags.BeginInit();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void EndInitialization()
	{
		VstoSmartTags.EndInit();
		CustomTaskPanes.EndInit();
		((AddInBase)this).EndInit();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void InitializeControls()
	{
		CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
		VstoSmartTags = Globals.Factory.CreateSmartTagCollection(null, null, "VstoSmartTags", "VstoSmartTags", this);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void InitializeComponents()
	{
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Advanced)]
	private bool NeedsFill(string MemberName)
	{
		return ((AddInBase)this).DataHost.NeedsFill(this, MemberName);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	protected override void OnShutdown()
	{
		VstoSmartTags.Dispose();
		CustomTaskPanes.Dispose();
		((AddInBase)this).OnShutdown();
	}
}
