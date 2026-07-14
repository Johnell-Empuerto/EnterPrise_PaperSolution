using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;
using Cimtops.Excel;
using iReporterExcelAddInCommon.Properties;

namespace iReporterExcelAddInCommon;

public class InputForm2 : Form
{
	public delegate void WriteCommentEventHandler(Tuple<int, string>[] rowAndTexts, bool focusExcel);

	public delegate void ClearCommentEventHandler();

	public delegate void ClusterJumpEventHandler(bool toPrev, bool onlyUnknown);

	public delegate void SelectCellEventHandler(string rangeName);

	public delegate string[] GetCellTextEventHandler(bool left, bool right, bool top, bool down);

	private class TextBoxInfo
	{
		public string PrevValue;

		public int RowNum;
	}

	private int updateStopper;

	private Dictionary<TextBox, TextBoxInfo> defValueDic = new Dictionary<TextBox, TextBoxInfo>();

	private bool suppressTypeSelection;

	private bool notCheckTable;

	private Dictionary<TextBox, string> _numValueDic = new Dictionary<TextBox, string>();

	private IContainer components;

	private CheckBox readOnlyCheckBox;

	private ListBox clusterTypeListBox;

	private TextBox inputParameterTextBox;

	private Label label16;

	private TextBox remarks10TextBox;

	private TextBox remarks9TextBox;

	private TextBox remarks8TextBox;

	private TextBox remarks7TextBox;

	private TextBox remarks6TextBox;

	private TextBox remarks5TextBox;

	private TextBox remarks4TextBox;

	private TextBox remarks3TextBox;

	private TextBox remarks2TextBox;

	private TextBox remarks1TextBox;

	private TextBox clusterIndexTextBox;

	private TextBox clusterTypeTextBox;

	private Label label14;

	private RadioButton windowsRadioButton;

	private RadioButton iOSRadioButton;

	private Label label13;

	private Label label12;

	private Label label11;

	private Label label10;

	private Label label9;

	private Label label8;

	private Label label7;

	private Label label6;

	private Label label5;

	private CheckBox extensionCheckBox;

	private Label label4;

	private Label label3;

	private Label label2;

	private TextBox clusterNameTextBox;

	private Label label1;

	private GroupBox groupBox1;

	private ComboBox languageComboBox;

	private TextBox kindNameTextBox;

	private Label label15;

	private Button btnDelete;

	private Button btnClose;

	private ListBox clusterListBox;

	private Label label17;

	private Button lnBtn;

	private Button leBtn;

	private Button reBtn;

	private Button rnBtn;

	private Button extractBtn;

	private Label label18;

	private TextBox titleTextBox;

	private ListBox clusterNameHintListBox;

	private CheckBox bottomHintCheckBox;

	private CheckBox topHintCheckBox;

	private CheckBox rightHintCheckBox;

	private CheckBox leftHintCheckBox;

	private Label label19;

	private Label confiLabel1;

	private Label confiLabel2;

	private Button CancelAllBtn;

	private Button toSheetNameBtn;

	private RadioButton androidRadioButton;

	private InputFormController2 Controller => InputFormController2.GetInstance();

	public ClusterList AllClusters { get; set; }

	public ClusterList Selected { get; private set; }

	public DeviceType SelectedDevice
	{
		get
		{
			if (windowsRadioButton.Checked)
			{
				return DeviceType.Windows;
			}
			if (iOSRadioButton.Checked)
			{
				return DeviceType.IOS;
			}
			if (androidRadioButton.Checked)
			{
				return DeviceType.Android;
			}
			return DeviceType.IOS;
		}
	}

	public static int GetWindowScale => (int)((double)(Screen.PrimaryScreen.Bounds.Width * 100) / SystemParameters.PrimaryScreenWidth);

	public event WriteCommentEventHandler WriteCommentEvent;

	public event ClearCommentEventHandler ClearCommentEvent;

	public event ClearCommentEventHandler LanguageChangeEvent;

	public event ClearCommentEventHandler CancelAllEvent;

	public event SelectCellEventHandler SelectCellEvent;

	public event ClearCommentEventHandler ExtractEvent;

	public event GetCellTextEventHandler GetCellTextEvent;

	public event SelectCellEventHandler SetSheetNameEvent;

	private void StartInitializing()
	{
		updateStopper++;
	}

	private void EndIntializing()
	{
		updateStopper--;
	}

	private bool IsInitializing()
	{
		return 0 < updateStopper;
	}

	public InputForm2()
	{
		InitializeComponent();
		SetEvent(remarks10TextBox);
		SetEvent(remarks9TextBox);
		SetEvent(remarks8TextBox);
		SetEvent(remarks7TextBox);
		SetEvent(remarks6TextBox);
		SetEvent(remarks5TextBox);
		SetEvent(remarks3TextBox);
		SetEvent(remarks2TextBox);
		SetEvent(clusterIndexTextBox);
		SetEvent(remarks1TextBox);
		SetEvent(remarks4TextBox);
		SetEvent(clusterNameTextBox);
		titleTextBox.LostFocus += titleTextBox_Leave;
	}

	private void SetEvent(TextBox box)
	{
		box.LostFocus += OnLeaveTextBox;
	}

	public void Setup()
	{
		UpdateUI();
	}

	private void InputForm_FormClosing(object sender, FormClosingEventArgs e)
	{
		UserConfig.GetInstance().SaveToFile();
	}

	private void InputForm_Load(object sender, EventArgs e)
	{
		StartInitializing();
		label3.Text = Resources.CLUSTER_KIND;
		InitializeLanguage();
		UpdateUI();
		EndIntializing();
		AssemblyName name = Assembly.GetExecutingAssembly().GetName();
		Text = Text + "(Version " + name.Version.ToString() + ")";
	}

	private void WriteComment(int rowIndex, string text)
	{
		WriteComment(new Tuple<int, string>[1] { Tuple.Create(rowIndex, text) }, rowIndex == 1);
	}

	private void WriteComment(Tuple<int, string>[] rowAndTexts, bool focusExcel)
	{
		this.WriteCommentEvent?.Invoke(rowAndTexts, focusExcel);
	}

	private void ClearComment()
	{
		this.ClearCommentEvent?.Invoke();
	}

	private void OnClusterListSelected(object sender, EventArgs e)
	{
		if (suppressTypeSelection)
		{
			return;
		}
		if (clusterTypeListBox.SelectedItem is ClusterTypeInfo clusterTypeInfo)
		{
			List<Tuple<int, string>> list = new List<Tuple<int, string>> { Tuple.Create(1, clusterTypeInfo.TypeKey) };
			if (clusterTypeTextBox.Text != clusterTypeInfo.TypeKey)
			{
				if (!notCheckTable && Selected.HasTable)
				{
					System.Windows.Forms.MessageBox.Show(Resources.CLUSTER_HAS_TABLE);
					try
					{
						suppressTypeSelection = true;
						clusterTypeListBox.SelectedItem = clusterTypeListBox.Items.OfType<ClusterTypeInfo>().FirstOrDefault((ClusterTypeInfo t) => t.TypeKey == Selected.ClusterTypeKey);
						return;
					}
					finally
					{
						suppressTypeSelection = false;
					}
				}
				list.Add(Tuple.Create(5, ""));
			}
			clusterTypeTextBox.Text = clusterTypeInfo.TypeKey;
			kindNameTextBox.Text = (clusterTypeInfo.TypeNames.TryGetValue(Controller.GetSelectedCulture(), out var value) ? value : "");
			if (!IsInitializing())
			{
				WriteComment(list.ToArray(), focusExcel: true);
			}
		}
		else
		{
			TextBox textBox = clusterTypeTextBox;
			string text = (kindNameTextBox.Text = "");
			textBox.Text = text;
		}
	}

	private void OnLeaveTextBox(object sender, EventArgs e)
	{
		if (sender is TextBox { ReadOnly: false } textBox && defValueDic.TryGetValue(textBox, out var value) && !(value.PrevValue == textBox.Text))
		{
			WriteComment(value.RowNum, textBox.Text);
			defValueDic[textBox].PrevValue = textBox.Text;
		}
	}

	private void OnExtentionCheckChanged(object sender, EventArgs e)
	{
		if (!IsInitializing())
		{
			extensionCheckBox.ThreeState = false;
			readOnlyCheckBox.Visible = extensionCheckBox.Checked;
			if (!extensionCheckBox.Checked)
			{
				readOnlyCheckBox.ThreeState = false;
				readOnlyCheckBox.Checked = true;
			}
			List<Tuple<int, string>> list = new List<Tuple<int, string>> { Tuple.Create(4, extensionCheckBox.Checked ? "1" : "0") };
			if (!readOnlyCheckBox.Visible || readOnlyCheckBox.CheckState == CheckState.Checked)
			{
				list.Add(Tuple.Create(3, "0"));
			}
			else if (readOnlyCheckBox.CheckState == CheckState.Unchecked)
			{
				list.Add(Tuple.Create(3, "1"));
			}
			WriteComment(list.ToArray(), focusExcel: false);
		}
	}

	private void OnReadOnlyCheckChanged(object sender, EventArgs e)
	{
		if (!IsInitializing())
		{
			readOnlyCheckBox.ThreeState = false;
			WriteComment(3, readOnlyCheckBox.Checked ? "0" : "1");
		}
	}

	private void OnOSCheckChanged(object sender, EventArgs e)
	{
		if (sender is RadioButton { Checked: not false } && !IsInitializing())
		{
			DeviceType selectedDevice = SelectedDevice;
			if (Controller.UserConfigData.SelectedDeviceType != selectedDevice)
			{
				Controller.UserConfigData.SelectedDeviceType = selectedDevice;
				AllClusters.CountChanged = true;
				SetupClusterList();
			}
		}
	}

	private void SetupClusterList()
	{
		DeviceType selectedDeviceType = Controller.UserConfigData.SelectedDeviceType;
		ClusterTypeInfo.Culture = Controller.GetSelectedCulture();
		Selected = AllClusters.GetSelected();
		ClusterTypeInfo[] clusterList = Controller.GetClusterList(Selected);
		clusterListBox.BeginUpdate();
		ClusterInfo[] array = AllClusters.SetupArray(clusterList, selectedDeviceType);
		if (AllClusters.CountChanged)
		{
			AllClusters.CountChanged = false;
			clusterListBox.Items.Clear();
			ListBox.ObjectCollection items = clusterListBox.Items;
			object[] items2 = array;
			items.AddRange(items2);
		}
		SetupConfident(array);
		for (int i = 0; i < clusterListBox.Items.Count; i++)
		{
			clusterListBox.SetSelected(i, (clusterListBox.Items[i] as ClusterInfo)?.IsSelected ?? false);
		}
		clusterListBox.EndUpdate();
		try
		{
			notCheckTable = true;
			string key = Controller.ApplicationConfigData.ClusterKeyForDevice(Selected.ClusterTypeKey, selectedDeviceType);
			clusterTypeListBox.Items.Clear();
			ListBox.ObjectCollection items3 = clusterTypeListBox.Items;
			object[] items2 = clusterList;
			items3.AddRange(items2);
			clusterTypeListBox.SelectedItem = clusterList.FirstOrDefault((ClusterTypeInfo c) => c.TypeKey == key);
			if (clusterTypeListBox.SelectedItem == null)
			{
				OnClusterListSelected(clusterTypeListBox, null);
			}
		}
		finally
		{
			notCheckTable = false;
		}
	}

	private void SetupConfident(ClusterInfo[] array)
	{
		ClusterInfo[] array2 = array.Where((ClusterInfo c) => c.UseAI).ToArray();
		Label label = confiLabel1;
		bool visible = (confiLabel2.Visible = array2.Length != 0);
		label.Visible = visible;
		confiLabel2.Text = $"{array2.Count((ClusterInfo c) => c.IsNoConfi)} / {array.Length}";
	}

	private void UpdateUI()
	{
		StartInitializing();
		_ = Controller.ApplicationConfigData;
		UserConfig userConfigData = Controller.UserConfigData;
		languageComboBox.SelectedIndex = userConfigData.Language.Index;
		iOSRadioButton.Checked = userConfigData.SelectedDeviceType == DeviceType.IOS;
		windowsRadioButton.Checked = userConfigData.SelectedDeviceType == DeviceType.Windows;
		androidRadioButton.Checked = userConfigData.SelectedDeviceType == DeviceType.Android;
		SetupClusterList();
		clusterIndexTextBox.ReadOnly = 1 < Selected.Count;
		SetDefValue(clusterNameTextBox, 0);
		SetDefValue(clusterIndexTextBox, 2, isNumber: true);
		string text = Selected.GetLine(4, null, "0")?.Trim();
		if (!(text == "1"))
		{
			if (text == null)
			{
				extensionCheckBox.ThreeState = true;
				extensionCheckBox.CheckState = CheckState.Indeterminate;
				readOnlyCheckBox.Visible = true;
			}
			else
			{
				extensionCheckBox.ThreeState = true;
				extensionCheckBox.Checked = false;
				readOnlyCheckBox.Visible = false;
			}
		}
		else
		{
			extensionCheckBox.ThreeState = false;
			extensionCheckBox.Checked = true;
			readOnlyCheckBox.Visible = true;
		}
		text = Selected.GetLine(3, null, "0")?.Trim();
		if (!(text == "1"))
		{
			if (text == null)
			{
				readOnlyCheckBox.ThreeState = true;
				readOnlyCheckBox.CheckState = CheckState.Indeterminate;
			}
			else
			{
				readOnlyCheckBox.ThreeState = false;
				readOnlyCheckBox.Checked = true;
			}
		}
		else
		{
			readOnlyCheckBox.ThreeState = false;
			readOnlyCheckBox.Checked = false;
		}
		extensionCheckBox.Visible = true;
		SetDefValue(remarks1TextBox, 6);
		SetDefValue(remarks2TextBox, 7);
		SetDefValue(remarks3TextBox, 8);
		SetDefValue(remarks4TextBox, 9);
		SetDefValue(remarks5TextBox, 10);
		SetDefValue(remarks6TextBox, 11);
		SetDefValue(remarks7TextBox, 12);
		SetDefValue(remarks8TextBox, 13);
		SetDefValue(remarks9TextBox, 14);
		SetDefValue(remarks10TextBox, 15);
		inputParameterTextBox.Text = Selected.GetLine(5);
		titleTextBox.Text = AllClusters.Title?.Text ?? "";
		TextBox textBox = titleTextBox;
		bool enabled = (toSheetNameBtn.Enabled = AllClusters.Title != null);
		textBox.Enabled = enabled;
		leftHintCheckBox.Checked = userConfigData.HintLeft;
		rightHintCheckBox.Checked = userConfigData.HintRight;
		topHintCheckBox.Checked = userConfigData.HintUp;
		bottomHintCheckBox.Checked = userConfigData.HintDown;
		InitHint();
		EndIntializing();
	}

	private void SetDefValue(TextBox box, int rowNum, bool isNumber = false)
	{
		box.Text = Selected.GetLine(rowNum);
		defValueDic[box] = new TextBoxInfo
		{
			PrevValue = box.Text,
			RowNum = rowNum
		};
		if (!isNumber)
		{
			return;
		}
		EventHandler value = delegate(object sender, EventArgs e)
		{
			TextBox textBox = (TextBox)sender;
			string text = _numValueDic[textBox];
			string text2 = textBox.Text;
			if (!string.IsNullOrEmpty(text2) && text != text2 && !int.TryParse(text2, out var _))
			{
				box.Text = text;
			}
			else
			{
				_numValueDic[textBox] = text2;
			}
		};
		if (!_numValueDic.ContainsKey(box))
		{
			box.TextChanged += value;
		}
		_numValueDic[box] = box.Text;
	}

	private void InitializeLanguage()
	{
		languageComboBox.Items.Clear();
		ComboBox.ObjectCollection items = languageComboBox.Items;
		object[] items2 = (from la in Language.GetAll()
			select la.Name).ToArray();
		items.AddRange(items2);
	}

	private void languageComboBox_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!IsInitializing())
		{
			string lang = languageComboBox.Items[languageComboBox.SelectedIndex].ToString();
			Controller.UserConfigData.SetLang(lang);
			this.LanguageChangeEvent?.Invoke();
		}
	}

	private void btnDelete_Click(object sender, EventArgs e)
	{
		ClearComment();
	}

	private void btnClose_Click(object sender, EventArgs e)
	{
		Close();
	}

	private void clusterListBox_DrawItem(object sender, DrawItemEventArgs e)
	{
		if (!((uint)e.Index <= clusterListBox.Items?.Count) || !(clusterListBox.Items[e.Index] is ClusterInfo clusterInfo))
		{
			return;
		}
		e.DrawBackground();
		string text = ((clusterInfo.ClusterIndex < 0) ? " -" : clusterInfo.ClusterIndex.ToString().PadLeft(2)) + ".  ";
		using (SolidBrush brush = new SolidBrush(e.ForeColor))
		{
			Rectangle bounds = e.Bounds;
			bounds.Height /= 2;
			e.Graphics.DrawString(text + clusterInfo.ClusterName, e.Font, brush, bounds);
			text = "  " + clusterInfo.TypeName;
			bounds = e.Bounds;
			bounds.Y += bounds.Height / 2;
			if (clusterInfo.IsUnknown && (e.State & DrawItemState.Selected) == 0)
			{
				Util.DrawString(e.Graphics, text, e.Font, Color.Red, bounds);
			}
			else if (clusterInfo.IsNoConfi && (e.State & DrawItemState.Selected) == 0)
			{
				Util.DrawString(e.Graphics, text, e.Font, Color.Orange, bounds);
			}
			else
			{
				e.Graphics.DrawString(text, e.Font, brush, bounds);
			}
		}
		e.DrawFocusRectangle();
	}

	private void clusterListBox_MeasureItem(object sender, MeasureItemEventArgs e)
	{
		e.ItemHeight *= 2 * GetWindowScale;
		e.ItemHeight /= 100;
	}

	private void clusterTypeListBox_MeasureItem(object sender, MeasureItemEventArgs e)
	{
		e.ItemHeight *= GetWindowScale;
		e.ItemHeight /= 100;
	}

	private void clusterListBox_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (IsInitializing())
		{
			return;
		}
		List<CellRect> list = new List<CellRect>();
		foreach (ClusterInfo item in clusterListBox.SelectedItems.OfType<ClusterInfo>())
		{
			CellRect rect = item.GetRect();
			int index = list.FindIndex((CellRect r) => r.Union(rect));
			if (index < 0)
			{
				list.Add(rect);
				continue;
			}
			while (true)
			{
				int num = list.FindIndex((CellRect r) => r.Union(list[index]));
				if (num < 0)
				{
					break;
				}
				if (index < num)
				{
					num--;
				}
				list.RemoveAt(index);
				index = num;
			}
		}
		if (list.Count != 0)
		{
			this.SelectCellEvent?.Invoke(string.Join(",", list));
		}
	}

	private void arrowBtn_Click(object sender, EventArgs e)
	{
		int arrowIndex = GetArrowIndex(sender);
		if (arrowIndex < 0)
		{
			System.Windows.Forms.MessageBox.Show(Resources.NOT_FOUND_CLUSTERS);
			return;
		}
		clusterListBox.ClearSelected();
		clusterListBox.SelectedIndex = arrowIndex;
	}

	private int GetArrowIndex(object sender)
	{
		int count = clusterListBox.Items.Count;
		if (count < 1)
		{
			return -1;
		}
		bool flag = sender == reBtn || sender == rnBtn;
		bool flag2 = sender == reBtn || sender == leBtn;
		int num = clusterListBox.SelectedIndex;
		if (num < 0)
		{
			num = (flag ? (count - 1) : 0);
		}
		int num2 = num;
		while (true)
		{
			if (flag)
			{
				if (count <= ++num2)
				{
					num2 = 0;
				}
			}
			else if (--num2 < 0)
			{
				num2 = count - 1;
			}
			if (flag2)
			{
				ClusterInfo obj = clusterListBox.Items[num2] as ClusterInfo;
				if (obj == null || !obj.IsUnknown)
				{
					if (num2 == num)
					{
						break;
					}
					continue;
				}
			}
			return num2;
		}
		return -1;
	}

	private void extractBtn_Click(object sender, EventArgs e)
	{
		this.ExtractEvent();
	}

	private void clusterTypeListBox_DrawItem(object sender, DrawItemEventArgs e)
	{
		if (e.Index >= 0)
		{
			e.DrawBackground();
			string text = "";
			if (clusterTypeListBox.Items[e.Index] is ClusterTypeInfo clusterTypeInfo && (e.State & DrawItemState.Selected) == 0 && 0.5 < clusterTypeInfo.Likelihood)
			{
				e.Graphics.FillRectangle(Brushes.SkyBlue, e.Bounds);
			}
			Util.DrawString(e.Graphics, clusterTypeListBox.Items[e.Index].ToString() + text, e.Font, e.ForeColor, e.Bounds);
			e.DrawFocusRectangle();
		}
	}

	private void titleTextBox_Leave(object sender, EventArgs e)
	{
		if (AllClusters.Title != null)
		{
			AllClusters.Title.Text = titleTextBox.Text;
		}
	}

	private void clusterNameHintListBox_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (clusterNameHintListBox.SelectedIndex >= 0)
		{
			clusterNameTextBox.Text = clusterNameHintListBox.Items[clusterNameHintListBox.SelectedIndex].ToString();
			clusterNameHintListBox.SelectedIndex = -1;
			OnLeaveTextBox(clusterNameTextBox, null);
		}
	}

	private void hintCheckBox_CheckedChanged(object sender, EventArgs e)
	{
		if (!IsInitializing())
		{
			UserConfig userConfigData = Controller.UserConfigData;
			userConfigData.HintUp = topHintCheckBox.Checked;
			userConfigData.HintDown = bottomHintCheckBox.Checked;
			userConfigData.HintLeft = leftHintCheckBox.Checked;
			userConfigData.HintRight = rightHintCheckBox.Checked;
			InitHint();
		}
	}

	private void InitHint()
	{
		clusterNameHintListBox.Items.Clear();
		ListBox.ObjectCollection items = clusterNameHintListBox.Items;
		object[] items2 = this.GetCellTextEvent?.Invoke(leftHintCheckBox.Checked, rightHintCheckBox.Checked, topHintCheckBox.Checked, bottomHintCheckBox.Checked);
		items.AddRange(items2);
		clusterNameHintListBox.SelectedIndex = -1;
	}

	private void CancelAllBtn_Click(object sender, EventArgs e)
	{
		this.CancelAllEvent?.Invoke();
	}

	private void toSheetNameBtn_Click(object sender, EventArgs e)
	{
		string text = titleTextBox.Text ?? "";
		string text2 = ":\\/?*[]";
		for (int i = 0; i < text2.Length; i++)
		{
			text = text.Replace(text2.Substring(i, 1), "");
		}
		if (31 < text.Length)
		{
			text = text.Substring(0, 31);
		}
		if (text == "")
		{
			System.Windows.Forms.MessageBox.Show(string.IsNullOrEmpty(titleTextBox.Text) ? Resources.SHEET_NAME_EMPTY : Resources.SHEET_NAME_INVALID);
		}
		else
		{
			this.SetSheetNameEvent?.Invoke(text);
		}
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		base.Dispose(disposing);
	}

	private void InitializeComponent()
	{
		System.ComponentModel.ComponentResourceManager componentResourceManager = new System.ComponentModel.ComponentResourceManager(typeof(iReporterExcelAddInCommon.InputForm2));
		this.readOnlyCheckBox = new System.Windows.Forms.CheckBox();
		this.clusterTypeListBox = new System.Windows.Forms.ListBox();
		this.inputParameterTextBox = new System.Windows.Forms.TextBox();
		this.label16 = new System.Windows.Forms.Label();
		this.remarks10TextBox = new System.Windows.Forms.TextBox();
		this.remarks9TextBox = new System.Windows.Forms.TextBox();
		this.remarks8TextBox = new System.Windows.Forms.TextBox();
		this.remarks7TextBox = new System.Windows.Forms.TextBox();
		this.remarks6TextBox = new System.Windows.Forms.TextBox();
		this.remarks5TextBox = new System.Windows.Forms.TextBox();
		this.remarks4TextBox = new System.Windows.Forms.TextBox();
		this.remarks3TextBox = new System.Windows.Forms.TextBox();
		this.remarks2TextBox = new System.Windows.Forms.TextBox();
		this.remarks1TextBox = new System.Windows.Forms.TextBox();
		this.clusterIndexTextBox = new System.Windows.Forms.TextBox();
		this.clusterTypeTextBox = new System.Windows.Forms.TextBox();
		this.label14 = new System.Windows.Forms.Label();
		this.windowsRadioButton = new System.Windows.Forms.RadioButton();
		this.iOSRadioButton = new System.Windows.Forms.RadioButton();
		this.label13 = new System.Windows.Forms.Label();
		this.label12 = new System.Windows.Forms.Label();
		this.label11 = new System.Windows.Forms.Label();
		this.label10 = new System.Windows.Forms.Label();
		this.label9 = new System.Windows.Forms.Label();
		this.label8 = new System.Windows.Forms.Label();
		this.label7 = new System.Windows.Forms.Label();
		this.label6 = new System.Windows.Forms.Label();
		this.label5 = new System.Windows.Forms.Label();
		this.extensionCheckBox = new System.Windows.Forms.CheckBox();
		this.label4 = new System.Windows.Forms.Label();
		this.label3 = new System.Windows.Forms.Label();
		this.label2 = new System.Windows.Forms.Label();
		this.clusterNameTextBox = new System.Windows.Forms.TextBox();
		this.label1 = new System.Windows.Forms.Label();
		this.groupBox1 = new System.Windows.Forms.GroupBox();
		this.androidRadioButton = new System.Windows.Forms.RadioButton();
		this.languageComboBox = new System.Windows.Forms.ComboBox();
		this.kindNameTextBox = new System.Windows.Forms.TextBox();
		this.label15 = new System.Windows.Forms.Label();
		this.btnDelete = new System.Windows.Forms.Button();
		this.btnClose = new System.Windows.Forms.Button();
		this.clusterListBox = new System.Windows.Forms.ListBox();
		this.label17 = new System.Windows.Forms.Label();
		this.lnBtn = new System.Windows.Forms.Button();
		this.leBtn = new System.Windows.Forms.Button();
		this.reBtn = new System.Windows.Forms.Button();
		this.rnBtn = new System.Windows.Forms.Button();
		this.extractBtn = new System.Windows.Forms.Button();
		this.label18 = new System.Windows.Forms.Label();
		this.titleTextBox = new System.Windows.Forms.TextBox();
		this.clusterNameHintListBox = new System.Windows.Forms.ListBox();
		this.bottomHintCheckBox = new System.Windows.Forms.CheckBox();
		this.topHintCheckBox = new System.Windows.Forms.CheckBox();
		this.rightHintCheckBox = new System.Windows.Forms.CheckBox();
		this.leftHintCheckBox = new System.Windows.Forms.CheckBox();
		this.label19 = new System.Windows.Forms.Label();
		this.confiLabel1 = new System.Windows.Forms.Label();
		this.confiLabel2 = new System.Windows.Forms.Label();
		this.CancelAllBtn = new System.Windows.Forms.Button();
		this.toSheetNameBtn = new System.Windows.Forms.Button();
		this.groupBox1.SuspendLayout();
		base.SuspendLayout();
		componentResourceManager.ApplyResources(this.readOnlyCheckBox, "readOnlyCheckBox");
		this.readOnlyCheckBox.Name = "readOnlyCheckBox";
		this.readOnlyCheckBox.ThreeState = true;
		this.readOnlyCheckBox.UseVisualStyleBackColor = true;
		this.readOnlyCheckBox.CheckedChanged += new System.EventHandler(OnReadOnlyCheckChanged);
		this.clusterTypeListBox.DisplayMember = "CultureName";
		this.clusterTypeListBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
		componentResourceManager.ApplyResources(this.clusterTypeListBox, "clusterTypeListBox");
		this.clusterTypeListBox.FormattingEnabled = true;
		this.clusterTypeListBox.Name = "clusterTypeListBox";
		this.clusterTypeListBox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(clusterTypeListBox_DrawItem);
		this.clusterTypeListBox.MeasureItem += new System.Windows.Forms.MeasureItemEventHandler(clusterTypeListBox_MeasureItem);
		this.clusterTypeListBox.SelectedIndexChanged += new System.EventHandler(OnClusterListSelected);
		componentResourceManager.ApplyResources(this.inputParameterTextBox, "inputParameterTextBox");
		this.inputParameterTextBox.Name = "inputParameterTextBox";
		this.inputParameterTextBox.ReadOnly = true;
		this.inputParameterTextBox.TabStop = false;
		componentResourceManager.ApplyResources(this.label16, "label16");
		this.label16.Name = "label16";
		componentResourceManager.ApplyResources(this.remarks10TextBox, "remarks10TextBox");
		this.remarks10TextBox.Name = "remarks10TextBox";
		componentResourceManager.ApplyResources(this.remarks9TextBox, "remarks9TextBox");
		this.remarks9TextBox.Name = "remarks9TextBox";
		componentResourceManager.ApplyResources(this.remarks8TextBox, "remarks8TextBox");
		this.remarks8TextBox.Name = "remarks8TextBox";
		componentResourceManager.ApplyResources(this.remarks7TextBox, "remarks7TextBox");
		this.remarks7TextBox.Name = "remarks7TextBox";
		componentResourceManager.ApplyResources(this.remarks6TextBox, "remarks6TextBox");
		this.remarks6TextBox.Name = "remarks6TextBox";
		componentResourceManager.ApplyResources(this.remarks5TextBox, "remarks5TextBox");
		this.remarks5TextBox.Name = "remarks5TextBox";
		componentResourceManager.ApplyResources(this.remarks4TextBox, "remarks4TextBox");
		this.remarks4TextBox.Name = "remarks4TextBox";
		componentResourceManager.ApplyResources(this.remarks3TextBox, "remarks3TextBox");
		this.remarks3TextBox.Name = "remarks3TextBox";
		componentResourceManager.ApplyResources(this.remarks2TextBox, "remarks2TextBox");
		this.remarks2TextBox.Name = "remarks2TextBox";
		componentResourceManager.ApplyResources(this.remarks1TextBox, "remarks1TextBox");
		this.remarks1TextBox.Name = "remarks1TextBox";
		componentResourceManager.ApplyResources(this.clusterIndexTextBox, "clusterIndexTextBox");
		this.clusterIndexTextBox.Name = "clusterIndexTextBox";
		componentResourceManager.ApplyResources(this.clusterTypeTextBox, "clusterTypeTextBox");
		this.clusterTypeTextBox.Name = "clusterTypeTextBox";
		this.clusterTypeTextBox.ReadOnly = true;
		componentResourceManager.ApplyResources(this.label14, "label14");
		this.label14.Name = "label14";
		componentResourceManager.ApplyResources(this.windowsRadioButton, "windowsRadioButton");
		this.windowsRadioButton.Name = "windowsRadioButton";
		this.windowsRadioButton.UseVisualStyleBackColor = true;
		this.windowsRadioButton.CheckedChanged += new System.EventHandler(OnOSCheckChanged);
		componentResourceManager.ApplyResources(this.iOSRadioButton, "iOSRadioButton");
		this.iOSRadioButton.Checked = true;
		this.iOSRadioButton.Name = "iOSRadioButton";
		this.iOSRadioButton.TabStop = true;
		this.iOSRadioButton.UseVisualStyleBackColor = true;
		this.iOSRadioButton.CheckedChanged += new System.EventHandler(OnOSCheckChanged);
		componentResourceManager.ApplyResources(this.label13, "label13");
		this.label13.Name = "label13";
		componentResourceManager.ApplyResources(this.label12, "label12");
		this.label12.Name = "label12";
		componentResourceManager.ApplyResources(this.label11, "label11");
		this.label11.Name = "label11";
		componentResourceManager.ApplyResources(this.label10, "label10");
		this.label10.Name = "label10";
		componentResourceManager.ApplyResources(this.label9, "label9");
		this.label9.Name = "label9";
		componentResourceManager.ApplyResources(this.label8, "label8");
		this.label8.Name = "label8";
		componentResourceManager.ApplyResources(this.label7, "label7");
		this.label7.Name = "label7";
		componentResourceManager.ApplyResources(this.label6, "label6");
		this.label6.Name = "label6";
		componentResourceManager.ApplyResources(this.label5, "label5");
		this.label5.Name = "label5";
		componentResourceManager.ApplyResources(this.extensionCheckBox, "extensionCheckBox");
		this.extensionCheckBox.Name = "extensionCheckBox";
		this.extensionCheckBox.ThreeState = true;
		this.extensionCheckBox.UseVisualStyleBackColor = true;
		this.extensionCheckBox.CheckedChanged += new System.EventHandler(OnExtentionCheckChanged);
		componentResourceManager.ApplyResources(this.label4, "label4");
		this.label4.Name = "label4";
		componentResourceManager.ApplyResources(this.label3, "label3");
		this.label3.Name = "label3";
		componentResourceManager.ApplyResources(this.label2, "label2");
		this.label2.Name = "label2";
		componentResourceManager.ApplyResources(this.clusterNameTextBox, "clusterNameTextBox");
		this.clusterNameTextBox.Name = "clusterNameTextBox";
		componentResourceManager.ApplyResources(this.label1, "label1");
		this.label1.Name = "label1";
		this.groupBox1.Controls.Add(this.androidRadioButton);
		this.groupBox1.Controls.Add(this.windowsRadioButton);
		this.groupBox1.Controls.Add(this.iOSRadioButton);
		componentResourceManager.ApplyResources(this.groupBox1, "groupBox1");
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.TabStop = false;
		componentResourceManager.ApplyResources(this.androidRadioButton, "androidRadioButton");
		this.androidRadioButton.Name = "androidRadioButton";
		this.androidRadioButton.TabStop = true;
		this.androidRadioButton.UseVisualStyleBackColor = true;
		this.androidRadioButton.CheckedChanged += new System.EventHandler(OnOSCheckChanged);
		this.languageComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.languageComboBox.FormattingEnabled = true;
		componentResourceManager.ApplyResources(this.languageComboBox, "languageComboBox");
		this.languageComboBox.Name = "languageComboBox";
		this.languageComboBox.SelectedIndexChanged += new System.EventHandler(languageComboBox_SelectedIndexChanged);
		componentResourceManager.ApplyResources(this.kindNameTextBox, "kindNameTextBox");
		this.kindNameTextBox.Name = "kindNameTextBox";
		this.kindNameTextBox.ReadOnly = true;
		componentResourceManager.ApplyResources(this.label15, "label15");
		this.label15.Name = "label15";
		componentResourceManager.ApplyResources(this.btnDelete, "btnDelete");
		this.btnDelete.Name = "btnDelete";
		this.btnDelete.UseVisualStyleBackColor = true;
		this.btnDelete.Click += new System.EventHandler(btnDelete_Click);
		this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		componentResourceManager.ApplyResources(this.btnClose, "btnClose");
		this.btnClose.Name = "btnClose";
		this.btnClose.UseVisualStyleBackColor = true;
		this.btnClose.Click += new System.EventHandler(btnClose_Click);
		this.clusterListBox.DisplayMember = "CultureName";
		this.clusterListBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
		componentResourceManager.ApplyResources(this.clusterListBox, "clusterListBox");
		this.clusterListBox.FormattingEnabled = true;
		this.clusterListBox.Name = "clusterListBox";
		this.clusterListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
		this.clusterListBox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(clusterListBox_DrawItem);
		this.clusterListBox.MeasureItem += new System.Windows.Forms.MeasureItemEventHandler(clusterListBox_MeasureItem);
		this.clusterListBox.SelectedIndexChanged += new System.EventHandler(clusterListBox_SelectedIndexChanged);
		componentResourceManager.ApplyResources(this.label17, "label17");
		this.label17.Name = "label17";
		componentResourceManager.ApplyResources(this.lnBtn, "lnBtn");
		this.lnBtn.Name = "lnBtn";
		this.lnBtn.UseVisualStyleBackColor = true;
		this.lnBtn.Click += new System.EventHandler(arrowBtn_Click);
		componentResourceManager.ApplyResources(this.leBtn, "leBtn");
		this.leBtn.Name = "leBtn";
		this.leBtn.UseVisualStyleBackColor = true;
		this.leBtn.Click += new System.EventHandler(arrowBtn_Click);
		componentResourceManager.ApplyResources(this.reBtn, "reBtn");
		this.reBtn.Name = "reBtn";
		this.reBtn.UseVisualStyleBackColor = true;
		this.reBtn.Click += new System.EventHandler(arrowBtn_Click);
		componentResourceManager.ApplyResources(this.rnBtn, "rnBtn");
		this.rnBtn.Name = "rnBtn";
		this.rnBtn.UseVisualStyleBackColor = true;
		this.rnBtn.Click += new System.EventHandler(arrowBtn_Click);
		componentResourceManager.ApplyResources(this.extractBtn, "extractBtn");
		this.extractBtn.Name = "extractBtn";
		this.extractBtn.UseVisualStyleBackColor = true;
		this.extractBtn.Click += new System.EventHandler(extractBtn_Click);
		componentResourceManager.ApplyResources(this.label18, "label18");
		this.label18.Name = "label18";
		componentResourceManager.ApplyResources(this.titleTextBox, "titleTextBox");
		this.titleTextBox.Name = "titleTextBox";
		componentResourceManager.ApplyResources(this.clusterNameHintListBox, "clusterNameHintListBox");
		this.clusterNameHintListBox.FormattingEnabled = true;
		this.clusterNameHintListBox.Name = "clusterNameHintListBox";
		this.clusterNameHintListBox.TabStop = false;
		this.clusterNameHintListBox.SelectedIndexChanged += new System.EventHandler(clusterNameHintListBox_SelectedIndexChanged);
		componentResourceManager.ApplyResources(this.bottomHintCheckBox, "bottomHintCheckBox");
		this.bottomHintCheckBox.Checked = true;
		this.bottomHintCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
		this.bottomHintCheckBox.Name = "bottomHintCheckBox";
		this.bottomHintCheckBox.UseVisualStyleBackColor = true;
		this.bottomHintCheckBox.CheckedChanged += new System.EventHandler(hintCheckBox_CheckedChanged);
		componentResourceManager.ApplyResources(this.topHintCheckBox, "topHintCheckBox");
		this.topHintCheckBox.Checked = true;
		this.topHintCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
		this.topHintCheckBox.Name = "topHintCheckBox";
		this.topHintCheckBox.UseVisualStyleBackColor = true;
		this.topHintCheckBox.CheckedChanged += new System.EventHandler(hintCheckBox_CheckedChanged);
		componentResourceManager.ApplyResources(this.rightHintCheckBox, "rightHintCheckBox");
		this.rightHintCheckBox.Checked = true;
		this.rightHintCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
		this.rightHintCheckBox.Name = "rightHintCheckBox";
		this.rightHintCheckBox.UseVisualStyleBackColor = true;
		this.rightHintCheckBox.CheckedChanged += new System.EventHandler(hintCheckBox_CheckedChanged);
		componentResourceManager.ApplyResources(this.leftHintCheckBox, "leftHintCheckBox");
		this.leftHintCheckBox.Checked = true;
		this.leftHintCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
		this.leftHintCheckBox.Name = "leftHintCheckBox";
		this.leftHintCheckBox.UseVisualStyleBackColor = true;
		this.leftHintCheckBox.CheckedChanged += new System.EventHandler(hintCheckBox_CheckedChanged);
		componentResourceManager.ApplyResources(this.label19, "label19");
		this.label19.Name = "label19";
		componentResourceManager.ApplyResources(this.confiLabel1, "confiLabel1");
		this.confiLabel1.Name = "confiLabel1";
		componentResourceManager.ApplyResources(this.confiLabel2, "confiLabel2");
		this.confiLabel2.Name = "confiLabel2";
		componentResourceManager.ApplyResources(this.CancelAllBtn, "CancelAllBtn");
		this.CancelAllBtn.Name = "CancelAllBtn";
		this.CancelAllBtn.UseVisualStyleBackColor = true;
		this.CancelAllBtn.Click += new System.EventHandler(CancelAllBtn_Click);
		componentResourceManager.ApplyResources(this.toSheetNameBtn, "toSheetNameBtn");
		this.toSheetNameBtn.Name = "toSheetNameBtn";
		this.toSheetNameBtn.UseVisualStyleBackColor = true;
		this.toSheetNameBtn.Click += new System.EventHandler(toSheetNameBtn_Click);
		base.AcceptButton = this.btnClose;
		componentResourceManager.ApplyResources(this, "$this");
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		base.CancelButton = this.btnClose;
		base.Controls.Add(this.toSheetNameBtn);
		base.Controls.Add(this.CancelAllBtn);
		base.Controls.Add(this.confiLabel2);
		base.Controls.Add(this.confiLabel1);
		base.Controls.Add(this.clusterNameHintListBox);
		base.Controls.Add(this.bottomHintCheckBox);
		base.Controls.Add(this.topHintCheckBox);
		base.Controls.Add(this.rightHintCheckBox);
		base.Controls.Add(this.leftHintCheckBox);
		base.Controls.Add(this.label19);
		base.Controls.Add(this.titleTextBox);
		base.Controls.Add(this.label18);
		base.Controls.Add(this.extractBtn);
		base.Controls.Add(this.rnBtn);
		base.Controls.Add(this.reBtn);
		base.Controls.Add(this.leBtn);
		base.Controls.Add(this.lnBtn);
		base.Controls.Add(this.clusterListBox);
		base.Controls.Add(this.label17);
		base.Controls.Add(this.btnDelete);
		base.Controls.Add(this.btnClose);
		base.Controls.Add(this.kindNameTextBox);
		base.Controls.Add(this.label15);
		base.Controls.Add(this.languageComboBox);
		base.Controls.Add(this.readOnlyCheckBox);
		base.Controls.Add(this.clusterTypeListBox);
		base.Controls.Add(this.inputParameterTextBox);
		base.Controls.Add(this.label16);
		base.Controls.Add(this.remarks10TextBox);
		base.Controls.Add(this.remarks9TextBox);
		base.Controls.Add(this.remarks8TextBox);
		base.Controls.Add(this.remarks7TextBox);
		base.Controls.Add(this.remarks6TextBox);
		base.Controls.Add(this.remarks5TextBox);
		base.Controls.Add(this.remarks4TextBox);
		base.Controls.Add(this.remarks3TextBox);
		base.Controls.Add(this.remarks2TextBox);
		base.Controls.Add(this.remarks1TextBox);
		base.Controls.Add(this.clusterIndexTextBox);
		base.Controls.Add(this.clusterTypeTextBox);
		base.Controls.Add(this.label14);
		base.Controls.Add(this.label13);
		base.Controls.Add(this.label12);
		base.Controls.Add(this.label11);
		base.Controls.Add(this.label10);
		base.Controls.Add(this.label9);
		base.Controls.Add(this.label8);
		base.Controls.Add(this.label7);
		base.Controls.Add(this.label6);
		base.Controls.Add(this.label5);
		base.Controls.Add(this.extensionCheckBox);
		base.Controls.Add(this.label4);
		base.Controls.Add(this.label3);
		base.Controls.Add(this.label2);
		base.Controls.Add(this.clusterNameTextBox);
		base.Controls.Add(this.label1);
		base.Controls.Add(this.groupBox1);
		base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		base.Name = "InputForm2";
		base.FormClosing += new System.Windows.Forms.FormClosingEventHandler(InputForm_FormClosing);
		base.Load += new System.EventHandler(InputForm_Load);
		this.groupBox1.ResumeLayout(false);
		this.groupBox1.PerformLayout();
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
