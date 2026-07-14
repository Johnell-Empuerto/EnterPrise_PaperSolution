using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Cimtops.Excel;
using Cimtops.R2Cluster;
using iReporterExcelAddInCommon.Properties;

namespace iReporterExcelAddInCommon;

public class ColorForm : Form
{
	public delegate void SetColorEventHandler(Color color);

	public delegate Color GetColorEventHandler();

	private class ActionKind
	{
		public string Kind;

		public string Caption;

		public ActionKind(string kind, string caption)
		{
			Kind = kind;
			Caption = caption;
		}

		public bool IsValid(DeviceType dev)
		{
			return ActionTypes.GetDevices(Kind)?.Contains(dev) ?? true;
		}
	}

	private RadioButton[] _captions;

	private readonly Dictionary<string, string> name2Type = new Dictionary<string, string>();

	private readonly ActionKind[] Actions = new ActionKind[15]
	{
		new ActionKind("document", Resources.ACTION_KIND_DOCUMENT),
		new ActionKind("sheetJump", Resources.ACTION_KIND_SHEET_JUMP),
		new ActionKind("menu", Resources.ACTION_KIND_MENU),
		new ActionKind("openURL", Resources.ACTION_KIND_OPEN_URL),
		new ActionKind("noNeedToFillOut", Resources.ACTION_KIND_NO_NTFO),
		new ActionKind("sheetCopy", Resources.ACTION_KIND_SHEET_COPY),
		new ActionKind("backToLibrary", Resources.ACTION_KIND_BTL),
		new ActionKind("conmasIotGateway", Resources.ACTION_KIND_CIOTG),
		new ActionKind("timerStart", Resources.ACTION_KIND_START_TIMER),
		new ActionKind("runCommand", Resources.ACTION_KIND_RUN_COMMAND),
		new ActionKind("outputText", Resources.ACTION_KIND_OUTPUT_TEXT),
		new ActionKind("autoInputCluster", Resources.ACTION_KIND_AUTO_INPUT_CLUSTER),
		new ActionKind("createQRCode", Resources.ACTION_KIND_QR_CODE),
		new ActionKind("biometrics", Resources.ACTION_KIND_BIOMETRICS),
		new ActionKind("batchClear", Resources.ACTION_KIND_BATCHCLEAR)
	};

	private const string KEY_ACTION = "Action";

	private Size _orgSize;

	private Dictionary<Control, Rectangle> _orgChildren = new Dictionary<Control, Rectangle>();

	private Dictionary<Tuple<string, float, FontStyle>, Font> _fontDic = new Dictionary<Tuple<string, float, FontStyle>, Font>();

	private const string Filter = "Xml Files(*.xml)|*.xml";

	private IContainer components;

	private DataGridView dataGridView1;

	private Label label1;

	private Label label2;

	private DataGridView dataGridView2;

	private Button btnOK;

	private Button btnCancel;

	private Button defaultButton;

	private Label label3;

	private Button setDefaultButton;

	private Button button1;

	private Button button2;

	private Button button3;

	private Button button4;

	private Button button5;

	private Button button10;

	private Button button9;

	private Button button8;

	private Button button7;

	private Button button6;

	private Label label4;

	private GroupBox groupBox1;

	private RadioButton nameRadioUp;

	private RadioButton nameRadioAbsUp;

	private RadioButton nameRadioAbsLeft;

	private RadioButton nameRadioLeft;

	private GroupBox groupBox2;

	private RadioButton bothRadio;

	private RadioButton staticRadio;

	private DataGridViewTextBoxColumn Column1;

	private DataGridViewTextBoxColumn Column2;

	private DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;

	private DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;

	private Label label5;

	private DataGridView capGrid;

	private Button btnDelCaption;

	private Button btnImport;

	private Button btnExport;

	private bool HideAI { get; }

	private bool IsDisableAI
	{
		get
		{
			if (!HideAI)
			{
				return !UserConfig.GetInstance().Language.CanUseAI;
			}
			return false;
		}
	}

	public event SetColorEventHandler SetColorEvent;

	public event GetColorEventHandler GetColorEvent;

	public ColorForm()
	{
		InitializeComponent();
		_captions = new RadioButton[4] { nameRadioUp, nameRadioAbsUp, nameRadioLeft, nameRadioAbsLeft };
		HideAI = true;
		groupBox2.Visible = (bothRadio.Visible = (staticRadio.Visible = false));
		int num = defaultButton.Location.Y - groupBox2.Location.Y;
		defaultButton.Top -= num;
		setDefaultButton.Top -= num;
		btnImport.Top -= num;
		btnExport.Top -= num;
		btnOK.Top -= num;
		btnCancel.Top -= num;
		base.Height -= num;
		btnDelCaption.Text = Resources.DELETE;
		nameRadioUp.Left = nameRadioAbsUp.Right + 4;
		nameRadioLeft.Left = nameRadioUp.Right + 4;
		nameRadioAbsLeft.Left = nameRadioLeft.Right + 4;
		staticRadio.Left = bothRadio.Right + 4;
	}

	private void LoadRadio(RadioButton[] array, int index)
	{
		for (int i = 0; i < array.Length; i++)
		{
			array[i].Checked = i == index;
		}
	}

	private int? SaveRadio(RadioButton[] array)
	{
		for (int i = 0; i < array.Length; i++)
		{
			if (array[i].Checked)
			{
				return i;
			}
		}
		return null;
	}

	private void LoadConfig(UserConfig config)
	{
		SetupGrid(dataGridView1, config.ClusterColors);
		SetupGrid(dataGridView2, config.NotClusterColors);
		SetupCapGrid(config.CapAndKinds);
		LoadRadio(_captions, (int)config.CaptionPriority);
		if (!HideAI)
		{
			GroupBox groupBox = groupBox1;
			GroupBox groupBox2 = this.groupBox2;
			RadioButton radioButton = bothRadio;
			bool flag = (staticRadio.Enabled = !IsDisableAI);
			bool flag3 = (radioButton.Enabled = flag);
			bool enabled = (groupBox2.Enabled = flag3);
			groupBox.Enabled = enabled;
			if (IsDisableAI)
			{
				staticRadio.Checked = true;
				return;
			}
			bothRadio.Checked = config.UseAI;
			staticRadio.Checked = !config.UseAI;
		}
	}

	private bool SaveConfig(UserConfig user)
	{
		if (!CheckColor())
		{
			return false;
		}
		PutCellColor(dataGridView1, user.ClusterColors);
		PutCellColor(dataGridView2, user.NotClusterColors);
		user.CapAndKinds.Clear();
		foreach (DataGridViewRow row in (IEnumerable)capGrid.Rows)
		{
			string text = row.Cells[0].Value?.ToString();
			string key = row.Cells[1].Value?.ToString() ?? "";
			string third = Actions.FirstOrDefault((ActionKind a) => a.Caption == row.Cells[2].Value?.ToString())?.Kind;
			if (!string.IsNullOrEmpty(text) && name2Type.TryGetValue(key, out var value))
			{
				user.CapAndKinds.Add(text, value, third);
			}
		}
		int? num = SaveRadio(_captions);
		if (num.HasValue)
		{
			user.CaptionPriority = (CaptionPriority)num.Value;
		}
		if (!HideAI && !IsDisableAI)
		{
			if (bothRadio.Checked)
			{
				user.UseAI = true;
			}
			else if (staticRadio.Checked)
			{
				user.UseAI = false;
			}
		}
		return true;
	}

	private void ColorForm_Load(object sender, EventArgs e)
	{
		LoadConfig(UserConfig.GetInstance());
		_orgSize = base.ClientSize;
		SaveLocations(this);
		ResizeGridCells();
	}

	private void SetupGrid(DataGridView grid, PairList colorDic)
	{
		while (grid.RowCount < 5)
		{
			grid.Rows.Add();
		}
		int i = 0;
		foreach (CimTuple item in colorDic)
		{
			grid.Rows[i].Cells[0].Style.BackColor = item.Item1.ToColor();
			grid.Rows[i].Cells[1].Style.BackColor = item.Item2.ToColor();
			if (grid.RowCount <= ++i)
			{
				break;
			}
		}
		for (; i < grid.RowCount; i++)
		{
			grid.Rows[i].Cells[0].Style.BackColor = Color.White;
			grid.Rows[i].Cells[1].Style.BackColor = Color.White;
		}
		grid.ClearSelection();
	}

	private void SetupCapGrid(PairList data)
	{
		DataGridView dataGridView = capGrid;
		dataGridView.Rows.Clear();
		name2Type.Clear();
		DeviceType device = UserConfig.GetInstance().SelectedDeviceType;
		List<ClusterTypeInfo> list = InputFormController2.GetInstance().ApplicationConfigData.ListClusters(device);
		foreach (ClusterTypeInfo item in list)
		{
			name2Type[item.CultureName] = item.TypeKey;
		}
		string[] array = list.Select((ClusterTypeInfo c) => c.CultureName).ToArray();
		if (dataGridView.Columns.Count < 1)
		{
			dataGridView.Columns.Add("caption", Resources.CLUSTER_CAPTION);
		}
		if (dataGridView.Columns.Count < 2)
		{
			DataGridViewComboBoxColumn dataGridViewComboBoxColumn = new DataGridViewComboBoxColumn
			{
				Name = "kind",
				HeaderText = Resources.CLUSTER_KIND
			};
			DataGridViewComboBoxCell.ObjectCollection items = dataGridViewComboBoxColumn.Items;
			object[] items2 = array;
			items.AddRange(items2);
			dataGridView.Columns.Add(dataGridViewComboBoxColumn);
		}
		ActionKind[] source = Actions.Where((ActionKind a) => a.IsValid(device)).ToArray();
		if (dataGridView.Columns.Count < 3)
		{
			DataGridViewComboBoxColumn dataGridViewComboBoxColumn2 = new DataGridViewComboBoxColumn
			{
				Name = "Action",
				HeaderText = Resources.ACTION_KIND
			};
			DataGridViewComboBoxCell.ObjectCollection items3 = dataGridViewComboBoxColumn2.Items;
			object[] items2 = source.Select((ActionKind a) => a.Caption).ToArray();
			items3.AddRange(items2);
			dataGridView.Columns.Add(dataGridViewComboBoxColumn2);
		}
		foreach (CimTuple t in data)
		{
			DataGridViewRow dataGridViewRow = dataGridView.Rows[dataGridView.Rows.Add()];
			dataGridViewRow.Cells[0].Value = t.Item1;
			dataGridViewRow.Cells[1].Value = list.FirstOrDefault((ClusterTypeInfo c) => c.TypeKey == t.Item2)?.CultureName ?? "";
			dataGridViewRow.Cells[2].Value = source.FirstOrDefault((ActionKind a) => a.Kind == t.Item3)?.Caption;
			SetActionKindReadOnly(dataGridViewRow, t.Item2 == "Action");
			dataGridViewRow.Height = dataGridView.RowTemplate.Height;
		}
	}

	private void ResizeGridCells()
	{
		DataGridView[] array = new DataGridView[3] { dataGridView1, dataGridView2, capGrid };
		foreach (DataGridView dataGridView in array)
		{
			int num = (dataGridView.RowHeadersVisible ? dataGridView.RowHeadersWidth : 0);
			int num2 = num;
			if (num < dataGridView.Width)
			{
				for (int j = 0; j < dataGridView.ColumnCount; j++)
				{
					dataGridView.Columns[j].Width = (dataGridView.Width - num) * (j + 1) / dataGridView.ColumnCount - num2 + num;
					num2 += dataGridView.Columns[j].Width;
				}
			}
			if (!dataGridView.AllowUserToAddRows)
			{
				dataGridView.ColumnHeadersHeight = Math.Max(4, Math.Min(32768, dataGridView.Height / (dataGridView.RowCount + 1)));
				num2 = dataGridView.ColumnHeadersHeight;
				for (int k = 0; k < dataGridView.RowCount; k++)
				{
					dataGridView.Rows[k].Height = dataGridView.Height * (k + 2) / (dataGridView.RowCount + 1) - num2;
					num2 += dataGridView.Rows[k].Height;
				}
				continue;
			}
			dataGridView.ColumnHeadersHeight = dataGridView1.ColumnHeadersHeight;
			if (0 >= dataGridView1.RowCount)
			{
				continue;
			}
			int num3 = (dataGridView.RowTemplate.Height = dataGridView1.Rows[0].Height);
			int num5 = num3;
			foreach (DataGridViewRow item in (IEnumerable)dataGridView.Rows)
			{
				item.Height = num5;
			}
		}
	}

	private bool IsWhite(Color color)
	{
		if (color.A != 0)
		{
			if (color.R == byte.MaxValue && color.G == byte.MaxValue)
			{
				return color.B == byte.MaxValue;
			}
			return false;
		}
		return true;
	}

	private void PutCellColor(DataGridView grid, PairList colorDic)
	{
		colorDic.Clear();
		foreach (DataGridViewRow item in (IEnumerable)grid.Rows)
		{
			Color backColor = item.Cells[0].Style.BackColor;
			if (!IsWhite(backColor))
			{
				colorDic.Add(backColor.ToText(), item.Cells[1].Style.BackColor.ToText(), null);
			}
		}
	}

	private bool CheckColor()
	{
		HashSet<int> hashSet = new HashSet<int>();
		DataGridView[] array = new DataGridView[2] { dataGridView1, dataGridView2 };
		for (int i = 0; i < array.Length; i++)
		{
			foreach (DataGridViewRow item in (IEnumerable)array[i].Rows)
			{
				Color backColor = item.Cells[0].Style.BackColor;
				if (!IsWhite(backColor) && !hashSet.Add(backColor.ToArgb()))
				{
					MessageBox.Show(label3.Text);
					return false;
				}
			}
		}
		return true;
	}

	private void btnOK_Click(object sender, EventArgs e)
	{
		if (SaveConfig(UserConfig.GetInstance()))
		{
			UserConfig.GetInstance().SaveToFile();
			Close();
		}
	}

	private void btnCancel_Click(object sender, EventArgs e)
	{
		Close();
	}

	private void defaultButton_Click(object sender, EventArgs e)
	{
		LoadConfig(UserConfig.GetInstance(ConfigKind.Default));
	}

	private void setDefaultButton_Click(object sender, EventArgs e)
	{
		if (SaveConfig(UserConfig.GetInstance(ConfigKind.Default)))
		{
			UserConfig.GetInstance().SaveToFile();
		}
	}

	private bool HasFlag(DataGridViewCellPaintingEventArgs e, DataGridViewPaintParts flg)
	{
		return (e.PaintParts & flg) == flg;
	}

	private bool HasFlag(DataGridViewCellPaintingEventArgs e, DataGridViewElementStates flg)
	{
		return (e.State & flg) == flg;
	}

	private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
	{
		if (e.ColumnIndex < 0 || e.RowIndex < 0 || !HasFlag(e, DataGridViewPaintParts.Background) || !(sender is DataGridView dataGridView))
		{
			return;
		}
		using (SolidBrush brush = new SolidBrush(e.CellStyle.BackColor))
		{
			e.Graphics.FillRectangle(brush, e.CellBounds);
		}
		List<Point> selectedCells = (from DataGridViewCell c in dataGridView.SelectedCells
			select new Point(c.ColumnIndex, c.RowIndex)).ToList();
		DataGridViewPaintParts dataGridViewPaintParts = e.PaintParts;
		if (HasFlag(e, DataGridViewPaintParts.SelectionBackground) && HasFlag(e, DataGridViewElementStates.Selected))
		{
			Func<int, int, bool> func = (int dx, int dy) => selectedCells.Contains(new Point(e.ColumnIndex + dx, e.RowIndex + dy));
			dataGridViewPaintParts &= ~(DataGridViewPaintParts.Background | DataGridViewPaintParts.SelectionBackground);
			if (HasFlag(e, DataGridViewPaintParts.Border))
			{
				e.Paint(e.CellBounds, DataGridViewPaintParts.Border);
				dataGridViewPaintParts &= ~DataGridViewPaintParts.Border;
			}
			Rectangle cellBounds = e.CellBounds;
			using Pen pen = new Pen(Color.Black, 2f);
			if (!func(0, -1))
			{
				e.Graphics.DrawLine(pen, cellBounds.X, cellBounds.Y + 1, cellBounds.Right, cellBounds.Y + 1);
			}
			else
			{
				if (!func(-1, -1))
				{
					e.Graphics.DrawLine(pen, cellBounds.X, cellBounds.Y + 1, cellBounds.X + 2, cellBounds.Y + 1);
				}
				if (!func(1, -1))
				{
					e.Graphics.DrawLine(pen, cellBounds.Right - 2, cellBounds.Y + 1, cellBounds.Right, cellBounds.Y + 1);
				}
			}
			if (!func(0, 1))
			{
				e.Graphics.DrawLine(pen, cellBounds.X, cellBounds.Bottom - 1, cellBounds.Right, cellBounds.Bottom - 1);
			}
			else
			{
				if (!func(-1, 1))
				{
					e.Graphics.DrawLine(pen, cellBounds.X, cellBounds.Bottom - 1, cellBounds.X + 2, cellBounds.Bottom - 1);
				}
				if (!func(1, 1))
				{
					e.Graphics.DrawLine(pen, cellBounds.Right - 2, cellBounds.Bottom - 1, cellBounds.Right, cellBounds.Bottom - 1);
				}
			}
			if (!func(-1, 0))
			{
				e.Graphics.DrawLine(pen, cellBounds.X + 1, cellBounds.Y, cellBounds.X + 1, cellBounds.Bottom);
			}
			if (!func(1, 0))
			{
				e.Graphics.DrawLine(pen, cellBounds.Right - 1, cellBounds.Y, cellBounds.Right - 1, cellBounds.Bottom);
			}
		}
		e.Paint(e.ClipBounds, dataGridViewPaintParts);
		e.Handled = true;
	}

	private void dataGridView1_SelectionChanged(object sender, EventArgs e)
	{
		(sender as DataGridView)?.Invalidate();
	}

	private void OnSettingButtonClick(object sender, EventArgs e)
	{
		int num;
		if (sender == button1)
		{
			num = 0;
		}
		else if (sender == button2)
		{
			num = 1;
		}
		else if (sender == button3)
		{
			num = 2;
		}
		else if (sender == button4)
		{
			num = 3;
		}
		else if (sender == button5)
		{
			num = 4;
		}
		else if (sender == button6)
		{
			num = 5;
		}
		else if (sender == button7)
		{
			num = 6;
		}
		else if (sender == button8)
		{
			num = 7;
		}
		else if (sender == button9)
		{
			num = 8;
		}
		else
		{
			if (sender != button10)
			{
				return;
			}
			num = 9;
		}
		this.SetColorEvent?.Invoke(((num < 5) ? dataGridView1 : dataGridView2).Rows[num % 5].Cells[0].Style.BackColor);
	}

	private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
	{
		if (sender is DataGridView dataGridView)
		{
			dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = this.GetColorEvent?.Invoke() ?? Color.Empty;
		}
	}

	private void ColorForm_Resize(object sender, EventArgs e)
	{
		if (0 < base.ClientSize.Width && 0 < base.ClientSize.Height)
		{
			SuspendLayout();
			LoadLocations(this);
			ResizeGridCells();
			ResumeLayout();
		}
	}

	private void SaveLocations(Control parent)
	{
		foreach (Control control in parent.Controls)
		{
			_orgChildren[control] = new Rectangle(control.Location, control.Size);
			SaveLocations(control);
		}
	}

	private void LoadLocations(Control parent)
	{
		foreach (Control control in parent.Controls)
		{
			if (_orgChildren.TryGetValue(control, out var value))
			{
				control.Location = new Point(value.X * base.ClientSize.Width / _orgSize.Width, value.Y * base.ClientSize.Height / _orgSize.Height);
				control.Width = value.Width * base.ClientSize.Width / _orgSize.Width;
				control.Height = value.Height * base.ClientSize.Height / _orgSize.Height;
				float size = 9f * Math.Min((float)base.ClientSize.Width / (float)_orgSize.Width, (float)base.ClientSize.Height / (float)_orgSize.Height);
				control.Font = GetFont(control.Font, size);
			}
			LoadLocations(control);
		}
	}

	private Font GetFont(Font org, float size)
	{
		if (org.Size == size)
		{
			return org;
		}
		Tuple<string, float, FontStyle> tuple = Tuple.Create(org.OriginalFontName, size, org.Style);
		if (!_fontDic.TryGetValue(tuple, out var value))
		{
			value = (_fontDic[tuple] = new Font(tuple.Item1, tuple.Item2, tuple.Item3));
		}
		return value;
	}

	protected override void OnClosed(EventArgs e)
	{
		base.OnClosed(e);
		base.Owner?.Focus();
	}

	private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void btnDelCaption_Click(object sender, EventArgs e)
	{
		foreach (int item in from r in (from c in capGrid.SelectedCells.OfType<DataGridViewCell>()
				select c.RowIndex).Distinct()
			orderby r descending
			select r)
		{
			if (item == capGrid.NewRowIndex)
			{
				foreach (DataGridViewCell item2 in capGrid.Rows[item].Cells.OfType<DataGridViewCell>())
				{
					item2.Value = "";
				}
			}
			else
			{
				capGrid.Rows.RemoveAt(item);
			}
		}
	}

	private void btnExport_Click(object sender, EventArgs e)
	{
		UserConfig instance = UserConfig.GetInstance(ConfigKind.Export);
		if (!SaveConfig(instance))
		{
			return;
		}
		SaveFileDialog saveFileDialog = new SaveFileDialog();
		saveFileDialog.Filter = "Xml Files(*.xml)|*.xml";
		saveFileDialog.FileName = "UserSetting.xml";
		saveFileDialog.RestoreDirectory = true;
		if (saveFileDialog.ShowDialog() != DialogResult.OK)
		{
			return;
		}
		try
		{
			instance.SaveToFile(saveFileDialog.FileName);
		}
		catch
		{
			MessageBox.Show(Resources.FAILED_FILE_SAVE);
		}
	}

	private void btnImport_Click(object sender, EventArgs e)
	{
		OpenFileDialog openFileDialog = new OpenFileDialog();
		openFileDialog.FileName = "UserSetting.xml";
		openFileDialog.Filter = "Xml Files(*.xml)|*.xml";
		openFileDialog.RestoreDirectory = true;
		if (openFileDialog.ShowDialog() != DialogResult.OK)
		{
			return;
		}
		try
		{
			LoadConfig(UserConfig.GetInstance(openFileDialog.FileName, ""));
		}
		catch
		{
			MessageBox.Show(Resources.FAILED_FILE_LOAD);
		}
	}

	private void capGrid_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
	{
		if (e.Control is DataGridViewTextBoxEditingControl dataGridViewTextBoxEditingControl)
		{
			dataGridViewTextBoxEditingControl.PreviewKeyDown -= HandleKey;
			dataGridViewTextBoxEditingControl.PreviewKeyDown += HandleKey;
		}
		else if (e.Control is DataGridViewComboBoxEditingControl dataGridViewComboBoxEditingControl)
		{
			dataGridViewComboBoxEditingControl.Height = capGrid.RowTemplate.Height;
		}
	}

	private void HandleKey(object sender, PreviewKeyDownEventArgs e)
	{
		if (e.KeyCode == Keys.Escape)
		{
			e.IsInputKey = true;
			capGrid.CancelEdit();
		}
	}

	private void SetActionKindReadOnly(DataGridViewRow row, bool? canEditN = null)
	{
		DataGridViewCell dataGridViewCell = row.Cells[2];
		string value;
		bool flag = canEditN ?? (name2Type.TryGetValue(row.Cells[1].Value?.ToString() ?? "", out value) && value == "Action");
		dataGridViewCell.ReadOnly = !flag;
		if (!flag)
		{
			dataGridViewCell.Value = null;
		}
		if (dataGridViewCell is DataGridViewComboBoxCell dataGridViewComboBoxCell)
		{
			dataGridViewComboBoxCell.FlatStyle = (flag ? FlatStyle.Standard : FlatStyle.Flat);
			dataGridViewComboBoxCell.DisplayStyle = (flag ? DataGridViewComboBoxDisplayStyle.DropDownButton : DataGridViewComboBoxDisplayStyle.Nothing);
		}
	}

	private void capGrid_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
	{
		for (int i = e.RowIndex; i < capGrid.RowCount; i++)
		{
			DataGridViewRow dataGridViewRow = capGrid.Rows[i];
			if (3 <= dataGridViewRow.Cells.Count)
			{
				SetActionKindReadOnly(dataGridViewRow);
			}
		}
	}

	private void capGrid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
	{
		if (e.ColumnIndex == 1 && 0 <= e.RowIndex)
		{
			SetActionKindReadOnly(capGrid.Rows[e.RowIndex]);
		}
	}

	private void capGrid_CellEnter(object sender, DataGridViewCellEventArgs e)
	{
		if ((sender as DataGridView)?.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn)
		{
			SendKeys.Send("{F4}");
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
		System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(iReporterExcelAddInCommon.ColorForm));
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle = new System.Windows.Forms.DataGridViewCellStyle();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
		this.dataGridView1 = new System.Windows.Forms.DataGridView();
		this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.label1 = new System.Windows.Forms.Label();
		this.label2 = new System.Windows.Forms.Label();
		this.dataGridView2 = new System.Windows.Forms.DataGridView();
		this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.btnOK = new System.Windows.Forms.Button();
		this.btnCancel = new System.Windows.Forms.Button();
		this.defaultButton = new System.Windows.Forms.Button();
		this.label3 = new System.Windows.Forms.Label();
		this.setDefaultButton = new System.Windows.Forms.Button();
		this.button1 = new System.Windows.Forms.Button();
		this.button2 = new System.Windows.Forms.Button();
		this.button3 = new System.Windows.Forms.Button();
		this.button4 = new System.Windows.Forms.Button();
		this.button5 = new System.Windows.Forms.Button();
		this.button10 = new System.Windows.Forms.Button();
		this.button9 = new System.Windows.Forms.Button();
		this.button8 = new System.Windows.Forms.Button();
		this.button7 = new System.Windows.Forms.Button();
		this.button6 = new System.Windows.Forms.Button();
		this.label4 = new System.Windows.Forms.Label();
		this.groupBox1 = new System.Windows.Forms.GroupBox();
		this.nameRadioAbsLeft = new System.Windows.Forms.RadioButton();
		this.nameRadioLeft = new System.Windows.Forms.RadioButton();
		this.nameRadioUp = new System.Windows.Forms.RadioButton();
		this.nameRadioAbsUp = new System.Windows.Forms.RadioButton();
		this.groupBox2 = new System.Windows.Forms.GroupBox();
		this.staticRadio = new System.Windows.Forms.RadioButton();
		this.bothRadio = new System.Windows.Forms.RadioButton();
		this.label5 = new System.Windows.Forms.Label();
		this.capGrid = new System.Windows.Forms.DataGridView();
		this.btnDelCaption = new System.Windows.Forms.Button();
		this.btnImport = new System.Windows.Forms.Button();
		this.btnExport = new System.Windows.Forms.Button();
		((System.ComponentModel.ISupportInitialize)this.dataGridView1).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.dataGridView2).BeginInit();
		this.groupBox1.SuspendLayout();
		this.groupBox2.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.capGrid).BeginInit();
		base.SuspendLayout();
		resources.ApplyResources(this.dataGridView1, "dataGridView1");
		this.dataGridView1.AllowUserToAddRows = false;
		this.dataGridView1.AllowUserToDeleteRows = false;
		this.dataGridView1.AllowUserToResizeColumns = false;
		this.dataGridView1.AllowUserToResizeRows = false;
		dataGridViewCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
		dataGridViewCellStyle.BackColor = System.Drawing.SystemColors.Control;
		dataGridViewCellStyle.Font = new System.Drawing.Font("MS UI Gothic", 9f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
		dataGridViewCellStyle.ForeColor = System.Drawing.SystemColors.WindowText;
		dataGridViewCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight;
		dataGridViewCellStyle.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
		dataGridViewCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
		this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle;
		this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
		this.dataGridView1.Columns.AddRange(this.Column1, this.Column2);
		this.dataGridView1.Name = "dataGridView1";
		this.dataGridView1.ReadOnly = true;
		this.dataGridView1.RowHeadersVisible = false;
		this.dataGridView1.RowTemplate.Height = 30;
		this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
		this.dataGridView1.TabStop = false;
		this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(dataGridView1_CellClick);
		this.dataGridView1.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(dataGridView1_CellPainting);
		this.dataGridView1.SelectionChanged += new System.EventHandler(dataGridView1_SelectionChanged);
		this.dataGridView1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(dataGridView1_KeyPress);
		this.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
		resources.ApplyResources(this.Column1, "Column1");
		this.Column1.Name = "Column1";
		this.Column1.ReadOnly = true;
		this.Column1.Resizable = System.Windows.Forms.DataGridViewTriState.False;
		this.Column1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
		resources.ApplyResources(this.Column2, "Column2");
		this.Column2.Name = "Column2";
		this.Column2.ReadOnly = true;
		this.Column2.Resizable = System.Windows.Forms.DataGridViewTriState.False;
		this.Column2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
		resources.ApplyResources(this.label1, "label1");
		this.label1.Name = "label1";
		resources.ApplyResources(this.label2, "label2");
		this.label2.Name = "label2";
		resources.ApplyResources(this.dataGridView2, "dataGridView2");
		this.dataGridView2.AllowUserToAddRows = false;
		this.dataGridView2.AllowUserToDeleteRows = false;
		this.dataGridView2.AllowUserToResizeColumns = false;
		this.dataGridView2.AllowUserToResizeRows = false;
		dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
		dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
		dataGridViewCellStyle2.Font = new System.Drawing.Font("MS UI Gothic", 9f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
		dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
		dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
		dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
		dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
		this.dataGridView2.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
		this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
		this.dataGridView2.Columns.AddRange(this.dataGridViewTextBoxColumn1, this.dataGridViewTextBoxColumn2);
		this.dataGridView2.Name = "dataGridView2";
		this.dataGridView2.ReadOnly = true;
		this.dataGridView2.RowHeadersVisible = false;
		this.dataGridView2.RowTemplate.Height = 30;
		this.dataGridView2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
		this.dataGridView2.TabStop = false;
		this.dataGridView2.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(dataGridView1_CellClick);
		this.dataGridView2.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(dataGridView1_CellPainting);
		this.dataGridView2.SelectionChanged += new System.EventHandler(dataGridView1_SelectionChanged);
		this.dataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
		resources.ApplyResources(this.dataGridViewTextBoxColumn1, "dataGridViewTextBoxColumn1");
		this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
		this.dataGridViewTextBoxColumn1.ReadOnly = true;
		this.dataGridViewTextBoxColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.False;
		this.dataGridViewTextBoxColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
		resources.ApplyResources(this.dataGridViewTextBoxColumn2, "dataGridViewTextBoxColumn2");
		this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
		this.dataGridViewTextBoxColumn2.ReadOnly = true;
		this.dataGridViewTextBoxColumn2.Resizable = System.Windows.Forms.DataGridViewTriState.False;
		this.dataGridViewTextBoxColumn2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
		resources.ApplyResources(this.btnOK, "btnOK");
		this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
		this.btnOK.Name = "btnOK";
		this.btnOK.UseVisualStyleBackColor = true;
		this.btnOK.Click += new System.EventHandler(btnOK_Click);
		resources.ApplyResources(this.btnCancel, "btnCancel");
		this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		this.btnCancel.Name = "btnCancel";
		this.btnCancel.UseVisualStyleBackColor = true;
		this.btnCancel.Click += new System.EventHandler(btnCancel_Click);
		resources.ApplyResources(this.defaultButton, "defaultButton");
		this.defaultButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		this.defaultButton.Name = "defaultButton";
		this.defaultButton.UseVisualStyleBackColor = true;
		this.defaultButton.Click += new System.EventHandler(defaultButton_Click);
		resources.ApplyResources(this.label3, "label3");
		this.label3.Name = "label3";
		resources.ApplyResources(this.setDefaultButton, "setDefaultButton");
		this.setDefaultButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		this.setDefaultButton.Name = "setDefaultButton";
		this.setDefaultButton.UseVisualStyleBackColor = true;
		this.setDefaultButton.Click += new System.EventHandler(setDefaultButton_Click);
		resources.ApplyResources(this.button1, "button1");
		this.button1.Name = "button1";
		this.button1.UseVisualStyleBackColor = true;
		this.button1.Click += new System.EventHandler(OnSettingButtonClick);
		resources.ApplyResources(this.button2, "button2");
		this.button2.Name = "button2";
		this.button2.UseVisualStyleBackColor = true;
		this.button2.Click += new System.EventHandler(OnSettingButtonClick);
		resources.ApplyResources(this.button3, "button3");
		this.button3.Name = "button3";
		this.button3.UseVisualStyleBackColor = true;
		this.button3.Click += new System.EventHandler(OnSettingButtonClick);
		resources.ApplyResources(this.button4, "button4");
		this.button4.Name = "button4";
		this.button4.UseVisualStyleBackColor = true;
		this.button4.Click += new System.EventHandler(OnSettingButtonClick);
		resources.ApplyResources(this.button5, "button5");
		this.button5.Name = "button5";
		this.button5.UseVisualStyleBackColor = true;
		this.button5.Click += new System.EventHandler(OnSettingButtonClick);
		resources.ApplyResources(this.button10, "button10");
		this.button10.Name = "button10";
		this.button10.UseVisualStyleBackColor = true;
		this.button10.Click += new System.EventHandler(OnSettingButtonClick);
		resources.ApplyResources(this.button9, "button9");
		this.button9.Name = "button9";
		this.button9.UseVisualStyleBackColor = true;
		this.button9.Click += new System.EventHandler(OnSettingButtonClick);
		resources.ApplyResources(this.button8, "button8");
		this.button8.Name = "button8";
		this.button8.UseVisualStyleBackColor = true;
		this.button8.Click += new System.EventHandler(OnSettingButtonClick);
		resources.ApplyResources(this.button7, "button7");
		this.button7.Name = "button7";
		this.button7.UseVisualStyleBackColor = true;
		this.button7.Click += new System.EventHandler(OnSettingButtonClick);
		resources.ApplyResources(this.button6, "button6");
		this.button6.Name = "button6";
		this.button6.UseVisualStyleBackColor = true;
		this.button6.Click += new System.EventHandler(OnSettingButtonClick);
		resources.ApplyResources(this.label4, "label4");
		this.label4.Name = "label4";
		resources.ApplyResources(this.groupBox1, "groupBox1");
		this.groupBox1.Controls.Add(this.nameRadioAbsLeft);
		this.groupBox1.Controls.Add(this.nameRadioLeft);
		this.groupBox1.Controls.Add(this.nameRadioUp);
		this.groupBox1.Controls.Add(this.nameRadioAbsUp);
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.TabStop = false;
		resources.ApplyResources(this.nameRadioAbsLeft, "nameRadioAbsLeft");
		this.nameRadioAbsLeft.Name = "nameRadioAbsLeft";
		this.nameRadioAbsLeft.UseVisualStyleBackColor = true;
		resources.ApplyResources(this.nameRadioLeft, "nameRadioLeft");
		this.nameRadioLeft.Name = "nameRadioLeft";
		this.nameRadioLeft.UseVisualStyleBackColor = true;
		resources.ApplyResources(this.nameRadioUp, "nameRadioUp");
		this.nameRadioUp.Name = "nameRadioUp";
		this.nameRadioUp.UseVisualStyleBackColor = true;
		resources.ApplyResources(this.nameRadioAbsUp, "nameRadioAbsUp");
		this.nameRadioAbsUp.Checked = true;
		this.nameRadioAbsUp.Name = "nameRadioAbsUp";
		this.nameRadioAbsUp.TabStop = true;
		this.nameRadioAbsUp.UseVisualStyleBackColor = true;
		resources.ApplyResources(this.groupBox2, "groupBox2");
		this.groupBox2.Controls.Add(this.staticRadio);
		this.groupBox2.Controls.Add(this.bothRadio);
		this.groupBox2.Name = "groupBox2";
		this.groupBox2.TabStop = false;
		resources.ApplyResources(this.staticRadio, "staticRadio");
		this.staticRadio.Name = "staticRadio";
		this.staticRadio.TabStop = true;
		this.staticRadio.UseVisualStyleBackColor = true;
		resources.ApplyResources(this.bothRadio, "bothRadio");
		this.bothRadio.Name = "bothRadio";
		this.bothRadio.TabStop = true;
		this.bothRadio.UseVisualStyleBackColor = true;
		resources.ApplyResources(this.label5, "label5");
		this.label5.Name = "label5";
		resources.ApplyResources(this.capGrid, "capGrid");
		this.capGrid.AllowUserToDeleteRows = false;
		this.capGrid.AllowUserToResizeColumns = false;
		this.capGrid.AllowUserToResizeRows = false;
		dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
		dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
		dataGridViewCellStyle3.Font = new System.Drawing.Font("MS UI Gothic", 9f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
		dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
		dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
		dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
		dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
		this.capGrid.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
		this.capGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
		this.capGrid.Name = "capGrid";
		this.capGrid.RowHeadersVisible = false;
		this.capGrid.RowTemplate.Height = 21;
		this.capGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
		this.capGrid.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(capGrid_CellEnter);
		this.capGrid.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(capGrid_CellValueChanged);
		this.capGrid.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(capGrid_EditingControlShowing);
		this.capGrid.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(capGrid_RowsAdded);
		resources.ApplyResources(this.btnDelCaption, "btnDelCaption");
		this.btnDelCaption.Name = "btnDelCaption";
		this.btnDelCaption.UseVisualStyleBackColor = true;
		this.btnDelCaption.Click += new System.EventHandler(btnDelCaption_Click);
		resources.ApplyResources(this.btnImport, "btnImport");
		this.btnImport.DialogResult = System.Windows.Forms.DialogResult.OK;
		this.btnImport.Name = "btnImport";
		this.btnImport.UseVisualStyleBackColor = true;
		this.btnImport.Click += new System.EventHandler(btnImport_Click);
		resources.ApplyResources(this.btnExport, "btnExport");
		this.btnExport.DialogResult = System.Windows.Forms.DialogResult.OK;
		this.btnExport.Name = "btnExport";
		this.btnExport.UseVisualStyleBackColor = true;
		this.btnExport.Click += new System.EventHandler(btnExport_Click);
		base.AcceptButton = this.btnOK;
		resources.ApplyResources(this, "$this");
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		base.CancelButton = this.btnCancel;
		base.Controls.Add(this.btnExport);
		base.Controls.Add(this.btnImport);
		base.Controls.Add(this.btnDelCaption);
		base.Controls.Add(this.capGrid);
		base.Controls.Add(this.label5);
		base.Controls.Add(this.groupBox2);
		base.Controls.Add(this.groupBox1);
		base.Controls.Add(this.label4);
		base.Controls.Add(this.button10);
		base.Controls.Add(this.button9);
		base.Controls.Add(this.button8);
		base.Controls.Add(this.button7);
		base.Controls.Add(this.button6);
		base.Controls.Add(this.button5);
		base.Controls.Add(this.button4);
		base.Controls.Add(this.button3);
		base.Controls.Add(this.button2);
		base.Controls.Add(this.button1);
		base.Controls.Add(this.setDefaultButton);
		base.Controls.Add(this.label3);
		base.Controls.Add(this.defaultButton);
		base.Controls.Add(this.btnOK);
		base.Controls.Add(this.btnCancel);
		base.Controls.Add(this.label2);
		base.Controls.Add(this.dataGridView2);
		base.Controls.Add(this.label1);
		base.Controls.Add(this.dataGridView1);
		base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		base.Name = "ColorForm";
		base.Load += new System.EventHandler(ColorForm_Load);
		base.Resize += new System.EventHandler(ColorForm_Resize);
		((System.ComponentModel.ISupportInitialize)this.dataGridView1).EndInit();
		((System.ComponentModel.ISupportInitialize)this.dataGridView2).EndInit();
		this.groupBox1.ResumeLayout(false);
		this.groupBox1.PerformLayout();
		this.groupBox2.ResumeLayout(false);
		this.groupBox2.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.capGrid).EndInit();
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
