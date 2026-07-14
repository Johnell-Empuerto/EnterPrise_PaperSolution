using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using iReporterExcelAddInCommon.Properties;
using iReporterExcelAddInCommon.Tables;

namespace iReporterExcelAddInCommon;

public class TableForm : Form
{
	public delegate Table ReadTableEventHandler();

	public delegate void WriteTableEventHandler(Table table);

	private class CtrlLine
	{
		protected readonly List<Control> boxes = new List<Control>();

		protected readonly List<Label> labels = new List<Label>();

		protected int Index { get; }

		public bool Visible => boxes.FirstOrDefault()?.Visible ?? false;

		private bool IsOrg => Index == 1;

		protected CtrlLine(CtrlLine src = null)
		{
			Index = (src?.Index ?? 0) + 1;
			if (src != null)
			{
				boxes.AddRange(src.boxes.Select(Util.CreateCopy));
				labels.AddRange(src.labels.Select(Util.CreateCopy));
			}
		}

		protected void SetVisible(bool show)
		{
			foreach (Control box in boxes)
			{
				box.Visible = show || IsOrg;
				box.TabStop = show && IsTabStop(box);
				if (!show)
				{
					box.Text = "";
				}
			}
			foreach (Label label in labels)
			{
				label.Visible = show || IsOrg;
			}
		}

		private bool IsTabStop(Control c)
		{
			if (c is TextBox textBox)
			{
				return !textBox.ReadOnly;
			}
			if (c is ComboBox comboBox)
			{
				return comboBox.Enabled;
			}
			return false;
		}

		public void Hide()
		{
			SetVisible(show: false);
		}
	}

	private class ColControls : CtrlLine
	{
		public const int Margin = -1;

		public string Name => boxes[0].Text;

		public TextBox NameBox => (TextBox)boxes[0];

		public TextBox Key => (TextBox)boxes[1];

		public ColType? Type => ((boxes[2] as ComboBox)?.SelectedItem as ColTypeItem)?.Type;

		public ColControls(TextBox name, TextBox key, TextBox type, TextBox kind, Label label)
		{
			boxes.Add(name);
			boxes.Add(key);
			boxes.Add(type);
			boxes.Add(kind);
			labels.Add(label);
			key.MaxLength = 5;
		}

		public ColControls(ColControls src)
			: base(src)
		{
			for (int i = 0; i < boxes.Count; i++)
			{
				boxes[i].Left = src.boxes[i].Right + -1;
			}
			labels[0].TextResize(base.Index.ToString());
			labels[0].Left = (boxes[0].Left + boxes[0].Right - labels[0].Width) / 2;
		}

		public void Show(TableCol col)
		{
			boxes[0].Text = (col.IsNotOutput ? col.Type.ToText() : col.Name);
			Key.Text = (col.IsNotOutput ? col.Type.ToText() : col.Key);
			if (!col.IsNotOutput)
			{
				Key.MaxLength = 5;
			}
			Control control = null;
			TextBox textBox = boxes[2] as TextBox;
			ComboBox comboBox = boxes[2] as ComboBox;
			if (col.IsCalc)
			{
				if (textBox != null)
				{
					control = textBox;
					List<Control> list = boxes;
					ComboBox obj = new ComboBox
					{
						DropDownStyle = ComboBoxStyle.DropDownList
					};
					comboBox = obj;
					list[2] = obj;
					comboBox.Items.Add(new ColTypeItem
					{
						Type = ColType.numeric
					});
					comboBox.Items.Add(new ColTypeItem
					{
						Type = ColType.date
					});
					comboBox.Items.Add(new ColTypeItem
					{
						Type = ColType.interval
					});
					comboBox.Items.Add(new ColTypeItem
					{
						Type = ColType.text
					});
				}
			}
			else if (comboBox != null)
			{
				control = comboBox;
				comboBox = null;
				List<Control> list2 = boxes;
				TextBox obj2 = new TextBox
				{
					ReadOnly = true,
					TabStop = false
				};
				textBox = obj2;
				list2[2] = obj2;
			}
			if (control != null)
			{
				boxes[2].Location = control.Location;
				boxes[2].Size = control.Size;
				boxes[2].TabIndex = control.TabIndex;
				control.Parent.Controls.Add(boxes[2]);
				control.Parent.Controls.Remove(control);
			}
			if (textBox != null)
			{
				textBox.Text = col.Type.ToText();
			}
			if (comboBox != null)
			{
				comboBox.Text = col.Type.ToText();
			}
			boxes[3].Text = col.Kind;
			TextBox nameBox = NameBox;
			bool readOnly = (Key.ReadOnly = col.IsNotOutput);
			nameBox.ReadOnly = readOnly;
			SetVisible(show: true);
		}
	}

	private class RowControls : CtrlLine
	{
		public const int Margin = -1;

		public string Name => boxes[0].Text;

		private int Right { get; }

		public RowControls(TextBox name, Label lblName, Label lblIndex, int right)
		{
			boxes.Add(name);
			labels.Add(lblName);
			labels.Add(lblIndex);
			Right = right;
			labels[1].TextResize("1");
		}

		public RowControls(RowControls src)
			: base(src)
		{
			for (int i = 0; i < boxes.Count; i++)
			{
				boxes[i].Top = src.boxes[i].Bottom + -1;
			}
			Right = src.Right;
			labels[1].TextResize(base.Index.ToString());
			labels[0].Top = (labels[1].Top = src.labels[0].Top + boxes[0].Top - src.boxes[0].Top);
			labels[1].Left = (boxes[0].Right + Right - labels[1].Width) / 2;
		}

		public void Show(TableRow row)
		{
			boxes[0].Text = row.Name;
			SetVisible(show: true);
		}
	}

	private readonly List<ColControls> cols = new List<ColControls>();

	private readonly List<RowControls> rows = new List<RowControls>();

	private readonly List<List<TextBox>> items = new List<List<TextBox>>();

	private bool stopTableChangedEvent;

	private IContainer components;

	private Label label17;

	private ListBox tableListBox;

	private Button btnRead;

	private Label label1;

	private TextBox boxTableNo;

	private Label label2;

	private TextBox boxTableName;

	private CheckBox checkOutput;

	private Label label3;

	private Panel panel1;

	private TextBox boxItem00;

	private Label lblColNo;

	private TextBox boxRowName;

	private Label lblRowName;

	private Label lblRowNo;

	private TextBox boxColKind;

	private Label label6;

	private TextBox boxColType;

	private Label label5;

	private TextBox boxColKey;

	private Label label4;

	private TextBox boxColName;

	private Label lblColName;

	private Button btnSave;

	private Button btmCancel;

	private Button btnDelete;

	private Button btnReload;

	private int ColCount => cols.Count;

	private int RowCount => rows.Count;

	private IEnumerable<ColControls> Cols => cols;

	private IEnumerable<RowControls> Rows => rows;

	private IEnumerable<Table> Tables => tableListBox.Items.OfType<Table>();

	public event ReadTableEventHandler ReadTable;

	public event WriteTableEventHandler WriteTable;

	public event EventHandler Reload;

	public TableForm()
	{
		InitializeComponent();
		lblRowNo.Left = (boxColKind.Left + boxColKind.Right - lblRowNo.Width) / 2;
		cols.Add(new ColControls(boxColName, boxColKey, boxColType, boxColKind, lblRowNo));
		rows.Add(new RowControls(boxRowName, lblRowName, lblColNo, boxItem00.Left));
		items.Add(new List<TextBox> { boxItem00 });
		MinimumSize = base.Size;
		btnDelete.Text = Resources.DELETE;
	}

	public void SetupTableList(IEnumerable<Table> tables)
	{
		tableListBox.BeginUpdate();
		tableListBox.Items.Clear();
		ListBox.ObjectCollection objectCollection = tableListBox.Items;
		object[] array = tables.ToArray();
		objectCollection.AddRange(array);
		tableListBox_SelectedIndexChanged(tableListBox, EventArgs.Empty);
		tableListBox.EndUpdate();
	}

	private void tableListBox_MeasureItem(object sender, MeasureItemEventArgs e)
	{
		e.ItemHeight *= InputForm2.GetWindowScale;
		e.ItemHeight /= 100;
	}

	private void tableListBox_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (stopTableChangedEvent)
		{
			return;
		}
		Table table = tableListBox.SelectedItem as Table;
		if (table == null)
		{
			table = Table.Empty;
			btnDelete.Enabled = false;
		}
		else
		{
			string text = table.Check();
			if (text != null)
			{
				ShowMessage(Resources.TABLE_INVALID + Environment.NewLine + Environment.NewLine + text);
				table = Table.Empty;
			}
			btnDelete.Enabled = true;
		}
		btnSave.Enabled = table != Table.Empty;
		try
		{
			SuspendLayout();
			boxTableNo.Text = table.No;
			boxTableName.Text = table.Name;
			checkOutput.Checked = table.OutputCollabo;
			foreach (TableCol col in table.Cols)
			{
				while (ColCount < col.Index)
				{
					cols.Add(new ColControls(Cols.Last()));
				}
				cols[col.Index - 1].Show(col);
			}
			foreach (ColControls item in Cols.Skip(table.ColCount))
			{
				item.Hide();
			}
			foreach (TableRow row in table.Rows)
			{
				while (RowCount < row.Index)
				{
					rows.Add(new RowControls(Rows.Last()));
				}
				rows[row.Index - 1].Show(row);
			}
			foreach (RowControls item2 in Rows.Skip(table.RowCount))
			{
				item2.Hide();
			}
			while (items[0].Count < ColCount)
			{
				foreach (List<TextBox> item3 in items)
				{
					TextBox textBox = item3.Last();
					TextBox textBox2 = textBox.CreateCopy();
					textBox2.Left = textBox.Right + -1;
					item3.Add(textBox2);
				}
			}
			while (items.Count < RowCount)
			{
				List<TextBox> list = items.Last();
				items.Add(new List<TextBox>());
				List<TextBox> list2 = items.Last();
				foreach (TextBox item4 in list)
				{
					TextBox textBox3 = item4.CreateCopy();
					textBox3.Top = item4.Bottom + -1;
					list2.Add(textBox3);
				}
			}
			for (int i = 0; i < items.Count; i++)
			{
				int j = ((table.RowCount > i) ? table.ColCount : 0);
				if (i == 0 && j == 0)
				{
					items[i][j++].Text = "";
				}
				for (; j < items[i].Count; j++)
				{
					items[i][j].Visible = false;
				}
			}
			foreach (TableRow row2 in table.Rows)
			{
				Dictionary<int, TableItem> dictionary = row2.Items.ToDictionary((TableItem tableItem) => tableItem.Col.Index);
				List<TextBox> list3 = items[row2.Index - 1];
				for (int num = 0; num < list3.Count; num++)
				{
					TextBox textBox4 = list3[num];
					dictionary.TryGetValue(num + 1, out var value);
					if (value == null)
					{
						textBox4.Visible = false;
						continue;
					}
					textBox4.Text = value.ClusterIndex + ".  " + value.ClusterName;
					textBox4.Visible = true;
				}
			}
		}
		catch (Exception ex)
		{
			Util.Catched(ex);
		}
		finally
		{
			ResumeLayout(performLayout: true);
		}
	}

	private void btnRead_Click(object sender, EventArgs e)
	{
		try
		{
			Table table = this.ReadTable?.Invoke();
			if (table == null)
			{
				ShowMessage(Resources.TABLE_NO_CLUSTERS);
				return;
			}
			string text = table.Check();
			if (text != null)
			{
				ShowMessage(text);
				return;
			}
			table.No = (Tables.Select((Table t) => t.No.TryParseInt()).MaxOr(0) + 1).ToString();
			tableListBox.Items.Add(table);
			tableListBox.SelectedItem = table;
		}
		catch (Exception ex)
		{
			Util.Catched(ex);
		}
	}

	private void ShowMessage(string text)
	{
		MessageBox.Show(text);
	}

	private void btmCancel_Click(object sender, EventArgs e)
	{
		Close();
	}

	private void btnSave_Click(object sender, EventArgs e)
	{
		try
		{
			Table table = tableListBox.SelectedItem as Table;
			if (table == null)
			{
				ShowMessage(Resources.NO_TABLE);
				return;
			}
			if (Tables.Any((Table t) => t != table && t.No == boxTableNo.Text))
			{
				ShowMessage(Resources.TABLE_NO_OVERLAP);
				return;
			}
			int.TryParse(boxTableNo.Text, out var result);
			if (result < 1)
			{
				ShowMessage(Resources.TABLE_NO_INVALID);
				return;
			}
			if (string.IsNullOrEmpty(boxTableName.Text))
			{
				ShowMessage(Resources.TABLE_NAME_EMPTY);
				return;
			}
			HashSet<string> hashSet = new HashSet<string>();
			foreach (ColControls item in from c in Cols.TakeWhile((ColControls c) => c.Key.Visible)
				where !c.Key.ReadOnly
				select c)
			{
				if (string.IsNullOrEmpty(item.Key.Text))
				{
					ShowMessage(Resources.TABLE_COL_KEY_EMPTY);
					return;
				}
				if (item.Key.Text.Any((char c) => ('a' > c || c > 'z') && ('A' > c || c > 'Z') && ('0' > c || c > '9') && c != '_'))
				{
					ShowMessage(Resources.TABLE_COL_KEY_INVALID);
					return;
				}
				if (!hashSet.Add(item.Key.Text))
				{
					ShowMessage(Resources.TABLE_COL_KEY_OVERLAP);
					return;
				}
				if (string.IsNullOrEmpty(item.Name))
				{
					ShowMessage(Resources.TABLE_COL_NAME_EMPTY);
					return;
				}
			}
			table.No = boxTableNo.Text;
			table.Name = boxTableName.Text;
			table.OutputCollabo = checkOutput.Checked;
			foreach (TableRow row in table.Rows)
			{
				row.Name = rows[row.Index - 1].Name;
			}
			foreach (TableCol col in table.Cols)
			{
				if (col.IsNotOutput)
				{
					string name = (col.Key = col.Type.ToString());
					col.Name = name;
				}
				else
				{
					col.Name = cols[col.Index - 1].Name;
					col.Key = cols[col.Index - 1].Key.Text;
					col.Type = cols[col.Index - 1].Type ?? col.Type;
				}
			}
			this.WriteTable?.Invoke(table);
			table.IsReal = true;
			stopTableChangedEvent = true;
			tableListBox.Items[tableListBox.SelectedIndex] = table;
			stopTableChangedEvent = false;
			ShowMessage(Resources.TABLE_SAVED);
		}
		catch (Exception ex)
		{
			Util.Catched(ex);
		}
	}

	private void btnDelete_Click(object sender, EventArgs e)
	{
		try
		{
			if (!(tableListBox.SelectedItem is Table table))
			{
				ShowMessage(Resources.NO_TABLE);
				return;
			}
			if (table.IsReal)
			{
				foreach (TableItem item in table.Rows.SelectMany((TableRow r) => r.Items))
				{
					item.RemoveTableInfos();
				}
				this.WriteTable?.Invoke(table);
			}
			tableListBox.Items.Remove(table);
		}
		catch (Exception ex)
		{
			Util.Catched(ex);
		}
	}

	private void btnReload_Click(object sender, EventArgs e)
	{
		this.Reload?.Invoke(sender, e);
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
		System.ComponentModel.ComponentResourceManager componentResourceManager = new System.ComponentModel.ComponentResourceManager(typeof(iReporterExcelAddInCommon.TableForm));
		this.label17 = new System.Windows.Forms.Label();
		this.tableListBox = new System.Windows.Forms.ListBox();
		this.btnRead = new System.Windows.Forms.Button();
		this.label1 = new System.Windows.Forms.Label();
		this.boxTableNo = new System.Windows.Forms.TextBox();
		this.label2 = new System.Windows.Forms.Label();
		this.boxTableName = new System.Windows.Forms.TextBox();
		this.checkOutput = new System.Windows.Forms.CheckBox();
		this.label3 = new System.Windows.Forms.Label();
		this.panel1 = new System.Windows.Forms.Panel();
		this.boxItem00 = new System.Windows.Forms.TextBox();
		this.lblColNo = new System.Windows.Forms.Label();
		this.boxRowName = new System.Windows.Forms.TextBox();
		this.lblRowName = new System.Windows.Forms.Label();
		this.lblRowNo = new System.Windows.Forms.Label();
		this.boxColKind = new System.Windows.Forms.TextBox();
		this.label6 = new System.Windows.Forms.Label();
		this.boxColType = new System.Windows.Forms.TextBox();
		this.label5 = new System.Windows.Forms.Label();
		this.boxColKey = new System.Windows.Forms.TextBox();
		this.label4 = new System.Windows.Forms.Label();
		this.boxColName = new System.Windows.Forms.TextBox();
		this.lblColName = new System.Windows.Forms.Label();
		this.btnSave = new System.Windows.Forms.Button();
		this.btmCancel = new System.Windows.Forms.Button();
		this.btnDelete = new System.Windows.Forms.Button();
		this.btnReload = new System.Windows.Forms.Button();
		this.panel1.SuspendLayout();
		base.SuspendLayout();
		componentResourceManager.ApplyResources(this.label17, "label17");
		this.label17.Name = "label17";
		componentResourceManager.ApplyResources(this.tableListBox, "tableListBox");
		this.tableListBox.DisplayMember = "CultureName";
		this.tableListBox.FormattingEnabled = true;
		this.tableListBox.Name = "tableListBox";
		this.tableListBox.MeasureItem += new System.Windows.Forms.MeasureItemEventHandler(tableListBox_MeasureItem);
		this.tableListBox.SelectedIndexChanged += new System.EventHandler(tableListBox_SelectedIndexChanged);
		componentResourceManager.ApplyResources(this.btnRead, "btnRead");
		this.btnRead.Name = "btnRead";
		this.btnRead.UseVisualStyleBackColor = true;
		this.btnRead.Click += new System.EventHandler(btnRead_Click);
		componentResourceManager.ApplyResources(this.label1, "label1");
		this.label1.Name = "label1";
		componentResourceManager.ApplyResources(this.boxTableNo, "boxTableNo");
		this.boxTableNo.Name = "boxTableNo";
		componentResourceManager.ApplyResources(this.label2, "label2");
		this.label2.Name = "label2";
		componentResourceManager.ApplyResources(this.boxTableName, "boxTableName");
		this.boxTableName.Name = "boxTableName";
		componentResourceManager.ApplyResources(this.checkOutput, "checkOutput");
		this.checkOutput.Name = "checkOutput";
		this.checkOutput.UseVisualStyleBackColor = true;
		componentResourceManager.ApplyResources(this.label3, "label3");
		this.label3.Name = "label3";
		componentResourceManager.ApplyResources(this.panel1, "panel1");
		this.panel1.Controls.Add(this.boxItem00);
		this.panel1.Controls.Add(this.lblColNo);
		this.panel1.Controls.Add(this.boxRowName);
		this.panel1.Controls.Add(this.lblRowName);
		this.panel1.Controls.Add(this.lblRowNo);
		this.panel1.Controls.Add(this.boxColKind);
		this.panel1.Controls.Add(this.label6);
		this.panel1.Controls.Add(this.boxColType);
		this.panel1.Controls.Add(this.label5);
		this.panel1.Controls.Add(this.boxColKey);
		this.panel1.Controls.Add(this.label4);
		this.panel1.Controls.Add(this.boxColName);
		this.panel1.Controls.Add(this.lblColName);
		this.panel1.Name = "panel1";
		componentResourceManager.ApplyResources(this.boxItem00, "boxItem00");
		this.boxItem00.Name = "boxItem00";
		this.boxItem00.ReadOnly = true;
		this.boxItem00.TabStop = false;
		componentResourceManager.ApplyResources(this.lblColNo, "lblColNo");
		this.lblColNo.Name = "lblColNo";
		componentResourceManager.ApplyResources(this.boxRowName, "boxRowName");
		this.boxRowName.Name = "boxRowName";
		componentResourceManager.ApplyResources(this.lblRowName, "lblRowName");
		this.lblRowName.Name = "lblRowName";
		componentResourceManager.ApplyResources(this.lblRowNo, "lblRowNo");
		this.lblRowNo.Name = "lblRowNo";
		componentResourceManager.ApplyResources(this.boxColKind, "boxColKind");
		this.boxColKind.Name = "boxColKind";
		this.boxColKind.ReadOnly = true;
		componentResourceManager.ApplyResources(this.label6, "label6");
		this.label6.Name = "label6";
		componentResourceManager.ApplyResources(this.boxColType, "boxColType");
		this.boxColType.Name = "boxColType";
		this.boxColType.ReadOnly = true;
		componentResourceManager.ApplyResources(this.label5, "label5");
		this.label5.Name = "label5";
		componentResourceManager.ApplyResources(this.boxColKey, "boxColKey");
		this.boxColKey.Name = "boxColKey";
		componentResourceManager.ApplyResources(this.label4, "label4");
		this.label4.Name = "label4";
		componentResourceManager.ApplyResources(this.boxColName, "boxColName");
		this.boxColName.Name = "boxColName";
		componentResourceManager.ApplyResources(this.lblColName, "lblColName");
		this.lblColName.Name = "lblColName";
		componentResourceManager.ApplyResources(this.btnSave, "btnSave");
		this.btnSave.Name = "btnSave";
		this.btnSave.UseVisualStyleBackColor = true;
		this.btnSave.Click += new System.EventHandler(btnSave_Click);
		componentResourceManager.ApplyResources(this.btmCancel, "btmCancel");
		this.btmCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		this.btmCancel.Name = "btmCancel";
		this.btmCancel.UseVisualStyleBackColor = true;
		this.btmCancel.Click += new System.EventHandler(btmCancel_Click);
		componentResourceManager.ApplyResources(this.btnDelete, "btnDelete");
		this.btnDelete.Name = "btnDelete";
		this.btnDelete.UseVisualStyleBackColor = true;
		this.btnDelete.Click += new System.EventHandler(btnDelete_Click);
		componentResourceManager.ApplyResources(this.btnReload, "btnReload");
		this.btnReload.Name = "btnReload";
		this.btnReload.UseVisualStyleBackColor = true;
		this.btnReload.Click += new System.EventHandler(btnReload_Click);
		base.AcceptButton = this.btnSave;
		componentResourceManager.ApplyResources(this, "$this");
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		base.CancelButton = this.btmCancel;
		base.Controls.Add(this.btnReload);
		base.Controls.Add(this.btnDelete);
		base.Controls.Add(this.btmCancel);
		base.Controls.Add(this.btnSave);
		base.Controls.Add(this.panel1);
		base.Controls.Add(this.label3);
		base.Controls.Add(this.checkOutput);
		base.Controls.Add(this.boxTableName);
		base.Controls.Add(this.label2);
		base.Controls.Add(this.boxTableNo);
		base.Controls.Add(this.label1);
		base.Controls.Add(this.btnRead);
		base.Controls.Add(this.tableListBox);
		base.Controls.Add(this.label17);
		base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		base.Name = "TableForm";
		this.panel1.ResumeLayout(false);
		this.panel1.PerformLayout();
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
