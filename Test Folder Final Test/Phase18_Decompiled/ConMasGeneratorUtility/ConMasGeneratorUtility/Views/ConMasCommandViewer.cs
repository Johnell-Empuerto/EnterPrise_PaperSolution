using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using ConMasGeneratorLib.Data;
using ConMasGeneratorUtility.Properties;

namespace ConMasGeneratorUtility.Views;

public class ConMasCommandViewer : Form
{
	private JobSettingView _parentView;

	private string _processType;

	private IContainer components = null;

	private DataGridView ConMasCommandView;

	private Button OK;

	private Button Cancel;

	private TableLayoutPanel baseTable;

	private TableLayoutPanel buttonTable;

	private CheckBox allCheckBox;

	public ConMasCommandViewer(JobSettingView view, string processType)
	{
		InitializeComponent();
		_parentView = view;
		_processType = processType;
		SetFormResouce();
		Type typeFromHandle = typeof(DataGridView);
		PropertyInfo property = typeFromHandle.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
		property.SetValue(ConMasCommandView, true, null);
	}

	private void ConMasProcessViewer_Load(object sender, EventArgs e)
	{
		try
		{
			SetGrid();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void SetGrid()
	{
		try
		{
			if (_parentView._conMasCommandList == null)
			{
				return;
			}
			DataGridViewColumn dataGridViewColumn = new DataGridViewCheckBoxColumn();
			dataGridViewColumn.Name = "CHECK";
			dataGridViewColumn.HeaderText = string.Empty;
			dataGridViewColumn.Width = 30;
			ConMasCommandView.Columns.Add(dataGridViewColumn);
			ConMasCommandView.Columns.Add("NAME", Resources.Name);
			ConMasCommandView.Columns["NAME"].Width = 180;
			ConMasCommandView.Columns["NAME"].ReadOnly = true;
			ConMasCommandView.Columns.Add("URL", Resources.Url);
			ConMasCommandView.Columns["URL"].Width = 250;
			ConMasCommandView.Columns["URL"].ReadOnly = true;
			ConMasCommandView.Columns.Add("IS_RESPONSE", Resources.IsResponse);
			ConMasCommandView.Columns["IS_RESPONSE"].Width = 100;
			ConMasCommandView.Columns["IS_RESPONSE"].ReadOnly = true;
			ConMasCommandView.Columns.Add("RESPONSE_FILE_NAME", Resources.ResponseFileName);
			ConMasCommandView.Columns["RESPONSE_FILE_NAME"].Width = 180;
			ConMasCommandView.Columns["RESPONSE_FILE_NAME"].ReadOnly = true;
			ConMasCommandView.Columns.Add("PATH", Resources.ExePath);
			ConMasCommandView.Columns["PATH"].Width = 250;
			ConMasCommandView.Columns["PATH"].ReadOnly = true;
			ConMasCommandView.Columns.Add("OBJECT", "");
			ConMasCommandView.Columns["OBJECT"].Visible = false;
			ConMasCommandView.Columns["OBJECT"].ReadOnly = true;
			if (_processType == "0")
			{
				ConMasCommandView.Columns["PATH"].Visible = false;
			}
			else
			{
				ConMasCommandView.Columns["OBJECT"].Visible = false;
				ConMasCommandView.Columns["URL"].Visible = false;
				ConMasCommandView.Columns["IS_RESPONSE"].Visible = false;
				ConMasCommandView.Columns["RESPONSE_FILE_NAME"].Visible = false;
			}
			int num = 0;
			ConMasCommandView.Rows.Clear();
			foreach (Command conMasCommand in _parentView._conMasCommandList)
			{
				num = ConMasCommandView.Rows.Add(1);
				ConMasCommandView.Rows[num].Cells["CHECK"].Value = false;
				ConMasCommandView.Rows[num].Cells["NAME"].Value = conMasCommand.Name;
				if (conMasCommand.IsResponse)
				{
					ConMasCommandView.Rows[num].Cells["IS_RESPONSE"].Value = Resources.ResponseOK;
				}
				else
				{
					ConMasCommandView.Rows[num].Cells["IS_RESPONSE"].Value = Resources.ResponseNO;
				}
				ConMasCommandView.Rows[num].Cells["RESPONSE_FILE_NAME"].Value = conMasCommand.ResponseFileName;
				ConMasCommandView.Rows[num].Cells["PATH"].Value = conMasCommand.Path;
				ConMasCommandView.Rows[num].Cells["OBJECT"].Value = conMasCommand;
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void OK_Click(object sender, EventArgs e)
	{
		try
		{
			bool flag = false;
			foreach (DataGridViewRow item in (IEnumerable)ConMasCommandView.Rows)
			{
				if ((bool)item.Cells["CHECK"].Value)
				{
					flag = true;
					break;
				}
			}
			if (flag)
			{
				SelectProcess();
				base.DialogResult = DialogResult.OK;
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void SelectProcess()
	{
		try
		{
			_parentView.SelectedConMasCommands = new List<Command>();
			foreach (DataGridViewRow item in (IEnumerable)ConMasCommandView.Rows)
			{
				if ((bool)item.Cells["CHECK"].Value)
				{
					_parentView.SelectedConMasCommands.Add(item.Cells["OBJECT"].Value as Command);
				}
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void SetFormResouce()
	{
		try
		{
			OK.Text = Resources.Add;
			Cancel.Text = Resources.Cancel;
			if (_processType == "0")
			{
				Text = Resources.RestCommandSelect;
			}
			else
			{
				Text = Resources.ExeCommandSelect;
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void ConMasProcessView_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
	{
		try
		{
			if (e.ColumnIndex != 0 || e.RowIndex != -1)
			{
				return;
			}
			using Bitmap bitmap = new Bitmap(100, 100);
			using (Graphics graphics = Graphics.FromImage(bitmap))
			{
				graphics.Clear(Color.Transparent);
			}
			Point point = new Point((bitmap.Width - allCheckBox.Width) / 2, (bitmap.Height - allCheckBox.Height) / 2);
			if (point.X < 0)
			{
				point.X = 0;
			}
			if (point.Y < 0)
			{
				point.Y = 0;
			}
			allCheckBox.DrawToBitmap(bitmap, new Rectangle(point.X, point.Y, bitmap.Width, bitmap.Height));
			int num = (e.CellBounds.Width - bitmap.Width) / 2;
			int num2 = (e.CellBounds.Height - bitmap.Height) / 2;
			Point point2 = new Point(e.CellBounds.Left + num, e.CellBounds.Top + num2);
			e.Paint(e.ClipBounds, e.PaintParts);
			e.Graphics.DrawImage(bitmap, point2);
			e.Handled = true;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void ConMasProcessView_CellClick(object sender, DataGridViewCellEventArgs e)
	{
		try
		{
			if (e.ColumnIndex == 0 && e.RowIndex == -1)
			{
				allCheckBox.Checked = !allCheckBox.Checked;
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void allCheckBox_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			ConMasCommandView.BeginEdit(selectAll: false);
			foreach (DataGridViewRow item in (IEnumerable)ConMasCommandView.Rows)
			{
				item.Cells[0].Value = allCheckBox.Checked;
			}
			ConMasCommandView.EndEdit();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
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
		this.ConMasCommandView = new System.Windows.Forms.DataGridView();
		this.OK = new System.Windows.Forms.Button();
		this.Cancel = new System.Windows.Forms.Button();
		this.baseTable = new System.Windows.Forms.TableLayoutPanel();
		this.buttonTable = new System.Windows.Forms.TableLayoutPanel();
		this.allCheckBox = new System.Windows.Forms.CheckBox();
		((System.ComponentModel.ISupportInitialize)this.ConMasCommandView).BeginInit();
		this.baseTable.SuspendLayout();
		this.buttonTable.SuspendLayout();
		base.SuspendLayout();
		this.ConMasCommandView.AllowUserToAddRows = false;
		this.ConMasCommandView.AllowUserToDeleteRows = false;
		this.ConMasCommandView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
		this.ConMasCommandView.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasCommandView.Location = new System.Drawing.Point(3, 3);
		this.ConMasCommandView.Name = "ConMasCommandView";
		this.ConMasCommandView.RowHeadersVisible = false;
		this.ConMasCommandView.RowTemplate.Height = 21;
		this.ConMasCommandView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
		this.ConMasCommandView.Size = new System.Drawing.Size(567, 304);
		this.ConMasCommandView.TabIndex = 10;
		this.ConMasCommandView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(ConMasProcessView_CellClick);
		this.ConMasCommandView.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(ConMasProcessView_CellPainting);
		this.OK.Dock = System.Windows.Forms.DockStyle.Left;
		this.OK.Location = new System.Drawing.Point(3, 3);
		this.OK.Name = "OK";
		this.OK.Size = new System.Drawing.Size(94, 39);
		this.OK.TabIndex = 11;
		this.OK.Text = "button1";
		this.OK.UseVisualStyleBackColor = true;
		this.OK.Click += new System.EventHandler(OK_Click);
		this.Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		this.Cancel.Dock = System.Windows.Forms.DockStyle.Right;
		this.Cancel.Location = new System.Drawing.Point(470, 3);
		this.Cancel.Name = "Cancel";
		this.Cancel.Size = new System.Drawing.Size(94, 39);
		this.Cancel.TabIndex = 12;
		this.Cancel.Text = "button2";
		this.Cancel.UseVisualStyleBackColor = true;
		this.baseTable.ColumnCount = 1;
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50f));
		this.baseTable.Controls.Add(this.ConMasCommandView, 0, 0);
		this.baseTable.Controls.Add(this.buttonTable, 0, 1);
		this.baseTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.baseTable.Location = new System.Drawing.Point(3, 3);
		this.baseTable.Name = "baseTable";
		this.baseTable.RowCount = 2;
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 86.10354f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 13.89646f));
		this.baseTable.Size = new System.Drawing.Size(573, 361);
		this.baseTable.TabIndex = 3;
		this.buttonTable.ColumnCount = 2;
		this.buttonTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50f));
		this.buttonTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 286f));
		this.buttonTable.Controls.Add(this.OK, 0, 0);
		this.buttonTable.Controls.Add(this.Cancel, 1, 0);
		this.buttonTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.buttonTable.Location = new System.Drawing.Point(3, 313);
		this.buttonTable.Name = "buttonTable";
		this.buttonTable.RowCount = 1;
		this.buttonTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
		this.buttonTable.Size = new System.Drawing.Size(567, 45);
		this.buttonTable.TabIndex = 1;
		this.allCheckBox.AutoSize = true;
		this.allCheckBox.Location = new System.Drawing.Point(210, 327);
		this.allCheckBox.Name = "allCheckBox";
		this.allCheckBox.Size = new System.Drawing.Size(15, 14);
		this.allCheckBox.TabIndex = 5;
		this.allCheckBox.UseVisualStyleBackColor = true;
		this.allCheckBox.Visible = false;
		this.allCheckBox.CheckedChanged += new System.EventHandler(allCheckBox_CheckedChanged);
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		base.CancelButton = this.Cancel;
		base.ClientSize = new System.Drawing.Size(579, 367);
		base.Controls.Add(this.allCheckBox);
		base.Controls.Add(this.baseTable);
		base.Name = "ConMasCommandViewer";
		base.Padding = new System.Windows.Forms.Padding(3);
		base.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
		this.Text = "ConMasProcessViewer";
		base.Load += new System.EventHandler(ConMasProcessViewer_Load);
		((System.ComponentModel.ISupportInitialize)this.ConMasCommandView).EndInit();
		this.baseTable.ResumeLayout(false);
		this.buttonTable.ResumeLayout(false);
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
