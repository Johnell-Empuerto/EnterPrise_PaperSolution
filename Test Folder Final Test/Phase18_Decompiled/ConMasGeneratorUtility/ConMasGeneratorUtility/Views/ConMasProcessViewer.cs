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

public class ConMasProcessViewer : Form
{
	private JobSettingView _parentView;

	private IContainer components = null;

	private DataGridView ConMasProcessView;

	private Button OK;

	private Button Cancel;

	private TableLayoutPanel baseTable;

	private TableLayoutPanel buttonTable;

	private CheckBox allCheckBox;

	public ConMasProcessViewer(JobSettingView view)
	{
		InitializeComponent();
		_parentView = view;
		SetFormResouce();
		Type typeFromHandle = typeof(DataGridView);
		PropertyInfo property = typeFromHandle.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
		property.SetValue(ConMasProcessView, true, null);
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
			if (_parentView._conMasProcessList == null)
			{
				return;
			}
			DataGridViewColumn dataGridViewColumn = new DataGridViewCheckBoxColumn();
			dataGridViewColumn.Name = "CHECK";
			dataGridViewColumn.HeaderText = string.Empty;
			dataGridViewColumn.Width = 30;
			ConMasProcessView.Columns.Add(dataGridViewColumn);
			ConMasProcessView.Columns.Add("NAME", Resources.Name);
			ConMasProcessView.Columns["NAME"].Width = 180;
			ConMasProcessView.Columns["NAME"].ReadOnly = true;
			ConMasProcessView.Columns.Add("REMARK", Resources.Remark);
			ConMasProcessView.Columns["REMARK"].Width = 180;
			ConMasProcessView.Columns["REMARK"].ReadOnly = true;
			ConMasProcessView.Columns.Add("TYPE", Resources.ProcessType);
			ConMasProcessView.Columns["TYPE"].Width = 180;
			ConMasProcessView.Columns["TYPE"].ReadOnly = true;
			ConMasProcessView.Columns.Add("OBJECT", "");
			ConMasProcessView.Columns["OBJECT"].Visible = false;
			ConMasProcessView.Columns["OBJECT"].ReadOnly = true;
			int num = 0;
			ConMasProcessView.Rows.Clear();
			foreach (ProcessData conMasProcess in _parentView._conMasProcessList)
			{
				num = ConMasProcessView.Rows.Add(1);
				ConMasProcessView.Rows[num].Cells["CHECK"].Value = false;
				ConMasProcessView.Rows[num].Cells["NAME"].Value = conMasProcess.Name;
				if (conMasProcess.Type == "0")
				{
					ConMasProcessView.Rows[num].Cells["TYPE"].Value = Resources.ProcessType01;
				}
				else
				{
					ConMasProcessView.Rows[num].Cells["TYPE"].Value = Resources.ProcessType02;
				}
				ConMasProcessView.Rows[num].Cells["REMARK"].Value = conMasProcess.Remark;
				ConMasProcessView.Rows[num].Cells["OBJECT"].Value = conMasProcess;
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
			foreach (DataGridViewRow item in (IEnumerable)ConMasProcessView.Rows)
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
			_parentView.SelectedConMasProcesses = new List<ProcessData>();
			foreach (DataGridViewRow item in (IEnumerable)ConMasProcessView.Rows)
			{
				if ((bool)item.Cells["CHECK"].Value)
				{
					_parentView.SelectedConMasProcesses.Add(item.Cells["OBJECT"].Value as ProcessData);
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
			Text = Resources.ProcessSelect;
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
			ConMasProcessView.BeginEdit(selectAll: false);
			foreach (DataGridViewRow item in (IEnumerable)ConMasProcessView.Rows)
			{
				item.Cells[0].Value = allCheckBox.Checked;
			}
			ConMasProcessView.EndEdit();
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
		this.ConMasProcessView = new System.Windows.Forms.DataGridView();
		this.OK = new System.Windows.Forms.Button();
		this.Cancel = new System.Windows.Forms.Button();
		this.baseTable = new System.Windows.Forms.TableLayoutPanel();
		this.buttonTable = new System.Windows.Forms.TableLayoutPanel();
		this.allCheckBox = new System.Windows.Forms.CheckBox();
		((System.ComponentModel.ISupportInitialize)this.ConMasProcessView).BeginInit();
		this.baseTable.SuspendLayout();
		this.buttonTable.SuspendLayout();
		base.SuspendLayout();
		this.ConMasProcessView.AllowUserToAddRows = false;
		this.ConMasProcessView.AllowUserToDeleteRows = false;
		this.ConMasProcessView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
		this.ConMasProcessView.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasProcessView.Location = new System.Drawing.Point(3, 3);
		this.ConMasProcessView.MultiSelect = false;
		this.ConMasProcessView.Name = "ConMasProcessView";
		this.ConMasProcessView.RowHeadersVisible = false;
		this.ConMasProcessView.RowTemplate.Height = 21;
		this.ConMasProcessView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
		this.ConMasProcessView.Size = new System.Drawing.Size(567, 304);
		this.ConMasProcessView.TabIndex = 10;
		this.ConMasProcessView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(ConMasProcessView_CellClick);
		this.ConMasProcessView.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(ConMasProcessView_CellPainting);
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
		this.baseTable.Controls.Add(this.ConMasProcessView, 0, 0);
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
		this.allCheckBox.Location = new System.Drawing.Point(222, 331);
		this.allCheckBox.Name = "allCheckBox";
		this.allCheckBox.Size = new System.Drawing.Size(15, 14);
		this.allCheckBox.TabIndex = 4;
		this.allCheckBox.UseVisualStyleBackColor = true;
		this.allCheckBox.Visible = false;
		this.allCheckBox.CheckedChanged += new System.EventHandler(allCheckBox_CheckedChanged);
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		base.CancelButton = this.Cancel;
		base.ClientSize = new System.Drawing.Size(579, 367);
		base.Controls.Add(this.allCheckBox);
		base.Controls.Add(this.baseTable);
		base.Name = "ConMasProcessViewer";
		base.Padding = new System.Windows.Forms.Padding(3);
		base.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
		this.Text = "ConMasProcessViewer";
		base.Load += new System.EventHandler(ConMasProcessViewer_Load);
		((System.ComponentModel.ISupportInitialize)this.ConMasProcessView).EndInit();
		this.baseTable.ResumeLayout(false);
		this.buttonTable.ResumeLayout(false);
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
