using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using ConMasGeneratorLib.Data;
using ConMasGeneratorUtility.Properties;

namespace ConMasGeneratorUtility.Views;

public class ProcessView : SettingViewBase
{
	private TreeNode _node;

	private ProcessData _processData;

	private IContainer components = null;

	private TextBox nameText;

	private TextBox remarkText;

	private TableLayoutPanel tableLayoutPanel1;

	private Label remarkLabel;

	private Label typeLabel;

	private Label nameLabel;

	private Label processType;

	public ProcessView(TreeNode node)
	{
		InitializeComponent();
		SetFormResouce();
		_node = node;
		_processData = node.Tag as ProcessData;
	}

	public override void SaveTemporary()
	{
		try
		{
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	public override bool IsValid()
	{
		try
		{
			bool flag = false;
			return true;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void ProcessView_Load(object sender, EventArgs e)
	{
		try
		{
			SetData();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void SetData()
	{
		try
		{
			nameText.Text = _processData.Name;
			remarkText.Text = _processData.Remark;
			if (_processData.Type == "1")
			{
				processType.Text = Resources.ProcessType02;
			}
			else
			{
				processType.Text = Resources.ProcessType01;
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void nameText_Leave(object sender, EventArgs e)
	{
		try
		{
			if (string.IsNullOrEmpty(nameText.Text.Trim()))
			{
				MessageBox.Show(Resources.NameEmptyMessage, "WARNING");
				nameText.Focus();
			}
			else
			{
				_processData.Name = nameText.Text;
				_node.Text = nameText.Text;
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void remarkText_Leave(object sender, EventArgs e)
	{
		try
		{
			_processData.Remark = remarkText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void SetFormResouce()
	{
		try
		{
			nameLabel.Text = Resources.Name;
			typeLabel.Text = Resources.ProcessType;
			remarkLabel.Text = Resources.Remark;
		}
		catch (Exception ex)
		{
			throw ex;
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
		this.nameText = new System.Windows.Forms.TextBox();
		this.remarkText = new System.Windows.Forms.TextBox();
		this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
		this.remarkLabel = new System.Windows.Forms.Label();
		this.typeLabel = new System.Windows.Forms.Label();
		this.nameLabel = new System.Windows.Forms.Label();
		this.processType = new System.Windows.Forms.Label();
		this.tableLayoutPanel1.SuspendLayout();
		base.SuspendLayout();
		this.nameText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.nameText.Location = new System.Drawing.Point(143, 8);
		this.nameText.Name = "nameText";
		this.nameText.Size = new System.Drawing.Size(467, 19);
		this.nameText.TabIndex = 0;
		this.nameText.Leave += new System.EventHandler(nameText_Leave);
		this.remarkText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.remarkText.Location = new System.Drawing.Point(143, 68);
		this.remarkText.Name = "remarkText";
		this.remarkText.Size = new System.Drawing.Size(467, 19);
		this.remarkText.TabIndex = 1;
		this.remarkText.Leave += new System.EventHandler(remarkText_Leave);
		this.tableLayoutPanel1.ColumnCount = 2;
		this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 140f));
		this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.tableLayoutPanel1.Controls.Add(this.remarkLabel, 0, 2);
		this.tableLayoutPanel1.Controls.Add(this.typeLabel, 0, 1);
		this.tableLayoutPanel1.Controls.Add(this.remarkText, 1, 2);
		this.tableLayoutPanel1.Controls.Add(this.nameText, 1, 0);
		this.tableLayoutPanel1.Controls.Add(this.nameLabel, 0, 0);
		this.tableLayoutPanel1.Controls.Add(this.processType, 1, 1);
		this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
		this.tableLayoutPanel1.Location = new System.Drawing.Point(10, 10);
		this.tableLayoutPanel1.Name = "tableLayoutPanel1";
		this.tableLayoutPanel1.RowCount = 4;
		this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.tableLayoutPanel1.Size = new System.Drawing.Size(613, 376);
		this.tableLayoutPanel1.TabIndex = 3;
		this.remarkLabel.AutoSize = true;
		this.remarkLabel.BackColor = System.Drawing.SystemColors.Control;
		this.remarkLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.remarkLabel.Location = new System.Drawing.Point(3, 63);
		this.remarkLabel.Margin = new System.Windows.Forms.Padding(3);
		this.remarkLabel.Name = "remarkLabel";
		this.remarkLabel.Size = new System.Drawing.Size(134, 24);
		this.remarkLabel.TabIndex = 5;
		this.remarkLabel.Text = "label3";
		this.remarkLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.typeLabel.AutoSize = true;
		this.typeLabel.BackColor = System.Drawing.SystemColors.Control;
		this.typeLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.typeLabel.Location = new System.Drawing.Point(3, 33);
		this.typeLabel.Margin = new System.Windows.Forms.Padding(3);
		this.typeLabel.Name = "typeLabel";
		this.typeLabel.Size = new System.Drawing.Size(134, 24);
		this.typeLabel.TabIndex = 4;
		this.typeLabel.Text = "label2";
		this.typeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.nameLabel.AutoSize = true;
		this.nameLabel.BackColor = System.Drawing.SystemColors.Control;
		this.nameLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.nameLabel.Location = new System.Drawing.Point(3, 3);
		this.nameLabel.Margin = new System.Windows.Forms.Padding(3);
		this.nameLabel.Name = "nameLabel";
		this.nameLabel.Size = new System.Drawing.Size(134, 24);
		this.nameLabel.TabIndex = 3;
		this.nameLabel.Text = "label1";
		this.nameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.processType.AutoSize = true;
		this.processType.Dock = System.Windows.Forms.DockStyle.Fill;
		this.processType.Location = new System.Drawing.Point(143, 30);
		this.processType.Name = "processType";
		this.processType.Size = new System.Drawing.Size(467, 30);
		this.processType.TabIndex = 6;
		this.processType.Text = "processType";
		this.processType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.White;
		base.Controls.Add(this.tableLayoutPanel1);
		base.Name = "ProcessView";
		base.Padding = new System.Windows.Forms.Padding(10);
		base.Size = new System.Drawing.Size(633, 396);
		base.Load += new System.EventHandler(ProcessView_Load);
		this.tableLayoutPanel1.ResumeLayout(false);
		this.tableLayoutPanel1.PerformLayout();
		base.ResumeLayout(false);
	}
}
