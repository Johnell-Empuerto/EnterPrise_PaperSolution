using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ConMasGeneratorLib.DB;
using ConMasGeneratorUtility.Properties;

namespace ConMasGeneratorUtility;

public class JobLogView : UserControl
{
	private ToolTip ToolTip;

	private string DbErrorMessage = Resources.NotAvailableLogViewMessage1 + Environment.NewLine + Resources.NotAvailableLogViewMessage2;

	private IContainer components = null;

	private DataGridView JobList;

	private TableLayoutPanel baseTable;

	private Button refreshButton;

	private Panel controllPanel;

	private Label line;

	private Label jobLabel;

	private TextBox jobSearchText;

	private RadioButton radioRow03;

	private RadioButton radioRow02;

	private RadioButton radioRow01;

	public JobLogView()
	{
		InitializeComponent();
		SetFormResouce();
	}

	private void JobListView_Load(object sender, EventArgs e)
	{
		try
		{
			radioRow01.Checked = true;
			SetGrid();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
		}
	}

	private bool SetGrid()
	{
		try
		{
			string jobSearchName = jobSearchText.Text.Trim();
			int rowCount = 10;
			if (radioRow01.Checked)
			{
				rowCount = 10;
			}
			else if (radioRow02.Checked)
			{
				rowCount = 100;
			}
			else if (radioRow03.Checked)
			{
				rowCount = 500;
			}
			DataTable dataTable = null;
			try
			{
				dataTable = SearchLog(rowCount, jobSearchName);
			}
			catch
			{
				return false;
			}
			if (dataTable != null)
			{
				JobList.DataSource = dataTable;
				JobList.Visible = true;
			}
			else
			{
				JobList.DataSource = null;
				JobList.Visible = false;
			}
			return true;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private DataTable SearchLog(int rowCount, string jobSearchName)
	{
		try
		{
			StringBuilder stringBuilder = new StringBuilder();
			string arg = string.Empty;
			if (rowCount > 0)
			{
				arg = " TOP (" + rowCount + ")";
			}
			string arg2 = string.Empty;
			if (!string.IsNullOrEmpty(jobSearchName))
			{
				arg2 = " WHERE JOB_NAME LIKE '%" + Regex.Replace(jobSearchName, "[%_\\[]", "[$0]") + "%'";
			}
			stringBuilder.Append(" SELECT {0}");
			stringBuilder.Append("   CONVERT(nvarchar(20), DATE_TIME, 120) AS DATE_TIME");
			stringBuilder.Append(" , JOB_NAME");
			stringBuilder.Append(" , JOB_TYPE");
			stringBuilder.Append(" , PROCESS_INDEX");
			stringBuilder.Append(" , PROCESS_NAME");
			stringBuilder.Append(" , PROCESS_TYPE");
			stringBuilder.Append(" , COMMAND_INDEX");
			stringBuilder.Append(" , COMMAND_NAME");
			stringBuilder.Append(" , COMMAND");
			stringBuilder.Append(" , COMMAND_RESULT");
			stringBuilder.Append(" , COMMAND_RESULT_MESSAGE");
			stringBuilder.Append(" , JOB_ID");
			stringBuilder.Append(" FROM JOB_LOG");
			stringBuilder.Append(" {1}");
			stringBuilder.Append(" ORDER BY DATE_TIME DESC, JOB_ID, PROCESS_INDEX DESC, COMMAND_INDEX DESC");
			string query = string.Format(stringBuilder.ToString(), arg, arg2);
			LocalDB localDB = new LocalDB();
			return localDB.Search(query);
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void button2_Click(object sender, EventArgs e)
	{
		try
		{
			if (!SetGrid())
			{
				MessageBox.Show(DbErrorMessage);
			}
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
			ToolTip = new ToolTip();
			ToolTip.InitialDelay = 500;
			ToolTip.ReshowDelay = 500;
			ToolTip.AutoPopDelay = 3000;
			ToolTip.SetToolTip(refreshButton, Resources.Refresh);
			radioRow01.Text = Resources.RecordCount01;
			radioRow02.Text = Resources.RecordCount02;
			radioRow03.Text = Resources.RecordCount03;
			jobLabel.Text = Resources.JobNameSearch;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void radioRow_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			if ((sender as RadioButton).Checked && !SetGrid())
			{
				MessageBox.Show(DbErrorMessage);
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void jobSearchText_KeyDown(object sender, KeyEventArgs e)
	{
		try
		{
			if (e.KeyData == Keys.Return && !SetGrid())
			{
				MessageBox.Show(DbErrorMessage);
			}
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
		System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConMasGeneratorUtility.JobLogView));
		this.JobList = new System.Windows.Forms.DataGridView();
		this.baseTable = new System.Windows.Forms.TableLayoutPanel();
		this.controllPanel = new System.Windows.Forms.Panel();
		this.jobLabel = new System.Windows.Forms.Label();
		this.jobSearchText = new System.Windows.Forms.TextBox();
		this.radioRow03 = new System.Windows.Forms.RadioButton();
		this.radioRow02 = new System.Windows.Forms.RadioButton();
		this.radioRow01 = new System.Windows.Forms.RadioButton();
		this.line = new System.Windows.Forms.Label();
		this.refreshButton = new System.Windows.Forms.Button();
		((System.ComponentModel.ISupportInitialize)this.JobList).BeginInit();
		this.baseTable.SuspendLayout();
		this.controllPanel.SuspendLayout();
		base.SuspendLayout();
		this.JobList.AllowUserToAddRows = false;
		this.JobList.AllowUserToDeleteRows = false;
		this.JobList.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
		this.JobList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
		this.JobList.Dock = System.Windows.Forms.DockStyle.Fill;
		this.JobList.Location = new System.Drawing.Point(6, 51);
		this.JobList.Name = "JobList";
		this.JobList.ReadOnly = true;
		this.JobList.RowTemplate.Height = 21;
		this.JobList.Size = new System.Drawing.Size(869, 406);
		this.JobList.TabIndex = 201;
		this.baseTable.BackColor = System.Drawing.Color.Transparent;
		this.baseTable.ColumnCount = 1;
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.baseTable.Controls.Add(this.JobList, 0, 1);
		this.baseTable.Controls.Add(this.controllPanel, 0, 0);
		this.baseTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.baseTable.Location = new System.Drawing.Point(0, 0);
		this.baseTable.Name = "baseTable";
		this.baseTable.Padding = new System.Windows.Forms.Padding(3);
		this.baseTable.RowCount = 2;
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.baseTable.Size = new System.Drawing.Size(881, 463);
		this.baseTable.TabIndex = 2;
		this.controllPanel.Controls.Add(this.jobLabel);
		this.controllPanel.Controls.Add(this.jobSearchText);
		this.controllPanel.Controls.Add(this.radioRow03);
		this.controllPanel.Controls.Add(this.radioRow02);
		this.controllPanel.Controls.Add(this.radioRow01);
		this.controllPanel.Controls.Add(this.line);
		this.controllPanel.Controls.Add(this.refreshButton);
		this.controllPanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.controllPanel.Location = new System.Drawing.Point(6, 6);
		this.controllPanel.Name = "controllPanel";
		this.controllPanel.Size = new System.Drawing.Size(869, 39);
		this.controllPanel.TabIndex = 2;
		this.jobLabel.Location = new System.Drawing.Point(345, 13);
		this.jobLabel.Name = "jobLabel";
		this.jobLabel.Size = new System.Drawing.Size(85, 14);
		this.jobLabel.TabIndex = 106;
		this.jobLabel.Text = "label1";
		this.jobSearchText.Location = new System.Drawing.Point(444, 9);
		this.jobSearchText.Name = "jobSearchText";
		this.jobSearchText.Size = new System.Drawing.Size(217, 19);
		this.jobSearchText.TabIndex = 105;
		this.jobSearchText.KeyDown += new System.Windows.Forms.KeyEventHandler(jobSearchText_KeyDown);
		this.radioRow03.Location = new System.Drawing.Point(230, 10);
		this.radioRow03.Name = "radioRow03";
		this.radioRow03.Size = new System.Drawing.Size(88, 16);
		this.radioRow03.TabIndex = 104;
		this.radioRow03.Text = "radioButton2";
		this.radioRow03.UseVisualStyleBackColor = true;
		this.radioRow03.CheckedChanged += new System.EventHandler(radioRow_CheckedChanged);
		this.radioRow02.Location = new System.Drawing.Point(138, 10);
		this.radioRow02.Name = "radioRow02";
		this.radioRow02.Size = new System.Drawing.Size(88, 16);
		this.radioRow02.TabIndex = 103;
		this.radioRow02.Text = "radioButton1";
		this.radioRow02.UseVisualStyleBackColor = true;
		this.radioRow02.CheckedChanged += new System.EventHandler(radioRow_CheckedChanged);
		this.radioRow01.Location = new System.Drawing.Point(47, 10);
		this.radioRow01.Name = "radioRow01";
		this.radioRow01.Size = new System.Drawing.Size(88, 16);
		this.radioRow01.TabIndex = 102;
		this.radioRow01.Text = "radioButton1";
		this.radioRow01.UseVisualStyleBackColor = true;
		this.radioRow01.CheckedChanged += new System.EventHandler(radioRow_CheckedChanged);
		this.line.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
		this.line.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.line.Location = new System.Drawing.Point(0, 36);
		this.line.Margin = new System.Windows.Forms.Padding(0);
		this.line.Name = "line";
		this.line.Size = new System.Drawing.Size(869, 3);
		this.line.TabIndex = 30;
		this.refreshButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("refreshButton.BackgroundImage");
		this.refreshButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.refreshButton.Location = new System.Drawing.Point(1, 1);
		this.refreshButton.Name = "refreshButton";
		this.refreshButton.Size = new System.Drawing.Size(37, 35);
		this.refreshButton.TabIndex = 101;
		this.refreshButton.UseVisualStyleBackColor = true;
		this.refreshButton.Click += new System.EventHandler(button2_Click);
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		base.Controls.Add(this.baseTable);
		base.Name = "JobLogView";
		base.Size = new System.Drawing.Size(881, 463);
		base.Load += new System.EventHandler(JobListView_Load);
		((System.ComponentModel.ISupportInitialize)this.JobList).EndInit();
		this.baseTable.ResumeLayout(false);
		this.controllPanel.ResumeLayout(false);
		this.controllPanel.PerformLayout();
		base.ResumeLayout(false);
	}
}
