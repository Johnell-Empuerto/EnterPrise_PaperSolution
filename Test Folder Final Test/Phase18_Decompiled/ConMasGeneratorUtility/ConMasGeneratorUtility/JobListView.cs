using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using ConMasGeneratorLib.Data;
using ConMasGeneratorLib.Job;

namespace ConMasGeneratorUtility;

public class JobListView : UserControl
{
	private List<JobData> jobDatas;

	private IContainer components = null;

	private DataGridView JobList;

	private TableLayoutPanel LayoutPanel;

	private Button button2;

	public JobListView()
	{
		InitializeComponent();
	}

	private void JobListView_Load(object sender, EventArgs e)
	{
		try
		{
			SetGrid();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
		}
	}

	private void SetGrid()
	{
		try
		{
			JobController jobController = new JobController();
			jobDatas = jobController.GetJobData();
			if (jobDatas != null)
			{
				JobList.DataSource = jobDatas;
				JobList.Visible = true;
			}
			else
			{
				JobList.DataSource = null;
				JobList.Visible = false;
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void button1_Click(object sender, EventArgs e)
	{
		try
		{
			SetGrid();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
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
		this.JobList = new System.Windows.Forms.DataGridView();
		this.LayoutPanel = new System.Windows.Forms.TableLayoutPanel();
		this.button2 = new System.Windows.Forms.Button();
		((System.ComponentModel.ISupportInitialize)this.JobList).BeginInit();
		this.LayoutPanel.SuspendLayout();
		base.SuspendLayout();
		this.JobList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
		this.JobList.Dock = System.Windows.Forms.DockStyle.Fill;
		this.JobList.Location = new System.Drawing.Point(6, 36);
		this.JobList.Name = "JobList";
		this.JobList.RowTemplate.Height = 21;
		this.JobList.Size = new System.Drawing.Size(869, 401);
		this.JobList.TabIndex = 0;
		this.LayoutPanel.BackColor = System.Drawing.Color.Transparent;
		this.LayoutPanel.ColumnCount = 1;
		this.LayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.LayoutPanel.Controls.Add(this.JobList, 0, 1);
		this.LayoutPanel.Controls.Add(this.button2, 0, 0);
		this.LayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.LayoutPanel.Location = new System.Drawing.Point(0, 0);
		this.LayoutPanel.Name = "LayoutPanel";
		this.LayoutPanel.Padding = new System.Windows.Forms.Padding(3);
		this.LayoutPanel.RowCount = 3;
		this.LayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.LayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.LayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20f));
		this.LayoutPanel.Size = new System.Drawing.Size(881, 463);
		this.LayoutPanel.TabIndex = 2;
		this.button2.Location = new System.Drawing.Point(6, 6);
		this.button2.Name = "button2";
		this.button2.Size = new System.Drawing.Size(55, 24);
		this.button2.TabIndex = 1;
		this.button2.Text = "更新";
		this.button2.UseVisualStyleBackColor = true;
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		base.Controls.Add(this.LayoutPanel);
		base.Name = "JobListView";
		base.Size = new System.Drawing.Size(881, 463);
		base.Load += new System.EventHandler(JobListView_Load);
		((System.ComponentModel.ISupportInitialize)this.JobList).EndInit();
		this.LayoutPanel.ResumeLayout(false);
		base.ResumeLayout(false);
	}
}
