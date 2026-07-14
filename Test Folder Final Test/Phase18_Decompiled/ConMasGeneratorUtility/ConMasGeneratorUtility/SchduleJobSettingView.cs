using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using ConMasGeneratorLib.Job;

namespace ConMasGeneratorUtility;

public class SchduleJobSettingView : UserControl
{
	private IContainer components = null;

	private TableLayoutPanel SettingTablePanel;

	private Panel panel1;

	private Button SubmitButton;

	public SchduleJobSettingView()
	{
		InitializeComponent();
	}

	private void SubmitButton_Click(object sender, EventArgs e)
	{
		try
		{
			JobController jobController = new JobController(JobController.JobType.SchduleType);
			jobController.RegistSchJob();
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.ToString());
		}
	}

	private void SchduleJobSettingView_Load(object sender, EventArgs e)
	{
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
		this.SettingTablePanel = new System.Windows.Forms.TableLayoutPanel();
		this.panel1 = new System.Windows.Forms.Panel();
		this.SubmitButton = new System.Windows.Forms.Button();
		base.SuspendLayout();
		this.SettingTablePanel.AutoScroll = true;
		this.SettingTablePanel.AutoSize = true;
		this.SettingTablePanel.ColumnCount = 2;
		this.SettingTablePanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 0f));
		this.SettingTablePanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 356f));
		this.SettingTablePanel.Dock = System.Windows.Forms.DockStyle.Top;
		this.SettingTablePanel.Location = new System.Drawing.Point(0, 0);
		this.SettingTablePanel.Name = "SettingTablePanel";
		this.SettingTablePanel.RowCount = 1;
		this.SettingTablePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
		this.SettingTablePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
		this.SettingTablePanel.Size = new System.Drawing.Size(356, 0);
		this.SettingTablePanel.TabIndex = 3;
		this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.panel1.Location = new System.Drawing.Point(0, 228);
		this.panel1.Name = "panel1";
		this.panel1.Size = new System.Drawing.Size(356, 33);
		this.panel1.TabIndex = 5;
		this.SubmitButton.Location = new System.Drawing.Point(76, 89);
		this.SubmitButton.Name = "SubmitButton";
		this.SubmitButton.Size = new System.Drawing.Size(196, 46);
		this.SubmitButton.TabIndex = 6;
		this.SubmitButton.Text = "タスクスケジューラに登録";
		this.SubmitButton.UseVisualStyleBackColor = true;
		this.SubmitButton.Click += new System.EventHandler(SubmitButton_Click);
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.Transparent;
		base.Controls.Add(this.SubmitButton);
		base.Controls.Add(this.panel1);
		base.Controls.Add(this.SettingTablePanel);
		base.Name = "SchduleJobSettingView";
		base.Size = new System.Drawing.Size(356, 261);
		base.Load += new System.EventHandler(SchduleJobSettingView_Load);
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
