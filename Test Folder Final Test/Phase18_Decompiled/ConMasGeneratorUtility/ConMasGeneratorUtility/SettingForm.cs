using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using ConMasGeneratorLib.Util;
using ConMasGeneratorUtility.Properties;
using ConMasGeneratorUtility.Views;

namespace ConMasGeneratorUtility;

public class SettingForm : Form
{
	private IContainer components = null;

	private TabControl SettingTab;

	private TabPage tabPage1;

	private TabPage tabWatch;

	private Button debugButton;

	private TabPage tabSchdule;

	private Panel BasePanel;

	private Panel DebugPanel;

	private TableLayoutPanel FormTableLayoutPanel;

	private ImageList ImgList;

	private TabPage tabJobLogList;

	private TabPage tabWatcherService;

	private TabPage tabMail;

	public SettingForm()
	{
		InitializeComponent();
		SetFormResouce();
		FileVersionInfo versionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
	}

	private void SettingForm_Load(object sender, EventArgs e)
	{
		try
		{
			if (Utility.DebugMode())
			{
				DebugPanel.Visible = true;
				DebugPanel.Height = 40;
				FormTableLayoutPanel.RowStyles[0].Height = 43f;
			}
			SettingTab.SelectedTab = tabPage1;
			ChangeTab();
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.ToString());
		}
	}

	private void button2_Click(object sender, EventArgs e)
	{
		TestForm testForm = new TestForm();
		testForm.ShowDialog();
		testForm.Dispose();
	}

	private void SettingTab_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			ChangeTab();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void ChangeTab()
	{
		try
		{
			if (SettingTab.SelectedTab == tabPage1)
			{
				if (tabPage1.Controls.Count == 0)
				{
					AppSettingView appSettingView = new AppSettingView();
					appSettingView.Dock = DockStyle.Fill;
					tabPage1.Controls.Add(appSettingView);
				}
			}
			else if (SettingTab.SelectedTab == tabWatch)
			{
				if (tabWatch.Controls.Count == 0)
				{
					JobSettingView jobSettingView = new JobSettingView("watch");
					jobSettingView.Dock = DockStyle.Fill;
					tabWatch.Controls.Add(jobSettingView);
				}
			}
			else if (SettingTab.SelectedTab == tabSchdule)
			{
				if (tabSchdule.Controls.Count == 0)
				{
					JobSettingView jobSettingView2 = new JobSettingView("schdule");
					jobSettingView2.Dock = DockStyle.Fill;
					tabSchdule.Controls.Add(jobSettingView2);
				}
			}
			else if (SettingTab.SelectedTab == tabJobLogList)
			{
				if (tabJobLogList.Controls.Count == 0)
				{
					JobLogView jobLogView = new JobLogView();
					jobLogView.Dock = DockStyle.Fill;
					tabJobLogList.Controls.Add(jobLogView);
				}
			}
			else if (SettingTab.SelectedTab == tabWatcherService)
			{
				if (tabWatcherService.Controls.Count == 0)
				{
					ServiceStatusView serviceStatusView = new ServiceStatusView();
					serviceStatusView.Dock = DockStyle.Fill;
					tabWatcherService.Controls.Add(serviceStatusView);
				}
			}
			else if (SettingTab.SelectedTab == tabMail && tabMail.Controls.Count == 0)
			{
				MailSettingView mailSettingView = new MailSettingView();
				mailSettingView.Dock = DockStyle.Fill;
				tabMail.Controls.Add(mailSettingView);
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
			Text = "ConMas Generator [Version:" + Application.ProductVersion + "]";
			tabPage1.Text = Resources.UtilityTabNameBaseSetting;
			tabWatch.Text = Resources.WatcherJobSetting;
			tabSchdule.Text = Resources.SchduleJobSetting;
			tabJobLogList.Text = Resources.UtilityTabNameJobLog;
			tabWatcherService.Text = Resources.WatcherServiceControll;
			tabMail.Text = Resources.MailSetting;
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
		this.components = new System.ComponentModel.Container();
		System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConMasGeneratorUtility.SettingForm));
		this.SettingTab = new System.Windows.Forms.TabControl();
		this.tabPage1 = new System.Windows.Forms.TabPage();
		this.tabWatch = new System.Windows.Forms.TabPage();
		this.tabSchdule = new System.Windows.Forms.TabPage();
		this.tabJobLogList = new System.Windows.Forms.TabPage();
		this.tabWatcherService = new System.Windows.Forms.TabPage();
		this.tabMail = new System.Windows.Forms.TabPage();
		this.ImgList = new System.Windows.Forms.ImageList(this.components);
		this.debugButton = new System.Windows.Forms.Button();
		this.BasePanel = new System.Windows.Forms.Panel();
		this.DebugPanel = new System.Windows.Forms.Panel();
		this.FormTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
		this.SettingTab.SuspendLayout();
		this.BasePanel.SuspendLayout();
		this.DebugPanel.SuspendLayout();
		this.FormTableLayoutPanel.SuspendLayout();
		base.SuspendLayout();
		this.SettingTab.Controls.Add(this.tabPage1);
		this.SettingTab.Controls.Add(this.tabWatch);
		this.SettingTab.Controls.Add(this.tabSchdule);
		this.SettingTab.Controls.Add(this.tabJobLogList);
		this.SettingTab.Controls.Add(this.tabWatcherService);
		this.SettingTab.Controls.Add(this.tabMail);
		resources.ApplyResources(this.SettingTab, "SettingTab");
		this.SettingTab.ImageList = this.ImgList;
		this.SettingTab.Name = "SettingTab";
		this.SettingTab.SelectedIndex = 0;
		this.SettingTab.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight;
		this.SettingTab.TabStop = false;
		this.SettingTab.SelectedIndexChanged += new System.EventHandler(SettingTab_SelectedIndexChanged);
		this.tabPage1.BackColor = System.Drawing.Color.Transparent;
		resources.ApplyResources(this.tabPage1, "tabPage1");
		this.tabPage1.Name = "tabPage1";
		this.tabWatch.BackColor = System.Drawing.Color.Transparent;
		resources.ApplyResources(this.tabWatch, "tabWatch");
		this.tabWatch.Name = "tabWatch";
		this.tabSchdule.BackColor = System.Drawing.Color.Transparent;
		resources.ApplyResources(this.tabSchdule, "tabSchdule");
		this.tabSchdule.Name = "tabSchdule";
		this.tabJobLogList.BackColor = System.Drawing.Color.Transparent;
		resources.ApplyResources(this.tabJobLogList, "tabJobLogList");
		this.tabJobLogList.Name = "tabJobLogList";
		this.tabWatcherService.BackColor = System.Drawing.Color.Transparent;
		resources.ApplyResources(this.tabWatcherService, "tabWatcherService");
		this.tabWatcherService.Name = "tabWatcherService";
		resources.ApplyResources(this.tabMail, "tabMail");
		this.tabMail.Name = "tabMail";
		this.tabMail.UseVisualStyleBackColor = true;
		this.ImgList.ImageStream = (System.Windows.Forms.ImageListStreamer)resources.GetObject("ImgList.ImageStream");
		this.ImgList.TransparentColor = System.Drawing.Color.Transparent;
		this.ImgList.Images.SetKeyName(0, "baseGear.png");
		this.ImgList.Images.SetKeyName(1, "clock.png");
		this.ImgList.Images.SetKeyName(2, "dbLog.png");
		this.ImgList.Images.SetKeyName(3, "FolderIcon.png");
		this.ImgList.Images.SetKeyName(4, "service.ico");
		this.ImgList.Images.SetKeyName(5, "mail.png");
		resources.ApplyResources(this.debugButton, "debugButton");
		this.debugButton.Name = "debugButton";
		this.debugButton.UseVisualStyleBackColor = true;
		this.debugButton.Click += new System.EventHandler(button2_Click);
		this.BasePanel.Controls.Add(this.SettingTab);
		resources.ApplyResources(this.BasePanel, "BasePanel");
		this.BasePanel.Name = "BasePanel";
		this.DebugPanel.Controls.Add(this.debugButton);
		resources.ApplyResources(this.DebugPanel, "DebugPanel");
		this.DebugPanel.Name = "DebugPanel";
		this.FormTableLayoutPanel.BackColor = System.Drawing.Color.Transparent;
		resources.ApplyResources(this.FormTableLayoutPanel, "FormTableLayoutPanel");
		this.FormTableLayoutPanel.Controls.Add(this.DebugPanel, 0, 0);
		this.FormTableLayoutPanel.Controls.Add(this.BasePanel, 0, 1);
		this.FormTableLayoutPanel.Name = "FormTableLayoutPanel";
		resources.ApplyResources(this, "$this");
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.DarkGray;
		base.Controls.Add(this.FormTableLayoutPanel);
		base.Name = "SettingForm";
		base.Load += new System.EventHandler(SettingForm_Load);
		this.SettingTab.ResumeLayout(false);
		this.BasePanel.ResumeLayout(false);
		this.DebugPanel.ResumeLayout(false);
		this.FormTableLayoutPanel.ResumeLayout(false);
		base.ResumeLayout(false);
	}
}
