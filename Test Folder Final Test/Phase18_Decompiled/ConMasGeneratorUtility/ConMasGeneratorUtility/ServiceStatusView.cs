using System;
using System.ComponentModel;
using System.Drawing;
using System.ServiceProcess;
using System.Windows.Forms;
using ConMasGeneratorUtility.Properties;
using ConMasGeneratorUtility.Utils;

namespace ConMasGeneratorUtility;

public class ServiceStatusView : UserControl
{
	private ToolTip ToolTip;

	private IContainer components = null;

	private Button ServiceStart;

	private Button ServiceStop;

	private Label ServiceStatus;

	private Button refreshButton;

	private TableLayoutPanel baseTable;

	private Panel watcherServiceControllPanel;

	private Panel panel;

	private Label line;

	private Label serviceStateLabel;

	public ServiceStatusView()
	{
		InitializeComponent();
		SetFormResouce();
	}

	private void ServiceStart_Click(object sender, EventArgs e)
	{
		try
		{
			FormUtility.StartService();
			SetServiceStatus();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void ServiceStop_Click(object sender, EventArgs e)
	{
		try
		{
			FormUtility.StoptService();
			SetServiceStatus();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void ServiceStatusView_Load(object sender, EventArgs e)
	{
		try
		{
			SetServiceStatus();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
		}
	}

	private void SetServiceStatus()
	{
		ServiceController serviceController = new ServiceController();
		try
		{
			try
			{
				serviceController.ServiceName = "ConMasGeneratorService";
			}
			catch (Exception ex)
			{
				Global.logger.Info((object)("None Service:" + Resources.NoneService), ex);
				ServiceStatus.Text = Resources.NoneService;
				return;
			}
			if (serviceController.Status == ServiceControllerStatus.StartPending)
			{
				ServiceStatus.Text = Resources.ServiceStartPending;
				ServiceStart.Enabled = false;
				ServiceStop.Enabled = false;
			}
			else if (serviceController.Status == ServiceControllerStatus.Running)
			{
				ServiceStatus.Text = Resources.ServiceRunning;
				ServiceStart.Enabled = false;
				ServiceStop.Enabled = true;
			}
			else if (serviceController.Status == ServiceControllerStatus.ContinuePending)
			{
				ServiceStatus.Text = Resources.ServiceContinuePending;
				ServiceStart.Enabled = false;
				ServiceStop.Enabled = false;
			}
			else if (serviceController.Status == ServiceControllerStatus.Paused)
			{
				ServiceStatus.Text = Resources.ServiceStoped;
				ServiceStart.Enabled = true;
				ServiceStop.Enabled = true;
			}
			else if (serviceController.Status == ServiceControllerStatus.PausePending)
			{
				ServiceStatus.Text = Resources.ServiceStoped;
				ServiceStart.Enabled = false;
				ServiceStop.Enabled = false;
			}
			else if (serviceController.Status == ServiceControllerStatus.StopPending)
			{
				ServiceStatus.Text = Resources.ServiceRunning;
				ServiceStart.Enabled = false;
				ServiceStop.Enabled = false;
			}
			else if (serviceController.Status == ServiceControllerStatus.Stopped)
			{
				ServiceStatus.Text = Resources.ServiceStoped;
				ServiceStart.Enabled = true;
				ServiceStop.Enabled = false;
			}
			else
			{
				ServiceStatus.Text = string.Empty;
				ServiceStart.Enabled = false;
				ServiceStop.Enabled = false;
			}
		}
		catch (Exception ex2)
		{
			throw ex2;
		}
		finally
		{
			if (serviceController.DisplayName != null)
			{
				serviceController.Dispose();
			}
		}
	}

	private void button1_Click(object sender, EventArgs e)
	{
		try
		{
			SetServiceStatus();
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
			ServiceStart.Text = Resources.ServiceStart;
			ServiceStop.Text = Resources.ServiceEnd;
			serviceStateLabel.Text = Resources.ServiceState;
			ToolTip = new ToolTip();
			ToolTip.InitialDelay = 500;
			ToolTip.ReshowDelay = 500;
			ToolTip.AutoPopDelay = 3000;
			ToolTip.SetToolTip(refreshButton, Resources.Refresh);
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
		System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConMasGeneratorUtility.ServiceStatusView));
		this.ServiceStart = new System.Windows.Forms.Button();
		this.ServiceStop = new System.Windows.Forms.Button();
		this.ServiceStatus = new System.Windows.Forms.Label();
		this.refreshButton = new System.Windows.Forms.Button();
		this.baseTable = new System.Windows.Forms.TableLayoutPanel();
		this.watcherServiceControllPanel = new System.Windows.Forms.Panel();
		this.line = new System.Windows.Forms.Label();
		this.panel = new System.Windows.Forms.Panel();
		this.serviceStateLabel = new System.Windows.Forms.Label();
		this.baseTable.SuspendLayout();
		this.watcherServiceControllPanel.SuspendLayout();
		this.panel.SuspendLayout();
		base.SuspendLayout();
		this.ServiceStart.ImeMode = System.Windows.Forms.ImeMode.NoControl;
		this.ServiceStart.Location = new System.Drawing.Point(5, 6);
		this.ServiceStart.Name = "ServiceStart";
		this.ServiceStart.Size = new System.Drawing.Size(107, 29);
		this.ServiceStart.TabIndex = 201;
		this.ServiceStart.Text = "開始";
		this.ServiceStart.UseVisualStyleBackColor = true;
		this.ServiceStart.Click += new System.EventHandler(ServiceStart_Click);
		this.ServiceStop.ImeMode = System.Windows.Forms.ImeMode.NoControl;
		this.ServiceStop.Location = new System.Drawing.Point(118, 6);
		this.ServiceStop.Name = "ServiceStop";
		this.ServiceStop.Size = new System.Drawing.Size(107, 29);
		this.ServiceStop.TabIndex = 202;
		this.ServiceStop.Text = "停止";
		this.ServiceStop.UseVisualStyleBackColor = true;
		this.ServiceStop.Click += new System.EventHandler(ServiceStop_Click);
		this.ServiceStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
		this.ServiceStatus.Location = new System.Drawing.Point(403, 7);
		this.ServiceStatus.Name = "ServiceStatus";
		this.ServiceStatus.Size = new System.Drawing.Size(209, 29);
		this.ServiceStatus.TabIndex = 3;
		this.ServiceStatus.Text = "…";
		this.ServiceStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
		this.refreshButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("refreshButton.BackgroundImage");
		this.refreshButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.refreshButton.Location = new System.Drawing.Point(2, 1);
		this.refreshButton.Name = "refreshButton";
		this.refreshButton.Size = new System.Drawing.Size(37, 35);
		this.refreshButton.TabIndex = 101;
		this.refreshButton.UseVisualStyleBackColor = true;
		this.refreshButton.Click += new System.EventHandler(button1_Click);
		this.baseTable.ColumnCount = 1;
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.baseTable.Controls.Add(this.watcherServiceControllPanel, 0, 0);
		this.baseTable.Controls.Add(this.panel, 0, 1);
		this.baseTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.baseTable.Location = new System.Drawing.Point(0, 0);
		this.baseTable.Name = "baseTable";
		this.baseTable.RowCount = 2;
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.baseTable.Size = new System.Drawing.Size(812, 373);
		this.baseTable.TabIndex = 4;
		this.watcherServiceControllPanel.BackColor = System.Drawing.SystemColors.Control;
		this.watcherServiceControllPanel.Controls.Add(this.line);
		this.watcherServiceControllPanel.Controls.Add(this.refreshButton);
		this.watcherServiceControllPanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.watcherServiceControllPanel.Location = new System.Drawing.Point(3, 3);
		this.watcherServiceControllPanel.Name = "watcherServiceControllPanel";
		this.watcherServiceControllPanel.Size = new System.Drawing.Size(806, 39);
		this.watcherServiceControllPanel.TabIndex = 100;
		this.line.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
		this.line.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.line.Location = new System.Drawing.Point(0, 36);
		this.line.Margin = new System.Windows.Forms.Padding(0);
		this.line.Name = "line";
		this.line.Size = new System.Drawing.Size(806, 3);
		this.line.TabIndex = 29;
		this.panel.Controls.Add(this.ServiceStatus);
		this.panel.Controls.Add(this.serviceStateLabel);
		this.panel.Controls.Add(this.ServiceStart);
		this.panel.Controls.Add(this.ServiceStop);
		this.panel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.panel.Location = new System.Drawing.Point(3, 48);
		this.panel.Name = "panel";
		this.panel.Size = new System.Drawing.Size(806, 322);
		this.panel.TabIndex = 1;
		this.serviceStateLabel.Location = new System.Drawing.Point(231, 7);
		this.serviceStateLabel.Name = "serviceStateLabel";
		this.serviceStateLabel.Size = new System.Drawing.Size(176, 28);
		this.serviceStateLabel.TabIndex = 4;
		this.serviceStateLabel.Text = "label1";
		this.serviceStateLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.White;
		base.Controls.Add(this.baseTable);
		base.Name = "ServiceStatusView";
		base.Size = new System.Drawing.Size(812, 373);
		base.Load += new System.EventHandler(ServiceStatusView_Load);
		this.baseTable.ResumeLayout(false);
		this.watcherServiceControllPanel.ResumeLayout(false);
		this.panel.ResumeLayout(false);
		base.ResumeLayout(false);
	}
}
