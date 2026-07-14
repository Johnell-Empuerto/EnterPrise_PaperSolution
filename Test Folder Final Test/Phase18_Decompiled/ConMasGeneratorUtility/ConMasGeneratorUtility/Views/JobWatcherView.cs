using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using ConMasGeneratorLib.Data;
using ConMasGeneratorUtility.Controls;
using ConMasGeneratorUtility.Properties;

namespace ConMasGeneratorUtility.Views;

public class JobWatcherView : SettingViewBase
{
	private JobData _jobData;

	private TreeNode _node;

	private IContainer components = null;

	private TextBox jobNameText;

	private FolderBrowserDialog folderBrowserDialog;

	private TextBox folderText;

	private TextBox normalEndFolderText;

	private Button normalEndFolderButton;

	private TextBox abnormalEndFolderText;

	private Button abnormalEndFolderButton;

	private Button workFolderButton;

	private TextBox workFolderText;

	private Button downLoadFolderButton;

	private TextBox downLoadFolderText;

	private TextBox remarkText;

	private CheckBox enableCheckBox;

	private TableLayoutPanel baseTable;

	private Button folderButton;

	private Label nameLabel;

	private Label watcherFolderLabel;

	private Label normalFolderLabel;

	private Label abnormalFolderLabel;

	private Label workFolderLabel;

	private Label downloadFolderLabel;

	private Label remarkLabel;

	private Label bufferLabel;

	private Label retryCountLabel;

	private NumericUpDown bufferSize;

	private Label retryTimingLabel;

	private NumericUpDown retry;

	private NumericTextBox retryTiming;

	public JobWatcherView(TreeNode node)
	{
		InitializeComponent();
		SetFormResouce();
		_jobData = node.Tag as JobData;
		_node = node;
		DoubleBuffered = true;
	}

	public override void SaveTemporary()
	{
		try
		{
			if (_jobData != null)
			{
				_jobData.Name = jobNameText.Text;
			}
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

	private void JobWatcherView_Load(object sender, EventArgs e)
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
			jobNameText.Text = _jobData.Name;
			folderText.Text = _jobData.Folder;
			normalEndFolderText.Text = _jobData.NormalEndFolder;
			abnormalEndFolderText.Text = _jobData.AbnomalEndFolder;
			workFolderText.Text = _jobData.WorkFolder;
			downLoadFolderText.Text = _jobData.DownLoadFolder;
			remarkText.Text = _jobData.Remark;
			if (!string.IsNullOrWhiteSpace(_jobData.InternalBufferSize))
			{
				bufferSize.Value = decimal.Parse(_jobData.InternalBufferSize);
			}
			if (!string.IsNullOrWhiteSpace(_jobData.JobRetryCount))
			{
				retry.Value = decimal.Parse(_jobData.JobRetryCount);
			}
			retryTiming.Text = "1000";
			if (!string.IsNullOrWhiteSpace(_jobData.JobRetryInterval))
			{
				retryTiming.Text = _jobData.JobRetryInterval;
			}
			if (_jobData.Enable == "0")
			{
				enableCheckBox.Checked = true;
			}
			else
			{
				enableCheckBox.Checked = false;
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void jobNameText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			if (string.IsNullOrEmpty(jobNameText.Text.Trim()))
			{
				MessageBox.Show(Resources.NameEmptyMessage, "WARNING");
				jobNameText.Focus();
			}
			else
			{
				_jobData.Name = jobNameText.Text;
				_node.Text = jobNameText.Text;
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void folderText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_jobData.Folder = folderText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void normalEndFolderText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_jobData.NormalEndFolder = normalEndFolderText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void abnormalEndFolderText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_jobData.AbnomalEndFolder = abnormalEndFolderText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void workFolderText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_jobData.WorkFolder = workFolderText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void downLoadFolderText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_jobData.DownLoadFolder = downLoadFolderText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void remarkText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_jobData.Remark = remarkText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void enableCheckBox_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			if (enableCheckBox.Checked)
			{
				_jobData.Enable = "0";
			}
			else
			{
				_jobData.Enable = "1";
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void folderButton_Click(object sender, EventArgs e)
	{
		try
		{
			folderBrowserDialog.Description = Resources.FolderMessageDialogTitle;
			folderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop;
			folderBrowserDialog.SelectedPath = "C:\\";
			folderBrowserDialog.ShowNewFolderButton = true;
			if ((sender as Button).Name == "folderButton")
			{
				if (!string.IsNullOrEmpty(folderText.Text))
				{
					folderBrowserDialog.SelectedPath = folderText.Text;
				}
				if (folderBrowserDialog.ShowDialog(this) == DialogResult.OK)
				{
					folderText.Text = folderBrowserDialog.SelectedPath;
					_jobData.Folder = folderText.Text;
				}
			}
			else if ((sender as Button).Name == "normalEndFolderButton")
			{
				if (!string.IsNullOrEmpty(normalEndFolderText.Text))
				{
					folderBrowserDialog.SelectedPath = normalEndFolderText.Text;
				}
				if (folderBrowserDialog.ShowDialog(this) == DialogResult.OK)
				{
					normalEndFolderText.Text = folderBrowserDialog.SelectedPath;
					_jobData.NormalEndFolder = normalEndFolderText.Text;
				}
			}
			else if ((sender as Button).Name == "abnormalEndFolderButton")
			{
				if (!string.IsNullOrEmpty(abnormalEndFolderText.Text))
				{
					folderBrowserDialog.SelectedPath = abnormalEndFolderText.Text;
				}
				if (folderBrowserDialog.ShowDialog(this) == DialogResult.OK)
				{
					abnormalEndFolderText.Text = folderBrowserDialog.SelectedPath;
					_jobData.AbnomalEndFolder = abnormalEndFolderText.Text;
				}
			}
			else if ((sender as Button).Name == "workFolderButton")
			{
				if (!string.IsNullOrEmpty(workFolderText.Text))
				{
					folderBrowserDialog.SelectedPath = workFolderText.Text;
				}
				if (folderBrowserDialog.ShowDialog(this) == DialogResult.OK)
				{
					workFolderText.Text = folderBrowserDialog.SelectedPath;
					_jobData.WorkFolder = workFolderText.Text;
				}
			}
			else if ((sender as Button).Name == "downLoadFolderButton")
			{
				if (!string.IsNullOrEmpty(downLoadFolderText.Text))
				{
					folderBrowserDialog.SelectedPath = downLoadFolderText.Text;
				}
				if (folderBrowserDialog.ShowDialog(this) == DialogResult.OK)
				{
					downLoadFolderText.Text = folderBrowserDialog.SelectedPath;
					_jobData.DownLoadFolder = downLoadFolderText.Text;
				}
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
			nameLabel.Text = Resources.Name;
			watcherFolderLabel.Text = Resources.WatchFolder;
			normalFolderLabel.Text = Resources.NormalFolder;
			abnormalFolderLabel.Text = Resources.AbnormalFolder;
			workFolderLabel.Text = Resources.WorkFolder;
			downloadFolderLabel.Text = Resources.DownLoadFolder;
			remarkLabel.Text = Resources.Remark;
			enableCheckBox.Text = Resources.Enable;
			bufferLabel.Text = Resources.BufferLabel;
			retryCountLabel.Text = Resources.RetryCountLabel;
			retryTimingLabel.Text = Resources.RetryTimingLabel;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void bufferSizeValidating(object sender, CancelEventArgs e)
	{
		_jobData.InternalBufferSize = int.Parse(bufferSize.Value.ToString()).ToString();
	}

	private void bufferSize_ValueChanged(object sender, EventArgs e)
	{
		_jobData.InternalBufferSize = int.Parse(bufferSize.Value.ToString()).ToString();
	}

	private void retryValidating(object sender, CancelEventArgs e)
	{
		_jobData.JobRetryCount = int.Parse(retry.Value.ToString()).ToString();
	}

	private void retry_ValueChanged(object sender, EventArgs e)
	{
		_jobData.JobRetryCount = int.Parse(retry.Value.ToString()).ToString();
	}

	private void retryTiming_Validating(object sender, CancelEventArgs e)
	{
		_jobData.JobRetryInterval = retryTiming.Text;
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
		System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConMasGeneratorUtility.Views.JobWatcherView));
		this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
		this.baseTable = new System.Windows.Forms.TableLayoutPanel();
		this.jobNameText = new System.Windows.Forms.TextBox();
		this.folderButton = new System.Windows.Forms.Button();
		this.remarkText = new System.Windows.Forms.TextBox();
		this.folderText = new System.Windows.Forms.TextBox();
		this.downLoadFolderButton = new System.Windows.Forms.Button();
		this.normalEndFolderText = new System.Windows.Forms.TextBox();
		this.downLoadFolderText = new System.Windows.Forms.TextBox();
		this.normalEndFolderButton = new System.Windows.Forms.Button();
		this.workFolderButton = new System.Windows.Forms.Button();
		this.abnormalEndFolderText = new System.Windows.Forms.TextBox();
		this.workFolderText = new System.Windows.Forms.TextBox();
		this.abnormalEndFolderButton = new System.Windows.Forms.Button();
		this.nameLabel = new System.Windows.Forms.Label();
		this.watcherFolderLabel = new System.Windows.Forms.Label();
		this.normalFolderLabel = new System.Windows.Forms.Label();
		this.abnormalFolderLabel = new System.Windows.Forms.Label();
		this.workFolderLabel = new System.Windows.Forms.Label();
		this.downloadFolderLabel = new System.Windows.Forms.Label();
		this.remarkLabel = new System.Windows.Forms.Label();
		this.bufferLabel = new System.Windows.Forms.Label();
		this.retryCountLabel = new System.Windows.Forms.Label();
		this.bufferSize = new System.Windows.Forms.NumericUpDown();
		this.enableCheckBox = new System.Windows.Forms.CheckBox();
		this.retryTimingLabel = new System.Windows.Forms.Label();
		this.retry = new System.Windows.Forms.NumericUpDown();
		this.retryTiming = new ConMasGeneratorUtility.Controls.NumericTextBox();
		this.baseTable.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.bufferSize).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.retry).BeginInit();
		base.SuspendLayout();
		this.baseTable.ColumnCount = 3;
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 150f));
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 37f));
		this.baseTable.Controls.Add(this.jobNameText, 1, 0);
		this.baseTable.Controls.Add(this.folderButton, 2, 1);
		this.baseTable.Controls.Add(this.remarkText, 1, 6);
		this.baseTable.Controls.Add(this.folderText, 1, 1);
		this.baseTable.Controls.Add(this.downLoadFolderButton, 2, 5);
		this.baseTable.Controls.Add(this.normalEndFolderText, 1, 2);
		this.baseTable.Controls.Add(this.downLoadFolderText, 1, 5);
		this.baseTable.Controls.Add(this.normalEndFolderButton, 2, 2);
		this.baseTable.Controls.Add(this.workFolderButton, 2, 4);
		this.baseTable.Controls.Add(this.abnormalEndFolderText, 1, 3);
		this.baseTable.Controls.Add(this.workFolderText, 1, 4);
		this.baseTable.Controls.Add(this.abnormalEndFolderButton, 2, 3);
		this.baseTable.Controls.Add(this.nameLabel, 0, 0);
		this.baseTable.Controls.Add(this.watcherFolderLabel, 0, 1);
		this.baseTable.Controls.Add(this.normalFolderLabel, 0, 2);
		this.baseTable.Controls.Add(this.abnormalFolderLabel, 0, 3);
		this.baseTable.Controls.Add(this.workFolderLabel, 0, 4);
		this.baseTable.Controls.Add(this.downloadFolderLabel, 0, 5);
		this.baseTable.Controls.Add(this.remarkLabel, 0, 6);
		this.baseTable.Controls.Add(this.bufferLabel, 0, 7);
		this.baseTable.Controls.Add(this.retryCountLabel, 0, 8);
		this.baseTable.Controls.Add(this.bufferSize, 1, 7);
		this.baseTable.Controls.Add(this.enableCheckBox, 1, 10);
		this.baseTable.Controls.Add(this.retryTimingLabel, 0, 9);
		this.baseTable.Controls.Add(this.retry, 1, 8);
		this.baseTable.Controls.Add(this.retryTiming, 1, 9);
		this.baseTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.baseTable.Location = new System.Drawing.Point(10, 10);
		this.baseTable.Name = "baseTable";
		this.baseTable.RowCount = 11;
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 23f));
		this.baseTable.Size = new System.Drawing.Size(797, 375);
		this.baseTable.TabIndex = 14;
		this.jobNameText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.jobNameText.Location = new System.Drawing.Point(153, 8);
		this.jobNameText.Name = "jobNameText";
		this.jobNameText.Size = new System.Drawing.Size(604, 19);
		this.jobNameText.TabIndex = 1;
		this.jobNameText.Validating += new System.ComponentModel.CancelEventHandler(jobNameText_Validating);
		this.folderButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("folderButton.BackgroundImage");
		this.folderButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.folderButton.Dock = System.Windows.Forms.DockStyle.Left;
		this.folderButton.Location = new System.Drawing.Point(763, 33);
		this.folderButton.Name = "folderButton";
		this.folderButton.Size = new System.Drawing.Size(27, 24);
		this.folderButton.TabIndex = 3;
		this.folderButton.UseVisualStyleBackColor = true;
		this.folderButton.Click += new System.EventHandler(folderButton_Click);
		this.remarkText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.remarkText.Location = new System.Drawing.Point(153, 188);
		this.remarkText.Name = "remarkText";
		this.remarkText.Size = new System.Drawing.Size(604, 19);
		this.remarkText.TabIndex = 12;
		this.remarkText.Validating += new System.ComponentModel.CancelEventHandler(remarkText_Validating);
		this.folderText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.folderText.Location = new System.Drawing.Point(153, 38);
		this.folderText.Name = "folderText";
		this.folderText.Size = new System.Drawing.Size(604, 19);
		this.folderText.TabIndex = 2;
		this.folderText.Validating += new System.ComponentModel.CancelEventHandler(folderText_Validating);
		this.downLoadFolderButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("downLoadFolderButton.BackgroundImage");
		this.downLoadFolderButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.downLoadFolderButton.Dock = System.Windows.Forms.DockStyle.Left;
		this.downLoadFolderButton.Location = new System.Drawing.Point(763, 153);
		this.downLoadFolderButton.Name = "downLoadFolderButton";
		this.downLoadFolderButton.Size = new System.Drawing.Size(27, 24);
		this.downLoadFolderButton.TabIndex = 11;
		this.downLoadFolderButton.UseVisualStyleBackColor = true;
		this.downLoadFolderButton.Click += new System.EventHandler(folderButton_Click);
		this.normalEndFolderText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.normalEndFolderText.Location = new System.Drawing.Point(153, 68);
		this.normalEndFolderText.Name = "normalEndFolderText";
		this.normalEndFolderText.Size = new System.Drawing.Size(604, 19);
		this.normalEndFolderText.TabIndex = 4;
		this.normalEndFolderText.Validating += new System.ComponentModel.CancelEventHandler(normalEndFolderText_Validating);
		this.downLoadFolderText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.downLoadFolderText.Location = new System.Drawing.Point(153, 158);
		this.downLoadFolderText.Name = "downLoadFolderText";
		this.downLoadFolderText.Size = new System.Drawing.Size(604, 19);
		this.downLoadFolderText.TabIndex = 10;
		this.downLoadFolderText.Validating += new System.ComponentModel.CancelEventHandler(downLoadFolderText_Validating);
		this.normalEndFolderButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("normalEndFolderButton.BackgroundImage");
		this.normalEndFolderButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.normalEndFolderButton.Dock = System.Windows.Forms.DockStyle.Left;
		this.normalEndFolderButton.Location = new System.Drawing.Point(763, 63);
		this.normalEndFolderButton.Name = "normalEndFolderButton";
		this.normalEndFolderButton.Size = new System.Drawing.Size(27, 24);
		this.normalEndFolderButton.TabIndex = 5;
		this.normalEndFolderButton.UseVisualStyleBackColor = true;
		this.normalEndFolderButton.Click += new System.EventHandler(folderButton_Click);
		this.workFolderButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("workFolderButton.BackgroundImage");
		this.workFolderButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.workFolderButton.Dock = System.Windows.Forms.DockStyle.Left;
		this.workFolderButton.Location = new System.Drawing.Point(763, 123);
		this.workFolderButton.Name = "workFolderButton";
		this.workFolderButton.Size = new System.Drawing.Size(27, 24);
		this.workFolderButton.TabIndex = 9;
		this.workFolderButton.UseVisualStyleBackColor = true;
		this.workFolderButton.Click += new System.EventHandler(folderButton_Click);
		this.abnormalEndFolderText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.abnormalEndFolderText.Location = new System.Drawing.Point(153, 98);
		this.abnormalEndFolderText.Name = "abnormalEndFolderText";
		this.abnormalEndFolderText.Size = new System.Drawing.Size(604, 19);
		this.abnormalEndFolderText.TabIndex = 6;
		this.abnormalEndFolderText.Validating += new System.ComponentModel.CancelEventHandler(abnormalEndFolderText_Validating);
		this.workFolderText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.workFolderText.Location = new System.Drawing.Point(153, 128);
		this.workFolderText.Name = "workFolderText";
		this.workFolderText.Size = new System.Drawing.Size(604, 19);
		this.workFolderText.TabIndex = 8;
		this.workFolderText.Validating += new System.ComponentModel.CancelEventHandler(workFolderText_Validating);
		this.abnormalEndFolderButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("abnormalEndFolderButton.BackgroundImage");
		this.abnormalEndFolderButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.abnormalEndFolderButton.Dock = System.Windows.Forms.DockStyle.Left;
		this.abnormalEndFolderButton.Location = new System.Drawing.Point(763, 93);
		this.abnormalEndFolderButton.Name = "abnormalEndFolderButton";
		this.abnormalEndFolderButton.Size = new System.Drawing.Size(27, 24);
		this.abnormalEndFolderButton.TabIndex = 7;
		this.abnormalEndFolderButton.UseVisualStyleBackColor = true;
		this.abnormalEndFolderButton.Click += new System.EventHandler(folderButton_Click);
		this.nameLabel.BackColor = System.Drawing.SystemColors.Control;
		this.nameLabel.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.nameLabel.Location = new System.Drawing.Point(3, 3);
		this.nameLabel.Margin = new System.Windows.Forms.Padding(3);
		this.nameLabel.Name = "nameLabel";
		this.nameLabel.Size = new System.Drawing.Size(144, 24);
		this.nameLabel.TabIndex = 14;
		this.nameLabel.Text = "label1";
		this.nameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.watcherFolderLabel.BackColor = System.Drawing.SystemColors.Control;
		this.watcherFolderLabel.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.watcherFolderLabel.Location = new System.Drawing.Point(3, 33);
		this.watcherFolderLabel.Margin = new System.Windows.Forms.Padding(3);
		this.watcherFolderLabel.Name = "watcherFolderLabel";
		this.watcherFolderLabel.Size = new System.Drawing.Size(144, 24);
		this.watcherFolderLabel.TabIndex = 15;
		this.watcherFolderLabel.Text = "label2";
		this.watcherFolderLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.normalFolderLabel.BackColor = System.Drawing.SystemColors.Control;
		this.normalFolderLabel.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.normalFolderLabel.Location = new System.Drawing.Point(3, 63);
		this.normalFolderLabel.Margin = new System.Windows.Forms.Padding(3);
		this.normalFolderLabel.Name = "normalFolderLabel";
		this.normalFolderLabel.Size = new System.Drawing.Size(144, 24);
		this.normalFolderLabel.TabIndex = 16;
		this.normalFolderLabel.Text = "label3";
		this.normalFolderLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.abnormalFolderLabel.BackColor = System.Drawing.SystemColors.Control;
		this.abnormalFolderLabel.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.abnormalFolderLabel.Location = new System.Drawing.Point(3, 93);
		this.abnormalFolderLabel.Margin = new System.Windows.Forms.Padding(3);
		this.abnormalFolderLabel.Name = "abnormalFolderLabel";
		this.abnormalFolderLabel.Size = new System.Drawing.Size(144, 24);
		this.abnormalFolderLabel.TabIndex = 17;
		this.abnormalFolderLabel.Text = "label4";
		this.abnormalFolderLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.workFolderLabel.BackColor = System.Drawing.SystemColors.Control;
		this.workFolderLabel.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.workFolderLabel.Location = new System.Drawing.Point(3, 123);
		this.workFolderLabel.Margin = new System.Windows.Forms.Padding(3);
		this.workFolderLabel.Name = "workFolderLabel";
		this.workFolderLabel.Size = new System.Drawing.Size(144, 24);
		this.workFolderLabel.TabIndex = 18;
		this.workFolderLabel.Text = "label5";
		this.workFolderLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.downloadFolderLabel.BackColor = System.Drawing.SystemColors.Control;
		this.downloadFolderLabel.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.downloadFolderLabel.Location = new System.Drawing.Point(3, 153);
		this.downloadFolderLabel.Margin = new System.Windows.Forms.Padding(3);
		this.downloadFolderLabel.Name = "downloadFolderLabel";
		this.downloadFolderLabel.Size = new System.Drawing.Size(144, 24);
		this.downloadFolderLabel.TabIndex = 19;
		this.downloadFolderLabel.Text = "label6";
		this.downloadFolderLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.remarkLabel.BackColor = System.Drawing.SystemColors.Control;
		this.remarkLabel.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.remarkLabel.Location = new System.Drawing.Point(3, 183);
		this.remarkLabel.Margin = new System.Windows.Forms.Padding(3);
		this.remarkLabel.Name = "remarkLabel";
		this.remarkLabel.Size = new System.Drawing.Size(144, 24);
		this.remarkLabel.TabIndex = 20;
		this.remarkLabel.Text = "label7";
		this.remarkLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.bufferLabel.BackColor = System.Drawing.SystemColors.Control;
		this.bufferLabel.Location = new System.Drawing.Point(3, 213);
		this.bufferLabel.Margin = new System.Windows.Forms.Padding(3);
		this.bufferLabel.Name = "bufferLabel";
		this.bufferLabel.Size = new System.Drawing.Size(144, 23);
		this.bufferLabel.TabIndex = 21;
		this.bufferLabel.Text = "label7";
		this.bufferLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.retryCountLabel.BackColor = System.Drawing.SystemColors.Control;
		this.retryCountLabel.Location = new System.Drawing.Point(3, 242);
		this.retryCountLabel.Margin = new System.Windows.Forms.Padding(3);
		this.retryCountLabel.Name = "retryCountLabel";
		this.retryCountLabel.Size = new System.Drawing.Size(144, 23);
		this.retryCountLabel.TabIndex = 22;
		this.retryCountLabel.Text = "label7";
		this.retryCountLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.bufferSize.Increment = new decimal(new int[4] { 4096, 0, 0, 0 });
		this.bufferSize.Location = new System.Drawing.Point(153, 213);
		this.bufferSize.Maximum = new decimal(new int[4] { 65536, 0, 0, 0 });
		this.bufferSize.Minimum = new decimal(new int[4] { 4096, 0, 0, 0 });
		this.bufferSize.Name = "bufferSize";
		this.bufferSize.Size = new System.Drawing.Size(120, 19);
		this.bufferSize.TabIndex = 25;
		this.bufferSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
		this.bufferSize.Value = new decimal(new int[4] { 8192, 0, 0, 0 });
		this.bufferSize.ValueChanged += new System.EventHandler(bufferSize_ValueChanged);
		this.bufferSize.Validating += new System.ComponentModel.CancelEventHandler(bufferSizeValidating);
		this.enableCheckBox.AutoSize = true;
		this.enableCheckBox.Checked = true;
		this.enableCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
		this.enableCheckBox.Location = new System.Drawing.Point(153, 306);
		this.enableCheckBox.Name = "enableCheckBox";
		this.enableCheckBox.Size = new System.Drawing.Size(80, 16);
		this.enableCheckBox.TabIndex = 13;
		this.enableCheckBox.Text = "checkBox1";
		this.enableCheckBox.UseVisualStyleBackColor = true;
		this.enableCheckBox.CheckedChanged += new System.EventHandler(enableCheckBox_CheckedChanged);
		this.retryTimingLabel.BackColor = System.Drawing.SystemColors.Control;
		this.retryTimingLabel.Location = new System.Drawing.Point(3, 271);
		this.retryTimingLabel.Margin = new System.Windows.Forms.Padding(3);
		this.retryTimingLabel.Name = "retryTimingLabel";
		this.retryTimingLabel.Size = new System.Drawing.Size(144, 15);
		this.retryTimingLabel.TabIndex = 28;
		this.retryTimingLabel.Text = "label7";
		this.retryTimingLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.retry.Location = new System.Drawing.Point(153, 242);
		this.retry.Maximum = new decimal(new int[4] { 9, 0, 0, 0 });
		this.retry.Name = "retry";
		this.retry.Size = new System.Drawing.Size(120, 19);
		this.retry.TabIndex = 26;
		this.retry.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
		this.retry.ValueChanged += new System.EventHandler(retry_ValueChanged);
		this.retry.Validating += new System.ComponentModel.CancelEventHandler(retryValidating);
		this.retryTiming.Location = new System.Drawing.Point(153, 271);
		this.retryTiming.Name = "retryTiming";
		this.retryTiming.Size = new System.Drawing.Size(100, 19);
		this.retryTiming.TabIndex = 29;
		this.retryTiming.Text = "1000";
		this.retryTiming.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
		this.retryTiming.Validating += new System.ComponentModel.CancelEventHandler(retryTiming_Validating);
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.White;
		base.Controls.Add(this.baseTable);
		base.Margin = new System.Windows.Forms.Padding(10);
		base.Name = "JobWatcherView";
		base.Padding = new System.Windows.Forms.Padding(10);
		base.Size = new System.Drawing.Size(817, 395);
		base.Load += new System.EventHandler(JobWatcherView_Load);
		this.baseTable.ResumeLayout(false);
		this.baseTable.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.bufferSize).EndInit();
		((System.ComponentModel.ISupportInitialize)this.retry).EndInit();
		base.ResumeLayout(false);
	}
}
