using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using ConMasGeneratorLib.Data;
using ConMasGeneratorLib.Job;
using ConMasGeneratorLib.Util;
using ConMasGeneratorUtility.Controls;
using ConMasGeneratorUtility.Properties;

namespace ConMasGeneratorUtility;

public class MailSettingView : UserControl
{
	private GeneratorData _generatorData;

	private IContainer components = null;

	private TableLayoutPanel baseTable;

	private Panel controllPanel;

	private Label line;

	private TableLayoutPanel mailTable;

	private Label maillServerLabel;

	private Label accountLabel;

	private Label portLabel;

	private Label passwordLabel;

	private TextBox accountText;

	private TextBox severText;

	private TextBox passwordText;

	private Label rePasswordLabel;

	private CheckBox sslEnable;

	private TextBox rePasswordText;

	private CheckBox mailEnable;

	private DataGridView toGrid;

	private DataGridViewTextBoxColumn address;

	private Label toAddressLabel;

	private NumericTextBox portText;

	private Panel buttonPanel;

	private Button SubmitButton;

	private Label sslLabel;

	private ImageList imageList1;

	public MailSettingView()
	{
		InitializeComponent();
		SetFormResouce();
	}

	private void JobListView_Load(object sender, EventArgs e)
	{
		try
		{
			SetData();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
		}
	}

	private void SetData()
	{
		try
		{
			using (JobController jobController = new JobController())
			{
				_generatorData = jobController.GetBaseData();
			}
			if (_generatorData == null || _generatorData.MailData == null)
			{
				return;
			}
			if (_generatorData.MailData.Enable == "1")
			{
				mailEnable.Checked = true;
			}
			else
			{
				mailEnable.Checked = false;
			}
			severText.Text = _generatorData.MailData.server;
			portText.Text = _generatorData.MailData.port.ToString();
			accountText.Text = _generatorData.MailData.from;
			passwordText.Text = _generatorData.MailData.pass;
			rePasswordText.Text = passwordText.Text;
			if (_generatorData.MailData.ssl == "1")
			{
				sslEnable.Checked = true;
			}
			else
			{
				sslEnable.Checked = false;
			}
			foreach (string item in _generatorData.MailData.to)
			{
				int index = toGrid.Rows.Add(1);
				toGrid.Rows[index].Cells["address"].Value = item;
			}
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.ToString());
			Global.logger.Error((object)ex);
		}
	}

	private void mailEnable_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			if (mailEnable.Checked)
			{
				severText.Enabled = true;
				portText.Enabled = true;
				accountText.Enabled = true;
				passwordText.Enabled = true;
				rePasswordText.Enabled = true;
				sslEnable.Enabled = true;
				toGrid.Enabled = true;
			}
			else
			{
				severText.Enabled = false;
				portText.Enabled = false;
				accountText.Enabled = false;
				passwordText.Enabled = false;
				rePasswordText.Enabled = false;
				sslEnable.Enabled = false;
				toGrid.Enabled = false;
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
			SubmitButton.Text = Resources.Enter;
			mailEnable.Text = Resources.IsErrorMail;
			maillServerLabel.Text = Resources.SmtpServer;
			portLabel.Text = Resources.Port;
			accountLabel.Text = Resources.MailFromAddress;
			passwordLabel.Text = Resources.MailPassword;
			rePasswordLabel.Text = Resources.MailRePassword;
			toAddressLabel.Text = Resources.MailToAddress;
			sslLabel.Text = Resources.Ssl;
			sslEnable.Text = Resources.IsSsl;
			if (toGrid.Columns.Count > 0)
			{
				toGrid.Columns[0].HeaderText = Resources.MailToAddress;
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void SubmitButton_Click(object sender, EventArgs e)
	{
		try
		{
			if (passwordText.Text.Trim() != rePasswordText.Text.Trim())
			{
				MessageBox.Show(Resources.PasswordErrorMessage, "WARNING");
			}
			else
			{
				if (MessageBox.Show(Resources.SaveMessage, "INFO", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
				{
					return;
				}
				int result = 0;
				if (int.TryParse(portText.Text, out result))
				{
					_generatorData.MailData.port = result;
				}
				if (mailEnable.Checked)
				{
					_generatorData.MailData.Enable = "1";
				}
				else
				{
					_generatorData.MailData.Enable = "0";
				}
				if (sslEnable.Checked)
				{
					_generatorData.MailData.ssl = "1";
				}
				else
				{
					_generatorData.MailData.ssl = "0";
				}
				_generatorData.MailData.server = severText.Text.Trim();
				_generatorData.MailData.from = accountText.Text.Trim();
				_generatorData.MailData.user = accountText.Text.Trim();
				_generatorData.MailData.pass = Utility.EncryptString(passwordText.Text, Utility.encryptKey);
				_generatorData.MailData.to.Clear();
				foreach (DataGridViewRow item in (IEnumerable)toGrid.Rows)
				{
					if (item.Cells["address"].Value != null && !string.IsNullOrEmpty(item.Cells["address"].Value.ToString()))
					{
						_generatorData.MailData.to.Add(item.Cells["address"].Value.ToString());
					}
				}
				using JobController jobController = new JobController();
				jobController.EditMailXml(_generatorData);
				return;
			}
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.ToString());
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
		this.components = new System.ComponentModel.Container();
		System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConMasGeneratorUtility.MailSettingView));
		this.baseTable = new System.Windows.Forms.TableLayoutPanel();
		this.controllPanel = new System.Windows.Forms.Panel();
		this.buttonPanel = new System.Windows.Forms.Panel();
		this.SubmitButton = new System.Windows.Forms.Button();
		this.imageList1 = new System.Windows.Forms.ImageList(this.components);
		this.line = new System.Windows.Forms.Label();
		this.mailTable = new System.Windows.Forms.TableLayoutPanel();
		this.sslLabel = new System.Windows.Forms.Label();
		this.portText = new ConMasGeneratorUtility.Controls.NumericTextBox();
		this.sslEnable = new System.Windows.Forms.CheckBox();
		this.rePasswordText = new System.Windows.Forms.TextBox();
		this.rePasswordLabel = new System.Windows.Forms.Label();
		this.passwordText = new System.Windows.Forms.TextBox();
		this.passwordLabel = new System.Windows.Forms.Label();
		this.accountText = new System.Windows.Forms.TextBox();
		this.accountLabel = new System.Windows.Forms.Label();
		this.portLabel = new System.Windows.Forms.Label();
		this.severText = new System.Windows.Forms.TextBox();
		this.maillServerLabel = new System.Windows.Forms.Label();
		this.mailEnable = new System.Windows.Forms.CheckBox();
		this.toGrid = new System.Windows.Forms.DataGridView();
		this.address = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.toAddressLabel = new System.Windows.Forms.Label();
		this.baseTable.SuspendLayout();
		this.controllPanel.SuspendLayout();
		this.buttonPanel.SuspendLayout();
		this.mailTable.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.toGrid).BeginInit();
		base.SuspendLayout();
		this.baseTable.BackColor = System.Drawing.Color.White;
		this.baseTable.ColumnCount = 1;
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.baseTable.Controls.Add(this.controllPanel, 0, 0);
		this.baseTable.Controls.Add(this.mailTable, 0, 1);
		this.baseTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.baseTable.Location = new System.Drawing.Point(0, 0);
		this.baseTable.Name = "baseTable";
		this.baseTable.Padding = new System.Windows.Forms.Padding(3);
		this.baseTable.RowCount = 2;
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.baseTable.Size = new System.Drawing.Size(881, 463);
		this.baseTable.TabIndex = 2;
		this.controllPanel.BackColor = System.Drawing.SystemColors.Control;
		this.controllPanel.Controls.Add(this.buttonPanel);
		this.controllPanel.Controls.Add(this.line);
		this.controllPanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.controllPanel.Location = new System.Drawing.Point(6, 6);
		this.controllPanel.Name = "controllPanel";
		this.controllPanel.Size = new System.Drawing.Size(869, 39);
		this.controllPanel.TabIndex = 2;
		this.buttonPanel.Controls.Add(this.SubmitButton);
		this.buttonPanel.Dock = System.Windows.Forms.DockStyle.Right;
		this.buttonPanel.Location = new System.Drawing.Point(780, 0);
		this.buttonPanel.Name = "buttonPanel";
		this.buttonPanel.Size = new System.Drawing.Size(89, 36);
		this.buttonPanel.TabIndex = 32;
		this.SubmitButton.Dock = System.Windows.Forms.DockStyle.Top;
		this.SubmitButton.Font = new System.Drawing.Font("MS UI Gothic", 11.25f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128);
		this.SubmitButton.ImageKey = "check.png";
		this.SubmitButton.ImageList = this.imageList1;
		this.SubmitButton.Location = new System.Drawing.Point(0, 0);
		this.SubmitButton.Name = "SubmitButton";
		this.SubmitButton.Size = new System.Drawing.Size(89, 33);
		this.SubmitButton.TabIndex = 201;
		this.SubmitButton.Text = "button1";
		this.SubmitButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
		this.SubmitButton.UseVisualStyleBackColor = true;
		this.SubmitButton.Click += new System.EventHandler(SubmitButton_Click);
		this.imageList1.ImageStream = (System.Windows.Forms.ImageListStreamer)resources.GetObject("imageList1.ImageStream");
		this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
		this.imageList1.Images.SetKeyName(0, "check.png");
		this.line.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
		this.line.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.line.Location = new System.Drawing.Point(0, 36);
		this.line.Margin = new System.Windows.Forms.Padding(0);
		this.line.Name = "line";
		this.line.Size = new System.Drawing.Size(869, 3);
		this.line.TabIndex = 30;
		this.mailTable.ColumnCount = 2;
		this.mailTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 150f));
		this.mailTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.mailTable.Controls.Add(this.sslLabel, 0, 6);
		this.mailTable.Controls.Add(this.portText, 1, 2);
		this.mailTable.Controls.Add(this.sslEnable, 1, 6);
		this.mailTable.Controls.Add(this.rePasswordText, 1, 5);
		this.mailTable.Controls.Add(this.rePasswordLabel, 0, 5);
		this.mailTable.Controls.Add(this.passwordText, 1, 4);
		this.mailTable.Controls.Add(this.passwordLabel, 0, 4);
		this.mailTable.Controls.Add(this.accountText, 1, 3);
		this.mailTable.Controls.Add(this.accountLabel, 0, 3);
		this.mailTable.Controls.Add(this.portLabel, 0, 2);
		this.mailTable.Controls.Add(this.severText, 1, 1);
		this.mailTable.Controls.Add(this.maillServerLabel, 0, 1);
		this.mailTable.Controls.Add(this.mailEnable, 1, 0);
		this.mailTable.Controls.Add(this.toGrid, 1, 7);
		this.mailTable.Controls.Add(this.toAddressLabel, 0, 7);
		this.mailTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.mailTable.Location = new System.Drawing.Point(6, 51);
		this.mailTable.Name = "mailTable";
		this.mailTable.RowCount = 8;
		this.mailTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.mailTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.mailTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.mailTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.mailTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.mailTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.mailTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.mailTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.mailTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20f));
		this.mailTable.Size = new System.Drawing.Size(869, 406);
		this.mailTable.TabIndex = 3;
		this.sslLabel.AutoSize = true;
		this.sslLabel.BackColor = System.Drawing.SystemColors.Control;
		this.sslLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.sslLabel.Location = new System.Drawing.Point(3, 183);
		this.sslLabel.Margin = new System.Windows.Forms.Padding(3);
		this.sslLabel.Name = "sslLabel";
		this.sslLabel.Size = new System.Drawing.Size(144, 24);
		this.sslLabel.TabIndex = 33;
		this.sslLabel.Text = "password";
		this.sslLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.portText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.portText.Location = new System.Drawing.Point(153, 68);
		this.portText.Name = "portText";
		this.portText.Size = new System.Drawing.Size(713, 19);
		this.portText.TabIndex = 103;
		this.sslEnable.AutoSize = true;
		this.sslEnable.Dock = System.Windows.Forms.DockStyle.Left;
		this.sslEnable.Location = new System.Drawing.Point(153, 183);
		this.sslEnable.Name = "sslEnable";
		this.sslEnable.Size = new System.Drawing.Size(80, 24);
		this.sslEnable.TabIndex = 107;
		this.sslEnable.Text = "checkBox1";
		this.sslEnable.UseVisualStyleBackColor = true;
		this.rePasswordText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.rePasswordText.Location = new System.Drawing.Point(153, 158);
		this.rePasswordText.Name = "rePasswordText";
		this.rePasswordText.PasswordChar = '*';
		this.rePasswordText.Size = new System.Drawing.Size(713, 19);
		this.rePasswordText.TabIndex = 106;
		this.rePasswordLabel.AutoSize = true;
		this.rePasswordLabel.BackColor = System.Drawing.SystemColors.Control;
		this.rePasswordLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.rePasswordLabel.Location = new System.Drawing.Point(3, 153);
		this.rePasswordLabel.Margin = new System.Windows.Forms.Padding(3);
		this.rePasswordLabel.Name = "rePasswordLabel";
		this.rePasswordLabel.Size = new System.Drawing.Size(144, 24);
		this.rePasswordLabel.TabIndex = 10;
		this.rePasswordLabel.Text = "password";
		this.rePasswordLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.passwordText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.passwordText.Location = new System.Drawing.Point(153, 128);
		this.passwordText.Name = "passwordText";
		this.passwordText.PasswordChar = '*';
		this.passwordText.Size = new System.Drawing.Size(713, 19);
		this.passwordText.TabIndex = 105;
		this.passwordLabel.AutoSize = true;
		this.passwordLabel.BackColor = System.Drawing.SystemColors.Control;
		this.passwordLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.passwordLabel.Location = new System.Drawing.Point(3, 123);
		this.passwordLabel.Margin = new System.Windows.Forms.Padding(3);
		this.passwordLabel.Name = "passwordLabel";
		this.passwordLabel.Size = new System.Drawing.Size(144, 24);
		this.passwordLabel.TabIndex = 3;
		this.passwordLabel.Text = "password";
		this.passwordLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.accountText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.accountText.Location = new System.Drawing.Point(153, 98);
		this.accountText.Name = "accountText";
		this.accountText.Size = new System.Drawing.Size(713, 19);
		this.accountText.TabIndex = 104;
		this.accountLabel.AutoSize = true;
		this.accountLabel.BackColor = System.Drawing.SystemColors.Control;
		this.accountLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.accountLabel.Location = new System.Drawing.Point(3, 93);
		this.accountLabel.Margin = new System.Windows.Forms.Padding(3);
		this.accountLabel.Name = "accountLabel";
		this.accountLabel.Size = new System.Drawing.Size(144, 24);
		this.accountLabel.TabIndex = 1;
		this.accountLabel.Text = "account";
		this.accountLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.portLabel.AutoSize = true;
		this.portLabel.BackColor = System.Drawing.SystemColors.Control;
		this.portLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.portLabel.Location = new System.Drawing.Point(3, 63);
		this.portLabel.Margin = new System.Windows.Forms.Padding(3);
		this.portLabel.Name = "portLabel";
		this.portLabel.Size = new System.Drawing.Size(144, 24);
		this.portLabel.TabIndex = 2;
		this.portLabel.Text = "port";
		this.portLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.severText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.severText.Location = new System.Drawing.Point(153, 38);
		this.severText.Name = "severText";
		this.severText.Size = new System.Drawing.Size(713, 19);
		this.severText.TabIndex = 102;
		this.maillServerLabel.AutoSize = true;
		this.maillServerLabel.BackColor = System.Drawing.SystemColors.Control;
		this.maillServerLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.maillServerLabel.Location = new System.Drawing.Point(3, 33);
		this.maillServerLabel.Margin = new System.Windows.Forms.Padding(3);
		this.maillServerLabel.Name = "maillServerLabel";
		this.maillServerLabel.Size = new System.Drawing.Size(144, 24);
		this.maillServerLabel.TabIndex = 0;
		this.maillServerLabel.Text = "server";
		this.maillServerLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.mailEnable.AutoSize = true;
		this.mailEnable.Checked = true;
		this.mailEnable.CheckState = System.Windows.Forms.CheckState.Checked;
		this.mailEnable.Dock = System.Windows.Forms.DockStyle.Left;
		this.mailEnable.Location = new System.Drawing.Point(153, 3);
		this.mailEnable.Name = "mailEnable";
		this.mailEnable.Size = new System.Drawing.Size(80, 24);
		this.mailEnable.TabIndex = 101;
		this.mailEnable.Text = "checkBox2";
		this.mailEnable.UseVisualStyleBackColor = true;
		this.mailEnable.CheckedChanged += new System.EventHandler(mailEnable_CheckedChanged);
		this.toGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
		this.toGrid.Columns.AddRange(this.address);
		this.toGrid.Dock = System.Windows.Forms.DockStyle.Fill;
		this.toGrid.Location = new System.Drawing.Point(153, 213);
		this.toGrid.Name = "toGrid";
		this.toGrid.RowTemplate.Height = 21;
		this.toGrid.Size = new System.Drawing.Size(713, 190);
		this.toGrid.TabIndex = 108;
		this.address.HeaderText = "address";
		this.address.Name = "address";
		this.address.Width = 500;
		this.toAddressLabel.AutoSize = true;
		this.toAddressLabel.BackColor = System.Drawing.SystemColors.Control;
		this.toAddressLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.toAddressLabel.Location = new System.Drawing.Point(3, 213);
		this.toAddressLabel.Margin = new System.Windows.Forms.Padding(3);
		this.toAddressLabel.Name = "toAddressLabel";
		this.toAddressLabel.Size = new System.Drawing.Size(144, 190);
		this.toAddressLabel.TabIndex = 14;
		this.toAddressLabel.Text = "to";
		this.toAddressLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		base.Controls.Add(this.baseTable);
		base.Name = "MailSettingView";
		base.Size = new System.Drawing.Size(881, 463);
		base.Load += new System.EventHandler(JobListView_Load);
		this.baseTable.ResumeLayout(false);
		this.controllPanel.ResumeLayout(false);
		this.buttonPanel.ResumeLayout(false);
		this.mailTable.ResumeLayout(false);
		this.mailTable.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.toGrid).EndInit();
		base.ResumeLayout(false);
	}
}
