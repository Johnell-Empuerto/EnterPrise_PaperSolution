using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using ConMasGeneratorLib.Util;
using ConMasGeneratorUtility.Controls;
using ConMasGeneratorUtility.Properties;

namespace ConMasGeneratorUtility;

public class AppSettingView : UserControl
{
	private IContainer components = null;

	private TableLayoutPanel SettingTablePanel;

	private Button SubmitButton;

	private TableLayoutPanel baseTable;

	private Panel controllPanel;

	private Label line;

	private Panel buttonPanel;

	private Label Expect100Continue;

	private Label JobLogDeleteDays;

	private Label LocalAdminPassword;

	private Label Password;

	private Label LocalAdminUser;

	private Label ServerUrl;

	private Label User;

	private TextBox localPasswordText;

	private TextBox localUserText;

	private TextBox passwordText;

	private TextBox userText;

	private TextBox serverUrlText;

	private CheckBox isExpect100Continue;

	private Panel panel1;

	private NumericTextBox deleteDaysText;

	private ImageList imageList1;

	private Label TimeOut;

	private NumericTextBox TimeoutText;

	private Label ProxyPasswordLabel;

	private CheckBox isProxy;

	private NumericTextBox proxyPortText;

	private Label proxyPortLabel;

	private TextBox proxyPasswordText;

	private TextBox proxyUserText;

	private TextBox proxyAddressText;

	private Label proxyUserLabel;

	private Label proxyAddressLabel;

	private Panel panel2;

	public AppSettingView()
	{
		InitializeComponent();
		SetFormResouce();
	}

	private void AppSettingView_Load(object sender, EventArgs e)
	{
		try
		{
			SetConfig();
			serverUrlText.Focus();
			ChangeProxyEnable();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void SetConfig()
	{
		try
		{
			serverUrlText.Text = Utility.GetServerUrl();
			userText.Text = Utility.GetUser();
			passwordText.Text = Utility.GetPassword();
			localUserText.Text = Utility.GetLocalAdminUser();
			localPasswordText.Text = Utility.GetLocalAdminPassword();
			deleteDaysText.Text = Utility.GetJobLogDeleteDays().ToString();
			isExpect100Continue.Checked = Utility.GetExpect100Continue();
			TimeoutText.Text = Utility.GetTimeoutValue();
			isProxy.Checked = Utility.GetIsProxy();
			proxyAddressText.Text = Utility.GetProxyAddress();
			proxyPortText.Text = Utility.GetProxyPort();
			proxyUserText.Text = Utility.GetProxyUser();
			proxyPasswordText.Text = Utility.GetProxyPassword();
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private string ReplaceResouceName(string value)
	{
		try
		{
			string empty = string.Empty;
			empty = Resources.ResourceManager.GetString(value);
			if (string.IsNullOrEmpty(empty))
			{
				empty = value;
			}
			return empty;
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
			if (string.IsNullOrEmpty(serverUrlText.Text.Trim()))
			{
				MessageBox.Show(string.Format(Resources.NotEnteredMessage, Resources.ServerUrl), "WARNING");
				serverUrlText.Focus();
				return;
			}
			if (string.IsNullOrEmpty(userText.Text.Trim()))
			{
				MessageBox.Show(string.Format(Resources.NotEnteredMessage, Resources.User), "WARNING");
				userText.Focus();
				return;
			}
			if (string.IsNullOrEmpty(passwordText.Text.Trim()))
			{
				MessageBox.Show(string.Format(Resources.NotEnteredMessage, Resources.Password), "WARNING");
				passwordText.Focus();
				return;
			}
			if (string.IsNullOrEmpty(localUserText.Text.Trim()))
			{
				MessageBox.Show(string.Format(Resources.NotEnteredMessage, Resources.LocalAdminUser), "WARNING");
				localUserText.Focus();
				return;
			}
			if (string.IsNullOrEmpty(localPasswordText.Text.Trim()))
			{
				MessageBox.Show(string.Format(Resources.NotEnteredMessage, Resources.LocalAdminPassword), "WARNING");
				localPasswordText.Focus();
				return;
			}
			if (string.IsNullOrEmpty(deleteDaysText.Text.Trim()))
			{
				MessageBox.Show(string.Format(Resources.NotEnteredMessage, Resources.JobLogDeleteDays), "WARNING");
				deleteDaysText.Focus();
				return;
			}
			int result = 0;
			if (!int.TryParse(deleteDaysText.Text, out result))
			{
				MessageBox.Show(string.Format(Resources.NoneNumberMessage, Resources.JobLogDeleteDays), "WARNING");
				deleteDaysText.Focus();
				return;
			}
			int result2 = 0;
			if (!string.IsNullOrEmpty(TimeoutText.Text.Trim()) && !int.TryParse(TimeoutText.Text, out result2))
			{
				MessageBox.Show(string.Format(Resources.NoneNumberMessage, "Timeout"), "WARNING");
				TimeoutText.Focus();
			}
			else if (MessageBox.Show(Resources.SaveMessage, "INFO", MessageBoxButtons.OKCancel) != DialogResult.Cancel)
			{
				Utility.SetClientSetting("ServerUrl", serverUrlText.Text.Trim());
				Utility.SetClientSetting("User", userText.Text.Trim());
				Utility.SetClientSetting("Password", Utility.EncryptString(passwordText.Text.Trim(), Utility.encryptKey));
				Utility.SetClientSetting("LocalAdminUser", localUserText.Text.Trim());
				Utility.SetClientSetting("LocalAdminPassword", Utility.EncryptString(localPasswordText.Text.Trim(), Utility.encryptKey));
				Utility.SetClientSetting("JobLogDeleteDays", result.ToString());
				Utility.SetClientSetting("Expect100Continue", isExpect100Continue.Checked.ToString());
				Utility.SetClientSetting("Timeout", TimeoutText.Text.Trim());
				Utility.SetClientSetting("IsProxy", isProxy.Checked.ToString().ToUpper());
				Utility.SetClientSetting("ProxyAddress", proxyAddressText.Text.Trim());
				Utility.SetClientSetting("ProxyPort", proxyPortText.Text.Trim());
				Utility.SetClientSetting("ProxyUser", proxyUserText.Text.Trim());
				Utility.SetClientSetting("ProxyPassword", proxyPasswordText.Text.Trim());
				SetConfig();
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
			ServerUrl.Text = Resources.ServerUrl;
			User.Text = Resources.User;
			Password.Text = Resources.Password;
			LocalAdminUser.Text = Resources.LocalAdminUser;
			LocalAdminPassword.Text = Resources.LocalAdminPassword;
			JobLogDeleteDays.Text = Resources.JobLogDeleteDays;
			Expect100Continue.Text = Resources.Expect100Continue;
			isExpect100Continue.Text = "";
			TimeOut.Text = "Timeout";
			isProxy.Text = Resources.IsProxy;
			proxyAddressLabel.Text = Resources.ProxyAddress;
			proxyPortLabel.Text = Resources.ProxyPort;
			proxyUserLabel.Text = Resources.ProxyUser;
			ProxyPasswordLabel.Text = Resources.ProxyPassword;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void SettingTablePanel_Paint(object sender, PaintEventArgs e)
	{
	}

	private void proxyUserText_TextChanged(object sender, EventArgs e)
	{
	}

	private void ChangeProxyEnable()
	{
		if (isProxy.Checked)
		{
			proxyAddressText.Enabled = true;
			proxyPortText.Enabled = true;
			proxyUserText.Enabled = true;
			proxyPasswordText.Enabled = true;
		}
		else
		{
			proxyAddressText.Enabled = false;
			proxyPortText.Enabled = false;
			proxyUserText.Enabled = false;
			proxyPasswordText.Enabled = false;
		}
	}

	private void isProxy_CheckedChanged(object sender, EventArgs e)
	{
		ChangeProxyEnable();
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
		System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConMasGeneratorUtility.AppSettingView));
		this.SettingTablePanel = new System.Windows.Forms.TableLayoutPanel();
		this.localPasswordText = new System.Windows.Forms.TextBox();
		this.localUserText = new System.Windows.Forms.TextBox();
		this.passwordText = new System.Windows.Forms.TextBox();
		this.userText = new System.Windows.Forms.TextBox();
		this.ServerUrl = new System.Windows.Forms.Label();
		this.User = new System.Windows.Forms.Label();
		this.Password = new System.Windows.Forms.Label();
		this.LocalAdminUser = new System.Windows.Forms.Label();
		this.LocalAdminPassword = new System.Windows.Forms.Label();
		this.JobLogDeleteDays = new System.Windows.Forms.Label();
		this.Expect100Continue = new System.Windows.Forms.Label();
		this.serverUrlText = new System.Windows.Forms.TextBox();
		this.panel1 = new System.Windows.Forms.Panel();
		this.deleteDaysText = new ConMasGeneratorUtility.Controls.NumericTextBox();
		this.TimeOut = new System.Windows.Forms.Label();
		this.isExpect100Continue = new System.Windows.Forms.CheckBox();
		this.TimeoutText = new ConMasGeneratorUtility.Controls.NumericTextBox();
		this.ProxyPasswordLabel = new System.Windows.Forms.Label();
		this.isProxy = new System.Windows.Forms.CheckBox();
		this.proxyPortLabel = new System.Windows.Forms.Label();
		this.proxyAddressText = new System.Windows.Forms.TextBox();
		this.proxyPasswordText = new System.Windows.Forms.TextBox();
		this.proxyUserText = new System.Windows.Forms.TextBox();
		this.proxyUserLabel = new System.Windows.Forms.Label();
		this.proxyAddressLabel = new System.Windows.Forms.Label();
		this.SubmitButton = new System.Windows.Forms.Button();
		this.imageList1 = new System.Windows.Forms.ImageList(this.components);
		this.baseTable = new System.Windows.Forms.TableLayoutPanel();
		this.controllPanel = new System.Windows.Forms.Panel();
		this.buttonPanel = new System.Windows.Forms.Panel();
		this.line = new System.Windows.Forms.Label();
		this.panel2 = new System.Windows.Forms.Panel();
		this.proxyPortText = new ConMasGeneratorUtility.Controls.NumericTextBox();
		this.SettingTablePanel.SuspendLayout();
		this.panel1.SuspendLayout();
		this.baseTable.SuspendLayout();
		this.controllPanel.SuspendLayout();
		this.buttonPanel.SuspendLayout();
		this.panel2.SuspendLayout();
		base.SuspendLayout();
		this.SettingTablePanel.AutoScroll = true;
		this.SettingTablePanel.AutoSize = true;
		this.SettingTablePanel.ColumnCount = 2;
		this.SettingTablePanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 170f));
		this.SettingTablePanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.SettingTablePanel.Controls.Add(this.localPasswordText, 1, 4);
		this.SettingTablePanel.Controls.Add(this.localUserText, 1, 3);
		this.SettingTablePanel.Controls.Add(this.passwordText, 1, 2);
		this.SettingTablePanel.Controls.Add(this.userText, 1, 1);
		this.SettingTablePanel.Controls.Add(this.ServerUrl, 0, 0);
		this.SettingTablePanel.Controls.Add(this.User, 0, 1);
		this.SettingTablePanel.Controls.Add(this.Password, 0, 2);
		this.SettingTablePanel.Controls.Add(this.LocalAdminUser, 0, 3);
		this.SettingTablePanel.Controls.Add(this.LocalAdminPassword, 0, 4);
		this.SettingTablePanel.Controls.Add(this.JobLogDeleteDays, 0, 5);
		this.SettingTablePanel.Controls.Add(this.Expect100Continue, 0, 6);
		this.SettingTablePanel.Controls.Add(this.serverUrlText, 1, 0);
		this.SettingTablePanel.Controls.Add(this.panel1, 1, 5);
		this.SettingTablePanel.Controls.Add(this.TimeOut, 0, 7);
		this.SettingTablePanel.Controls.Add(this.isExpect100Continue, 1, 6);
		this.SettingTablePanel.Controls.Add(this.TimeoutText, 1, 7);
		this.SettingTablePanel.Location = new System.Drawing.Point(3, 48);
		this.SettingTablePanel.Name = "SettingTablePanel";
		this.SettingTablePanel.Padding = new System.Windows.Forms.Padding(3);
		this.SettingTablePanel.RowCount = 8;
		this.SettingTablePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.SettingTablePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.SettingTablePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.SettingTablePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.SettingTablePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.SettingTablePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32f));
		this.SettingTablePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 33f));
		this.SettingTablePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 25f));
		this.SettingTablePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20f));
		this.SettingTablePanel.Size = new System.Drawing.Size(685, 271);
		this.SettingTablePanel.TabIndex = 3;
		this.SettingTablePanel.Paint += new System.Windows.Forms.PaintEventHandler(SettingTablePanel_Paint);
		this.localPasswordText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.localPasswordText.Location = new System.Drawing.Point(176, 156);
		this.localPasswordText.Name = "localPasswordText";
		this.localPasswordText.PasswordChar = '*';
		this.localPasswordText.Size = new System.Drawing.Size(503, 19);
		this.localPasswordText.TabIndex = 105;
		this.localUserText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.localUserText.Location = new System.Drawing.Point(176, 121);
		this.localUserText.Name = "localUserText";
		this.localUserText.Size = new System.Drawing.Size(503, 19);
		this.localUserText.TabIndex = 104;
		this.passwordText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.passwordText.Location = new System.Drawing.Point(176, 86);
		this.passwordText.Name = "passwordText";
		this.passwordText.PasswordChar = '*';
		this.passwordText.Size = new System.Drawing.Size(503, 19);
		this.passwordText.TabIndex = 103;
		this.userText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.userText.Location = new System.Drawing.Point(176, 51);
		this.userText.Name = "userText";
		this.userText.Size = new System.Drawing.Size(503, 19);
		this.userText.TabIndex = 102;
		this.ServerUrl.AutoSize = true;
		this.ServerUrl.BackColor = System.Drawing.SystemColors.Control;
		this.ServerUrl.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ServerUrl.Location = new System.Drawing.Point(6, 6);
		this.ServerUrl.Margin = new System.Windows.Forms.Padding(3);
		this.ServerUrl.Name = "ServerUrl";
		this.ServerUrl.Size = new System.Drawing.Size(164, 29);
		this.ServerUrl.TabIndex = 0;
		this.ServerUrl.Text = "label1";
		this.ServerUrl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.User.AutoSize = true;
		this.User.BackColor = System.Drawing.SystemColors.Control;
		this.User.Dock = System.Windows.Forms.DockStyle.Fill;
		this.User.Location = new System.Drawing.Point(6, 41);
		this.User.Margin = new System.Windows.Forms.Padding(3);
		this.User.Name = "User";
		this.User.Size = new System.Drawing.Size(164, 29);
		this.User.TabIndex = 1;
		this.User.Text = "label2";
		this.User.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Password.AutoSize = true;
		this.Password.BackColor = System.Drawing.SystemColors.Control;
		this.Password.Dock = System.Windows.Forms.DockStyle.Fill;
		this.Password.Location = new System.Drawing.Point(6, 76);
		this.Password.Margin = new System.Windows.Forms.Padding(3);
		this.Password.Name = "Password";
		this.Password.Size = new System.Drawing.Size(164, 29);
		this.Password.TabIndex = 32;
		this.Password.Text = "label3";
		this.Password.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.LocalAdminUser.AutoSize = true;
		this.LocalAdminUser.BackColor = System.Drawing.SystemColors.Control;
		this.LocalAdminUser.Dock = System.Windows.Forms.DockStyle.Fill;
		this.LocalAdminUser.Location = new System.Drawing.Point(6, 111);
		this.LocalAdminUser.Margin = new System.Windows.Forms.Padding(3);
		this.LocalAdminUser.Name = "LocalAdminUser";
		this.LocalAdminUser.Size = new System.Drawing.Size(164, 29);
		this.LocalAdminUser.TabIndex = 33;
		this.LocalAdminUser.Text = "label4";
		this.LocalAdminUser.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.LocalAdminPassword.AutoSize = true;
		this.LocalAdminPassword.BackColor = System.Drawing.SystemColors.Control;
		this.LocalAdminPassword.Dock = System.Windows.Forms.DockStyle.Fill;
		this.LocalAdminPassword.Location = new System.Drawing.Point(6, 146);
		this.LocalAdminPassword.Margin = new System.Windows.Forms.Padding(3);
		this.LocalAdminPassword.Name = "LocalAdminPassword";
		this.LocalAdminPassword.Size = new System.Drawing.Size(164, 29);
		this.LocalAdminPassword.TabIndex = 34;
		this.LocalAdminPassword.Text = "label5";
		this.LocalAdminPassword.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.JobLogDeleteDays.AutoSize = true;
		this.JobLogDeleteDays.BackColor = System.Drawing.SystemColors.Control;
		this.JobLogDeleteDays.Dock = System.Windows.Forms.DockStyle.Fill;
		this.JobLogDeleteDays.Location = new System.Drawing.Point(6, 181);
		this.JobLogDeleteDays.Margin = new System.Windows.Forms.Padding(3);
		this.JobLogDeleteDays.Name = "JobLogDeleteDays";
		this.JobLogDeleteDays.Size = new System.Drawing.Size(164, 26);
		this.JobLogDeleteDays.TabIndex = 35;
		this.JobLogDeleteDays.Text = "label6";
		this.JobLogDeleteDays.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Expect100Continue.AutoSize = true;
		this.Expect100Continue.BackColor = System.Drawing.SystemColors.Control;
		this.Expect100Continue.Dock = System.Windows.Forms.DockStyle.Fill;
		this.Expect100Continue.Location = new System.Drawing.Point(6, 213);
		this.Expect100Continue.Margin = new System.Windows.Forms.Padding(3);
		this.Expect100Continue.Name = "Expect100Continue";
		this.Expect100Continue.Size = new System.Drawing.Size(164, 27);
		this.Expect100Continue.TabIndex = 36;
		this.Expect100Continue.Text = "label7";
		this.Expect100Continue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.serverUrlText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.serverUrlText.Location = new System.Drawing.Point(176, 16);
		this.serverUrlText.Name = "serverUrlText";
		this.serverUrlText.Size = new System.Drawing.Size(503, 19);
		this.serverUrlText.TabIndex = 101;
		this.panel1.Controls.Add(this.deleteDaysText);
		this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
		this.panel1.Location = new System.Drawing.Point(173, 178);
		this.panel1.Margin = new System.Windows.Forms.Padding(0);
		this.panel1.Name = "panel1";
		this.panel1.Size = new System.Drawing.Size(509, 32);
		this.panel1.TabIndex = 106;
		this.deleteDaysText.Location = new System.Drawing.Point(3, 7);
		this.deleteDaysText.Name = "deleteDaysText";
		this.deleteDaysText.Size = new System.Drawing.Size(77, 19);
		this.deleteDaysText.TabIndex = 106;
		this.deleteDaysText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
		this.TimeOut.AutoSize = true;
		this.TimeOut.BackColor = System.Drawing.SystemColors.Control;
		this.TimeOut.Dock = System.Windows.Forms.DockStyle.Fill;
		this.TimeOut.Location = new System.Drawing.Point(6, 246);
		this.TimeOut.Margin = new System.Windows.Forms.Padding(3);
		this.TimeOut.Name = "TimeOut";
		this.TimeOut.Size = new System.Drawing.Size(164, 19);
		this.TimeOut.TabIndex = 109;
		this.TimeOut.Text = "T";
		this.TimeOut.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.isExpect100Continue.AutoSize = true;
		this.isExpect100Continue.Dock = System.Windows.Forms.DockStyle.Left;
		this.isExpect100Continue.Location = new System.Drawing.Point(176, 213);
		this.isExpect100Continue.Name = "isExpect100Continue";
		this.isExpect100Continue.Size = new System.Drawing.Size(80, 27);
		this.isExpect100Continue.TabIndex = 107;
		this.isExpect100Continue.Text = "checkBox1";
		this.isExpect100Continue.UseVisualStyleBackColor = true;
		this.TimeoutText.Location = new System.Drawing.Point(176, 246);
		this.TimeoutText.Name = "TimeoutText";
		this.TimeoutText.Size = new System.Drawing.Size(121, 19);
		this.TimeoutText.TabIndex = 110;
		this.TimeoutText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
		this.ProxyPasswordLabel.BackColor = System.Drawing.SystemColors.Control;
		this.ProxyPasswordLabel.Location = new System.Drawing.Point(262, 63);
		this.ProxyPasswordLabel.Margin = new System.Windows.Forms.Padding(3);
		this.ProxyPasswordLabel.Name = "ProxyPasswordLabel";
		this.ProxyPasswordLabel.Padding = new System.Windows.Forms.Padding(3);
		this.ProxyPasswordLabel.Size = new System.Drawing.Size(76, 18);
		this.ProxyPasswordLabel.TabIndex = 116;
		this.ProxyPasswordLabel.Text = "T";
		this.ProxyPasswordLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
		this.isProxy.AutoSize = true;
		this.isProxy.Location = new System.Drawing.Point(8, 12);
		this.isProxy.Name = "isProxy";
		this.isProxy.Size = new System.Drawing.Size(80, 16);
		this.isProxy.TabIndex = 113;
		this.isProxy.Text = "checkBox1";
		this.isProxy.UseVisualStyleBackColor = true;
		this.isProxy.CheckedChanged += new System.EventHandler(isProxy_CheckedChanged);
		this.proxyPortLabel.BackColor = System.Drawing.SystemColors.Control;
		this.proxyPortLabel.Location = new System.Drawing.Point(260, 34);
		this.proxyPortLabel.Margin = new System.Windows.Forms.Padding(3);
		this.proxyPortLabel.Name = "proxyPortLabel";
		this.proxyPortLabel.Padding = new System.Windows.Forms.Padding(3);
		this.proxyPortLabel.Size = new System.Drawing.Size(79, 18);
		this.proxyPortLabel.TabIndex = 121;
		this.proxyPortLabel.Text = "T";
		this.proxyPortLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
		this.proxyAddressText.Location = new System.Drawing.Point(92, 34);
		this.proxyAddressText.Name = "proxyAddressText";
		this.proxyAddressText.Size = new System.Drawing.Size(142, 19);
		this.proxyAddressText.TabIndex = 114;
		this.proxyPasswordText.Location = new System.Drawing.Point(344, 63);
		this.proxyPasswordText.Name = "proxyPasswordText";
		this.proxyPasswordText.PasswordChar = '*';
		this.proxyPasswordText.Size = new System.Drawing.Size(142, 19);
		this.proxyPasswordText.TabIndex = 116;
		this.proxyUserText.Location = new System.Drawing.Point(92, 63);
		this.proxyUserText.Name = "proxyUserText";
		this.proxyUserText.Size = new System.Drawing.Size(142, 19);
		this.proxyUserText.TabIndex = 117;
		this.proxyUserText.TextChanged += new System.EventHandler(proxyUserText_TextChanged);
		this.proxyUserLabel.BackColor = System.Drawing.SystemColors.Control;
		this.proxyUserLabel.Location = new System.Drawing.Point(10, 63);
		this.proxyUserLabel.Margin = new System.Windows.Forms.Padding(3);
		this.proxyUserLabel.Name = "proxyUserLabel";
		this.proxyUserLabel.Padding = new System.Windows.Forms.Padding(3);
		this.proxyUserLabel.Size = new System.Drawing.Size(77, 18);
		this.proxyUserLabel.TabIndex = 126;
		this.proxyUserLabel.Text = "T";
		this.proxyUserLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
		this.proxyAddressLabel.BackColor = System.Drawing.SystemColors.Control;
		this.proxyAddressLabel.Location = new System.Drawing.Point(8, 34);
		this.proxyAddressLabel.Margin = new System.Windows.Forms.Padding(3);
		this.proxyAddressLabel.Name = "proxyAddressLabel";
		this.proxyAddressLabel.Padding = new System.Windows.Forms.Padding(3);
		this.proxyAddressLabel.Size = new System.Drawing.Size(79, 18);
		this.proxyAddressLabel.TabIndex = 127;
		this.proxyAddressLabel.Text = "T";
		this.proxyAddressLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
		this.SubmitButton.Dock = System.Windows.Forms.DockStyle.Top;
		this.SubmitButton.Font = new System.Drawing.Font("MS UI Gothic", 11.25f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128);
		this.SubmitButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.SubmitButton.ImageKey = "check.png";
		this.SubmitButton.ImageList = this.imageList1;
		this.SubmitButton.Location = new System.Drawing.Point(0, 0);
		this.SubmitButton.Name = "SubmitButton";
		this.SubmitButton.Padding = new System.Windows.Forms.Padding(3, 0, 0, 0);
		this.SubmitButton.Size = new System.Drawing.Size(89, 33);
		this.SubmitButton.TabIndex = 301;
		this.SubmitButton.Text = "button1";
		this.SubmitButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
		this.SubmitButton.UseVisualStyleBackColor = true;
		this.SubmitButton.Click += new System.EventHandler(SubmitButton_Click);
		this.imageList1.ImageStream = (System.Windows.Forms.ImageListStreamer)resources.GetObject("imageList1.ImageStream");
		this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
		this.imageList1.Images.SetKeyName(0, "check.png");
		this.baseTable.ColumnCount = 1;
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.baseTable.Controls.Add(this.SettingTablePanel, 0, 1);
		this.baseTable.Controls.Add(this.controllPanel, 0, 0);
		this.baseTable.Controls.Add(this.panel2, 0, 2);
		this.baseTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.baseTable.Location = new System.Drawing.Point(0, 0);
		this.baseTable.Name = "baseTable";
		this.baseTable.RowCount = 3;
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 277f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 86f));
		this.baseTable.Size = new System.Drawing.Size(691, 458);
		this.baseTable.TabIndex = 6;
		this.controllPanel.BackColor = System.Drawing.SystemColors.Control;
		this.controllPanel.Controls.Add(this.buttonPanel);
		this.controllPanel.Controls.Add(this.line);
		this.controllPanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.controllPanel.Location = new System.Drawing.Point(3, 3);
		this.controllPanel.Name = "controllPanel";
		this.controllPanel.Size = new System.Drawing.Size(685, 39);
		this.controllPanel.TabIndex = 4;
		this.buttonPanel.Controls.Add(this.SubmitButton);
		this.buttonPanel.Dock = System.Windows.Forms.DockStyle.Right;
		this.buttonPanel.Location = new System.Drawing.Point(596, 0);
		this.buttonPanel.Name = "buttonPanel";
		this.buttonPanel.Size = new System.Drawing.Size(89, 36);
		this.buttonPanel.TabIndex = 31;
		this.line.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
		this.line.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.line.Location = new System.Drawing.Point(0, 36);
		this.line.Margin = new System.Windows.Forms.Padding(0);
		this.line.Name = "line";
		this.line.Size = new System.Drawing.Size(685, 3);
		this.line.TabIndex = 30;
		this.panel2.Controls.Add(this.proxyAddressLabel);
		this.panel2.Controls.Add(this.proxyPortLabel);
		this.panel2.Controls.Add(this.ProxyPasswordLabel);
		this.panel2.Controls.Add(this.proxyUserText);
		this.panel2.Controls.Add(this.proxyPortText);
		this.panel2.Controls.Add(this.isProxy);
		this.panel2.Controls.Add(this.proxyAddressText);
		this.panel2.Controls.Add(this.proxyPasswordText);
		this.panel2.Controls.Add(this.proxyUserLabel);
		this.panel2.Location = new System.Drawing.Point(3, 325);
		this.panel2.Name = "panel2";
		this.panel2.Size = new System.Drawing.Size(679, 108);
		this.panel2.TabIndex = 5;
		this.proxyPortText.Location = new System.Drawing.Point(343, 34);
		this.proxyPortText.Name = "proxyPortText";
		this.proxyPortText.Size = new System.Drawing.Size(142, 19);
		this.proxyPortText.TabIndex = 115;
		this.proxyPortText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.White;
		base.Controls.Add(this.baseTable);
		base.Name = "AppSettingView";
		base.Size = new System.Drawing.Size(691, 458);
		base.Load += new System.EventHandler(AppSettingView_Load);
		this.SettingTablePanel.ResumeLayout(false);
		this.SettingTablePanel.PerformLayout();
		this.panel1.ResumeLayout(false);
		this.panel1.PerformLayout();
		this.baseTable.ResumeLayout(false);
		this.baseTable.PerformLayout();
		this.controllPanel.ResumeLayout(false);
		this.buttonPanel.ResumeLayout(false);
		this.panel2.ResumeLayout(false);
		this.panel2.PerformLayout();
		base.ResumeLayout(false);
	}
}
