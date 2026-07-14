using System;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using ConMasGeneratorUtility.Properties;

namespace ConMasGeneratorUtility.Views;

public class ConMasLegendViewer : Form
{
	private IContainer components = null;

	private TableLayoutPanel baseTable;

	private Button button0;

	private Label ConMasJobId;

	private Button button1;

	private Button button2;

	private Button button5;

	private Button button4;

	private Button button3;

	private Label ConMasJobName;

	private Button button8;

	private Label ConMasProcessIndex;

	private Label ConMasProcessName;

	private Label ConMasCommandIndex;

	private Button button6;

	private Button button14;

	private Button button7;

	private Button button13;

	private Label ConMasCommandName;

	private Label ConMasWatcherFile;

	private Label DateTimeOffsetFormat;

	private Label ConMasServerUrl;

	private Label ConMasUser;

	private Label DateTimeFormat;

	private Button button9;

	private Label ConMasPassword;

	private Button button11;

	private Label ConMasVersion;

	private Button button10;

	private Label ConMasClient;

	private Button button12;

	private Label ConMasTerminalId;

	public ConMasLegendViewer()
	{
		InitializeComponent();
		SetFormResouce();
		Type typeFromHandle = typeof(DataGridView);
		PropertyInfo property = typeFromHandle.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
		property.SetValue(this, true, null);
	}

	private void ConMasProcessViewer_Load(object sender, EventArgs e)
	{
		try
		{
			SetGrid();
			baseTable.RowStyles[11].Height = 0f;
			baseTable.RowStyles[12].Height = 0f;
			baseTable.RowStyles[13].Height = 0f;
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
			Text = Resources.LegendViewer;
			ConMasJobId.Text = Resources.ConMasJobId;
			ConMasJobName.Text = Resources.ConMasJobName;
			ConMasProcessIndex.Text = Resources.ConMasProcessIndex;
			ConMasProcessName.Text = Resources.ConMasProcessName;
			ConMasCommandIndex.Text = Resources.ConMasCommandIndex;
			ConMasCommandName.Text = Resources.ConMasCommandName;
			DateTimeFormat.Text = Resources.DateTimeFormat;
			ConMasWatcherFile.Text = Resources.ConMasWatcherFile;
			DateTimeOffsetFormat.Text = Resources.DateTimeOffsetFormat;
			ConMasServerUrl.Text = Resources.ConMasServerUrl;
			ConMasUser.Text = Resources.ConMasUser;
			ConMasPassword.Text = Resources.ConMasPassword;
			ConMasClient.Text = Resources.ConMasClient;
			ConMasVersion.Text = Resources.ConMasVersion;
			ConMasTerminalId.Text = Resources.ConMasTerminalId;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void copyButtonClick(object sender, EventArgs e)
	{
		try
		{
			Clipboard.SetDataObject((sender as Button).Text, copy: true);
			MessageBox.Show("Save Clipboard", "INFO");
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
		this.baseTable = new System.Windows.Forms.TableLayoutPanel();
		this.button12 = new System.Windows.Forms.Button();
		this.ConMasTerminalId = new System.Windows.Forms.Label();
		this.button0 = new System.Windows.Forms.Button();
		this.ConMasJobId = new System.Windows.Forms.Label();
		this.button1 = new System.Windows.Forms.Button();
		this.button2 = new System.Windows.Forms.Button();
		this.button5 = new System.Windows.Forms.Button();
		this.button4 = new System.Windows.Forms.Button();
		this.button3 = new System.Windows.Forms.Button();
		this.ConMasJobName = new System.Windows.Forms.Label();
		this.ConMasProcessIndex = new System.Windows.Forms.Label();
		this.ConMasProcessName = new System.Windows.Forms.Label();
		this.ConMasCommandIndex = new System.Windows.Forms.Label();
		this.button6 = new System.Windows.Forms.Button();
		this.ConMasCommandName = new System.Windows.Forms.Label();
		this.ConMasWatcherFile = new System.Windows.Forms.Label();
		this.button11 = new System.Windows.Forms.Button();
		this.ConMasVersion = new System.Windows.Forms.Label();
		this.button10 = new System.Windows.Forms.Button();
		this.ConMasClient = new System.Windows.Forms.Label();
		this.button9 = new System.Windows.Forms.Button();
		this.ConMasPassword = new System.Windows.Forms.Label();
		this.button8 = new System.Windows.Forms.Button();
		this.ConMasUser = new System.Windows.Forms.Label();
		this.button7 = new System.Windows.Forms.Button();
		this.ConMasServerUrl = new System.Windows.Forms.Label();
		this.button14 = new System.Windows.Forms.Button();
		this.DateTimeOffsetFormat = new System.Windows.Forms.Label();
		this.button13 = new System.Windows.Forms.Button();
		this.DateTimeFormat = new System.Windows.Forms.Label();
		this.baseTable.SuspendLayout();
		base.SuspendLayout();
		this.baseTable.ColumnCount = 2;
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 316f));
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.baseTable.Controls.Add(this.button12, 0, 14);
		this.baseTable.Controls.Add(this.ConMasTerminalId, 0, 14);
		this.baseTable.Controls.Add(this.button0, 0, 0);
		this.baseTable.Controls.Add(this.ConMasJobId, 1, 0);
		this.baseTable.Controls.Add(this.button1, 0, 1);
		this.baseTable.Controls.Add(this.button2, 0, 2);
		this.baseTable.Controls.Add(this.button5, 0, 3);
		this.baseTable.Controls.Add(this.button4, 0, 4);
		this.baseTable.Controls.Add(this.button3, 0, 5);
		this.baseTable.Controls.Add(this.ConMasJobName, 1, 1);
		this.baseTable.Controls.Add(this.ConMasProcessIndex, 1, 2);
		this.baseTable.Controls.Add(this.ConMasProcessName, 1, 3);
		this.baseTable.Controls.Add(this.ConMasCommandIndex, 1, 4);
		this.baseTable.Controls.Add(this.button6, 0, 6);
		this.baseTable.Controls.Add(this.ConMasCommandName, 1, 5);
		this.baseTable.Controls.Add(this.ConMasWatcherFile, 1, 6);
		this.baseTable.Controls.Add(this.button11, 0, 13);
		this.baseTable.Controls.Add(this.ConMasVersion, 1, 13);
		this.baseTable.Controls.Add(this.button10, 0, 12);
		this.baseTable.Controls.Add(this.ConMasClient, 1, 12);
		this.baseTable.Controls.Add(this.button9, 0, 11);
		this.baseTable.Controls.Add(this.ConMasPassword, 1, 11);
		this.baseTable.Controls.Add(this.button8, 0, 10);
		this.baseTable.Controls.Add(this.ConMasUser, 1, 10);
		this.baseTable.Controls.Add(this.button7, 0, 9);
		this.baseTable.Controls.Add(this.ConMasServerUrl, 1, 9);
		this.baseTable.Controls.Add(this.button14, 0, 8);
		this.baseTable.Controls.Add(this.DateTimeOffsetFormat, 1, 8);
		this.baseTable.Controls.Add(this.button13, 0, 7);
		this.baseTable.Controls.Add(this.DateTimeFormat, 1, 7);
		this.baseTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.baseTable.Location = new System.Drawing.Point(3, 3);
		this.baseTable.Name = "baseTable";
		this.baseTable.RowCount = 16;
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20f));
		this.baseTable.Size = new System.Drawing.Size(857, 388);
		this.baseTable.TabIndex = 3;
		this.button12.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button12.Location = new System.Drawing.Point(3, 493);
		this.button12.Name = "button12";
		this.button12.Size = new System.Drawing.Size(310, 29);
		this.button12.TabIndex = 35;
		this.button12.Text = "{ConMasTerminalId}";
		this.button12.UseVisualStyleBackColor = true;
		this.button12.Visible = false;
		this.ConMasTerminalId.AutoSize = true;
		this.ConMasTerminalId.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasTerminalId.Location = new System.Drawing.Point(319, 490);
		this.ConMasTerminalId.Name = "ConMasTerminalId";
		this.ConMasTerminalId.Size = new System.Drawing.Size(535, 35);
		this.ConMasTerminalId.TabIndex = 34;
		this.ConMasTerminalId.Text = "label11";
		this.ConMasTerminalId.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.ConMasTerminalId.Visible = false;
		this.button0.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button0.Location = new System.Drawing.Point(3, 3);
		this.button0.Name = "button0";
		this.button0.Size = new System.Drawing.Size(310, 29);
		this.button0.TabIndex = 0;
		this.button0.Text = "{ConMasJobId}";
		this.button0.UseVisualStyleBackColor = true;
		this.button0.Click += new System.EventHandler(copyButtonClick);
		this.ConMasJobId.AutoSize = true;
		this.ConMasJobId.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasJobId.Location = new System.Drawing.Point(319, 0);
		this.ConMasJobId.Name = "ConMasJobId";
		this.ConMasJobId.Size = new System.Drawing.Size(535, 35);
		this.ConMasJobId.TabIndex = 1;
		this.ConMasJobId.Text = "label1";
		this.ConMasJobId.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.button1.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button1.Location = new System.Drawing.Point(3, 38);
		this.button1.Name = "button1";
		this.button1.Size = new System.Drawing.Size(310, 29);
		this.button1.TabIndex = 2;
		this.button1.Text = "{ConMasJobName}";
		this.button1.UseVisualStyleBackColor = true;
		this.button1.Click += new System.EventHandler(copyButtonClick);
		this.button2.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button2.Location = new System.Drawing.Point(3, 73);
		this.button2.Name = "button2";
		this.button2.Size = new System.Drawing.Size(310, 29);
		this.button2.TabIndex = 3;
		this.button2.Text = "{ConMasProcessIndex}";
		this.button2.UseVisualStyleBackColor = true;
		this.button2.Click += new System.EventHandler(copyButtonClick);
		this.button5.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button5.Location = new System.Drawing.Point(3, 108);
		this.button5.Name = "button5";
		this.button5.Size = new System.Drawing.Size(310, 29);
		this.button5.TabIndex = 6;
		this.button5.Text = "{ConMasProcessName}";
		this.button5.UseVisualStyleBackColor = true;
		this.button5.Click += new System.EventHandler(copyButtonClick);
		this.button4.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button4.Location = new System.Drawing.Point(3, 143);
		this.button4.Name = "button4";
		this.button4.Size = new System.Drawing.Size(310, 29);
		this.button4.TabIndex = 8;
		this.button4.Text = "{ConMasCommandIndex}";
		this.button4.UseVisualStyleBackColor = true;
		this.button4.Click += new System.EventHandler(copyButtonClick);
		this.button3.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button3.Location = new System.Drawing.Point(3, 178);
		this.button3.Name = "button3";
		this.button3.Size = new System.Drawing.Size(310, 29);
		this.button3.TabIndex = 7;
		this.button3.Text = "{ConMasCommandName}";
		this.button3.UseVisualStyleBackColor = true;
		this.button3.Click += new System.EventHandler(copyButtonClick);
		this.ConMasJobName.AutoSize = true;
		this.ConMasJobName.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasJobName.Location = new System.Drawing.Point(319, 35);
		this.ConMasJobName.Name = "ConMasJobName";
		this.ConMasJobName.Size = new System.Drawing.Size(535, 35);
		this.ConMasJobName.TabIndex = 9;
		this.ConMasJobName.Text = "label1";
		this.ConMasJobName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.ConMasProcessIndex.AutoSize = true;
		this.ConMasProcessIndex.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasProcessIndex.Location = new System.Drawing.Point(319, 70);
		this.ConMasProcessIndex.Name = "ConMasProcessIndex";
		this.ConMasProcessIndex.Size = new System.Drawing.Size(535, 35);
		this.ConMasProcessIndex.TabIndex = 10;
		this.ConMasProcessIndex.Text = "label1";
		this.ConMasProcessIndex.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.ConMasProcessName.AutoSize = true;
		this.ConMasProcessName.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasProcessName.Location = new System.Drawing.Point(319, 105);
		this.ConMasProcessName.Name = "ConMasProcessName";
		this.ConMasProcessName.Size = new System.Drawing.Size(535, 35);
		this.ConMasProcessName.TabIndex = 11;
		this.ConMasProcessName.Text = "label2";
		this.ConMasProcessName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.ConMasCommandIndex.AutoSize = true;
		this.ConMasCommandIndex.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasCommandIndex.Location = new System.Drawing.Point(319, 140);
		this.ConMasCommandIndex.Name = "ConMasCommandIndex";
		this.ConMasCommandIndex.Size = new System.Drawing.Size(535, 35);
		this.ConMasCommandIndex.TabIndex = 12;
		this.ConMasCommandIndex.Text = "label3";
		this.ConMasCommandIndex.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.button6.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button6.Location = new System.Drawing.Point(3, 213);
		this.button6.Name = "button6";
		this.button6.Size = new System.Drawing.Size(310, 29);
		this.button6.TabIndex = 13;
		this.button6.Text = "{ConMasWatcherFile}";
		this.button6.UseVisualStyleBackColor = true;
		this.button6.Click += new System.EventHandler(copyButtonClick);
		this.ConMasCommandName.AutoSize = true;
		this.ConMasCommandName.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasCommandName.Location = new System.Drawing.Point(319, 175);
		this.ConMasCommandName.Name = "ConMasCommandName";
		this.ConMasCommandName.Size = new System.Drawing.Size(535, 35);
		this.ConMasCommandName.TabIndex = 25;
		this.ConMasCommandName.Text = "label4";
		this.ConMasCommandName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.ConMasWatcherFile.AutoSize = true;
		this.ConMasWatcherFile.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasWatcherFile.Location = new System.Drawing.Point(319, 210);
		this.ConMasWatcherFile.Name = "ConMasWatcherFile";
		this.ConMasWatcherFile.Size = new System.Drawing.Size(535, 35);
		this.ConMasWatcherFile.TabIndex = 26;
		this.ConMasWatcherFile.Text = "label5";
		this.ConMasWatcherFile.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.button11.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button11.Location = new System.Drawing.Point(3, 458);
		this.button11.Name = "button11";
		this.button11.Size = new System.Drawing.Size(310, 29);
		this.button11.TabIndex = 19;
		this.button11.Text = "{ConMasVersion}";
		this.button11.UseVisualStyleBackColor = true;
		this.button11.Visible = false;
		this.button11.Click += new System.EventHandler(copyButtonClick);
		this.ConMasVersion.AutoSize = true;
		this.ConMasVersion.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasVersion.Location = new System.Drawing.Point(319, 455);
		this.ConMasVersion.Name = "ConMasVersion";
		this.ConMasVersion.Size = new System.Drawing.Size(535, 35);
		this.ConMasVersion.TabIndex = 32;
		this.ConMasVersion.Text = "label10";
		this.ConMasVersion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.ConMasVersion.Visible = false;
		this.button10.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button10.Location = new System.Drawing.Point(3, 423);
		this.button10.Name = "button10";
		this.button10.Size = new System.Drawing.Size(310, 29);
		this.button10.TabIndex = 18;
		this.button10.Text = "{ConMasClient}";
		this.button10.UseVisualStyleBackColor = true;
		this.button10.Visible = false;
		this.button10.Click += new System.EventHandler(copyButtonClick);
		this.ConMasClient.AutoSize = true;
		this.ConMasClient.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasClient.Location = new System.Drawing.Point(319, 420);
		this.ConMasClient.Name = "ConMasClient";
		this.ConMasClient.Size = new System.Drawing.Size(535, 35);
		this.ConMasClient.TabIndex = 31;
		this.ConMasClient.Text = "label9";
		this.ConMasClient.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.ConMasClient.Visible = false;
		this.button9.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button9.Location = new System.Drawing.Point(3, 388);
		this.button9.Name = "button9";
		this.button9.Size = new System.Drawing.Size(310, 29);
		this.button9.TabIndex = 17;
		this.button9.Text = "{ConMasPassword}";
		this.button9.UseVisualStyleBackColor = true;
		this.button9.Visible = false;
		this.button9.Click += new System.EventHandler(copyButtonClick);
		this.ConMasPassword.AutoSize = true;
		this.ConMasPassword.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasPassword.Location = new System.Drawing.Point(319, 385);
		this.ConMasPassword.Name = "ConMasPassword";
		this.ConMasPassword.Size = new System.Drawing.Size(535, 35);
		this.ConMasPassword.TabIndex = 30;
		this.ConMasPassword.Text = "label8";
		this.ConMasPassword.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.ConMasPassword.Visible = false;
		this.button8.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button8.Location = new System.Drawing.Point(3, 353);
		this.button8.Name = "button8";
		this.button8.Size = new System.Drawing.Size(310, 29);
		this.button8.TabIndex = 16;
		this.button8.Text = "{ConMasUser}";
		this.button8.UseVisualStyleBackColor = true;
		this.button8.Click += new System.EventHandler(copyButtonClick);
		this.ConMasUser.AutoSize = true;
		this.ConMasUser.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasUser.Location = new System.Drawing.Point(319, 350);
		this.ConMasUser.Name = "ConMasUser";
		this.ConMasUser.Size = new System.Drawing.Size(535, 35);
		this.ConMasUser.TabIndex = 29;
		this.ConMasUser.Text = "label7";
		this.ConMasUser.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.button7.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button7.Location = new System.Drawing.Point(3, 318);
		this.button7.Name = "button7";
		this.button7.Size = new System.Drawing.Size(310, 29);
		this.button7.TabIndex = 15;
		this.button7.Text = "{ConMasServerUrl}";
		this.button7.UseVisualStyleBackColor = true;
		this.button7.Click += new System.EventHandler(copyButtonClick);
		this.ConMasServerUrl.AutoSize = true;
		this.ConMasServerUrl.Dock = System.Windows.Forms.DockStyle.Fill;
		this.ConMasServerUrl.Location = new System.Drawing.Point(319, 315);
		this.ConMasServerUrl.Name = "ConMasServerUrl";
		this.ConMasServerUrl.Size = new System.Drawing.Size(535, 35);
		this.ConMasServerUrl.TabIndex = 28;
		this.ConMasServerUrl.Text = "label6";
		this.ConMasServerUrl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.button14.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button14.Font = new System.Drawing.Font("MS UI Gothic", 7.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
		this.button14.Location = new System.Drawing.Point(3, 283);
		this.button14.Name = "button14";
		this.button14.Size = new System.Drawing.Size(310, 29);
		this.button14.TabIndex = 14;
		this.button14.Text = "{DateTime(format=\"yyyy/MM/dd\",offsetType=M,offsetValue=-2)}";
		this.button14.UseVisualStyleBackColor = true;
		this.button14.Click += new System.EventHandler(copyButtonClick);
		this.DateTimeOffsetFormat.AutoSize = true;
		this.DateTimeOffsetFormat.Dock = System.Windows.Forms.DockStyle.Fill;
		this.DateTimeOffsetFormat.Location = new System.Drawing.Point(319, 280);
		this.DateTimeOffsetFormat.Name = "DateTimeOffsetFormat";
		this.DateTimeOffsetFormat.Size = new System.Drawing.Size(535, 35);
		this.DateTimeOffsetFormat.TabIndex = 27;
		this.DateTimeOffsetFormat.Text = "label13";
		this.DateTimeOffsetFormat.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.button13.Dock = System.Windows.Forms.DockStyle.Fill;
		this.button13.Location = new System.Drawing.Point(3, 248);
		this.button13.Name = "button13";
		this.button13.Size = new System.Drawing.Size(310, 29);
		this.button13.TabIndex = 20;
		this.button13.Text = "{DateTime(format=\"yyyy/MM/dd\")}";
		this.button13.UseVisualStyleBackColor = true;
		this.button13.Click += new System.EventHandler(copyButtonClick);
		this.DateTimeFormat.AutoSize = true;
		this.DateTimeFormat.Dock = System.Windows.Forms.DockStyle.Fill;
		this.DateTimeFormat.Location = new System.Drawing.Point(319, 245);
		this.DateTimeFormat.Name = "DateTimeFormat";
		this.DateTimeFormat.Size = new System.Drawing.Size(535, 35);
		this.DateTimeFormat.TabIndex = 33;
		this.DateTimeFormat.Text = "label12";
		this.DateTimeFormat.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		base.ClientSize = new System.Drawing.Size(863, 394);
		base.Controls.Add(this.baseTable);
		base.Name = "ConMasLegendViewer";
		base.Padding = new System.Windows.Forms.Padding(3);
		base.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
		this.Text = "ConMasLegendViewer";
		base.Load += new System.EventHandler(ConMasProcessViewer_Load);
		this.baseTable.ResumeLayout(false);
		this.baseTable.PerformLayout();
		base.ResumeLayout(false);
	}
}
