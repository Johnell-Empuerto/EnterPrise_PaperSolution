using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using ConMasGeneratorLib.Data;
using ConMasGeneratorUtility.Controls;
using ConMasGeneratorUtility.Properties;

namespace ConMasGeneratorUtility.Views;

public class JobSchduleView : SettingViewBase
{
	private TreeNode _node;

	private JobData _jobData;

	private const int GRID_HEIGHT_ONE_TYPE = 35;

	private const int GRID_HEIGHT_DAY_TYPE = 75;

	private const int GRID_HEIGHT_WEEK_TYPE = 100;

	private const int GRID_HEIGHT_MONTH_TYPE = 260;

	private IContainer components = null;

	private DateTimePicker startDate;

	private DateTimePicker endDate;

	private RadioButton typeOne;

	private GroupBox schSettingGroup;

	private RadioButton typeWeek;

	private RadioButton typeMonth;

	private RadioButton typeDay;

	private CheckBox retryCheck;

	private NumericTextBox retryMinitesText;

	private Panel retryPanel;

	private CheckBox endCheck;

	private Panel weekCheck;

	private NumericUpDown startHour;

	private Panel monthPanel;

	private NumericUpDown startMinites;

	private Label label1;

	private NumericUpDown endMinites;

	private NumericUpDown endHour;

	private CheckedListBox dayCheckList;

	private CheckedListBox monthCheckList;

	private CheckedListBox weekCheckList;

	private CheckBox allMonthCheck;

	private CheckBox allWeekCheck;

	private Panel timingPanel;

	private Panel endTimePanel;

	private Label label3;

	private TextBox nameText;

	private TextBox remarkText;

	private Button downLoadFolderButton;

	private TextBox downLoadFolderText;

	private Button workFolderButton;

	private TextBox workFolderText;

	private FolderBrowserDialog folderBrowserDialog;

	private NumericTextBox timingText;

	private CheckBox allDayCheck;

	private TableLayoutPanel schduleJobPanel;

	private TableLayoutPanel typeSettingTable;

	private Panel startDatePanel;

	private Label startDateLabel;

	private Label timingLabel;

	private Label weekLabel;

	private Label dayLabel;

	private Label monthLabel;

	private TableLayoutPanel baseTable;

	private Label nameLabel;

	private Label workFolderLabel;

	private Label downloadFolderLabel;

	private Label remarkLabel;

	private TableLayoutPanel schBaseTable;

	private Panel selectPanel;

	private CheckBox enableCheckBox;

	private Panel otherPanel;

	private FlowLayoutPanel flowLayout;

	private Label line;

	private Button moveFolder;

	private Label workLabel;

	public JobSchduleView(TreeNode node)
	{
		InitializeComponent();
		SetFormResouce();
		_node = node;
		_jobData = node.Tag as JobData;
		DoubleBuffered = true;
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

	private void JobSchduleView_Load(object sender, EventArgs e)
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
			allMonthCheck.Text = Resources.All;
			allWeekCheck.Text = Resources.All;
			CheckedListBox.ObjectCollection items = monthCheckList.Items;
			object[] items2 = new string[12]
			{
				"1" + Resources.Month,
				"2" + Resources.Month,
				"3" + Resources.Month,
				"4" + Resources.Month,
				"5" + Resources.Month,
				"6" + Resources.Month,
				"7" + Resources.Month,
				"8" + Resources.Month,
				"9" + Resources.Month,
				"10" + Resources.Month,
				"11" + Resources.Month,
				"12" + Resources.Month
			};
			items.AddRange(items2);
			CheckedListBox.ObjectCollection items3 = weekCheckList.Items;
			items2 = new string[7]
			{
				Resources.Sunday,
				Resources.Monday,
				Resources.Tuesday,
				Resources.Wednesday,
				Resources.Thursday,
				Resources.Friday,
				Resources.Saturday
			};
			items3.AddRange(items2);
			CheckedListBox.ObjectCollection items4 = dayCheckList.Items;
			items2 = new string[31]
			{
				"1" + Resources.Day,
				"2" + Resources.Day,
				"3" + Resources.Day,
				"4" + Resources.Day,
				"5" + Resources.Day,
				"6" + Resources.Day,
				"7" + Resources.Day,
				"8" + Resources.Day,
				"9" + Resources.Day,
				"10" + Resources.Day,
				"11" + Resources.Day,
				"12" + Resources.Day,
				"13" + Resources.Day,
				"14" + Resources.Day,
				"15" + Resources.Day,
				"16" + Resources.Day,
				"17" + Resources.Day,
				"18" + Resources.Day,
				"19" + Resources.Day,
				"20" + Resources.Day,
				"21" + Resources.Day,
				"22" + Resources.Day,
				"23" + Resources.Day,
				"24" + Resources.Day,
				"25" + Resources.Day,
				"26" + Resources.Day,
				"27" + Resources.Day,
				"28" + Resources.Day,
				"29" + Resources.Day,
				"30" + Resources.Day,
				"31" + Resources.Day
			};
			items4.AddRange(items2);
			if (_jobData.Type == "4")
			{
				typeOne.Checked = true;
			}
			else if (_jobData.Type == "1")
			{
				typeDay.Checked = true;
			}
			else if (_jobData.Type == "2")
			{
				typeWeek.Checked = true;
			}
			else if (_jobData.Type == "3")
			{
				typeMonth.Checked = true;
			}
			if (!string.IsNullOrEmpty(_jobData.Timing))
			{
				timingText.Text = _jobData.Timing;
			}
			if (!string.IsNullOrEmpty(_jobData.Week))
			{
				string[] array = _jobData.Week.Split(',');
				foreach (string s in array)
				{
					weekCheckList.SetItemChecked(int.Parse(s), value: true);
				}
			}
			if (!string.IsNullOrEmpty(_jobData.Month))
			{
				string[] array2 = _jobData.Month.Split(',');
				foreach (string s2 in array2)
				{
					monthCheckList.SetItemChecked(int.Parse(s2) - 1, value: true);
				}
			}
			if (!string.IsNullOrEmpty(_jobData.Day))
			{
				string[] array3 = _jobData.Day.Split(',');
				foreach (string s3 in array3)
				{
					dayCheckList.SetItemChecked(int.Parse(s3) - 1, value: true);
				}
			}
			retryCheck.Checked = _jobData.IsRetry;
			retryMinitesText.Text = _jobData.RetryTiming;
			if (!string.IsNullOrEmpty(_jobData.StartTime))
			{
				startDate.Text = _jobData.StartTime;
				DateTime dateTime = DateTime.Parse(_jobData.StartTime);
				startDate.Text = dateTime.ToString("yyyy/MM/dd");
				startHour.Text = dateTime.ToString("HH");
				startMinites.Text = dateTime.ToString("mm");
			}
			else
			{
				DateTime now = DateTime.Now;
				startDate.Text = now.ToString("yyyy/MM/dd");
				startHour.Text = now.ToString("HH");
				startMinites.Text = now.ToString("mm");
				_jobData.StartTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
			}
			endCheck.Checked = _jobData.IsEndDate;
			if (!string.IsNullOrEmpty(_jobData.EndTime))
			{
				DateTime dateTime2 = DateTime.Parse(_jobData.EndTime);
				endDate.Text = dateTime2.ToString("yyyy/MM/dd");
				endHour.Text = dateTime2.ToString("HH");
				endMinites.Text = dateTime2.ToString("mm");
			}
			else
			{
				endCheck.Checked = false;
				DateTime now2 = DateTime.Now;
				endDate.Text = now2.ToString("yyyy/MM/dd");
				endHour.Text = "23";
				endMinites.Text = "59";
			}
			nameText.Text = _jobData.Name;
			workFolderText.Text = _jobData.WorkFolder;
			downLoadFolderText.Text = _jobData.DownLoadFolder;
			remarkText.Text = _jobData.Remark;
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

	private void type_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			if ((sender as RadioButton).Checked)
			{
				if ((sender as RadioButton).Name == "typeOne")
				{
					timingPanel.Visible = false;
					retryPanel.Visible = false;
					weekCheck.Visible = false;
					monthPanel.Visible = false;
					typeSettingTable.RowStyles[1].Height = 0f;
					typeSettingTable.RowStyles[2].Height = 0f;
					typeSettingTable.RowStyles[3].Height = 0f;
					_jobData.Type = "4";
				}
				else if ((sender as RadioButton).Name == "typeDay")
				{
					timingPanel.Visible = true;
					retryPanel.Visible = true;
					weekCheck.Visible = false;
					monthPanel.Visible = false;
					typeSettingTable.RowStyles[1].Height = 75f;
					typeSettingTable.RowStyles[2].Height = 0f;
					typeSettingTable.RowStyles[3].Height = 0f;
					_jobData.Type = "1";
				}
				else if ((sender as RadioButton).Name == "typeWeek")
				{
					timingPanel.Visible = false;
					retryPanel.Visible = false;
					weekCheck.Visible = true;
					monthPanel.Visible = false;
					typeSettingTable.RowStyles[1].Height = 0f;
					typeSettingTable.RowStyles[2].Height = 100f;
					typeSettingTable.RowStyles[3].Height = 0f;
					_jobData.Type = "2";
				}
				else if ((sender as RadioButton).Name == "typeMonth")
				{
					timingPanel.Visible = false;
					retryPanel.Visible = false;
					weekCheck.Visible = false;
					monthPanel.Visible = true;
					typeSettingTable.RowStyles[1].Height = 0f;
					typeSettingTable.RowStyles[2].Height = 0f;
					typeSettingTable.RowStyles[3].Height = 260f;
					_jobData.Type = "3";
				}
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void retryCheck_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			if (retryCheck.Checked)
			{
				retryMinitesText.Enabled = true;
				timingText.Enabled = false;
			}
			else
			{
				retryMinitesText.Enabled = false;
				timingText.Enabled = true;
			}
			_jobData.IsRetry = retryCheck.Checked;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void endCheck_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			if (endCheck.Checked)
			{
				endTimePanel.Enabled = true;
			}
			else
			{
				endTimePanel.Enabled = false;
			}
			if (endTimePanel.Enabled && string.IsNullOrEmpty(_jobData.EndTime))
			{
				endHour.Text = endHour.Value.ToString("00");
				endMinites.Text = endMinites.Value.ToString("00");
				_jobData.EndTime = endDate.Value.ToString("yyyy/MM/dd") + " " + endHour.Value.ToString("00") + ":" + endMinites.Value.ToString("00");
			}
			_jobData.IsEndDate = endCheck.Checked;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void nameText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			if (string.IsNullOrEmpty(nameText.Text.Trim()))
			{
				MessageBox.Show(Resources.NameEmptyMessage, "WARNING");
				e.Cancel = true;
			}
			_jobData.Name = nameText.Text;
			_node.Text = nameText.Text;
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
			if ((sender as Button).Name == "moveFolder")
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

	private void startDate_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			startHour.Text = startHour.Value.ToString("00");
			startMinites.Text = startMinites.Value.ToString("00");
			_jobData.StartTime = startDate.Value.ToString("yyyy/MM/dd") + " " + startHour.Value.ToString("00") + ":" + startMinites.Value.ToString("00");
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void timingText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			int result = 0;
			if (int.TryParse(timingText.Text, out result))
			{
				_jobData.Timing = result.ToString();
			}
			else
			{
				timingText.Text = _jobData.Timing;
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void allMonthCheck_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			for (int i = 0; i < monthCheckList.Items.Count; i++)
			{
				monthCheckList.SetItemChecked(i, allMonthCheck.Checked);
			}
			if (allMonthCheck.Checked)
			{
				_jobData.Timing = "1,2,3,4,5,6,7,8,9,10,11,12";
			}
			else
			{
				_jobData.Timing = "";
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void allDayCheck_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			for (int i = 0; i < dayCheckList.Items.Count; i++)
			{
				dayCheckList.SetItemChecked(i, allDayCheck.Checked);
			}
			if (allDayCheck.Checked)
			{
				_jobData.Day = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31";
			}
			else
			{
				_jobData.Day = "";
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void allWeekCheck_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			for (int i = 0; i < weekCheckList.Items.Count; i++)
			{
				weekCheckList.SetItemChecked(i, allWeekCheck.Checked);
			}
			if (allWeekCheck.Checked)
			{
				_jobData.Timing = "0,1,2,3,4,5,6";
			}
			else
			{
				_jobData.Timing = "";
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void monthCheckList_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			string text = string.Empty;
			string text2 = string.Empty;
			foreach (int checkedIndex in monthCheckList.CheckedIndices)
			{
				text = text + text2 + (checkedIndex + 1);
				text2 = ",";
			}
			_jobData.Month = text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void dayCheckList_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			string text = string.Empty;
			string text2 = string.Empty;
			foreach (int checkedIndex in dayCheckList.CheckedIndices)
			{
				text = text + text2 + (checkedIndex + 1);
				text2 = ",";
			}
			_jobData.Day = text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void weekCheckList_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			string text = string.Empty;
			string text2 = string.Empty;
			foreach (int checkedIndex in weekCheckList.CheckedIndices)
			{
				text = text + text2 + checkedIndex;
				text2 = ",";
			}
			_jobData.Week = text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void retryMinitesText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			int result = 0;
			if (int.TryParse(retryMinitesText.Text, out result))
			{
				_jobData.RetryTiming = result.ToString();
			}
			else
			{
				retryMinitesText.Text = _jobData.RetryTiming;
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void endDate_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			endHour.Text = endHour.Value.ToString("00");
			endMinites.Text = endMinites.Value.ToString("00");
			_jobData.EndTime = endDate.Value.ToString("yyyy/MM/dd") + " " + endHour.Value.ToString("00") + ":" + endMinites.Value.ToString("00");
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

	private void SetFormResouce()
	{
		try
		{
			nameLabel.Text = Resources.Name;
			workFolderLabel.Text = Resources.WorkFolder;
			downloadFolderLabel.Text = Resources.DownLoadFolder;
			remarkLabel.Text = Resources.Remark;
			typeOne.Text = Resources.SchduleTypeOne;
			typeDay.Text = Resources.SchduleTypeDay;
			typeWeek.Text = Resources.SchduleTypeWeek;
			typeMonth.Text = Resources.SchduleTypeMonth;
			startDateLabel.Text = Resources.SchduleStartDate;
			timingLabel.Text = Resources.SchduleTimingDay;
			retryCheck.Text = Resources.SchduleRetryTiming;
			weekLabel.Text = Resources.WeekSelect;
			monthLabel.Text = Resources.MonthSelect;
			dayLabel.Text = Resources.DaySelect;
			allDayCheck.Text = Resources.All;
			allMonthCheck.Text = Resources.All;
			allWeekCheck.Text = Resources.All;
			enableCheckBox.Text = Resources.Enable;
			endCheck.Text = Resources.EnableDate;
			workLabel.Text = Resources.WorkFolder;
		}
		catch (Exception ex)
		{
			throw ex;
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
		System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConMasGeneratorUtility.Views.JobSchduleView));
		this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
		this.baseTable = new System.Windows.Forms.TableLayoutPanel();
		this.schSettingGroup = new System.Windows.Forms.GroupBox();
		this.schBaseTable = new System.Windows.Forms.TableLayoutPanel();
		this.selectPanel = new System.Windows.Forms.Panel();
		this.line = new System.Windows.Forms.Label();
		this.flowLayout = new System.Windows.Forms.FlowLayoutPanel();
		this.typeOne = new System.Windows.Forms.RadioButton();
		this.typeDay = new System.Windows.Forms.RadioButton();
		this.typeWeek = new System.Windows.Forms.RadioButton();
		this.typeMonth = new System.Windows.Forms.RadioButton();
		this.typeSettingTable = new System.Windows.Forms.TableLayoutPanel();
		this.startDatePanel = new System.Windows.Forms.Panel();
		this.startDate = new System.Windows.Forms.DateTimePicker();
		this.startMinites = new System.Windows.Forms.NumericUpDown();
		this.startDateLabel = new System.Windows.Forms.Label();
		this.startHour = new System.Windows.Forms.NumericUpDown();
		this.label1 = new System.Windows.Forms.Label();
		this.weekCheck = new System.Windows.Forms.Panel();
		this.weekLabel = new System.Windows.Forms.Label();
		this.weekCheckList = new System.Windows.Forms.CheckedListBox();
		this.allWeekCheck = new System.Windows.Forms.CheckBox();
		this.monthPanel = new System.Windows.Forms.Panel();
		this.dayLabel = new System.Windows.Forms.Label();
		this.monthLabel = new System.Windows.Forms.Label();
		this.allDayCheck = new System.Windows.Forms.CheckBox();
		this.allMonthCheck = new System.Windows.Forms.CheckBox();
		this.dayCheckList = new System.Windows.Forms.CheckedListBox();
		this.monthCheckList = new System.Windows.Forms.CheckedListBox();
		this.timingPanel = new System.Windows.Forms.Panel();
		this.timingText = new ConMasGeneratorUtility.Controls.NumericTextBox();
		this.timingLabel = new System.Windows.Forms.Label();
		this.retryPanel = new System.Windows.Forms.Panel();
		this.retryCheck = new System.Windows.Forms.CheckBox();
		this.retryMinitesText = new ConMasGeneratorUtility.Controls.NumericTextBox();
		this.otherPanel = new System.Windows.Forms.Panel();
		this.endTimePanel = new System.Windows.Forms.Panel();
		this.label3 = new System.Windows.Forms.Label();
		this.endDate = new System.Windows.Forms.DateTimePicker();
		this.endMinites = new System.Windows.Forms.NumericUpDown();
		this.endHour = new System.Windows.Forms.NumericUpDown();
		this.enableCheckBox = new System.Windows.Forms.CheckBox();
		this.endCheck = new System.Windows.Forms.CheckBox();
		this.schduleJobPanel = new System.Windows.Forms.TableLayoutPanel();
		this.nameLabel = new System.Windows.Forms.Label();
		this.nameText = new System.Windows.Forms.TextBox();
		this.remarkText = new System.Windows.Forms.TextBox();
		this.downLoadFolderText = new System.Windows.Forms.TextBox();
		this.downLoadFolderButton = new System.Windows.Forms.Button();
		this.workFolderButton = new System.Windows.Forms.Button();
		this.workFolderLabel = new System.Windows.Forms.Label();
		this.downloadFolderLabel = new System.Windows.Forms.Label();
		this.remarkLabel = new System.Windows.Forms.Label();
		this.workFolderText = new System.Windows.Forms.TextBox();
		this.moveFolder = new System.Windows.Forms.Button();
		this.workLabel = new System.Windows.Forms.Label();
		this.baseTable.SuspendLayout();
		this.schSettingGroup.SuspendLayout();
		this.schBaseTable.SuspendLayout();
		this.selectPanel.SuspendLayout();
		this.flowLayout.SuspendLayout();
		this.typeSettingTable.SuspendLayout();
		this.startDatePanel.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.startMinites).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.startHour).BeginInit();
		this.weekCheck.SuspendLayout();
		this.monthPanel.SuspendLayout();
		this.timingPanel.SuspendLayout();
		this.retryPanel.SuspendLayout();
		this.otherPanel.SuspendLayout();
		this.endTimePanel.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.endMinites).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.endHour).BeginInit();
		this.schduleJobPanel.SuspendLayout();
		base.SuspendLayout();
		this.baseTable.AutoScroll = true;
		this.baseTable.ColumnCount = 1;
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.baseTable.Controls.Add(this.schSettingGroup, 0, 1);
		this.baseTable.Controls.Add(this.schduleJobPanel, 0, 0);
		this.baseTable.Dock = System.Windows.Forms.DockStyle.Top;
		this.baseTable.Location = new System.Drawing.Point(10, 10);
		this.baseTable.Name = "baseTable";
		this.baseTable.RowCount = 2;
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 128f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 96.95341f));
		this.baseTable.Size = new System.Drawing.Size(590, 686);
		this.baseTable.TabIndex = 25;
		this.schSettingGroup.Controls.Add(this.schBaseTable);
		this.schSettingGroup.Dock = System.Windows.Forms.DockStyle.Top;
		this.schSettingGroup.Location = new System.Drawing.Point(3, 131);
		this.schSettingGroup.Name = "schSettingGroup";
		this.schSettingGroup.Size = new System.Drawing.Size(584, 552);
		this.schSettingGroup.TabIndex = 105;
		this.schSettingGroup.TabStop = false;
		this.schSettingGroup.Text = "設定";
		this.schBaseTable.AutoScroll = true;
		this.schBaseTable.ColumnCount = 1;
		this.schBaseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.schBaseTable.Controls.Add(this.selectPanel, 0, 0);
		this.schBaseTable.Controls.Add(this.typeSettingTable, 0, 1);
		this.schBaseTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.schBaseTable.Location = new System.Drawing.Point(3, 15);
		this.schBaseTable.Name = "schBaseTable";
		this.schBaseTable.Padding = new System.Windows.Forms.Padding(3);
		this.schBaseTable.RowCount = 2;
		this.schBaseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.schBaseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.schBaseTable.Size = new System.Drawing.Size(578, 534);
		this.schBaseTable.TabIndex = 28;
		this.selectPanel.Controls.Add(this.line);
		this.selectPanel.Controls.Add(this.flowLayout);
		this.selectPanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.selectPanel.Location = new System.Drawing.Point(6, 6);
		this.selectPanel.Name = "selectPanel";
		this.selectPanel.Size = new System.Drawing.Size(566, 29);
		this.selectPanel.TabIndex = 0;
		this.line.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
		this.line.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.line.Location = new System.Drawing.Point(0, 26);
		this.line.Margin = new System.Windows.Forms.Padding(0);
		this.line.Name = "line";
		this.line.Size = new System.Drawing.Size(566, 3);
		this.line.TabIndex = 29;
		this.flowLayout.Controls.Add(this.typeOne);
		this.flowLayout.Controls.Add(this.typeDay);
		this.flowLayout.Controls.Add(this.typeWeek);
		this.flowLayout.Controls.Add(this.typeMonth);
		this.flowLayout.Dock = System.Windows.Forms.DockStyle.Top;
		this.flowLayout.Location = new System.Drawing.Point(0, 0);
		this.flowLayout.Name = "flowLayout";
		this.flowLayout.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
		this.flowLayout.Size = new System.Drawing.Size(566, 27);
		this.flowLayout.TabIndex = 5;
		this.typeOne.AutoSize = true;
		this.typeOne.Location = new System.Drawing.Point(8, 3);
		this.typeOne.Name = "typeOne";
		this.typeOne.Size = new System.Drawing.Size(68, 16);
		this.typeOne.TabIndex = 105;
		this.typeOne.TabStop = true;
		this.typeOne.Text = "一回のみ";
		this.typeOne.UseVisualStyleBackColor = true;
		this.typeOne.CheckedChanged += new System.EventHandler(type_CheckedChanged);
		this.typeDay.AutoSize = true;
		this.typeDay.Location = new System.Drawing.Point(82, 3);
		this.typeDay.Name = "typeDay";
		this.typeDay.Size = new System.Drawing.Size(47, 16);
		this.typeDay.TabIndex = 106;
		this.typeDay.TabStop = true;
		this.typeDay.Text = "毎日";
		this.typeDay.UseVisualStyleBackColor = true;
		this.typeDay.CheckedChanged += new System.EventHandler(type_CheckedChanged);
		this.typeWeek.AutoSize = true;
		this.typeWeek.Location = new System.Drawing.Point(135, 3);
		this.typeWeek.Name = "typeWeek";
		this.typeWeek.Size = new System.Drawing.Size(47, 16);
		this.typeWeek.TabIndex = 107;
		this.typeWeek.TabStop = true;
		this.typeWeek.Text = "毎週";
		this.typeWeek.UseVisualStyleBackColor = true;
		this.typeWeek.CheckedChanged += new System.EventHandler(type_CheckedChanged);
		this.typeMonth.AutoSize = true;
		this.typeMonth.Location = new System.Drawing.Point(188, 3);
		this.typeMonth.Name = "typeMonth";
		this.typeMonth.Size = new System.Drawing.Size(47, 16);
		this.typeMonth.TabIndex = 108;
		this.typeMonth.TabStop = true;
		this.typeMonth.Text = "毎月";
		this.typeMonth.UseVisualStyleBackColor = true;
		this.typeMonth.CheckedChanged += new System.EventHandler(type_CheckedChanged);
		this.typeSettingTable.ColumnCount = 1;
		this.typeSettingTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.typeSettingTable.Controls.Add(this.startDatePanel, 0, 0);
		this.typeSettingTable.Controls.Add(this.weekCheck, 0, 2);
		this.typeSettingTable.Controls.Add(this.monthPanel, 0, 3);
		this.typeSettingTable.Controls.Add(this.timingPanel, 0, 1);
		this.typeSettingTable.Controls.Add(this.otherPanel, 0, 4);
		this.typeSettingTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.typeSettingTable.Location = new System.Drawing.Point(6, 41);
		this.typeSettingTable.Name = "typeSettingTable";
		this.typeSettingTable.RowCount = 5;
		this.typeSettingTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35f));
		this.typeSettingTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 64f));
		this.typeSettingTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 91f));
		this.typeSettingTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 218f));
		this.typeSettingTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20f));
		this.typeSettingTable.Size = new System.Drawing.Size(566, 487);
		this.typeSettingTable.TabIndex = 27;
		this.startDatePanel.Controls.Add(this.startDate);
		this.startDatePanel.Controls.Add(this.startMinites);
		this.startDatePanel.Controls.Add(this.startDateLabel);
		this.startDatePanel.Controls.Add(this.startHour);
		this.startDatePanel.Controls.Add(this.label1);
		this.startDatePanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.startDatePanel.Location = new System.Drawing.Point(3, 3);
		this.startDatePanel.Name = "startDatePanel";
		this.startDatePanel.Size = new System.Drawing.Size(560, 29);
		this.startDatePanel.TabIndex = 28;
		this.startDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
		this.startDate.Location = new System.Drawing.Point(140, 5);
		this.startDate.Name = "startDate";
		this.startDate.Size = new System.Drawing.Size(119, 19);
		this.startDate.TabIndex = 201;
		this.startDate.Validating += new System.ComponentModel.CancelEventHandler(startDate_Validating);
		this.startMinites.Location = new System.Drawing.Point(319, 5);
		this.startMinites.Maximum = new decimal(new int[4] { 59, 0, 0, 0 });
		this.startMinites.Name = "startMinites";
		this.startMinites.Size = new System.Drawing.Size(38, 19);
		this.startMinites.TabIndex = 203;
		this.startMinites.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
		this.startMinites.Validating += new System.ComponentModel.CancelEventHandler(startDate_Validating);
		this.startDateLabel.Location = new System.Drawing.Point(3, 3);
		this.startDateLabel.Name = "startDateLabel";
		this.startDateLabel.Size = new System.Drawing.Size(132, 23);
		this.startDateLabel.TabIndex = 26;
		this.startDateLabel.Text = "label4";
		this.startDateLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.startHour.Location = new System.Drawing.Point(264, 5);
		this.startHour.Maximum = new decimal(new int[4] { 23, 0, 0, 0 });
		this.startHour.Name = "startHour";
		this.startHour.Size = new System.Drawing.Size(38, 19);
		this.startHour.TabIndex = 202;
		this.startHour.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
		this.startHour.Validating += new System.ComponentModel.CancelEventHandler(startDate_Validating);
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(306, 8);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(11, 12);
		this.label1.TabIndex = 21;
		this.label1.Text = "：";
		this.weekCheck.Controls.Add(this.weekLabel);
		this.weekCheck.Controls.Add(this.weekCheckList);
		this.weekCheck.Controls.Add(this.allWeekCheck);
		this.weekCheck.Dock = System.Windows.Forms.DockStyle.Fill;
		this.weekCheck.Location = new System.Drawing.Point(3, 102);
		this.weekCheck.Name = "weekCheck";
		this.weekCheck.Size = new System.Drawing.Size(560, 85);
		this.weekCheck.TabIndex = 15;
		this.weekCheck.Visible = false;
		this.weekLabel.Location = new System.Drawing.Point(5, 2);
		this.weekLabel.Name = "weekLabel";
		this.weekLabel.Size = new System.Drawing.Size(132, 27);
		this.weekLabel.TabIndex = 28;
		this.weekLabel.Text = "label5";
		this.weekLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.weekCheckList.FormattingEnabled = true;
		this.weekCheckList.Location = new System.Drawing.Point(141, 25);
		this.weekCheckList.MultiColumn = true;
		this.weekCheckList.Name = "weekCheckList";
		this.weekCheckList.Size = new System.Drawing.Size(329, 46);
		this.weekCheckList.TabIndex = 312;
		this.weekCheckList.Validating += new System.ComponentModel.CancelEventHandler(weekCheckList_Validating);
		this.allWeekCheck.AutoSize = true;
		this.allWeekCheck.Location = new System.Drawing.Point(143, 4);
		this.allWeekCheck.Name = "allWeekCheck";
		this.allWeekCheck.Size = new System.Drawing.Size(80, 16);
		this.allWeekCheck.TabIndex = 311;
		this.allWeekCheck.Text = "checkBox4";
		this.allWeekCheck.UseVisualStyleBackColor = true;
		this.allWeekCheck.CheckedChanged += new System.EventHandler(allWeekCheck_CheckedChanged);
		this.monthPanel.Controls.Add(this.dayLabel);
		this.monthPanel.Controls.Add(this.monthLabel);
		this.monthPanel.Controls.Add(this.allDayCheck);
		this.monthPanel.Controls.Add(this.allMonthCheck);
		this.monthPanel.Controls.Add(this.dayCheckList);
		this.monthPanel.Controls.Add(this.monthCheckList);
		this.monthPanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.monthPanel.Location = new System.Drawing.Point(3, 193);
		this.monthPanel.Name = "monthPanel";
		this.monthPanel.Size = new System.Drawing.Size(560, 212);
		this.monthPanel.TabIndex = 19;
		this.monthPanel.Visible = false;
		this.dayLabel.Location = new System.Drawing.Point(4, 96);
		this.dayLabel.Name = "dayLabel";
		this.dayLabel.Size = new System.Drawing.Size(118, 23);
		this.dayLabel.TabIndex = 30;
		this.dayLabel.Text = "label7";
		this.dayLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.monthLabel.Location = new System.Drawing.Point(5, 3);
		this.monthLabel.Name = "monthLabel";
		this.monthLabel.Size = new System.Drawing.Size(131, 23);
		this.monthLabel.TabIndex = 29;
		this.monthLabel.Text = "label6";
		this.monthLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.allDayCheck.AutoSize = true;
		this.allDayCheck.Location = new System.Drawing.Point(143, 96);
		this.allDayCheck.Name = "allDayCheck";
		this.allDayCheck.Size = new System.Drawing.Size(80, 16);
		this.allDayCheck.TabIndex = 323;
		this.allDayCheck.Text = "checkBox4";
		this.allDayCheck.UseVisualStyleBackColor = true;
		this.allDayCheck.CheckedChanged += new System.EventHandler(allDayCheck_CheckedChanged);
		this.allMonthCheck.AutoSize = true;
		this.allMonthCheck.Location = new System.Drawing.Point(143, 5);
		this.allMonthCheck.Name = "allMonthCheck";
		this.allMonthCheck.Size = new System.Drawing.Size(80, 16);
		this.allMonthCheck.TabIndex = 321;
		this.allMonthCheck.Text = "checkBox4";
		this.allMonthCheck.UseVisualStyleBackColor = true;
		this.allMonthCheck.CheckedChanged += new System.EventHandler(allMonthCheck_CheckedChanged);
		this.dayCheckList.FormattingEnabled = true;
		this.dayCheckList.Location = new System.Drawing.Point(141, 117);
		this.dayCheckList.MultiColumn = true;
		this.dayCheckList.Name = "dayCheckList";
		this.dayCheckList.ScrollAlwaysVisible = true;
		this.dayCheckList.Size = new System.Drawing.Size(454, 88);
		this.dayCheckList.TabIndex = 324;
		this.dayCheckList.Validating += new System.ComponentModel.CancelEventHandler(dayCheckList_Validating);
		this.monthCheckList.FormattingEnabled = true;
		this.monthCheckList.Location = new System.Drawing.Point(141, 28);
		this.monthCheckList.MultiColumn = true;
		this.monthCheckList.Name = "monthCheckList";
		this.monthCheckList.Size = new System.Drawing.Size(454, 60);
		this.monthCheckList.TabIndex = 322;
		this.monthCheckList.Validating += new System.ComponentModel.CancelEventHandler(monthCheckList_Validating);
		this.timingPanel.Controls.Add(this.timingText);
		this.timingPanel.Controls.Add(this.timingLabel);
		this.timingPanel.Controls.Add(this.retryPanel);
		this.timingPanel.Location = new System.Drawing.Point(3, 38);
		this.timingPanel.Name = "timingPanel";
		this.timingPanel.Size = new System.Drawing.Size(270, 58);
		this.timingPanel.TabIndex = 301;
		this.timingPanel.Visible = false;
		this.timingText.Location = new System.Drawing.Point(141, 6);
		this.timingText.Name = "timingText";
		this.timingText.Size = new System.Drawing.Size(49, 19);
		this.timingText.TabIndex = 301;
		this.timingText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
		this.timingText.Validating += new System.ComponentModel.CancelEventHandler(timingText_Validating);
		this.timingLabel.Location = new System.Drawing.Point(5, 4);
		this.timingLabel.Name = "timingLabel";
		this.timingLabel.Size = new System.Drawing.Size(138, 23);
		this.timingLabel.TabIndex = 27;
		this.timingLabel.Text = "label4";
		this.timingLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.retryPanel.Controls.Add(this.retryCheck);
		this.retryPanel.Controls.Add(this.retryMinitesText);
		this.retryPanel.Location = new System.Drawing.Point(2, 28);
		this.retryPanel.Name = "retryPanel";
		this.retryPanel.Size = new System.Drawing.Size(201, 27);
		this.retryPanel.TabIndex = 310;
		this.retryPanel.Visible = false;
		this.retryCheck.AutoSize = true;
		this.retryCheck.Location = new System.Drawing.Point(4, 3);
		this.retryCheck.Name = "retryCheck";
		this.retryCheck.Size = new System.Drawing.Size(30, 16);
		this.retryCheck.TabIndex = 311;
		this.retryCheck.Text = "...";
		this.retryCheck.UseVisualStyleBackColor = true;
		this.retryCheck.CheckedChanged += new System.EventHandler(retryCheck_CheckedChanged);
		this.retryMinitesText.Enabled = false;
		this.retryMinitesText.Location = new System.Drawing.Point(139, 2);
		this.retryMinitesText.Name = "retryMinitesText";
		this.retryMinitesText.Size = new System.Drawing.Size(49, 19);
		this.retryMinitesText.TabIndex = 312;
		this.retryMinitesText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
		this.retryMinitesText.Validating += new System.ComponentModel.CancelEventHandler(retryMinitesText_Validating);
		this.otherPanel.Controls.Add(this.endTimePanel);
		this.otherPanel.Controls.Add(this.enableCheckBox);
		this.otherPanel.Controls.Add(this.endCheck);
		this.otherPanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.otherPanel.Location = new System.Drawing.Point(3, 411);
		this.otherPanel.Name = "otherPanel";
		this.otherPanel.Size = new System.Drawing.Size(560, 73);
		this.otherPanel.TabIndex = 29;
		this.endTimePanel.Controls.Add(this.label3);
		this.endTimePanel.Controls.Add(this.endDate);
		this.endTimePanel.Controls.Add(this.endMinites);
		this.endTimePanel.Controls.Add(this.endHour);
		this.endTimePanel.Enabled = false;
		this.endTimePanel.Location = new System.Drawing.Point(132, 24);
		this.endTimePanel.Name = "endTimePanel";
		this.endTimePanel.Size = new System.Drawing.Size(230, 33);
		this.endTimePanel.TabIndex = 26;
		this.label3.AutoSize = true;
		this.label3.Location = new System.Drawing.Point(173, 8);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(11, 12);
		this.label3.TabIndex = 26;
		this.label3.Text = "：";
		this.endDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
		this.endDate.Location = new System.Drawing.Point(8, 6);
		this.endDate.Name = "endDate";
		this.endDate.Size = new System.Drawing.Size(119, 19);
		this.endDate.TabIndex = 333;
		this.endDate.Validating += new System.ComponentModel.CancelEventHandler(endDate_Validating);
		this.endMinites.Location = new System.Drawing.Point(186, 6);
		this.endMinites.Maximum = new decimal(new int[4] { 59, 0, 0, 0 });
		this.endMinites.Name = "endMinites";
		this.endMinites.Size = new System.Drawing.Size(38, 19);
		this.endMinites.TabIndex = 335;
		this.endMinites.Value = new decimal(new int[4] { 59, 0, 0, 0 });
		this.endMinites.Validating += new System.ComponentModel.CancelEventHandler(endDate_Validating);
		this.endMinites.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
		this.endHour.Location = new System.Drawing.Point(131, 6);
		this.endHour.Maximum = new decimal(new int[4] { 23, 0, 0, 0 });
		this.endHour.Name = "endHour";
		this.endHour.Size = new System.Drawing.Size(38, 19);
		this.endHour.TabIndex = 334;
		this.endHour.Value = new decimal(new int[4] { 23, 0, 0, 0 });
		this.endHour.Validating += new System.ComponentModel.CancelEventHandler(endDate_Validating);
		this.endHour.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
		this.enableCheckBox.AutoSize = true;
		this.enableCheckBox.Location = new System.Drawing.Point(6, 6);
		this.enableCheckBox.Name = "enableCheckBox";
		this.enableCheckBox.Size = new System.Drawing.Size(80, 16);
		this.enableCheckBox.TabIndex = 331;
		this.enableCheckBox.Text = "checkBox4";
		this.enableCheckBox.UseVisualStyleBackColor = true;
		this.enableCheckBox.CheckedChanged += new System.EventHandler(enableCheckBox_CheckedChanged);
		this.endCheck.AutoSize = true;
		this.endCheck.Location = new System.Drawing.Point(6, 31);
		this.endCheck.Name = "endCheck";
		this.endCheck.Size = new System.Drawing.Size(30, 16);
		this.endCheck.TabIndex = 332;
		this.endCheck.Text = "...";
		this.endCheck.UseVisualStyleBackColor = true;
		this.endCheck.CheckedChanged += new System.EventHandler(endCheck_CheckedChanged);
		this.schduleJobPanel.ColumnCount = 3;
		this.schduleJobPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 150f));
		this.schduleJobPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.schduleJobPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 52f));
		this.schduleJobPanel.Controls.Add(this.nameLabel, 0, 0);
		this.schduleJobPanel.Controls.Add(this.nameText, 1, 0);
		this.schduleJobPanel.Controls.Add(this.remarkText, 1, 3);
		this.schduleJobPanel.Controls.Add(this.downLoadFolderText, 1, 2);
		this.schduleJobPanel.Controls.Add(this.downLoadFolderButton, 2, 2);
		this.schduleJobPanel.Controls.Add(this.workFolderButton, 2, 1);
		this.schduleJobPanel.Controls.Add(this.workFolderLabel, 0, 1);
		this.schduleJobPanel.Controls.Add(this.downloadFolderLabel, 0, 2);
		this.schduleJobPanel.Controls.Add(this.remarkLabel, 0, 3);
		this.schduleJobPanel.Controls.Add(this.workFolderText, 1, 4);
		this.schduleJobPanel.Controls.Add(this.moveFolder, 2, 4);
		this.schduleJobPanel.Controls.Add(this.workLabel, 0, 4);
		this.schduleJobPanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.schduleJobPanel.Location = new System.Drawing.Point(3, 3);
		this.schduleJobPanel.Name = "schduleJobPanel";
		this.schduleJobPanel.RowCount = 5;
		this.schduleJobPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.schduleJobPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 0f));
		this.schduleJobPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 31f));
		this.schduleJobPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.schduleJobPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.schduleJobPanel.Size = new System.Drawing.Size(584, 122);
		this.schduleJobPanel.TabIndex = 22;
		this.nameLabel.BackColor = System.Drawing.SystemColors.Control;
		this.nameLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.nameLabel.Location = new System.Drawing.Point(3, 3);
		this.nameLabel.Margin = new System.Windows.Forms.Padding(3);
		this.nameLabel.Name = "nameLabel";
		this.nameLabel.Size = new System.Drawing.Size(144, 24);
		this.nameLabel.TabIndex = 26;
		this.nameLabel.Text = "label2";
		this.nameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.nameText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.nameText.Location = new System.Drawing.Point(153, 8);
		this.nameText.Name = "nameText";
		this.nameText.Size = new System.Drawing.Size(376, 19);
		this.nameText.TabIndex = 101;
		this.nameText.Validating += new System.ComponentModel.CancelEventHandler(nameText_Validating);
		this.remarkText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.remarkText.Location = new System.Drawing.Point(153, 69);
		this.remarkText.Name = "remarkText";
		this.remarkText.Size = new System.Drawing.Size(376, 19);
		this.remarkText.TabIndex = 104;
		this.remarkText.Validating += new System.ComponentModel.CancelEventHandler(remarkText_Validating);
		this.downLoadFolderText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.downLoadFolderText.Location = new System.Drawing.Point(153, 39);
		this.downLoadFolderText.Name = "downLoadFolderText";
		this.downLoadFolderText.Size = new System.Drawing.Size(376, 19);
		this.downLoadFolderText.TabIndex = 102;
		this.downLoadFolderText.Validating += new System.ComponentModel.CancelEventHandler(downLoadFolderText_Validating);
		this.downLoadFolderButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("downLoadFolderButton.BackgroundImage");
		this.downLoadFolderButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.downLoadFolderButton.Dock = System.Windows.Forms.DockStyle.Left;
		this.downLoadFolderButton.Location = new System.Drawing.Point(535, 33);
		this.downLoadFolderButton.Name = "downLoadFolderButton";
		this.downLoadFolderButton.Size = new System.Drawing.Size(27, 25);
		this.downLoadFolderButton.TabIndex = 103;
		this.downLoadFolderButton.UseVisualStyleBackColor = true;
		this.downLoadFolderButton.Click += new System.EventHandler(folderButton_Click);
		this.workFolderButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("workFolderButton.BackgroundImage");
		this.workFolderButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.workFolderButton.Dock = System.Windows.Forms.DockStyle.Left;
		this.workFolderButton.Location = new System.Drawing.Point(535, 33);
		this.workFolderButton.Name = "workFolderButton";
		this.workFolderButton.Size = new System.Drawing.Size(27, 1);
		this.workFolderButton.TabIndex = 18;
		this.workFolderButton.UseVisualStyleBackColor = true;
		this.workFolderButton.Click += new System.EventHandler(folderButton_Click);
		this.workFolderLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.workFolderLabel.Location = new System.Drawing.Point(3, 30);
		this.workFolderLabel.Name = "workFolderLabel";
		this.workFolderLabel.Size = new System.Drawing.Size(144, 1);
		this.workFolderLabel.TabIndex = 27;
		this.workFolderLabel.Text = "label8";
		this.workFolderLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.workFolderLabel.Visible = false;
		this.downloadFolderLabel.BackColor = System.Drawing.SystemColors.Control;
		this.downloadFolderLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.downloadFolderLabel.Location = new System.Drawing.Point(3, 33);
		this.downloadFolderLabel.Margin = new System.Windows.Forms.Padding(3);
		this.downloadFolderLabel.Name = "downloadFolderLabel";
		this.downloadFolderLabel.Size = new System.Drawing.Size(144, 25);
		this.downloadFolderLabel.TabIndex = 28;
		this.downloadFolderLabel.Text = "label9";
		this.downloadFolderLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.remarkLabel.BackColor = System.Drawing.SystemColors.Control;
		this.remarkLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.remarkLabel.Location = new System.Drawing.Point(3, 64);
		this.remarkLabel.Margin = new System.Windows.Forms.Padding(3);
		this.remarkLabel.Name = "remarkLabel";
		this.remarkLabel.Size = new System.Drawing.Size(144, 24);
		this.remarkLabel.TabIndex = 29;
		this.remarkLabel.Text = "label10";
		this.remarkLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.workFolderText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.workFolderText.Location = new System.Drawing.Point(153, 100);
		this.workFolderText.Name = "workFolderText";
		this.workFolderText.Size = new System.Drawing.Size(376, 19);
		this.workFolderText.TabIndex = 10000;
		this.workFolderText.TabStop = false;
		this.workFolderText.Validating += new System.ComponentModel.CancelEventHandler(workFolderText_Validating);
		this.moveFolder.BackgroundImage = (System.Drawing.Image)resources.GetObject("moveFolder.BackgroundImage");
		this.moveFolder.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.moveFolder.Dock = System.Windows.Forms.DockStyle.Left;
		this.moveFolder.Location = new System.Drawing.Point(535, 94);
		this.moveFolder.Name = "moveFolder";
		this.moveFolder.Size = new System.Drawing.Size(27, 25);
		this.moveFolder.TabIndex = 10001;
		this.moveFolder.UseVisualStyleBackColor = true;
		this.moveFolder.Click += new System.EventHandler(folderButton_Click);
		this.workLabel.AutoSize = true;
		this.workLabel.BackColor = System.Drawing.SystemColors.Control;
		this.workLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.workLabel.Location = new System.Drawing.Point(3, 94);
		this.workLabel.Margin = new System.Windows.Forms.Padding(3);
		this.workLabel.Name = "workLabel";
		this.workLabel.Size = new System.Drawing.Size(144, 25);
		this.workLabel.TabIndex = 10002;
		this.workLabel.Text = "label2";
		this.workLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.AutoScroll = true;
		this.BackColor = System.Drawing.Color.White;
		base.Controls.Add(this.baseTable);
		base.Name = "JobSchduleView";
		base.Padding = new System.Windows.Forms.Padding(10);
		base.Size = new System.Drawing.Size(610, 482);
		base.Load += new System.EventHandler(JobSchduleView_Load);
		this.baseTable.ResumeLayout(false);
		this.schSettingGroup.ResumeLayout(false);
		this.schBaseTable.ResumeLayout(false);
		this.selectPanel.ResumeLayout(false);
		this.flowLayout.ResumeLayout(false);
		this.flowLayout.PerformLayout();
		this.typeSettingTable.ResumeLayout(false);
		this.startDatePanel.ResumeLayout(false);
		this.startDatePanel.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.startMinites).EndInit();
		((System.ComponentModel.ISupportInitialize)this.startHour).EndInit();
		this.weekCheck.ResumeLayout(false);
		this.weekCheck.PerformLayout();
		this.monthPanel.ResumeLayout(false);
		this.monthPanel.PerformLayout();
		this.timingPanel.ResumeLayout(false);
		this.timingPanel.PerformLayout();
		this.retryPanel.ResumeLayout(false);
		this.retryPanel.PerformLayout();
		this.otherPanel.ResumeLayout(false);
		this.otherPanel.PerformLayout();
		this.endTimePanel.ResumeLayout(false);
		this.endTimePanel.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.endMinites).EndInit();
		((System.ComponentModel.ISupportInitialize)this.endHour).EndInit();
		this.schduleJobPanel.ResumeLayout(false);
		this.schduleJobPanel.PerformLayout();
		base.ResumeLayout(false);
	}
}
