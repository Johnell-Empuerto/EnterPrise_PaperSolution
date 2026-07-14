using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ConMasGeneratorLib.Data;
using ConMasGeneratorUtility.Properties;

namespace ConMasGeneratorUtility.Views;

public class CommandView : SettingViewBase
{
	private TreeNode _node;

	private Command _command;

	private string _processType;

	private string _jobType;

	private const int GRID_HEIGHT_REST = 520;

	private const int GRID_HEIGHT_REST_WATCH = 400;

	private const int GRID_HEIGHT_EXE = 40;

	private IContainer components = null;

	private DataGridView parameterGridView;

	private TextBox nameText;

	private TextBox urlText;

	private TextBox pathText;

	private Button pathFileButton;

	private FolderBrowserDialog folderBrowserDialog;

	private CheckBox isResponseCheckBox;

	private TextBox responseFileNameText;

	private Panel exePanel;

	private TableLayoutPanel baseTable;

	private Label UrlLabel;

	private Label pathLabel;

	private DataGridViewTextBoxColumn ParameterName;

	private DataGridViewTextBoxColumn ParameterValue;

	private Label restLabel;

	private Label nameLabel;

	private Label parameterLabel;

	private Label responseNameLabel;

	private TableLayoutPanel exeTable;

	private Label waitLabel;

	private TextBox uploadFileNameText;

	private CheckBox isWaitCheck;

	private Label parameterLabel2;

	private TextBox paramText;

	private Label dounloadLabel;

	private CheckBox isUpload;

	private TableLayoutPanel unZipTable;

	private Button unZipFolderButton;

	private Label frozenLabel;

	private Panel panel1;

	private TableLayoutPanel downloadTable;

	private Label downloadStdLabel;

	private CheckBox isDownloadFolder;

	private TextBox downloadFolderCommandText;

	private CheckBox isZip;

	private TextBox unZipFolderText;

	private TableLayoutPanel downLoadSettingTable;

	private Button downloadCommandButton;

	private TableLayoutPanel urlTabel;

	private TableLayoutPanel restPanel;

	private TableLayoutPanel uploadBaseTable;

	private GroupBox downloadGroup;

	private TableLayoutPanel downloadBaseTable;

	private Button uploadFileButton;

	private GroupBox uploadGroup;

	private Button moveButton;

	private RadioButton uploadRadio1;

	private RadioButton uploadRadio2;

	private RadioButton uploadRadio3;

	private Label workFolderLabel;

	private TextBox moveUploadFolderText;

	private Label uploadAfterLabel;

	public CommandView(TreeNode node)
	{
		InitializeComponent();
		_node = node;
		_command = node.Tag as Command;
		SetFormResouce();
		_processType = (_node.Parent.Tag as ProcessData).Type;
		_jobType = (_node.Parent.Parent.Tag as JobData).Type;
	}

	private void CommandView_Load(object sender, EventArgs e)
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
			if (_processType == "0")
			{
				restPanel.Visible = true;
				exePanel.Visible = false;
				baseTable.RowStyles[1].Height = 520f;
				baseTable.RowStyles[2].Height = 0f;
				baseTable.RowStyles[3].Height = 0f;
				baseTable.RowStyles[4].Height = 0f;
				baseTable.RowStyles[5].SizeType = SizeType.Percent;
				baseTable.RowStyles[5].Height = 100f;
				if (_jobType == "0")
				{
					baseTable.RowStyles[1].Height = 400f;
					moveButton.Visible = false;
					uploadRadio1.Visible = false;
					uploadRadio2.Visible = false;
					uploadRadio3.Visible = false;
					moveUploadFolderText.Visible = false;
					uploadBaseTable.RowStyles[1].Height = 0f;
					uploadBaseTable.RowStyles[2].Height = 0f;
					uploadBaseTable.RowStyles[3].Height = 0f;
					uploadBaseTable.RowStyles[4].Height = 0f;
				}
			}
			else
			{
				restPanel.Visible = false;
				exePanel.Visible = true;
				baseTable.RowStyles[1].Height = 0f;
				baseTable.RowStyles[2].Height = 40f;
				baseTable.RowStyles[3].Height = 40f;
				baseTable.RowStyles[4].Height = 40f;
				baseTable.RowStyles[5].Height = 0f;
				if (_command.Parameters == null)
				{
					_command.Parameters = new List<ParameterData>();
				}
				if (_command.Parameters.Where((ParameterData child) => child.Name == "mode").Count() == 0)
				{
					ParameterData parameterData = new ParameterData();
					parameterData.Name = "mode";
					parameterData.Value = "0";
					_command.Parameters.Add(parameterData);
				}
				if (_command.Parameters.Where((ParameterData child) => child.Name == "param").Count() == 0)
				{
					ParameterData parameterData2 = new ParameterData();
					parameterData2.Name = "param";
					parameterData2.Value = "";
					_command.Parameters.Add(parameterData2);
				}
			}
			nameText.Text = _command.Name;
			urlText.Text = _command.UrlOrg;
			pathText.Text = _command.Path;
			isResponseCheckBox.Checked = _command.IsResponse;
			responseFileNameText.Text = _command.ResponseFileName;
			uploadFileNameText.Text = _command.UploadFilePath;
			isUpload.Checked = _command.IsUpload;
			isZip.Checked = _command.IsZip;
			unZipFolderText.Text = _command.UnZipFolder;
			isDownloadFolder.Checked = _command.IsDownloadFolder;
			downloadFolderCommandText.Text = _command.DownloadFolder;
			workFolderLabel.Text = (_node.Parent.Parent.Tag as JobData).WorkFolder;
			if (_command.MoveKbn == "2")
			{
				uploadRadio3.Checked = true;
			}
			else if (_command.MoveKbn == "1")
			{
				uploadRadio2.Checked = true;
			}
			else
			{
				uploadRadio1.Checked = true;
			}
			moveUploadFolderText.Text = _command.MovePath;
			if (_processType == "1")
			{
				if (_command.Parameters.Where((ParameterData child) => child.Name == "mode").First().Value == "1")
				{
					isWaitCheck.Checked = true;
				}
				else
				{
					isWaitCheck.Checked = false;
				}
				paramText.Text = _command.Parameters.Where((ParameterData child) => child.Name == "param").First().Value;
				return;
			}
			foreach (ParameterData parameter in _command.Parameters)
			{
				if (!string.IsNullOrEmpty(parameter.Name.Trim()))
				{
					int index = parameterGridView.Rows.Add(1);
					parameterGridView.Rows[index].Cells["ParameterName"].Value = parameter.Name;
					parameterGridView.Rows[index].Cells["ParameterValue"].Value = parameter.ValueOrg;
				}
			}
		}
		catch (Exception ex)
		{
			throw ex;
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
			_command.Name = nameText.Text;
			_node.Text = nameText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void urlText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_command.UrlOrg = urlText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void isResponseCheckBox_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			if (isResponseCheckBox.Checked)
			{
				responseFileNameText.Enabled = true;
				dounloadLabel.Text = (_node.Parent.Parent.Tag as JobData).DownLoadFolder;
				isZip.Enabled = true;
				isDownloadFolder.Enabled = true;
			}
			else
			{
				responseFileNameText.Enabled = false;
				dounloadLabel.Text = string.Empty;
				isZip.Enabled = false;
				isDownloadFolder.Enabled = false;
			}
			if (isZip.Checked && isResponseCheckBox.Checked)
			{
				unZipFolderButton.Enabled = true;
				unZipFolderText.Enabled = true;
			}
			else
			{
				unZipFolderButton.Enabled = false;
				unZipFolderText.Enabled = false;
			}
			if (isDownloadFolder.Checked && isResponseCheckBox.Checked)
			{
				downloadFolderCommandText.Enabled = true;
				downloadCommandButton.Enabled = true;
			}
			else
			{
				downloadFolderCommandText.Enabled = false;
				downloadCommandButton.Enabled = false;
			}
			_command.IsResponse = isResponseCheckBox.Checked;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void responseFileNameText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_command.ResponseFileName = responseFileNameText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void uploadFileNameText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_command.UploadFilePath = uploadFileNameText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void pathText_TextChanged(object sender, EventArgs e)
	{
		try
		{
			_command.Path = pathText.Text;
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
			if ((sender as Button).Name == "moveButton")
			{
				if (!string.IsNullOrEmpty(moveUploadFolderText.Text))
				{
					folderBrowserDialog.SelectedPath = moveUploadFolderText.Text;
				}
				if (folderBrowserDialog.ShowDialog(this) == DialogResult.OK)
				{
					moveUploadFolderText.Text = folderBrowserDialog.SelectedPath;
					_command.MovePath = moveUploadFolderText.Text;
				}
			}
			else
			{
				if (!string.IsNullOrEmpty(pathText.Text))
				{
					folderBrowserDialog.SelectedPath = pathText.Text;
				}
				if (folderBrowserDialog.ShowDialog(this) == DialogResult.OK)
				{
					pathText.Text = folderBrowserDialog.SelectedPath;
					_command.Path = pathText.Text;
				}
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void parameterGridView_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			if (_processType == "1")
			{
				return;
			}
			_command.Parameters.Clear();
			foreach (DataGridViewRow item in (IEnumerable)parameterGridView.Rows)
			{
				ParameterData parameterData = new ParameterData();
				if (item.Cells["ParameterName"].Value != null)
				{
					parameterData.Name = item.Cells["ParameterName"].Value.ToString();
					if (item.Cells["ParameterValue"].Value == null)
					{
						item.Cells["ParameterValue"].Value = string.Empty;
					}
					parameterData.Value = item.Cells["ParameterValue"].Value.ToString();
					if (!string.IsNullOrEmpty(parameterData.Name))
					{
						_command.Parameters.Add(parameterData);
					}
				}
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void pathText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_command.Path = pathText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void fileButton_Click(object sender, EventArgs e)
	{
		try
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.FileName = "";
			openFileDialog.InitialDirectory = "C:\\";
			openFileDialog.Filter = Resources.HtmlFile + "(*.html;*.htm)|*.html;*.htm|" + Resources.AllFile + "(*.*) | *.* ";
			openFileDialog.FilterIndex = 2;
			openFileDialog.Title = Resources.FileSelectTitle;
			openFileDialog.RestoreDirectory = true;
			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				if ((sender as Button).Name == "uploadFileButton")
				{
					uploadFileNameText.Text = openFileDialog.FileName;
					_command.UploadFilePath = openFileDialog.FileName;
				}
				else if ((sender as Button).Name == "pathFileButton")
				{
					pathText.Text = openFileDialog.FileName;
					_command.Path = openFileDialog.FileName;
				}
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void isWaitCheck_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			if (isWaitCheck.Checked)
			{
				_command.Parameters.Where((ParameterData child) => child.Name == "mode").First().Value = "1";
			}
			else
			{
				_command.Parameters.Where((ParameterData child) => child.Name == "mode").First().Value = "0";
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void paramText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_command.Parameters.Where((ParameterData child) => child.Name == "param").First().Value = paramText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void isUpload_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			_command.IsUpload = isUpload.Checked;
			if (isUpload.Checked)
			{
				uploadFileNameText.Enabled = true;
				uploadFileButton.Enabled = true;
				uploadRadio1.Enabled = true;
				uploadRadio2.Enabled = true;
				uploadRadio3.Enabled = true;
			}
			else
			{
				uploadFileNameText.Enabled = false;
				uploadFileButton.Enabled = false;
				uploadRadio1.Enabled = false;
				uploadRadio2.Enabled = false;
				uploadRadio3.Enabled = false;
			}
			if (isUpload.Checked && uploadRadio3.Checked)
			{
				moveButton.Enabled = true;
				moveUploadFolderText.Enabled = true;
			}
			else
			{
				moveButton.Enabled = false;
				moveUploadFolderText.Enabled = false;
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void unZipFolderButton_Click(object sender, EventArgs e)
	{
		try
		{
			folderBrowserDialog.Description = Resources.FolderMessageDialogTitle;
			folderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop;
			folderBrowserDialog.SelectedPath = "C:\\";
			folderBrowserDialog.ShowNewFolderButton = true;
			if ((sender as Button).Name == "downloadCommandButton")
			{
				if (!string.IsNullOrEmpty(downloadFolderCommandText.Text))
				{
					folderBrowserDialog.SelectedPath = downloadFolderCommandText.Text;
				}
				if (folderBrowserDialog.ShowDialog(this) == DialogResult.OK)
				{
					downloadFolderCommandText.Text = folderBrowserDialog.SelectedPath;
					_command.DownloadFolder = downloadFolderCommandText.Text;
				}
			}
			else
			{
				if (!string.IsNullOrEmpty(unZipFolderText.Text))
				{
					folderBrowserDialog.SelectedPath = unZipFolderText.Text;
				}
				if (folderBrowserDialog.ShowDialog(this) == DialogResult.OK)
				{
					unZipFolderText.Text = folderBrowserDialog.SelectedPath;
					_command.UnZipFolder = unZipFolderText.Text;
				}
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void unZipFolderText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_command.UnZipFolder = unZipFolderText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void isZip_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			_command.IsZip = isZip.Checked;
			if (isZip.Checked && isResponseCheckBox.Checked)
			{
				unZipFolderButton.Enabled = true;
				unZipFolderText.Enabled = true;
			}
			else
			{
				unZipFolderButton.Enabled = false;
				unZipFolderText.Enabled = false;
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void downloadFolderCommandText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_command.DownloadFolder = downloadFolderCommandText.Text;
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void isDownloadFolder_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			_command.IsDownloadFolder = isDownloadFolder.Checked;
			if (isDownloadFolder.Checked && isResponseCheckBox.Checked)
			{
				downloadFolderCommandText.Enabled = true;
				downloadCommandButton.Enabled = true;
			}
			else
			{
				downloadFolderCommandText.Enabled = false;
				downloadCommandButton.Enabled = false;
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
			restLabel.Text = Resources.RestSetting;
			pathLabel.Text = Resources.ExePath;
			parameterLabel.Text = Resources.ParameterSetting;
			UrlLabel.Text = Resources.Url;
			isResponseCheckBox.Text = Resources.IsResponse;
			responseNameLabel.Text = Resources.ResponseFileName;
			isUpload.Text = Resources.UploadFilePath;
			parameterLabel2.Text = Resources.ParameterSetting;
			isWaitCheck.Text = Resources.IsWait;
			waitLabel.Text = Resources.WaitSetting;
			isZip.Text = Resources.IsZip;
			frozenLabel.Text = Resources.FrozenLabel;
			downloadGroup.Text = Resources.DownloadSetting;
			downloadStdLabel.Text = Resources.DownloadStd;
			isDownloadFolder.Text = Resources.IsChangeDownloadFolder;
			uploadAfterLabel.Text = Resources.AfterUpload;
			uploadRadio1.Text = Resources.moveNone;
			uploadRadio2.Text = Resources.moveDefault;
			uploadRadio3.Text = Resources.moveCustom;
			uploadGroup.Text = Resources.UploadFilePath;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void uploadSelectChanged(object sender, EventArgs e)
	{
		try
		{
			if (uploadRadio3.Checked)
			{
				moveUploadFolderText.Enabled = true;
				moveButton.Enabled = true;
				_command.MoveKbn = "2";
			}
			else if (uploadRadio2.Checked)
			{
				moveUploadFolderText.Enabled = false;
				moveButton.Enabled = false;
				_command.MoveKbn = "1";
			}
			else if (uploadRadio1.Checked)
			{
				moveUploadFolderText.Enabled = false;
				moveButton.Enabled = false;
				_command.MoveKbn = "0";
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void moveUploadFolderText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			_command.MovePath = moveUploadFolderText.Text;
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
		System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConMasGeneratorUtility.Views.CommandView));
		this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
		this.baseTable = new System.Windows.Forms.TableLayoutPanel();
		this.waitLabel = new System.Windows.Forms.Label();
		this.restLabel = new System.Windows.Forms.Label();
		this.nameLabel = new System.Windows.Forms.Label();
		this.pathLabel = new System.Windows.Forms.Label();
		this.exePanel = new System.Windows.Forms.Panel();
		this.exeTable = new System.Windows.Forms.TableLayoutPanel();
		this.pathText = new System.Windows.Forms.TextBox();
		this.pathFileButton = new System.Windows.Forms.Button();
		this.nameText = new System.Windows.Forms.TextBox();
		this.parameterGridView = new System.Windows.Forms.DataGridView();
		this.ParameterName = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.ParameterValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.parameterLabel = new System.Windows.Forms.Label();
		this.isWaitCheck = new System.Windows.Forms.CheckBox();
		this.parameterLabel2 = new System.Windows.Forms.Label();
		this.paramText = new System.Windows.Forms.TextBox();
		this.restPanel = new System.Windows.Forms.TableLayoutPanel();
		this.urlTabel = new System.Windows.Forms.TableLayoutPanel();
		this.urlText = new System.Windows.Forms.TextBox();
		this.UrlLabel = new System.Windows.Forms.Label();
		this.downloadGroup = new System.Windows.Forms.GroupBox();
		this.downloadBaseTable = new System.Windows.Forms.TableLayoutPanel();
		this.downloadTable = new System.Windows.Forms.TableLayoutPanel();
		this.dounloadLabel = new System.Windows.Forms.Label();
		this.responseFileNameText = new System.Windows.Forms.TextBox();
		this.downloadStdLabel = new System.Windows.Forms.Label();
		this.responseNameLabel = new System.Windows.Forms.Label();
		this.isDownloadFolder = new System.Windows.Forms.CheckBox();
		this.isZip = new System.Windows.Forms.CheckBox();
		this.frozenLabel = new System.Windows.Forms.Label();
		this.unZipTable = new System.Windows.Forms.TableLayoutPanel();
		this.panel1 = new System.Windows.Forms.Panel();
		this.unZipFolderButton = new System.Windows.Forms.Button();
		this.unZipFolderText = new System.Windows.Forms.TextBox();
		this.downLoadSettingTable = new System.Windows.Forms.TableLayoutPanel();
		this.downloadCommandButton = new System.Windows.Forms.Button();
		this.downloadFolderCommandText = new System.Windows.Forms.TextBox();
		this.isResponseCheckBox = new System.Windows.Forms.CheckBox();
		this.uploadGroup = new System.Windows.Forms.GroupBox();
		this.uploadBaseTable = new System.Windows.Forms.TableLayoutPanel();
		this.uploadFileButton = new System.Windows.Forms.Button();
		this.isUpload = new System.Windows.Forms.CheckBox();
		this.uploadFileNameText = new System.Windows.Forms.TextBox();
		this.uploadRadio3 = new System.Windows.Forms.RadioButton();
		this.uploadRadio2 = new System.Windows.Forms.RadioButton();
		this.uploadRadio1 = new System.Windows.Forms.RadioButton();
		this.uploadAfterLabel = new System.Windows.Forms.Label();
		this.moveUploadFolderText = new System.Windows.Forms.TextBox();
		this.workFolderLabel = new System.Windows.Forms.Label();
		this.moveButton = new System.Windows.Forms.Button();
		this.baseTable.SuspendLayout();
		this.exePanel.SuspendLayout();
		this.exeTable.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.parameterGridView).BeginInit();
		this.restPanel.SuspendLayout();
		this.urlTabel.SuspendLayout();
		this.downloadGroup.SuspendLayout();
		this.downloadBaseTable.SuspendLayout();
		this.downloadTable.SuspendLayout();
		this.unZipTable.SuspendLayout();
		this.panel1.SuspendLayout();
		this.downLoadSettingTable.SuspendLayout();
		this.uploadGroup.SuspendLayout();
		this.uploadBaseTable.SuspendLayout();
		base.SuspendLayout();
		this.baseTable.ColumnCount = 2;
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 95f));
		this.baseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.baseTable.Controls.Add(this.waitLabel, 0, 3);
		this.baseTable.Controls.Add(this.restLabel, 0, 1);
		this.baseTable.Controls.Add(this.nameLabel, 0, 0);
		this.baseTable.Controls.Add(this.pathLabel, 0, 2);
		this.baseTable.Controls.Add(this.exePanel, 1, 2);
		this.baseTable.Controls.Add(this.nameText, 1, 0);
		this.baseTable.Controls.Add(this.parameterGridView, 1, 5);
		this.baseTable.Controls.Add(this.parameterLabel, 0, 5);
		this.baseTable.Controls.Add(this.isWaitCheck, 1, 3);
		this.baseTable.Controls.Add(this.parameterLabel2, 0, 4);
		this.baseTable.Controls.Add(this.paramText, 1, 4);
		this.baseTable.Controls.Add(this.restPanel, 1, 1);
		this.baseTable.Dock = System.Windows.Forms.DockStyle.Top;
		this.baseTable.Location = new System.Drawing.Point(10, 10);
		this.baseTable.Margin = new System.Windows.Forms.Padding(2);
		this.baseTable.Name = "baseTable";
		this.baseTable.RowCount = 7;
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 412f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 33f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 135f));
		this.baseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8f));
		this.baseTable.Size = new System.Drawing.Size(517, 648);
		this.baseTable.TabIndex = 10;
		this.waitLabel.AutoSize = true;
		this.waitLabel.BackColor = System.Drawing.SystemColors.Control;
		this.waitLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.waitLabel.Location = new System.Drawing.Point(3, 481);
		this.waitLabel.Margin = new System.Windows.Forms.Padding(3);
		this.waitLabel.Name = "waitLabel";
		this.waitLabel.Padding = new System.Windows.Forms.Padding(2, 0, 0, 0);
		this.waitLabel.Size = new System.Drawing.Size(89, 22);
		this.waitLabel.TabIndex = 17;
		this.waitLabel.Text = "PATH";
		this.waitLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.restLabel.AutoSize = true;
		this.restLabel.BackColor = System.Drawing.SystemColors.Control;
		this.restLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.restLabel.Location = new System.Drawing.Point(3, 33);
		this.restLabel.Margin = new System.Windows.Forms.Padding(3);
		this.restLabel.Name = "restLabel";
		this.restLabel.Padding = new System.Windows.Forms.Padding(2, 0, 0, 0);
		this.restLabel.Size = new System.Drawing.Size(89, 406);
		this.restLabel.TabIndex = 14;
		this.restLabel.Text = "PATH";
		this.restLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.nameLabel.AutoSize = true;
		this.nameLabel.BackColor = System.Drawing.SystemColors.Control;
		this.nameLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.nameLabel.Location = new System.Drawing.Point(3, 3);
		this.nameLabel.Margin = new System.Windows.Forms.Padding(3);
		this.nameLabel.Name = "nameLabel";
		this.nameLabel.Padding = new System.Windows.Forms.Padding(2, 0, 0, 0);
		this.nameLabel.Size = new System.Drawing.Size(89, 24);
		this.nameLabel.TabIndex = 13;
		this.nameLabel.Text = "PATH";
		this.nameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.pathLabel.AutoSize = true;
		this.pathLabel.BackColor = System.Drawing.SystemColors.Control;
		this.pathLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.pathLabel.Location = new System.Drawing.Point(3, 445);
		this.pathLabel.Margin = new System.Windows.Forms.Padding(3);
		this.pathLabel.Name = "pathLabel";
		this.pathLabel.Padding = new System.Windows.Forms.Padding(2, 0, 0, 0);
		this.pathLabel.Size = new System.Drawing.Size(89, 30);
		this.pathLabel.TabIndex = 10;
		this.pathLabel.Text = "PATH";
		this.pathLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.exePanel.Controls.Add(this.exeTable);
		this.exePanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.exePanel.Location = new System.Drawing.Point(98, 445);
		this.exePanel.Name = "exePanel";
		this.exePanel.Size = new System.Drawing.Size(416, 30);
		this.exePanel.TabIndex = 9;
		this.exeTable.ColumnCount = 2;
		this.exeTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 94.32433f));
		this.exeTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 5.675676f));
		this.exeTable.Controls.Add(this.pathText, 0, 0);
		this.exeTable.Controls.Add(this.pathFileButton, 1, 0);
		this.exeTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.exeTable.Location = new System.Drawing.Point(0, 0);
		this.exeTable.Name = "exeTable";
		this.exeTable.RowCount = 1;
		this.exeTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
		this.exeTable.Size = new System.Drawing.Size(416, 30);
		this.exeTable.TabIndex = 5;
		this.pathText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.pathText.Location = new System.Drawing.Point(3, 8);
		this.pathText.Name = "pathText";
		this.pathText.Size = new System.Drawing.Size(386, 19);
		this.pathText.TabIndex = 214;
		this.pathText.Validating += new System.ComponentModel.CancelEventHandler(pathText_Validating);
		this.pathFileButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("pathFileButton.BackgroundImage");
		this.pathFileButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.pathFileButton.Dock = System.Windows.Forms.DockStyle.Left;
		this.pathFileButton.Location = new System.Drawing.Point(395, 3);
		this.pathFileButton.Name = "pathFileButton";
		this.pathFileButton.Size = new System.Drawing.Size(15, 24);
		this.pathFileButton.TabIndex = 215;
		this.pathFileButton.UseVisualStyleBackColor = true;
		this.pathFileButton.Click += new System.EventHandler(fileButton_Click);
		this.nameText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.nameText.Location = new System.Drawing.Point(98, 8);
		this.nameText.Name = "nameText";
		this.nameText.Size = new System.Drawing.Size(416, 19);
		this.nameText.TabIndex = 201;
		this.nameText.Validating += new System.ComponentModel.CancelEventHandler(nameText_Validating);
		this.parameterGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
		this.parameterGridView.Columns.AddRange(this.ParameterName, this.ParameterValue);
		this.parameterGridView.Dock = System.Windows.Forms.DockStyle.Fill;
		this.parameterGridView.Location = new System.Drawing.Point(98, 542);
		this.parameterGridView.Name = "parameterGridView";
		this.parameterGridView.RowTemplate.Height = 21;
		this.parameterGridView.Size = new System.Drawing.Size(416, 129);
		this.parameterGridView.TabIndex = 220;
		this.parameterGridView.Validating += new System.ComponentModel.CancelEventHandler(parameterGridView_Validating);
		this.ParameterName.DataPropertyName = "ParameterName";
		this.ParameterName.HeaderText = "ParameterName";
		this.ParameterName.Name = "ParameterName";
		this.ParameterName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
		this.ParameterName.Width = 200;
		this.ParameterValue.DataPropertyName = "ParameterValue";
		this.ParameterValue.HeaderText = "ParameterValue";
		this.ParameterValue.Name = "ParameterValue";
		this.ParameterValue.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
		this.ParameterValue.Width = 200;
		this.parameterLabel.AutoSize = true;
		this.parameterLabel.BackColor = System.Drawing.SystemColors.Control;
		this.parameterLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.parameterLabel.Location = new System.Drawing.Point(3, 542);
		this.parameterLabel.Margin = new System.Windows.Forms.Padding(3);
		this.parameterLabel.Name = "parameterLabel";
		this.parameterLabel.Padding = new System.Windows.Forms.Padding(2, 0, 0, 0);
		this.parameterLabel.Size = new System.Drawing.Size(89, 129);
		this.parameterLabel.TabIndex = 12;
		this.parameterLabel.Text = "PATH";
		this.parameterLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.isWaitCheck.AutoSize = true;
		this.isWaitCheck.Dock = System.Windows.Forms.DockStyle.Left;
		this.isWaitCheck.Location = new System.Drawing.Point(98, 481);
		this.isWaitCheck.Name = "isWaitCheck";
		this.isWaitCheck.Size = new System.Drawing.Size(80, 22);
		this.isWaitCheck.TabIndex = 216;
		this.isWaitCheck.Text = "checkBox1";
		this.isWaitCheck.UseVisualStyleBackColor = true;
		this.isWaitCheck.CheckedChanged += new System.EventHandler(isWaitCheck_CheckedChanged);
		this.parameterLabel2.AutoSize = true;
		this.parameterLabel2.BackColor = System.Drawing.SystemColors.Control;
		this.parameterLabel2.Dock = System.Windows.Forms.DockStyle.Fill;
		this.parameterLabel2.Location = new System.Drawing.Point(3, 509);
		this.parameterLabel2.Margin = new System.Windows.Forms.Padding(3);
		this.parameterLabel2.Name = "parameterLabel2";
		this.parameterLabel2.Padding = new System.Windows.Forms.Padding(2, 0, 0, 0);
		this.parameterLabel2.Size = new System.Drawing.Size(89, 27);
		this.parameterLabel2.TabIndex = 16;
		this.parameterLabel2.Text = "PATH";
		this.parameterLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.paramText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.paramText.Location = new System.Drawing.Point(98, 517);
		this.paramText.Name = "paramText";
		this.paramText.Size = new System.Drawing.Size(416, 19);
		this.paramText.TabIndex = 217;
		this.paramText.Validating += new System.ComponentModel.CancelEventHandler(paramText_Validating);
		this.restPanel.ColumnCount = 1;
		this.restPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.restPanel.Controls.Add(this.urlTabel, 0, 0);
		this.restPanel.Controls.Add(this.downloadGroup, 0, 1);
		this.restPanel.Controls.Add(this.uploadGroup, 0, 2);
		this.restPanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.restPanel.Location = new System.Drawing.Point(98, 33);
		this.restPanel.Name = "restPanel";
		this.restPanel.RowCount = 3;
		this.restPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 37f));
		this.restPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 210f));
		this.restPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 33f));
		this.restPanel.Size = new System.Drawing.Size(416, 406);
		this.restPanel.TabIndex = 15;
		this.urlTabel.ColumnCount = 2;
		this.urlTabel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 77f));
		this.urlTabel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.urlTabel.Controls.Add(this.urlText, 1, 0);
		this.urlTabel.Controls.Add(this.UrlLabel, 0, 0);
		this.urlTabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.urlTabel.Location = new System.Drawing.Point(3, 3);
		this.urlTabel.Name = "urlTabel";
		this.urlTabel.RowCount = 1;
		this.urlTabel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 31f));
		this.urlTabel.Size = new System.Drawing.Size(410, 31);
		this.urlTabel.TabIndex = 13;
		this.urlText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.urlText.Location = new System.Drawing.Point(80, 9);
		this.urlText.Name = "urlText";
		this.urlText.Size = new System.Drawing.Size(327, 19);
		this.urlText.TabIndex = 202;
		this.urlText.Validating += new System.ComponentModel.CancelEventHandler(urlText_Validating);
		this.UrlLabel.BackColor = System.Drawing.Color.Transparent;
		this.UrlLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.UrlLabel.Location = new System.Drawing.Point(1, 0);
		this.UrlLabel.Margin = new System.Windows.Forms.Padding(1, 0, 3, 0);
		this.UrlLabel.Name = "UrlLabel";
		this.UrlLabel.Size = new System.Drawing.Size(73, 31);
		this.UrlLabel.TabIndex = 8;
		this.UrlLabel.Text = "URL";
		this.UrlLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.downloadGroup.Controls.Add(this.downloadBaseTable);
		this.downloadGroup.Dock = System.Windows.Forms.DockStyle.Fill;
		this.downloadGroup.Location = new System.Drawing.Point(3, 40);
		this.downloadGroup.Name = "downloadGroup";
		this.downloadGroup.Size = new System.Drawing.Size(410, 204);
		this.downloadGroup.TabIndex = 15;
		this.downloadGroup.TabStop = false;
		this.downloadGroup.Text = "groupBox1";
		this.downloadBaseTable.ColumnCount = 1;
		this.downloadBaseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50f));
		this.downloadBaseTable.Controls.Add(this.downloadTable, 0, 1);
		this.downloadBaseTable.Controls.Add(this.isResponseCheckBox, 0, 0);
		this.downloadBaseTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.downloadBaseTable.Location = new System.Drawing.Point(3, 15);
		this.downloadBaseTable.Margin = new System.Windows.Forms.Padding(1);
		this.downloadBaseTable.Name = "downloadBaseTable";
		this.downloadBaseTable.RowCount = 2;
		this.downloadBaseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.89362f));
		this.downloadBaseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 85.10638f));
		this.downloadBaseTable.Size = new System.Drawing.Size(404, 186);
		this.downloadBaseTable.TabIndex = 13;
		this.downloadTable.ColumnCount = 2;
		this.downloadTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 130f));
		this.downloadTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.downloadTable.Controls.Add(this.dounloadLabel, 1, 0);
		this.downloadTable.Controls.Add(this.responseFileNameText, 1, 2);
		this.downloadTable.Controls.Add(this.downloadStdLabel, 0, 0);
		this.downloadTable.Controls.Add(this.responseNameLabel, 0, 2);
		this.downloadTable.Controls.Add(this.isDownloadFolder, 0, 1);
		this.downloadTable.Controls.Add(this.isZip, 0, 3);
		this.downloadTable.Controls.Add(this.frozenLabel, 1, 3);
		this.downloadTable.Controls.Add(this.unZipTable, 1, 4);
		this.downloadTable.Controls.Add(this.downLoadSettingTable, 1, 1);
		this.downloadTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.downloadTable.Location = new System.Drawing.Point(1, 28);
		this.downloadTable.Margin = new System.Windows.Forms.Padding(1);
		this.downloadTable.Name = "downloadTable";
		this.downloadTable.RowCount = 5;
		this.downloadTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 31f));
		this.downloadTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 37f));
		this.downloadTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 27f));
		this.downloadTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29f));
		this.downloadTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8f));
		this.downloadTable.Size = new System.Drawing.Size(402, 157);
		this.downloadTable.TabIndex = 215;
		this.dounloadLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.dounloadLabel.Location = new System.Drawing.Point(133, 0);
		this.dounloadLabel.Name = "dounloadLabel";
		this.dounloadLabel.Size = new System.Drawing.Size(266, 31);
		this.dounloadLabel.TabIndex = 12;
		this.dounloadLabel.Text = "label1";
		this.dounloadLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.responseFileNameText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.responseFileNameText.Enabled = false;
		this.responseFileNameText.Location = new System.Drawing.Point(133, 73);
		this.responseFileNameText.Name = "responseFileNameText";
		this.responseFileNameText.Size = new System.Drawing.Size(266, 19);
		this.responseFileNameText.TabIndex = 210;
		this.responseFileNameText.Validating += new System.ComponentModel.CancelEventHandler(responseFileNameText_Validating);
		this.downloadStdLabel.AutoSize = true;
		this.downloadStdLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.downloadStdLabel.Location = new System.Drawing.Point(3, 0);
		this.downloadStdLabel.Name = "downloadStdLabel";
		this.downloadStdLabel.Size = new System.Drawing.Size(124, 31);
		this.downloadStdLabel.TabIndex = 12;
		this.downloadStdLabel.Text = "既定のダウン";
		this.downloadStdLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.responseNameLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.responseNameLabel.Location = new System.Drawing.Point(3, 68);
		this.responseNameLabel.Name = "responseNameLabel";
		this.responseNameLabel.Size = new System.Drawing.Size(124, 27);
		this.responseNameLabel.TabIndex = 11;
		this.responseNameLabel.Text = "response";
		this.responseNameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.isDownloadFolder.AutoSize = true;
		this.isDownloadFolder.Dock = System.Windows.Forms.DockStyle.Left;
		this.isDownloadFolder.Location = new System.Drawing.Point(3, 34);
		this.isDownloadFolder.Name = "isDownloadFolder";
		this.isDownloadFolder.Size = new System.Drawing.Size(80, 31);
		this.isDownloadFolder.TabIndex = 222;
		this.isDownloadFolder.Text = "checkBox1";
		this.isDownloadFolder.UseVisualStyleBackColor = true;
		this.isDownloadFolder.CheckedChanged += new System.EventHandler(isDownloadFolder_CheckedChanged);
		this.isZip.AutoSize = true;
		this.isZip.Dock = System.Windows.Forms.DockStyle.Left;
		this.isZip.Enabled = false;
		this.isZip.Location = new System.Drawing.Point(3, 98);
		this.isZip.Name = "isZip";
		this.isZip.Size = new System.Drawing.Size(80, 23);
		this.isZip.TabIndex = 212;
		this.isZip.Text = "checkBox1";
		this.isZip.UseVisualStyleBackColor = true;
		this.isZip.CheckedChanged += new System.EventHandler(isZip_CheckedChanged);
		this.frozenLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.frozenLabel.Location = new System.Drawing.Point(133, 95);
		this.frozenLabel.Name = "frozenLabel";
		this.frozenLabel.Size = new System.Drawing.Size(266, 29);
		this.frozenLabel.TabIndex = 221;
		this.frozenLabel.Text = "label1";
		this.frozenLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.unZipTable.ColumnCount = 2;
		this.unZipTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50f));
		this.unZipTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 34f));
		this.unZipTable.Controls.Add(this.panel1, 1, 0);
		this.unZipTable.Controls.Add(this.unZipFolderText, 0, 0);
		this.unZipTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.unZipTable.Location = new System.Drawing.Point(130, 124);
		this.unZipTable.Margin = new System.Windows.Forms.Padding(0);
		this.unZipTable.Name = "unZipTable";
		this.unZipTable.RowCount = 1;
		this.unZipTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
		this.unZipTable.Size = new System.Drawing.Size(272, 33);
		this.unZipTable.TabIndex = 213;
		this.panel1.Controls.Add(this.unZipFolderButton);
		this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
		this.panel1.Location = new System.Drawing.Point(241, 3);
		this.panel1.Name = "panel1";
		this.panel1.Size = new System.Drawing.Size(28, 27);
		this.panel1.TabIndex = 215;
		this.unZipFolderButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("unZipFolderButton.BackgroundImage");
		this.unZipFolderButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.unZipFolderButton.Dock = System.Windows.Forms.DockStyle.Left;
		this.unZipFolderButton.Enabled = false;
		this.unZipFolderButton.Location = new System.Drawing.Point(0, 0);
		this.unZipFolderButton.Name = "unZipFolderButton";
		this.unZipFolderButton.Size = new System.Drawing.Size(27, 27);
		this.unZipFolderButton.TabIndex = 213;
		this.unZipFolderButton.UseVisualStyleBackColor = true;
		this.unZipFolderButton.Click += new System.EventHandler(unZipFolderButton_Click);
		this.unZipFolderText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.unZipFolderText.Location = new System.Drawing.Point(0, 11);
		this.unZipFolderText.Margin = new System.Windows.Forms.Padding(0, 3, 3, 3);
		this.unZipFolderText.Name = "unZipFolderText";
		this.unZipFolderText.Size = new System.Drawing.Size(235, 19);
		this.unZipFolderText.TabIndex = 216;
		this.unZipFolderText.Validating += new System.ComponentModel.CancelEventHandler(unZipFolderText_Validating);
		this.downLoadSettingTable.ColumnCount = 2;
		this.downLoadSettingTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50f));
		this.downLoadSettingTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 32f));
		this.downLoadSettingTable.Controls.Add(this.downloadCommandButton, 1, 0);
		this.downLoadSettingTable.Controls.Add(this.downloadFolderCommandText, 0, 0);
		this.downLoadSettingTable.Dock = System.Windows.Forms.DockStyle.Fill;
		this.downLoadSettingTable.Location = new System.Drawing.Point(133, 34);
		this.downLoadSettingTable.Name = "downLoadSettingTable";
		this.downLoadSettingTable.RowCount = 1;
		this.downLoadSettingTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
		this.downLoadSettingTable.Size = new System.Drawing.Size(266, 31);
		this.downLoadSettingTable.TabIndex = 223;
		this.downloadCommandButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("downloadCommandButton.BackgroundImage");
		this.downloadCommandButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.downloadCommandButton.Dock = System.Windows.Forms.DockStyle.Left;
		this.downloadCommandButton.Enabled = false;
		this.downloadCommandButton.Location = new System.Drawing.Point(237, 3);
		this.downloadCommandButton.Name = "downloadCommandButton";
		this.downloadCommandButton.Size = new System.Drawing.Size(26, 25);
		this.downloadCommandButton.TabIndex = 214;
		this.downloadCommandButton.UseVisualStyleBackColor = true;
		this.downloadCommandButton.Click += new System.EventHandler(unZipFolderButton_Click);
		this.downloadFolderCommandText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.downloadFolderCommandText.Enabled = false;
		this.downloadFolderCommandText.Location = new System.Drawing.Point(0, 9);
		this.downloadFolderCommandText.Margin = new System.Windows.Forms.Padding(0, 3, 3, 3);
		this.downloadFolderCommandText.Name = "downloadFolderCommandText";
		this.downloadFolderCommandText.Size = new System.Drawing.Size(231, 19);
		this.downloadFolderCommandText.TabIndex = 223;
		this.downloadFolderCommandText.Validating += new System.ComponentModel.CancelEventHandler(downloadFolderCommandText_Validating);
		this.isResponseCheckBox.AutoSize = true;
		this.isResponseCheckBox.Location = new System.Drawing.Point(3, 3);
		this.isResponseCheckBox.Name = "isResponseCheckBox";
		this.isResponseCheckBox.Size = new System.Drawing.Size(80, 16);
		this.isResponseCheckBox.TabIndex = 203;
		this.isResponseCheckBox.Text = "checkBox1";
		this.isResponseCheckBox.UseVisualStyleBackColor = true;
		this.isResponseCheckBox.CheckedChanged += new System.EventHandler(isResponseCheckBox_CheckedChanged);
		this.uploadGroup.Controls.Add(this.uploadBaseTable);
		this.uploadGroup.Dock = System.Windows.Forms.DockStyle.Fill;
		this.uploadGroup.Location = new System.Drawing.Point(3, 250);
		this.uploadGroup.Name = "uploadGroup";
		this.uploadGroup.Size = new System.Drawing.Size(410, 153);
		this.uploadGroup.TabIndex = 16;
		this.uploadGroup.TabStop = false;
		this.uploadGroup.Text = "groupBox1";
		this.uploadBaseTable.ColumnCount = 3;
		this.uploadBaseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 173f));
		this.uploadBaseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));
		this.uploadBaseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 36f));
		this.uploadBaseTable.Controls.Add(this.uploadFileButton, 2, 0);
		this.uploadBaseTable.Controls.Add(this.isUpload, 0, 0);
		this.uploadBaseTable.Controls.Add(this.uploadFileNameText, 1, 0);
		this.uploadBaseTable.Controls.Add(this.uploadRadio3, 0, 4);
		this.uploadBaseTable.Controls.Add(this.uploadRadio2, 0, 3);
		this.uploadBaseTable.Controls.Add(this.uploadRadio1, 0, 2);
		this.uploadBaseTable.Controls.Add(this.uploadAfterLabel, 0, 1);
		this.uploadBaseTable.Controls.Add(this.moveUploadFolderText, 1, 4);
		this.uploadBaseTable.Controls.Add(this.workFolderLabel, 1, 3);
		this.uploadBaseTable.Controls.Add(this.moveButton, 2, 4);
		this.uploadBaseTable.Dock = System.Windows.Forms.DockStyle.Top;
		this.uploadBaseTable.Location = new System.Drawing.Point(3, 15);
		this.uploadBaseTable.Name = "uploadBaseTable";
		this.uploadBaseTable.RowCount = 5;
		this.uploadBaseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32f));
		this.uploadBaseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 22f));
		this.uploadBaseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 21f));
		this.uploadBaseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24f));
		this.uploadBaseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 23f));
		this.uploadBaseTable.Size = new System.Drawing.Size(404, 130);
		this.uploadBaseTable.TabIndex = 14;
		this.uploadFileButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("uploadFileButton.BackgroundImage");
		this.uploadFileButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.uploadFileButton.Location = new System.Drawing.Point(371, 3);
		this.uploadFileButton.Name = "uploadFileButton";
		this.uploadFileButton.Size = new System.Drawing.Size(25, 25);
		this.uploadFileButton.TabIndex = 216;
		this.uploadFileButton.UseVisualStyleBackColor = true;
		this.uploadFileButton.Click += new System.EventHandler(fileButton_Click);
		this.isUpload.AutoSize = true;
		this.isUpload.Dock = System.Windows.Forms.DockStyle.Left;
		this.isUpload.Location = new System.Drawing.Point(3, 3);
		this.isUpload.Name = "isUpload";
		this.isUpload.Size = new System.Drawing.Size(80, 26);
		this.isUpload.TabIndex = 211;
		this.isUpload.Text = "checkBox1";
		this.isUpload.UseVisualStyleBackColor = true;
		this.isUpload.CheckedChanged += new System.EventHandler(isUpload_CheckedChanged);
		this.uploadFileNameText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.uploadFileNameText.Enabled = false;
		this.uploadFileNameText.Location = new System.Drawing.Point(176, 10);
		this.uploadFileNameText.Name = "uploadFileNameText";
		this.uploadFileNameText.Size = new System.Drawing.Size(189, 19);
		this.uploadFileNameText.TabIndex = 212;
		this.uploadFileNameText.Validating += new System.ComponentModel.CancelEventHandler(uploadFileNameText_Validating);
		this.uploadRadio3.AutoSize = true;
		this.uploadRadio3.Enabled = false;
		this.uploadRadio3.Location = new System.Drawing.Point(3, 102);
		this.uploadRadio3.Name = "uploadRadio3";
		this.uploadRadio3.Size = new System.Drawing.Size(88, 16);
		this.uploadRadio3.TabIndex = 219;
		this.uploadRadio3.TabStop = true;
		this.uploadRadio3.Text = "radioButton3";
		this.uploadRadio3.UseVisualStyleBackColor = true;
		this.uploadRadio3.CheckedChanged += new System.EventHandler(uploadSelectChanged);
		this.uploadRadio2.AutoSize = true;
		this.uploadRadio2.Enabled = false;
		this.uploadRadio2.Location = new System.Drawing.Point(3, 78);
		this.uploadRadio2.Name = "uploadRadio2";
		this.uploadRadio2.Size = new System.Drawing.Size(88, 16);
		this.uploadRadio2.TabIndex = 218;
		this.uploadRadio2.TabStop = true;
		this.uploadRadio2.Text = "radioButton2";
		this.uploadRadio2.UseVisualStyleBackColor = true;
		this.uploadRadio2.CheckedChanged += new System.EventHandler(uploadSelectChanged);
		this.uploadRadio1.AutoSize = true;
		this.uploadRadio1.Checked = true;
		this.uploadRadio1.Enabled = false;
		this.uploadRadio1.Location = new System.Drawing.Point(3, 57);
		this.uploadRadio1.Name = "uploadRadio1";
		this.uploadRadio1.Size = new System.Drawing.Size(88, 15);
		this.uploadRadio1.TabIndex = 217;
		this.uploadRadio1.TabStop = true;
		this.uploadRadio1.Text = "radioButton1";
		this.uploadRadio1.UseVisualStyleBackColor = true;
		this.uploadRadio1.CheckedChanged += new System.EventHandler(uploadSelectChanged);
		this.uploadAfterLabel.AutoSize = true;
		this.uploadAfterLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.uploadAfterLabel.Location = new System.Drawing.Point(3, 32);
		this.uploadAfterLabel.Name = "uploadAfterLabel";
		this.uploadAfterLabel.Size = new System.Drawing.Size(167, 22);
		this.uploadAfterLabel.TabIndex = 223;
		this.uploadAfterLabel.Text = "label1";
		this.uploadAfterLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.moveUploadFolderText.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.moveUploadFolderText.Location = new System.Drawing.Point(176, 108);
		this.moveUploadFolderText.Name = "moveUploadFolderText";
		this.moveUploadFolderText.Size = new System.Drawing.Size(189, 19);
		this.moveUploadFolderText.TabIndex = 222;
		this.moveUploadFolderText.Validating += new System.ComponentModel.CancelEventHandler(moveUploadFolderText_Validating);
		this.workFolderLabel.AutoSize = true;
		this.workFolderLabel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.workFolderLabel.Location = new System.Drawing.Point(176, 75);
		this.workFolderLabel.Name = "workFolderLabel";
		this.workFolderLabel.Size = new System.Drawing.Size(189, 24);
		this.workFolderLabel.TabIndex = 221;
		this.workFolderLabel.Text = "label2";
		this.workFolderLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.moveButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("moveButton.BackgroundImage");
		this.moveButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.moveButton.Dock = System.Windows.Forms.DockStyle.Left;
		this.moveButton.Enabled = false;
		this.moveButton.Location = new System.Drawing.Point(371, 102);
		this.moveButton.Name = "moveButton";
		this.moveButton.Size = new System.Drawing.Size(26, 25);
		this.moveButton.TabIndex = 215;
		this.moveButton.UseVisualStyleBackColor = true;
		this.moveButton.Click += new System.EventHandler(folderButton_Click);
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.AutoScroll = true;
		this.BackColor = System.Drawing.Color.White;
		base.Controls.Add(this.baseTable);
		base.Name = "CommandView";
		base.Padding = new System.Windows.Forms.Padding(10);
		base.Size = new System.Drawing.Size(537, 484);
		base.Load += new System.EventHandler(CommandView_Load);
		this.baseTable.ResumeLayout(false);
		this.baseTable.PerformLayout();
		this.exePanel.ResumeLayout(false);
		this.exeTable.ResumeLayout(false);
		this.exeTable.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.parameterGridView).EndInit();
		this.restPanel.ResumeLayout(false);
		this.urlTabel.ResumeLayout(false);
		this.urlTabel.PerformLayout();
		this.downloadGroup.ResumeLayout(false);
		this.downloadBaseTable.ResumeLayout(false);
		this.downloadBaseTable.PerformLayout();
		this.downloadTable.ResumeLayout(false);
		this.downloadTable.PerformLayout();
		this.unZipTable.ResumeLayout(false);
		this.unZipTable.PerformLayout();
		this.panel1.ResumeLayout(false);
		this.downLoadSettingTable.ResumeLayout(false);
		this.downLoadSettingTable.PerformLayout();
		this.uploadGroup.ResumeLayout(false);
		this.uploadBaseTable.ResumeLayout(false);
		this.uploadBaseTable.PerformLayout();
		base.ResumeLayout(false);
	}
}
