using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ConMasGeneratorLib.Data;
using ConMasGeneratorLib.Job;
using ConMasGeneratorUtility.Properties;
using ConMasGeneratorUtility.Utils;

namespace ConMasGeneratorUtility.Views;

public class JobSettingView : UserControl
{
	private List<JobData> _jobDatas;

	internal List<ProcessData> _conMasProcessList;

	internal List<Command> _conMasCommandList;

	private ToolTip ToolTip;

	private string _jobMode = string.Empty;

	private IContainer components = null;

	private SplitContainer splitContainer;

	private TreeView JobTreeView;

	private Panel variablePanel;

	private Button saveButton;

	private Panel controllPanel;

	private Button addButton;

	private Button deleteButton;

	private Button up;

	private Button down;

	private Label line;

	private Button refreshButton;

	private Panel buttonPanel;

	private ContextMenuStrip contextMenu;

	private ToolStripMenuItem addItem;

	private ToolStripMenuItem deleteItem;

	private ImageList imageList1;

	private Button legendButton;

	internal List<ProcessData> SelectedConMasProcesses { get; set; }

	internal List<Command> SelectedConMasCommands { get; set; }

	public JobSettingView(string mode)
	{
		InitializeComponent();
		SetFormResouce();
		_jobDatas = new List<JobData>();
		_jobMode = mode;
		DoubleBuffered = true;
	}

	private void SetJobTree()
	{
		try
		{
			JobController jobController = new JobController();
			_jobDatas = jobController.GetJobData();
			JobTreeView.Nodes.Clear();
			List<JobData> list;
			if (_jobMode == "watch")
			{
				list = _jobDatas.Where((JobData child) => child.JobType == "0").ToList();
				JobTreeView.Nodes.Add(Resources.WatcherJob);
			}
			else
			{
				list = _jobDatas.Where((JobData child) => child.JobType == "1").ToList();
				JobTreeView.Nodes.Add(Resources.SchduleJob);
			}
			foreach (JobData item in list)
			{
				TreeNode treeNode = new TreeNode();
				treeNode.Text = item.Name;
				treeNode.Tag = item;
				foreach (ProcessData process in item.Processes)
				{
					TreeNode treeNode2 = new TreeNode();
					treeNode2.Text = process.Name;
					treeNode2.Tag = process;
					foreach (Command command in process.Commands)
					{
						TreeNode treeNode3 = new TreeNode();
						treeNode3.Text = command.Name;
						treeNode3.Tag = command;
						treeNode2.Nodes.Add(treeNode3);
					}
					treeNode.Nodes.Add(treeNode2);
				}
				JobTreeView.Nodes[0].Nodes.Add(treeNode);
			}
			if (JobTreeView.Nodes[0].Nodes.Count > 0)
			{
				JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[0];
			}
			else
			{
				JobTreeView.SelectedNode = JobTreeView.Nodes[0];
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void JobSettingView_Load(object sender, EventArgs e)
	{
		try
		{
			init();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void init()
	{
		try
		{
			SetLayout();
			SetJobTree();
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void Tab_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void SetLayout()
	{
		try
		{
			splitContainer.Location = new Point(0, controllPanel.Height);
			splitContainer.Height = base.Height - controllPanel.Height;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void JobSettingView_Resize(object sender, EventArgs e)
	{
		try
		{
			SetLayout();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void JobTreeView_AfterSelect(object sender, TreeViewEventArgs e)
	{
		try
		{
			variablePanel.Controls.Clear();
			legendButton.Visible = false;
			if (e.Node.Tag == null)
			{
				return;
			}
			if (e.Node.Tag.GetType() == typeof(JobData))
			{
				if (_jobMode == "watch")
				{
					JobWatcherView jobWatcherView = new JobWatcherView(e.Node);
					jobWatcherView.Dock = DockStyle.Fill;
					variablePanel.Controls.Add(jobWatcherView);
				}
				else
				{
					JobSchduleView jobSchduleView = new JobSchduleView(e.Node);
					jobSchduleView.Dock = DockStyle.Fill;
					variablePanel.Controls.Add(jobSchduleView);
				}
			}
			else if (e.Node.Tag.GetType() == typeof(ProcessData))
			{
				ProcessView processView = new ProcessView(e.Node);
				processView.Dock = DockStyle.Fill;
				variablePanel.Controls.Add(processView);
			}
			else if (e.Node.Tag.GetType() == typeof(Command))
			{
				legendButton.Visible = true;
				CommandView commandView = new CommandView(e.Node);
				commandView.Dock = DockStyle.Fill;
				variablePanel.Controls.Add(commandView);
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void jobAddButton_Click(object sender, EventArgs e)
	{
		try
		{
			AddNewJob();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void AddNewJob()
	{
		try
		{
			if (JobTreeView.Nodes[0].Tag == null)
			{
				TreeNode treeNode = new TreeNode();
				treeNode.Text = "New Job";
				JobData jobData = new JobData();
				jobData.Name = treeNode.Text;
				jobData.Id = CreateJobId();
				jobData.Enable = "0";
				if (_jobMode == "watch")
				{
					jobData.JobType = "0";
					jobData.Type = "0";
					treeNode.Tag = jobData;
					_jobDatas.Add(jobData);
					JobTreeView.Nodes[0].Nodes.Add(treeNode);
					JobTreeView.SelectedNode = treeNode;
				}
				else
				{
					jobData.JobType = "0";
					jobData.Type = "1";
					jobData.Timing = "1";
					jobData.Enable = "0";
					treeNode.Tag = jobData;
					_jobDatas.Add(jobData);
					JobTreeView.Nodes[0].Nodes.Add(treeNode);
					JobTreeView.SelectedNode = treeNode;
				}
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void AddNewProcess()
	{
		try
		{
			using (JobController jobController = new JobController())
			{
				_conMasProcessList = jobController.GetConMasProcessDatas();
			}
			using ConMasProcessViewer conMasProcessViewer = new ConMasProcessViewer(this);
			if (conMasProcessViewer.ShowDialog() != DialogResult.OK || SelectedConMasProcesses == null || SelectedConMasProcesses.Count <= 0)
			{
				return;
			}
			string jobId = (JobTreeView.SelectedNode.Tag as JobData).Id;
			foreach (ProcessData selectedConMasProcess in SelectedConMasProcesses)
			{
				TreeNode selectedNode = JobTreeView.SelectedNode;
				TreeNode treeNode = new TreeNode();
				treeNode.Text = selectedConMasProcess.Name;
				treeNode.Tag = selectedConMasProcess;
				foreach (Command command in selectedConMasProcess.Commands)
				{
					TreeNode treeNode2 = new TreeNode();
					treeNode2.Text = command.Name;
					treeNode2.Tag = command;
					treeNode.Nodes.Add(treeNode2);
				}
				selectedNode.Nodes.Add(treeNode);
				_jobDatas.Where((JobData child) => child.Id == jobId).First().Processes.Add(selectedConMasProcess);
			}
			JobTreeView.SelectedNode.Expand();
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void AddNewCommand()
	{
		try
		{
			string type = (JobTreeView.SelectedNode.Tag as ProcessData).Type;
			using (JobController jobController = new JobController())
			{
				_conMasCommandList = jobController.GetConMasCommands(type);
			}
			using ConMasCommandViewer conMasCommandViewer = new ConMasCommandViewer(this, type);
			if (conMasCommandViewer.ShowDialog() != DialogResult.OK || SelectedConMasCommands == null || SelectedConMasCommands.Count <= 0)
			{
				return;
			}
			foreach (Command selectedConMasCommand in SelectedConMasCommands)
			{
				TreeNode selectedNode = JobTreeView.SelectedNode;
				TreeNode treeNode = new TreeNode();
				treeNode.Text = selectedConMasCommand.Name;
				treeNode.Tag = selectedConMasCommand;
				selectedNode.Nodes.Add(treeNode);
				(JobTreeView.SelectedNode.Tag as ProcessData).Commands.Add(selectedConMasCommand);
			}
			JobTreeView.SelectedNode.Expand();
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private string CreateJobId()
	{
		try
		{
			return Guid.NewGuid().ToString();
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void AddNode()
	{
		try
		{
			if (JobTreeView.SelectedNode != null)
			{
				if (JobTreeView.SelectedNode.Tag == null)
				{
					AddNewJob();
				}
				else if (JobTreeView.SelectedNode.Tag.GetType() == typeof(JobData))
				{
					AddNewProcess();
				}
				else if (JobTreeView.SelectedNode.Tag.GetType() == typeof(ProcessData))
				{
					AddNewCommand();
				}
				else if (!(JobTreeView.SelectedNode.Tag.GetType() == typeof(Command)))
				{
				}
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void RemoveNode()
	{
		try
		{
			if (JobTreeView.SelectedNode.Tag != null)
			{
				DialogResult dialogResult = MessageBox.Show(Resources.DeleteMessage, "INFOMATION", MessageBoxButtons.OKCancel);
				if (dialogResult != DialogResult.Cancel)
				{
					JobTreeView.Nodes.Remove(JobTreeView.SelectedNode);
				}
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void addButton_Click(object sender, EventArgs e)
	{
		try
		{
			AddNode();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void deleteButton_Click(object sender, EventArgs e)
	{
		try
		{
			RemoveNode();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void down_Click(object sender, EventArgs e)
	{
		try
		{
			if (JobTreeView.SelectedNode == null || JobTreeView.SelectedNode.Tag == null)
			{
				return;
			}
			int index = JobTreeView.SelectedNode.Index;
			TreeNode treeNode = JobTreeView.SelectedNode.Parent;
			int num = ((treeNode != null) ? (treeNode.Nodes.Count - 1) : (JobTreeView.Nodes.Count - 1));
			if (index != num)
			{
				TreeNode selectedNode = JobTreeView.SelectedNode;
				JobTreeView.SelectedNode.Remove();
				if (treeNode == null)
				{
					JobTreeView.Nodes.Insert(index + 1, selectedNode);
				}
				else
				{
					treeNode.Nodes.Insert(index + 1, selectedNode);
				}
				JobTreeView.SelectedNode = selectedNode;
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void up_Click(object sender, EventArgs e)
	{
		try
		{
			if (JobTreeView.SelectedNode == null || JobTreeView.SelectedNode.Tag == null)
			{
				return;
			}
			int index = JobTreeView.SelectedNode.Index;
			TreeNode treeNode = JobTreeView.SelectedNode.Parent;
			if (index != 0)
			{
				TreeNode selectedNode = JobTreeView.SelectedNode;
				JobTreeView.SelectedNode.Remove();
				if (treeNode == null)
				{
					JobTreeView.Nodes.Insert(index - 1, selectedNode);
				}
				else
				{
					treeNode.Nodes.Insert(index - 1, selectedNode);
				}
				JobTreeView.SelectedNode = selectedNode;
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private List<JobData> CreateTreeNodeJobData()
	{
		try
		{
			List<JobData> list = new List<JobData>();
			foreach (TreeNode node in JobTreeView.Nodes[0].Nodes)
			{
				JobData jobData = node.Tag as JobData;
				jobData.Processes.Clear();
				foreach (TreeNode node2 in node.Nodes)
				{
					ProcessData processData = node2.Tag as ProcessData;
					processData.Commands.Clear();
					foreach (TreeNode node3 in node2.Nodes)
					{
						Command item = node3.Tag as Command;
						processData.Commands.Add(item);
					}
					jobData.Processes.Add(processData);
				}
				list.Add(jobData);
			}
			return list;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void saveButton_Click(object sender, EventArgs e)
	{
		try
		{
			List<JobData> list = CreateTreeNodeJobData();
			string message = string.Empty;
			if (!IsValidJob(list, ref message))
			{
				MessageBox.Show(message, "WARNING");
				return;
			}
			using JobController jobController = new JobController();
			if (_jobMode == "watch")
			{
				if (MessageBox.Show(Resources.WatcherJobUpdateMessage, "", MessageBoxButtons.OKCancel) != DialogResult.Cancel)
				{
					Cursor.Current = Cursors.WaitCursor;
					if (jobController.SaveJob(JobController.JobType.WatchType, list))
					{
						FormUtility.StartService();
					}
				}
			}
			else
			{
				if (MessageBox.Show(Resources.SchduleJobUpdateMessage, "", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
				{
					return;
				}
				Cursor.Current = Cursors.WaitCursor;
				if (jobController.SaveJob(JobController.JobType.SchduleType, list))
				{
					using (JobController jobController2 = new JobController(JobController.JobType.SchduleType))
					{
						jobController2.RegistSchJob();
						return;
					}
				}
			}
		}
		catch (Exception ex)
		{
			Cursor.Current = Cursors.Default;
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
		finally
		{
			Cursor.Current = Cursors.Default;
		}
	}

	private bool IsValidJob(List<JobData> checkJobs, ref string message)
	{
		try
		{
			bool result = false;
			int num = 0;
			int num2 = 0;
			int num3 = 0;
			foreach (JobData checkJob in checkJobs)
			{
				if (string.IsNullOrEmpty(checkJob.Name))
				{
					JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
					message = Resources.NameEmptyMessage;
					return result;
				}
				if (_jobMode == "watch")
				{
					if (string.IsNullOrEmpty(checkJob.Folder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NotSetMessage, Resources.WatchFolder);
						return result;
					}
					if (string.IsNullOrEmpty(checkJob.NormalEndFolder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NotSetMessage, Resources.NormalFolder);
						return result;
					}
					if (string.IsNullOrEmpty(checkJob.AbnomalEndFolder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NotSetMessage, Resources.AbnormalFolder);
						return result;
					}
					if (string.IsNullOrEmpty(checkJob.WorkFolder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NotSetMessage, Resources.WorkFolder);
						return result;
					}
					if (string.IsNullOrEmpty(checkJob.DownLoadFolder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NotSetMessage, Resources.DownLoadFolder);
						return result;
					}
					if (!Directory.Exists(checkJob.Folder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NoExsistFolderMessage, Resources.WatchFolder);
						return result;
					}
					if (!Directory.Exists(checkJob.NormalEndFolder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NoExsistFolderMessage, Resources.NormalFolder);
						return result;
					}
					if (!Directory.Exists(checkJob.AbnomalEndFolder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NoExsistFolderMessage, Resources.AbnormalFolder);
						return result;
					}
					if (!Directory.Exists(checkJob.WorkFolder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NoExsistFolderMessage, Resources.WorkFolder);
						return result;
					}
					if (!Directory.Exists(checkJob.DownLoadFolder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NoExsistFolderMessage, Resources.DownLoadFolder);
						return result;
					}
				}
				else
				{
					if (string.IsNullOrEmpty(checkJob.StartTime))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NotEnteredMessage, Resources.SchduleStartDate);
						return result;
					}
					if (checkJob.Type == "1" && !checkJob.IsRetry && string.IsNullOrEmpty(checkJob.Timing))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NotEnteredMessage, Resources.SchduleTiming);
						return result;
					}
					if (checkJob.Type == "2" && string.IsNullOrEmpty(checkJob.Week))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NotEnteredMessage, Resources.Week);
						return result;
					}
					if (checkJob.Type == "1" && checkJob.IsRetry && string.IsNullOrEmpty(checkJob.RetryTiming))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NotEnteredMessage, Resources.SchduleRetryTiming);
						return result;
					}
					if (checkJob.Type == "3")
					{
						if (string.IsNullOrEmpty(checkJob.Month))
						{
							JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
							message = string.Format(Resources.NotEnteredMessage, Resources.Month);
							return result;
						}
						if (string.IsNullOrEmpty(checkJob.Day))
						{
							JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
							message = string.Format(Resources.NotEnteredMessage, Resources.SchduleMonthDays);
							return result;
						}
					}
					if (string.IsNullOrEmpty(checkJob.WorkFolder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NotSetMessage, Resources.WorkFolder);
						return result;
					}
					if (string.IsNullOrEmpty(checkJob.DownLoadFolder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NotSetMessage, Resources.DownLoadFolder);
						return result;
					}
					if (!Directory.Exists(checkJob.WorkFolder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NoExsistFolderMessage, Resources.WorkFolder);
						return result;
					}
					if (!Directory.Exists(checkJob.DownLoadFolder))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
						message = string.Format(Resources.NoExsistFolderMessage, Resources.DownLoadFolder);
						return result;
					}
				}
				if (checkJob.Processes == null || checkJob.Processes.Count == 0)
				{
					JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num];
					message = string.Format(Resources.NotExsistJobItemMessage, Resources.Job, Resources.Process);
					return result;
				}
				num2 = 0;
				num3 = 0;
				foreach (ProcessData process in checkJob.Processes)
				{
					if (string.IsNullOrEmpty(process.Name))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num].Nodes[num2];
						message = Resources.NameEmptyMessage;
						return result;
					}
					if (string.IsNullOrEmpty(process.Type))
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num].Nodes[num2];
						message = string.Format(Resources.NotEnteredMessage, Resources.ProcessType);
						return result;
					}
					if (process.Commands == null || process.Commands.Count == 0)
					{
						JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num].Nodes[num2];
						message = string.Format(Resources.NotExsistJobItemMessage, Resources.Process, Resources.Command);
						return result;
					}
					num3 = 0;
					foreach (Command command in process.Commands)
					{
						if (string.IsNullOrEmpty(command.Name))
						{
							JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num].Nodes[num2].Nodes[num3];
							message = Resources.NameEmptyMessage;
							return result;
						}
						if (process.Type == "0")
						{
							if (string.IsNullOrEmpty(command.UrlOrg))
							{
								JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num].Nodes[num2].Nodes[num3];
								message = string.Format(Resources.NotEnteredMessage, Resources.Url);
								return result;
							}
							if (command.IsResponse && string.IsNullOrEmpty(command.ResponseFileName))
							{
								JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num].Nodes[num2].Nodes[num3];
								message = string.Format(Resources.NotEnteredMessage, Resources.ResponseFileName);
								return result;
							}
							if (command.IsUpload && string.IsNullOrEmpty(command.UploadFilePath))
							{
								JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num].Nodes[num2].Nodes[num3];
								message = string.Format(Resources.NotEnteredMessage, Resources.UploadFilePath);
								return result;
							}
							if (command.IsDownloadFolder)
							{
								if (string.IsNullOrEmpty(command.DownloadFolder))
								{
									JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num].Nodes[num2].Nodes[num3];
									message = string.Format(Resources.NotEnteredMessage, Resources.SaveFolderName);
									return result;
								}
								if (!Directory.Exists(command.DownloadFolder))
								{
									JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num].Nodes[num2].Nodes[num3];
									message = string.Format(Resources.NoExsistFolderMessage, Resources.SaveFolderName);
									return result;
								}
							}
							if (command.IsUpload && command.MoveKbn == "2" && string.IsNullOrEmpty(command.MovePath))
							{
								JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num].Nodes[num2].Nodes[num3];
								message = string.Format(Resources.NotEnteredMessage, Resources.MoveFolderName);
								return result;
							}
						}
						else if (string.IsNullOrEmpty(command.Path))
						{
							JobTreeView.SelectedNode = JobTreeView.Nodes[0].Nodes[num].Nodes[num2].Nodes[num3];
							message = string.Format(Resources.NotEnteredMessage, Resources.Path);
							return result;
						}
						num3++;
					}
					num2++;
				}
				num++;
			}
			return true;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void processSelect_Click(object sender, EventArgs e)
	{
		try
		{
			if (JobTreeView.SelectedNode == null || JobTreeView.SelectedNode.Tag == null || JobTreeView.SelectedNode.Tag.GetType() != typeof(JobData))
			{
				MessageBox.Show(Resources.UnSelectedJobNode, "WARNING");
				return;
			}
			using (JobController jobController = new JobController())
			{
				_conMasProcessList = jobController.GetConMasProcessDatas();
			}
			using ConMasProcessViewer conMasProcessViewer = new ConMasProcessViewer(this);
			if (conMasProcessViewer.ShowDialog() != DialogResult.OK || SelectedConMasProcesses == null || SelectedConMasProcesses.Count <= 0)
			{
				return;
			}
			string jobId = (JobTreeView.SelectedNode.Tag as JobData).Id;
			foreach (ProcessData selectedConMasProcess in SelectedConMasProcesses)
			{
				TreeNode selectedNode = JobTreeView.SelectedNode;
				TreeNode treeNode = new TreeNode();
				treeNode.Text = selectedConMasProcess.Name;
				treeNode.Tag = selectedConMasProcess;
				foreach (Command command in selectedConMasProcess.Commands)
				{
					TreeNode treeNode2 = new TreeNode();
					treeNode2.Text = command.Name;
					treeNode2.Tag = command;
					treeNode.Nodes.Add(treeNode2);
				}
				selectedNode.Nodes.Add(treeNode);
				_jobDatas.Where((JobData child) => child.Id == jobId).First().Processes.Add(selectedConMasProcess);
			}
			JobTreeView.SelectedNode.Expand();
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void refreshButton_Click(object sender, EventArgs e)
	{
		try
		{
			if (MessageBox.Show(Resources.JobSettingRefreshMessage, "", MessageBoxButtons.OKCancel) != DialogResult.Cancel)
			{
				_jobDatas.Clear();
				init();
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
			saveButton.Text = Resources.Enter;
			ToolTip = new ToolTip();
			ToolTip.InitialDelay = 500;
			ToolTip.ReshowDelay = 500;
			ToolTip.AutoPopDelay = 3000;
			ToolTip.SetToolTip(addButton, Resources.Add);
			ToolTip.SetToolTip(deleteButton, Resources.Delete);
			ToolTip.SetToolTip(refreshButton, Resources.Refresh);
			legendButton.Text = Resources.LegendViewer;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void JobTreeView_MouseUp(object sender, MouseEventArgs e)
	{
		try
		{
			if (e.Button != MouseButtons.Right)
			{
				return;
			}
			Point point = new Point(e.X, e.Y);
			TreeNode nodeAt = JobTreeView.GetNodeAt(point);
			if (nodeAt != null)
			{
				JobTreeView.SelectedNode = nodeAt;
				if (JobTreeView.SelectedNode.Tag == null)
				{
					addItem.Visible = true;
					addItem.Text = Resources.AddJob;
					deleteItem.Visible = false;
					contextMenu.Show(JobTreeView, point);
				}
				else if (JobTreeView.SelectedNode.Tag.GetType() == typeof(JobData))
				{
					addItem.Visible = true;
					addItem.Text = Resources.AddProcess;
					deleteItem.Visible = true;
					deleteItem.Text = Resources.DeleteJob;
					contextMenu.Show(JobTreeView, point);
				}
				else if (JobTreeView.SelectedNode.Tag.GetType() == typeof(ProcessData))
				{
					addItem.Visible = true;
					addItem.Text = Resources.AddCommand;
					deleteItem.Visible = true;
					deleteItem.Text = Resources.DeleteProcess;
					contextMenu.Show(JobTreeView, point);
				}
				else if (JobTreeView.SelectedNode.Tag.GetType() == typeof(Command))
				{
					addItem.Visible = false;
					deleteItem.Visible = true;
					deleteItem.Text = Resources.DeleteCommand;
					contextMenu.Show(JobTreeView, point);
				}
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void contextMenu_Click(object sender, EventArgs e)
	{
	}

	private void contextMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
	{
		try
		{
			if (e.ClickedItem.Name == "addItem")
			{
				AddNode();
			}
			else if (e.ClickedItem.Name == "deleteItem")
			{
				RemoveNode();
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void legendButton_Click(object sender, EventArgs e)
	{
		try
		{
			using ConMasLegendViewer conMasLegendViewer = new ConMasLegendViewer();
			conMasLegendViewer.ShowDialog();
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
		this.components = new System.ComponentModel.Container();
		System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConMasGeneratorUtility.Views.JobSettingView));
		this.splitContainer = new System.Windows.Forms.SplitContainer();
		this.JobTreeView = new System.Windows.Forms.TreeView();
		this.variablePanel = new System.Windows.Forms.Panel();
		this.saveButton = new System.Windows.Forms.Button();
		this.imageList1 = new System.Windows.Forms.ImageList(this.components);
		this.controllPanel = new System.Windows.Forms.Panel();
		this.legendButton = new System.Windows.Forms.Button();
		this.line = new System.Windows.Forms.Label();
		this.buttonPanel = new System.Windows.Forms.Panel();
		this.refreshButton = new System.Windows.Forms.Button();
		this.up = new System.Windows.Forms.Button();
		this.deleteButton = new System.Windows.Forms.Button();
		this.down = new System.Windows.Forms.Button();
		this.addButton = new System.Windows.Forms.Button();
		this.contextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
		this.addItem = new System.Windows.Forms.ToolStripMenuItem();
		this.deleteItem = new System.Windows.Forms.ToolStripMenuItem();
		((System.ComponentModel.ISupportInitialize)this.splitContainer).BeginInit();
		this.splitContainer.Panel1.SuspendLayout();
		this.splitContainer.Panel2.SuspendLayout();
		this.splitContainer.SuspendLayout();
		this.controllPanel.SuspendLayout();
		this.buttonPanel.SuspendLayout();
		this.contextMenu.SuspendLayout();
		base.SuspendLayout();
		this.splitContainer.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.splitContainer.Location = new System.Drawing.Point(0, 39);
		this.splitContainer.Margin = new System.Windows.Forms.Padding(5);
		this.splitContainer.Name = "splitContainer";
		this.splitContainer.Panel1.Controls.Add(this.JobTreeView);
		this.splitContainer.Panel2.Controls.Add(this.variablePanel);
		this.splitContainer.Size = new System.Drawing.Size(660, 404);
		this.splitContainer.SplitterDistance = 164;
		this.splitContainer.TabIndex = 0;
		this.JobTreeView.Dock = System.Windows.Forms.DockStyle.Fill;
		this.JobTreeView.HideSelection = false;
		this.JobTreeView.HotTracking = true;
		this.JobTreeView.Location = new System.Drawing.Point(0, 0);
		this.JobTreeView.Name = "JobTreeView";
		this.JobTreeView.Size = new System.Drawing.Size(164, 404);
		this.JobTreeView.TabIndex = 201;
		this.JobTreeView.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(JobTreeView_AfterSelect);
		this.JobTreeView.MouseUp += new System.Windows.Forms.MouseEventHandler(JobTreeView_MouseUp);
		this.variablePanel.BackColor = System.Drawing.Color.Transparent;
		this.variablePanel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
		this.variablePanel.Dock = System.Windows.Forms.DockStyle.Fill;
		this.variablePanel.Location = new System.Drawing.Point(0, 0);
		this.variablePanel.Name = "variablePanel";
		this.variablePanel.Size = new System.Drawing.Size(492, 404);
		this.variablePanel.TabIndex = 301;
		this.saveButton.Dock = System.Windows.Forms.DockStyle.Top;
		this.saveButton.Font = new System.Drawing.Font("MS UI Gothic", 11.25f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128);
		this.saveButton.ImageKey = "check.png";
		this.saveButton.ImageList = this.imageList1;
		this.saveButton.Location = new System.Drawing.Point(0, 0);
		this.saveButton.Name = "saveButton";
		this.saveButton.Size = new System.Drawing.Size(89, 33);
		this.saveButton.TabIndex = 106;
		this.saveButton.Text = "Save";
		this.saveButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
		this.saveButton.UseVisualStyleBackColor = true;
		this.saveButton.Click += new System.EventHandler(saveButton_Click);
		this.imageList1.ImageStream = (System.Windows.Forms.ImageListStreamer)resources.GetObject("imageList1.ImageStream");
		this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
		this.imageList1.Images.SetKeyName(0, "check.png");
		this.controllPanel.Controls.Add(this.legendButton);
		this.controllPanel.Controls.Add(this.line);
		this.controllPanel.Controls.Add(this.buttonPanel);
		this.controllPanel.Controls.Add(this.refreshButton);
		this.controllPanel.Controls.Add(this.up);
		this.controllPanel.Controls.Add(this.deleteButton);
		this.controllPanel.Controls.Add(this.down);
		this.controllPanel.Controls.Add(this.addButton);
		this.controllPanel.Dock = System.Windows.Forms.DockStyle.Top;
		this.controllPanel.Location = new System.Drawing.Point(0, 0);
		this.controllPanel.Name = "controllPanel";
		this.controllPanel.Size = new System.Drawing.Size(660, 42);
		this.controllPanel.TabIndex = 100;
		this.legendButton.Font = new System.Drawing.Font("MS UI Gothic", 11.25f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128);
		this.legendButton.Location = new System.Drawing.Point(238, 2);
		this.legendButton.Name = "legendButton";
		this.legendButton.Size = new System.Drawing.Size(87, 35);
		this.legendButton.TabIndex = 106;
		this.legendButton.Text = "...";
		this.legendButton.UseVisualStyleBackColor = true;
		this.legendButton.Visible = false;
		this.legendButton.Click += new System.EventHandler(legendButton_Click);
		this.line.BackColor = System.Drawing.SystemColors.Control;
		this.line.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
		this.line.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.line.Location = new System.Drawing.Point(0, 39);
		this.line.Margin = new System.Windows.Forms.Padding(0);
		this.line.Name = "line";
		this.line.Size = new System.Drawing.Size(571, 3);
		this.line.TabIndex = 28;
		this.buttonPanel.BackColor = System.Drawing.Color.Transparent;
		this.buttonPanel.Controls.Add(this.saveButton);
		this.buttonPanel.Dock = System.Windows.Forms.DockStyle.Right;
		this.buttonPanel.Location = new System.Drawing.Point(571, 0);
		this.buttonPanel.Name = "buttonPanel";
		this.buttonPanel.Size = new System.Drawing.Size(89, 42);
		this.buttonPanel.TabIndex = 30;
		this.refreshButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("refreshButton.BackgroundImage");
		this.refreshButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.refreshButton.Location = new System.Drawing.Point(1, 2);
		this.refreshButton.Name = "refreshButton";
		this.refreshButton.Size = new System.Drawing.Size(37, 35);
		this.refreshButton.TabIndex = 105;
		this.refreshButton.UseVisualStyleBackColor = true;
		this.refreshButton.Click += new System.EventHandler(refreshButton_Click);
		this.up.Font = new System.Drawing.Font("MS UI Gothic", 11.25f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
		this.up.Location = new System.Drawing.Point(141, 2);
		this.up.Name = "up";
		this.up.Size = new System.Drawing.Size(37, 35);
		this.up.TabIndex = 102;
		this.up.Text = "▲";
		this.up.UseVisualStyleBackColor = true;
		this.up.Click += new System.EventHandler(up_Click);
		this.deleteButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("deleteButton.BackgroundImage");
		this.deleteButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.deleteButton.Location = new System.Drawing.Point(85, 2);
		this.deleteButton.Name = "deleteButton";
		this.deleteButton.Size = new System.Drawing.Size(37, 35);
		this.deleteButton.TabIndex = 104;
		this.deleteButton.UseVisualStyleBackColor = true;
		this.deleteButton.Click += new System.EventHandler(deleteButton_Click);
		this.down.Font = new System.Drawing.Font("MS UI Gothic", 11.25f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
		this.down.Location = new System.Drawing.Point(179, 2);
		this.down.Name = "down";
		this.down.Size = new System.Drawing.Size(37, 35);
		this.down.TabIndex = 101;
		this.down.Text = "▼";
		this.down.UseVisualStyleBackColor = true;
		this.down.Click += new System.EventHandler(down_Click);
		this.addButton.BackgroundImage = (System.Drawing.Image)resources.GetObject("addButton.BackgroundImage");
		this.addButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.addButton.Location = new System.Drawing.Point(43, 2);
		this.addButton.Name = "addButton";
		this.addButton.Size = new System.Drawing.Size(37, 35);
		this.addButton.TabIndex = 103;
		this.addButton.UseVisualStyleBackColor = true;
		this.addButton.Click += new System.EventHandler(addButton_Click);
		this.contextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[2] { this.addItem, this.deleteItem });
		this.contextMenu.Name = "contextMenu";
		this.contextMenu.Size = new System.Drawing.Size(113, 48);
		this.contextMenu.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(contextMenu_ItemClicked);
		this.contextMenu.Click += new System.EventHandler(contextMenu_Click);
		this.addItem.Name = "addItem";
		this.addItem.Size = new System.Drawing.Size(152, 22);
		this.addItem.Text = "add";
		this.deleteItem.Name = "deleteItem";
		this.deleteItem.Size = new System.Drawing.Size(152, 22);
		this.deleteItem.Text = "delete";
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		base.Controls.Add(this.controllPanel);
		base.Controls.Add(this.splitContainer);
		base.Name = "JobSettingView";
		base.Size = new System.Drawing.Size(660, 443);
		base.Load += new System.EventHandler(JobSettingView_Load);
		base.Resize += new System.EventHandler(JobSettingView_Resize);
		this.splitContainer.Panel1.ResumeLayout(false);
		this.splitContainer.Panel2.ResumeLayout(false);
		((System.ComponentModel.ISupportInitialize)this.splitContainer).EndInit();
		this.splitContainer.ResumeLayout(false);
		this.controllPanel.ResumeLayout(false);
		this.buttonPanel.ResumeLayout(false);
		this.contextMenu.ResumeLayout(false);
		base.ResumeLayout(false);
	}
}
