using System;
using System.ComponentModel;
using System.Threading;
using System.Windows.Forms;

namespace iReporterExcelAddIn2019;

public class WaitDialog : Form
{
	public Label labelMainMsg;

	private Button button1;

	private Container components;

	private CancellationTokenSource Src { get; }

	public WaitDialog(CancellationTokenSource src)
	{
		InitializeComponent();
		Src = src;
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
		System.ComponentModel.ComponentResourceManager componentResourceManager = new System.ComponentModel.ComponentResourceManager(typeof(iReporterExcelAddIn2019.WaitDialog));
		this.labelMainMsg = new System.Windows.Forms.Label();
		this.button1 = new System.Windows.Forms.Button();
		base.SuspendLayout();
		componentResourceManager.ApplyResources(this.labelMainMsg, "labelMainMsg");
		this.labelMainMsg.Name = "labelMainMsg";
		componentResourceManager.ApplyResources(this.button1, "button1");
		this.button1.Name = "button1";
		this.button1.UseVisualStyleBackColor = true;
		this.button1.Click += new System.EventHandler(button1_Click);
		componentResourceManager.ApplyResources(this, "$this");
		base.ControlBox = false;
		base.Controls.Add(this.button1);
		base.Controls.Add(this.labelMainMsg);
		base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		base.Name = "WaitDialog";
		base.ShowInTaskbar = false;
		base.ResumeLayout(false);
	}

	private void button1_Click(object sender, EventArgs e)
	{
		Src.Cancel();
		Close();
	}
}
