using System.ComponentModel;
using System.Windows.Forms;

namespace iReporterExcelAddInCommon;

public class ColorGrid : UserControl
{
	private IContainer components;

	public ColorGrid()
	{
		InitializeComponent();
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
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
	}
}
