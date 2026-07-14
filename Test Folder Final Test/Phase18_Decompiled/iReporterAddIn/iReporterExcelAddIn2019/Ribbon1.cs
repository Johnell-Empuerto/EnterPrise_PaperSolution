using System.ComponentModel;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

namespace iReporterExcelAddIn2019;

public class Ribbon1 : RibbonBase
{
	private IContainer components;

	internal RibbonTab tab1;

	internal RibbonGroup group1;

	internal RibbonButton autoJudgeButton;

	internal RibbonButton clusterSettingButton;

	internal RibbonButton colorSettingButton;

	internal RibbonButton tableSettingButton;

	internal RibbonButton checkSheetNameButton;

	private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
	{
	}

	private void autoJudgeButton_Click(object sender, RibbonControlEventArgs e)
	{
		Globals.ThisAddIn.AutoJudge(openDlg: true);
	}

	private void clusterSettingButton_Click(object sender, RibbonControlEventArgs e)
	{
		Globals.ThisAddIn.ManualCluster();
	}

	private void colorSettingButton_Click(object sender, RibbonControlEventArgs e)
	{
		Globals.ThisAddIn.SetClusterColors();
	}

	private void tableSettingButton_Click(object sender, RibbonControlEventArgs e)
	{
		Globals.ThisAddIn.ShowTableForm();
	}

	private void checkSheetNameButton_Click(object sender, RibbonControlEventArgs e)
	{
		Globals.ThisAddIn.CheckAllSheetName();
	}

	public Ribbon1()
		: base(Globals.Factory.GetRibbonFactory())
	{
		InitializeComponent();
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		((RibbonBase)this).Dispose(disposing);
	}

	private void InitializeComponent()
	{
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(Ribbon1));
		tab1 = ((RibbonBase)this).Factory.CreateRibbonTab();
		group1 = ((RibbonBase)this).Factory.CreateRibbonGroup();
		autoJudgeButton = ((RibbonBase)this).Factory.CreateRibbonButton();
		clusterSettingButton = ((RibbonBase)this).Factory.CreateRibbonButton();
		colorSettingButton = ((RibbonBase)this).Factory.CreateRibbonButton();
		tableSettingButton = ((RibbonBase)this).Factory.CreateRibbonButton();
		checkSheetNameButton = ((RibbonBase)this).Factory.CreateRibbonButton();
		tab1.SuspendLayout();
		group1.SuspendLayout();
		((RibbonBase)this).SuspendLayout();
		tab1.ControlId.ControlIdType = RibbonControlIdType.Office;
		tab1.Groups.Add(group1);
		componentResourceManager.ApplyResources(tab1, "tab1");
		tab1.Name = "tab1";
		group1.Items.Add(autoJudgeButton);
		group1.Items.Add(clusterSettingButton);
		group1.Items.Add(colorSettingButton);
		group1.Items.Add(tableSettingButton);
		group1.Items.Add(checkSheetNameButton);
		componentResourceManager.ApplyResources(group1, "group1");
		group1.Name = "group1";
		autoJudgeButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		componentResourceManager.ApplyResources(autoJudgeButton, "autoJudgeButton");
		autoJudgeButton.Name = "autoJudgeButton";
		autoJudgeButton.ShowImage = true;
		autoJudgeButton.Click += autoJudgeButton_Click;
		clusterSettingButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		componentResourceManager.ApplyResources(clusterSettingButton, "clusterSettingButton");
		clusterSettingButton.Name = "clusterSettingButton";
		clusterSettingButton.ShowImage = true;
		clusterSettingButton.Click += clusterSettingButton_Click;
		colorSettingButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		componentResourceManager.ApplyResources(colorSettingButton, "colorSettingButton");
		colorSettingButton.Name = "colorSettingButton";
		colorSettingButton.ShowImage = true;
		colorSettingButton.Click += colorSettingButton_Click;
		tableSettingButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		componentResourceManager.ApplyResources(tableSettingButton, "tableSettingButton");
		tableSettingButton.Name = "tableSettingButton";
		tableSettingButton.ShowImage = true;
		tableSettingButton.Click += tableSettingButton_Click;
		checkSheetNameButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		componentResourceManager.ApplyResources(checkSheetNameButton, "checkSheetNameButton");
		checkSheetNameButton.Name = "checkSheetNameButton";
		checkSheetNameButton.ShowImage = true;
		checkSheetNameButton.Click += checkSheetNameButton_Click;
		((RibbonBase)this).Name = "Ribbon1";
		((RibbonBase)this).RibbonType = "Microsoft.Excel.Workbook";
		((RibbonBase)this).Tabs.Add(tab1);
		((RibbonBase)this).Load += Ribbon1_Load;
		tab1.ResumeLayout(performLayout: false);
		tab1.PerformLayout();
		group1.ResumeLayout(performLayout: false);
		group1.PerformLayout();
		((RibbonBase)this).ResumeLayout(false);
	}
}
