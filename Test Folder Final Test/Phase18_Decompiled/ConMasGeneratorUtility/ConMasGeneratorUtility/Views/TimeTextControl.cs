using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using ConMasGeneratorUtility.Controls;
using ConMasGeneratorUtility.Properties;

namespace ConMasGeneratorUtility.Views;

public class TimeTextControl : UserControl
{
	private IContainer components = null;

	private NumericTextBox timeText;

	private NumericTextBox minitesText;

	private Label label1;

	public TimeTextControl()
	{
		InitializeComponent();
	}

	private void timeText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			if (string.IsNullOrEmpty(timeText.Text.Trim()))
			{
				timeText.Text = string.Empty;
				return;
			}
			int result = 0;
			if (int.TryParse(timeText.Text, out result))
			{
				if (int.Parse(timeText.Text) < 0 || int.Parse(timeText.Text) > 23)
				{
					MessageBox.Show(Resources.NoTimeMessage, "WARNING");
					e.Cancel = true;
				}
			}
			else
			{
				e.Cancel = true;
			}
		}
		catch (Exception ex)
		{
			Global.logger.Error((object)ex);
			MessageBox.Show(ex.ToString());
		}
	}

	private void minitesText_Validating(object sender, CancelEventArgs e)
	{
		try
		{
			if (string.IsNullOrEmpty(minitesText.Text.Trim()))
			{
				minitesText.Text = string.Empty;
				return;
			}
			int result = 0;
			if (int.TryParse(minitesText.Text, out result))
			{
				if (int.Parse(minitesText.Text) < 0 || int.Parse(minitesText.Text) > 23)
				{
					MessageBox.Show(Resources.NoMinitesMessage, "WARNING");
					e.Cancel = true;
				}
			}
			else
			{
				e.Cancel = true;
			}
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
		this.label1 = new System.Windows.Forms.Label();
		this.minitesText = new ConMasGeneratorUtility.Controls.NumericTextBox();
		this.timeText = new ConMasGeneratorUtility.Controls.NumericTextBox();
		base.SuspendLayout();
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(29, 6);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(11, 12);
		this.label1.TabIndex = 2;
		this.label1.Text = "：";
		this.minitesText.Location = new System.Drawing.Point(40, 3);
		this.minitesText.MaxLength = 2;
		this.minitesText.Name = "minitesText";
		this.minitesText.Size = new System.Drawing.Size(22, 19);
		this.minitesText.TabIndex = 1;
		this.minitesText.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
		this.minitesText.Validating += new System.ComponentModel.CancelEventHandler(minitesText_Validating);
		this.timeText.Location = new System.Drawing.Point(3, 3);
		this.timeText.MaxLength = 2;
		this.timeText.Name = "timeText";
		this.timeText.Size = new System.Drawing.Size(22, 19);
		this.timeText.TabIndex = 0;
		this.timeText.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
		this.timeText.Validating += new System.ComponentModel.CancelEventHandler(timeText_Validating);
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.Transparent;
		base.Controls.Add(this.timeText);
		base.Controls.Add(this.minitesText);
		base.Controls.Add(this.label1);
		base.Name = "TimeTextControl";
		base.Size = new System.Drawing.Size(70, 25);
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
