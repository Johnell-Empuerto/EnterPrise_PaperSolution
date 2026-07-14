using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using ConMasGeneratorLib;
using ConMasGeneratorLib.Util;

namespace ConMasGeneratorUtility;

public class TestForm : Form
{
	private GeneratorMain generator;

	private IContainer components = null;

	private Button EncodingPassword;

	private TextBox PasswordText;

	private Button OnStart;

	private Button OnEnd;

	private Button button1;

	public TestForm()
	{
		InitializeComponent();
	}

	private void EncodingPassword_Click(object sender, EventArgs e)
	{
		try
		{
			PasswordText.Text = Utility.SetPassword(PasswordText.Text);
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.ToString());
		}
	}

	private void OnStart_Click(object sender, EventArgs e)
	{
		try
		{
			if (generator != null)
			{
				generator.Dispose();
			}
			generator = new GeneratorMain();
			generator.ExecuteInit();
			OnStart.Enabled = false;
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.ToString());
		}
	}

	private void OnEnd_Click(object sender, EventArgs e)
	{
		try
		{
			if (generator != null)
			{
				generator.Dispose();
			}
			OnStart.Enabled = true;
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.ToString());
		}
	}

	private void button1_Click(object sender, EventArgs e)
	{
		try
		{
		}
		catch (Exception ex)
		{
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
		this.EncodingPassword = new System.Windows.Forms.Button();
		this.PasswordText = new System.Windows.Forms.TextBox();
		this.OnStart = new System.Windows.Forms.Button();
		this.OnEnd = new System.Windows.Forms.Button();
		this.button1 = new System.Windows.Forms.Button();
		base.SuspendLayout();
		this.EncodingPassword.Location = new System.Drawing.Point(12, 12);
		this.EncodingPassword.Name = "EncodingPassword";
		this.EncodingPassword.Size = new System.Drawing.Size(141, 33);
		this.EncodingPassword.TabIndex = 2;
		this.EncodingPassword.Text = "パスワード暗号化";
		this.EncodingPassword.UseVisualStyleBackColor = true;
		this.EncodingPassword.Click += new System.EventHandler(EncodingPassword_Click);
		this.PasswordText.Location = new System.Drawing.Point(174, 19);
		this.PasswordText.Name = "PasswordText";
		this.PasswordText.Size = new System.Drawing.Size(184, 19);
		this.PasswordText.TabIndex = 3;
		this.OnStart.Font = new System.Drawing.Font("メイリオ", 14.25f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128);
		this.OnStart.Location = new System.Drawing.Point(12, 81);
		this.OnStart.Name = "OnStart";
		this.OnStart.Size = new System.Drawing.Size(159, 75);
		this.OnStart.TabIndex = 4;
		this.OnStart.Text = "サービス開始";
		this.OnStart.UseVisualStyleBackColor = true;
		this.OnStart.Click += new System.EventHandler(OnStart_Click);
		this.OnEnd.Font = new System.Drawing.Font("メイリオ", 14.25f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128);
		this.OnEnd.Location = new System.Drawing.Point(203, 81);
		this.OnEnd.Name = "OnEnd";
		this.OnEnd.Size = new System.Drawing.Size(159, 75);
		this.OnEnd.TabIndex = 5;
		this.OnEnd.Text = "サービス終了";
		this.OnEnd.UseVisualStyleBackColor = true;
		this.OnEnd.Click += new System.EventHandler(OnEnd_Click);
		this.button1.Font = new System.Drawing.Font("メイリオ", 14.25f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128);
		this.button1.Location = new System.Drawing.Point(454, 81);
		this.button1.Name = "button1";
		this.button1.Size = new System.Drawing.Size(159, 75);
		this.button1.TabIndex = 6;
		this.button1.Text = "test";
		this.button1.UseVisualStyleBackColor = true;
		this.button1.Click += new System.EventHandler(button1_Click);
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 12f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		base.ClientSize = new System.Drawing.Size(719, 175);
		base.Controls.Add(this.button1);
		base.Controls.Add(this.OnEnd);
		base.Controls.Add(this.OnStart);
		base.Controls.Add(this.PasswordText);
		base.Controls.Add(this.EncodingPassword);
		base.Name = "TestForm";
		this.Text = "TestForm";
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
