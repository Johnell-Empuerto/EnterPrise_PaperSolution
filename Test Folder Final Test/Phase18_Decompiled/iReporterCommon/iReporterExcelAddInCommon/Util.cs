using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace iReporterExcelAddInCommon;

public static class Util
{
	internal static void DrawString(Graphics graphics, string text, Font font, Color color, RectangleF rect)
	{
		using SolidBrush brush = new SolidBrush(color);
		graphics.DrawString(text, font, brush, rect);
	}

	public static TextBox CreateCopy(this Control c)
	{
		TextBox textBox2;
		if (c is TextBox textBox)
		{
			textBox2 = new TextBox
			{
				ReadOnly = textBox.ReadOnly,
				TabStop = !textBox.ReadOnly
			};
		}
		else
		{
			if (!(c is ComboBox))
			{
				return null;
			}
			textBox2 = new TextBox
			{
				ReadOnly = true,
				TabStop = false
			};
		}
		textBox2.Location = c.Location;
		textBox2.Size = c.Size;
		textBox2.TabIndex = c.TabIndex;
		c.Parent.Controls.Add(textBox2);
		return textBox2;
	}

	public static Label CreateCopy(this Label src)
	{
		Label label = new Label
		{
			Text = src.Text,
			Location = src.Location,
			Size = src.Size,
			Font = src.Font,
			TabIndex = src.TabIndex + 1
		};
		src.Parent.Controls.Add(label);
		return label;
	}

	public static void TextResize(this Label label, string text)
	{
		label.Text = text;
		label.Size = TextRenderer.MeasureText(text, label.Font);
	}

	public static int? TryParseInt(this string text)
	{
		if (!int.TryParse(text, out var result))
		{
			return null;
		}
		return result;
	}

	public static int MaxOr(this IEnumerable<int?> src, int min)
	{
		foreach (int? item in src)
		{
			if (item.HasValue)
			{
				min = Math.Max(item.Value, min);
			}
		}
		return min;
	}

	public static void Catched(Exception ex)
	{
		MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
	}

	public static string ToText(this Color c)
	{
		return c.A.ToString("X2") + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
	}

	public static Color ToColor(this string text)
	{
		Color result = default(Color);
		if (text == null || text.Length != 8)
		{
			return result;
		}
		if (!int.TryParse(text.Substring(0, 2), NumberStyles.HexNumber, null, out var result2))
		{
			return result;
		}
		if (!int.TryParse(text.Substring(2, 2), NumberStyles.HexNumber, null, out var result3))
		{
			return result;
		}
		if (!int.TryParse(text.Substring(4, 2), NumberStyles.HexNumber, null, out var result4))
		{
			return result;
		}
		if (!int.TryParse(text.Substring(6, 2), NumberStyles.HexNumber, null, out var result5))
		{
			return result;
		}
		return Color.FromArgb(result2, result3, result4, result5);
	}
}
