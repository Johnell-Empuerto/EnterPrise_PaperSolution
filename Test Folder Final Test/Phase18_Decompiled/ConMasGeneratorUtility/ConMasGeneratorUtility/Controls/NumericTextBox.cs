using System.Security.Permissions;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ConMasGeneratorUtility.Controls;

internal class NumericTextBox : TextBox
{
	private const int WM_PASTE = 770;

	private const int ES_NUMBER = 8192;

	protected override CreateParams CreateParams
	{
		[SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
		get
		{
			CreateParams createParams = base.CreateParams;
			createParams.Style |= 8192;
			return createParams;
		}
	}

	[SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
	protected override void WndProc(ref Message m)
	{
		if (m.Msg == 770)
		{
			IDataObject dataObject = Clipboard.GetDataObject();
			if (dataObject != null && dataObject.GetDataPresent(DataFormats.Text))
			{
				string input = (string)dataObject.GetData(DataFormats.Text);
				if (!Regex.IsMatch(input, "^[0-9]+$"))
				{
					return;
				}
			}
		}
		base.WndProc(ref m);
	}
}
