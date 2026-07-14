using System;
using System.Diagnostics;
using System.Net;
using System.Windows.Forms;

namespace ConMasGeneratorUtility;

internal static class Program
{
	[STAThread]
	private static void Main()
	{
		if (!PrevInstance())
		{
			ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(defaultValue: false);
			Application.Run(new SettingForm());
		}
	}

	private static bool PrevInstance()
	{
		string processName = Process.GetCurrentProcess().ProcessName;
		if (Process.GetProcessesByName(processName).Length > 1)
		{
			return true;
		}
		return false;
	}
}
