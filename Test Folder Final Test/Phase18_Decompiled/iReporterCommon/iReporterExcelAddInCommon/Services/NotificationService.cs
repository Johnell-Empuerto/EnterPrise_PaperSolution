using System.Windows;
using iReporterExcelAddInCommon.Domain.ValueDefinitions;

namespace iReporterExcelAddInCommon.Services;

public class NotificationService : INotificationService
{
	private string Title { get; set; } = CommonResource.MessageGridCopied;

	public void ShowMessage(string message)
	{
		MessageBox.Show(message);
	}
}
