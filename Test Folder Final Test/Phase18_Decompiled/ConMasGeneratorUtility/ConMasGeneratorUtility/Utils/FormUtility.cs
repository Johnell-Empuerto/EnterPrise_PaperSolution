using System;
using System.ServiceProcess;

namespace ConMasGeneratorUtility.Utils;

internal class FormUtility
{
	public const string ConMasGeneratorServiceName = "ConMasGeneratorService";

	public static void StartService()
	{
		try
		{
			using ServiceController serviceController = new ServiceController("ConMasGeneratorService");
			if (serviceController.Status != ServiceControllerStatus.Stopped && serviceController.Status != ServiceControllerStatus.StopPending)
			{
				serviceController.Stop();
				serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
			}
			serviceController.Start();
			serviceController.WaitForStatus(ServiceControllerStatus.Running);
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	public static void StoptService()
	{
		try
		{
			using ServiceController serviceController = new ServiceController("ConMasGeneratorService");
			if (serviceController.Status != ServiceControllerStatus.Stopped && serviceController.Status != ServiceControllerStatus.StopPending)
			{
				serviceController.Stop();
				serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	public static ServiceControllerStatus GetServiceStatus(ServiceController sc, string serviceName)
	{
		try
		{
			return sc.Status;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}
}
