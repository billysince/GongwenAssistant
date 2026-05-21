using System;
using System.Management;

namespace Local_Wps_Vsto
{
	internal class MachineCode
	{
		private static MachineCode machineCode = null;

		private static string machineCodeString = string.Empty;

		public static string GetMachineCodeString()
		{
			return FingerPrint.Value();
		}

		public string GetCpuInfo()
		{
			string text = "";
			try
			{
				using (ManagementClass managementClass = new ManagementClass("Win32_Processor"))
				{
					foreach (ManagementObject instance in managementClass.GetInstances())
					{
						text = instance.Properties["ProcessorId"].Value.ToString();
						instance.Dispose();
					}
				}
			}
			catch (Exception)
			{
				throw;
			}
			return text.ToString();
		}

		public string GetHDid()
		{
			string text = "";
			try
			{
				using (ManagementClass managementClass = new ManagementClass("Win32_DiskDrive"))
				{
					foreach (ManagementObject instance in managementClass.GetInstances())
					{
						text = (string)instance.Properties["Model"].Value;
						instance.Dispose();
					}
				}
			}
			catch (Exception)
			{
				throw;
			}
			return text.ToString();
		}

		public string GetMoAddress()
		{
			string text = "";
			try
			{
				using (ManagementClass managementClass = new ManagementClass("Win32_NetworkAdapterConfiguration"))
				{
					foreach (ManagementObject instance in managementClass.GetInstances())
					{
						if ((bool)instance["IPEnabled"])
						{
							text = instance["MacAddress"].ToString();
						}
						instance.Dispose();
					}
				}
			}
			catch (Exception)
			{
				throw;
			}
			return text.ToString();
		}
	}
}
