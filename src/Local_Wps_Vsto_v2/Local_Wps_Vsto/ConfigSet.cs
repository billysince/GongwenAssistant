using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Local_Wps_Vsto
{
	public class ConfigSet
	{
		private string configFilePath;

		private List<string> configName = new List<string>();

		private List<string> configValue = new List<string>();

		public ConfigSet(string configFilePath, bool isRead = true)
		{
			this.configFilePath = configFilePath;
			if (File.Exists(this.configFilePath) && isRead)
			{
				ReadConfig();
			}
		}

		public void SetConfigValue(string cName, string cValue)
		{
			try
			{
				bool flag = false;
				if (configName.Count != 0)
				{
					for (int i = 0; i < configName.Count; i++)
					{
						if (configName[i].Equals(cName.ToLower()))
						{
							configValue[i] = cValue;
							flag = true;
						}
					}
				}
				if (!flag)
				{
					configName.Add(cName);
					configValue.Add(cValue);
				}
			}
			catch (Exception)
			{
			}
		}

		public bool WriteConfigToFile(ConfigFile cf = ConfigFile.newFile)
		{
			try
			{
				StreamWriter streamWriter;
				switch (cf)
				{
				case ConfigFile.newFile:
					streamWriter = new StreamWriter(configFilePath, false);
					break;
				case ConfigFile.appendFile:
					streamWriter = new StreamWriter(configFilePath, true);
					break;
				default:
					streamWriter = new StreamWriter(configFilePath);
					break;
				}
				try
				{
					for (int i = 0; i < configName.Count; i++)
					{
						streamWriter.WriteLine("{0}={1}", configName[i].ToLower(), configValue[i]);
					}
				}
				catch
				{
					return false;
				}
				finally
				{
					streamWriter.Close();
				}
				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}

		private bool ReadConfig()
		{
			try
			{
				StreamReader streamReader = new StreamReader(configFilePath, Encoding.Default);
				string text;
				while ((text = streamReader.ReadLine()) != null)
				{
					text = text.Trim();
					string[] array = text.Split('=');
					if (array.Length == 2)
					{
						string item = array[0].ToLower();
						string item2 = array[1].ToLower();
						configName.Add(item);
						configValue.Add(item2);
					}
				}
				streamReader.Close();
				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}
	}
}
