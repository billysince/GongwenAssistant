using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Local_Wps_Vsto
{
	internal class ConfigLoad
	{
		private string conFilePath;

		public List<string> configName = new List<string>();

		public List<string> configValue = new List<string>();

		public string ConFilePath
		{
			get
			{
				return conFilePath;
			}
			set
			{
				conFilePath = value;
			}
		}

		public ConfigLoad(string conFilePath)
		{
			this.conFilePath = conFilePath;
			ReadConfig();
		}

		public bool ReadConfig()
		{
			try
			{
				if (!File.Exists(conFilePath))
				{
					return false;
				}
				StreamReader streamReader = new StreamReader(conFilePath, Encoding.Default);
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

		public string GetStringValue(string cName)
		{
			for (int i = 0; i < configName.Count; i++)
			{
				if (configName[i].Equals(cName.ToLower()))
				{
					return configValue[i];
				}
			}
			return "";
		}

		public int GetIntValue(string cName)
		{
			for (int i = 0; i < configName.Count; i++)
			{
				int result;
				if (configName[i].Equals(cName.ToLower()) && int.TryParse(configValue[i], out result))
				{
					return result;
				}
			}
			return 0;
		}

		public float GetFloatValue(string cName)
		{
			for (int i = 0; i < configName.Count; i++)
			{
				float result;
				if (configName[i].Equals(cName.ToLower()) && float.TryParse(configValue[i], out result))
				{
					return result;
				}
			}
			return 0f;
		}
	}
}
