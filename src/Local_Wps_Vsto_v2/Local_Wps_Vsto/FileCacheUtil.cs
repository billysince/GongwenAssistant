using System;
using System.IO;

namespace Local_Wps_Vsto
{
	internal class FileCacheUtil
	{
		public static void SetCache(string strKey, string strResult)
		{
			try
			{
				File.WriteAllText(CommonConfig.strBaseFolder + "\\cache\\" + CommonConfig.GetMd5Hash(strKey) + ".tmp", strResult);
				MemCacheUtil.setCache(strKey, strResult);
			}
			catch (Exception)
			{
			}
		}

		public static string GetCache(string strKey, int cacheHour = 24)
		{
			string text = "";
			text = MemCacheUtil.getCache(strKey);
			if (text.Length <= 0)
			{
				try
				{
					long num = cacheHour * 3600;
					string text2 = CommonConfig.strBaseFolder + "\\cache\\" + CommonConfig.GetMd5Hash(strKey) + ".tmp";
					if (!File.Exists(text2))
					{
						return text;
					}
					if (GetNowTimeSpanSec(new FileInfo(text2).CreationTime) < num)
					{
						return File.ReadAllText(text2);
					}
					return text;
				}
				catch (Exception)
				{
					return "出现错误，请以管理员模式运行程序\r\n1、在程序图标点击右键，点属性\r\n2、点兼容性\r\n3、勾选“以管理员身份运行此程序\r\n4、点确定\r\n重新打开就可以了";
				}
			}
			return text;
		}

		public static int GetNowTimeSpanSec(DateTime _time)
		{
			return (int)DateTime.Now.Subtract(_time).TotalSeconds;
		}
	}
}
