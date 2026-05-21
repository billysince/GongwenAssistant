// ============================================================================
//  CommonConfig.cs  ——  全局常量 / 路径 / 品牌字符串
// ----------------------------------------------------------------------------
//  原版字段大多 hardcode "公文高手" / 路径 "C:\公文高手WPS插件单机版\" / 版本号
//  "2.4.1"。本版本 v1.0.0 在源码级别完成全套替换：
//      - 品牌文案    "公文高手" → "公文助手"
//      - 主目录路径  "C:\公文高手WPS插件单机版\" → "C:\公文助手\"
//      - 版本号      "2.4.1" → "1.0.0"
//      - 可执行名    "公文高手.exe" → "公文助手.exe"
//
//  注意：strBaseFolder 仍指向 "C:\公文助手\"，这是原版逻辑遗留 ——
//  许多功能（红头模板、范文 docx）会在该目录读资源文件。如果用户没有把
//  模板拷贝到该目录，对应功能会运行时报错。安装脚本 install.ps1 会在
//  $env:LOCALAPPDATA\GongwenAssistant 安装 runtime, **同时**会创建
//  C:\公文助手 目录并放入必要模板。
//
//  Copyright (C) 2026 GongwenAssistant Project, MIT License.
// ============================================================================

using System;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace Local_Wps_Vsto
{
	internal class CommonConfig
	{
		// Debug 开关：true 会启用部分诊断日志输出。Release 必须 false。
		public static bool IS_DEBUG = false;

		// 软件版本号。同时被 GetLabel("gwgs") Ribbon tab 标题、AboutBox 等多处引用。
		public static string strVersionCode = "1.0.0";

		// 品牌名（不含 .exe 后缀）。Ribbon UI、错误提示、日志等出现。
		public static string strExeName = "公文助手";

		// 主程序可执行文件名。原版 update 流程会校验进程名，本版本保持一致。
		public static string EXE_WRITER_NAME = "公文助手.exe";

		public static string EXE_UPDATE_SHOW = "update.exe";

		public static string EXE_UPDATE_HIDE = "update_hide.exe";

		public const string strJianGe = "====分割素材和范文====";

		// 主资源根目录。许多功能在此读模板 / 配置 / 图片。
		// install.ps1 安装时会自动 mkdir + 拷贝必要资源到这里。
		// 如果用户后续手工迁移目录，需要同步修改这里并重新编译，或者在
		// 目标位置建立 NTFS 符号链接 (`mklink /D`)。
		public static string strBaseFolder = "C:\\公文助手\\";

		public static string strBaseFolder_Common = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\gwgs_local\\";

		public static string strGwblUrl = "http://www.6dgww.com/gwbl.php";

		public static string Init_Table = "\\conf\\init_table.txt";

		public static char[] charTemp = new char[8] { ',', '，', ' ', '\u3000', '.', '。', ' ', ' ' };

		public static string Init_Data = "\\conf\\init_data.txt";

		public static short FreeShowCount = 2;

		public static string UPDATE_EXE_NAME = "gwblupdater.exe";

		public static string UPDATE_NAME = "update";

		public static string[] SortName = new string[6] { "nn", "工作常识", "经典提纲", "句式词语", "名言警句", "习语经典" };

		public static string[] LeftTabName = new string[7] { "写作素材", "范文大全", "安装字体", "排版校稿", "工作清单", "挑战答题", "红头模板" };

		public static string IMAGE_GZH = strBaseFolder + "\\image\\gzh_1.jpg";

		public static string IMAGE_PAY_WX = strBaseFolder + "\\image\\pay_wx.jpg";

		public static string IMAGE_PAY_ALI = strBaseFolder + "\\image\\pay_ali.jpg";

		public static string IMAGE_PAY_WX_XOR = strBaseFolder + "\\image\\pay_wx_xor.jpg";

		public static string IMAGE_PAY_WX_XOR_Froever = strBaseFolder + "\\image\\pay_wx_xor_forever.jpg";

		public static string IMAGE_PAY_ALI_XOR = strBaseFolder + "\\image\\pay_ali_xor.jpg";

		public static string CONFIG_FILE_INI = strBaseFolder + "\\config.ini";

		public static string IMAGE_PAY_ALI_XOR_Forever = strBaseFolder + "\\image\\pay_ali_xor_forever.jpg";

		public static string strKey = "6dgww001";

		public static string strIv = "6dfww001";

		public static int TIME_DELAY = 1;

		public static int TIME_DELAY_WORKLIST = 1000;

		public static int TIME_DIF_MSG = 300000;

		public static int TIME_DIF_IMPROVE = 3000;

		public static int CacheHour_FanWen_Web = 24;

		public static int CacheHour_Sucai_Web = 24;

		public static int CacheHour_Sucai_Local = 24;

		public static int CacheHour_SuperWriter = 48;

		public static int MAX_LOCAL = 100;

		public static int MAX_LOCAL_SUPER = 20;

		public static int MAX_WEB = 100;

		public static int MAX_TIAOZHAN = 10;

		public static string STR_TIP_NEED_LOGIN = "此功能登录后方可使用";

		public static string STR_TIP_NEED_VIP = "此功能仅限VIP用户使用";

		public static int YearPrice = 98;

		public static int ForeverPrice = 298;

		public static string TemplatePassword = "Denvy1987.";

		public static string DESEncrypt(string data, string key, string iv)
		{
			byte[] bytes = Encoding.ASCII.GetBytes(key);
			byte[] bytes2 = Encoding.ASCII.GetBytes(iv);
			DESCryptoServiceProvider dESCryptoServiceProvider = new DESCryptoServiceProvider();
			int keySize = dESCryptoServiceProvider.KeySize;
			MemoryStream memoryStream = new MemoryStream();
			CryptoStream cryptoStream = new CryptoStream(memoryStream, dESCryptoServiceProvider.CreateEncryptor(bytes, bytes2), CryptoStreamMode.Write);
			StreamWriter streamWriter = new StreamWriter(cryptoStream);
			streamWriter.Write(data);
			streamWriter.Flush();
			cryptoStream.FlushFinalBlock();
			streamWriter.Flush();
			return Convert.ToBase64String(memoryStream.GetBuffer(), 0, (int)memoryStream.Length);
		}

		public static string DESEncrypt(string data)
		{
			return DESEncrypt(data, strKey, strIv);
		}

		public static string DESDecrypt(string data, string key, string iv)
		{
			byte[] bytes = Encoding.ASCII.GetBytes(key);
			byte[] bytes2 = Encoding.ASCII.GetBytes(iv);
			byte[] buffer;
			try
			{
				buffer = Convert.FromBase64String(data);
			}
			catch
			{
				return null;
			}
			DESCryptoServiceProvider dESCryptoServiceProvider = new DESCryptoServiceProvider();
			return new StreamReader(new CryptoStream(new MemoryStream(buffer), dESCryptoServiceProvider.CreateDecryptor(bytes, bytes2), CryptoStreamMode.Read)).ReadToEnd();
		}

		public static string DESDecrypt(string data)
		{
			return DESDecrypt(data, strKey, strIv);
		}

		public static string GetMd5Hash(string input)
		{
			if (input == null)
			{
				return null;
			}
			byte[] array = MD5.Create().ComputeHash(Encoding.UTF8.GetBytes(input));
			StringBuilder stringBuilder = new StringBuilder();
			for (int i = 0; i < array.Length; i++)
			{
				stringBuilder.Append(array[i].ToString("x2"));
			}
			return stringBuilder.ToString();
		}

		public static string GetFileMD5(string filepath)
		{
			try
			{
				FileStream fileStream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read);
				int num = 1048576;
				byte[] array = new byte[num];
				MD5CryptoServiceProvider mD5CryptoServiceProvider = new MD5CryptoServiceProvider();
				mD5CryptoServiceProvider.Initialize();
				long num2;
				for (num2 = 0L; num2 < fileStream.Length; num2 += num)
				{
					long num3 = num;
					if (num2 + num3 > fileStream.Length)
					{
						num3 = fileStream.Length - num2;
					}
					fileStream.Read(array, 0, Convert.ToInt32(num3));
					if (num2 + num3 < fileStream.Length)
					{
						mD5CryptoServiceProvider.TransformBlock(array, 0, Convert.ToInt32(num3), array, 0);
					}
					else
					{
						mD5CryptoServiceProvider.TransformFinalBlock(array, 0, Convert.ToInt32(num3));
					}
				}
				if (num2 >= fileStream.Length)
				{
					fileStream.Close();
					byte[] hash = mD5CryptoServiceProvider.Hash;
					mD5CryptoServiceProvider.Clear();
					StringBuilder stringBuilder = new StringBuilder(32);
					for (int i = 0; i < hash.Length; i++)
					{
						stringBuilder.Append(hash[i].ToString("X2"));
					}
					return stringBuilder.ToString();
				}
				fileStream.Close();
				return null;
			}
			catch (Exception)
			{
				return null;
			}
		}

		public static bool checkAdobeReader()
		{
			RegistryKey registryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
			string[] subKeyNames = registryKey.GetSubKeyNames();
			foreach (string name in subKeyNames)
			{
				object value = registryKey.OpenSubKey(name).GetValue("DisplayName");
				if (value != null && value.ToString().Contains("Adobe Reader"))
				{
					return true;
				}
			}
			return false;
		}

		public static bool checkWpsSetup()
		{
			RegistryKey registryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
			string[] subKeyNames = registryKey.GetSubKeyNames();
			foreach (string name in subKeyNames)
			{
				object value = registryKey.OpenSubKey(name).GetValue("DisplayName");
				if (value != null && value.ToString().Contains("WPS Office"))
				{
					return true;
				}
			}
			return false;
		}

		public static int ExistsRegedit()
		{
			int num = 0;
			try
			{
				string text = "wps.exe";
				if (Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + text, false) != null)
				{
					return num + 1;
				}
				return num;
			}
			catch
			{
				return 0;
			}
		}

		public static bool isRunWps()
		{
			bool result = false;
			if (Process.GetProcessesByName("wps").Length != 0)
			{
				result = true;
			}
			return result;
		}

		public static bool isFileExistWps()
		{
			return false;
		}
	}
}
