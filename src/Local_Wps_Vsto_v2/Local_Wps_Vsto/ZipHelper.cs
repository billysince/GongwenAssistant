using System;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Text;
using ICSharpCode.SharpZipLib.Zip;

namespace Local_Wps_Vsto
{
	internal class ZipHelper
	{
		public static DataSet GetDatasetByString(string Value)
		{
			DataSet dataSet = new DataSet();
			StringReader reader = new StringReader(GZipDecompressString(Value));
			dataSet.ReadXml((TextReader)reader);
			return dataSet;
		}

		public static string GetStringByDataset(string ds)
		{
			return GZipCompressString(ds);
		}

		public static string GZipCompressString(string rawString)
		{
			if (string.IsNullOrEmpty(rawString) || rawString.Length == 0)
			{
				return "";
			}
			return Convert.ToBase64String(Compress(Encoding.UTF8.GetBytes(rawString.ToString())));
		}

		public static byte[] Compress(byte[] rawData)
		{
			MemoryStream memoryStream = new MemoryStream();
			GZipStream gZipStream = new GZipStream(memoryStream, CompressionMode.Compress, true);
			gZipStream.Write(rawData, 0, rawData.Length);
			gZipStream.Close();
			return memoryStream.ToArray();
		}

		public static string GZipDecompressString(string zippedString)
		{
			if (string.IsNullOrEmpty(zippedString) || zippedString.Length == 0)
			{
				return "";
			}
			byte[] zippedData = Convert.FromBase64String(zippedString.ToString());
			return Encoding.UTF8.GetString(Decompress(zippedData));
		}

		public static byte[] Decompress(byte[] zippedData)
		{
			GZipStream gZipStream = new GZipStream(new MemoryStream(zippedData), CompressionMode.Decompress);
			MemoryStream memoryStream = new MemoryStream();
			byte[] array = new byte[1024];
			while (true)
			{
				int num = gZipStream.Read(array, 0, array.Length);
				if (num <= 0)
				{
					break;
				}
				memoryStream.Write(array, 0, num);
			}
			gZipStream.Close();
			return memoryStream.ToArray();
		}

		public static bool PackFiles(string filename, string directory)
		{
			try
			{
				directory = directory.Replace("/", "\\");
				if (!directory.EndsWith("\\"))
				{
					directory += "\\";
				}
				if (!Directory.Exists(directory))
				{
					Directory.CreateDirectory(directory);
				}
				if (File.Exists(filename))
				{
					File.Delete(filename);
				}
				FastZip fastZip = new FastZip();
				fastZip.CreateEmptyDirectories = true;
				fastZip.CreateZip(filename, directory, true, "");
				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}

		public static void UnZipFile(string zipFilePath, string targetDir)
		{
			new FastZip(new FastZipEvents()).ExtractZip(zipFilePath, targetDir, "");
		}

		public static void DeleteDir(string strFolder)
		{
			try
			{
				new DirectoryInfo(strFolder).Attributes = (FileAttributes)0;
				File.SetAttributes(strFolder, FileAttributes.Normal);
				if (!Directory.Exists(strFolder))
				{
					return;
				}
				string[] fileSystemEntries = Directory.GetFileSystemEntries(strFolder);
				foreach (string text in fileSystemEntries)
				{
					if (File.Exists(text))
					{
						File.Delete(text);
						Console.WriteLine(text);
					}
					else
					{
						DeleteDir(text);
					}
				}
				Directory.Delete(strFolder);
			}
			catch (Exception)
			{
			}
		}

		public static void DeleteOneFile(string filename)
		{
			FileInfo fileInfo = new FileInfo(filename);
			if (fileInfo.Attributes.ToString().IndexOf("ReadOnly") != -1)
			{
				fileInfo.Attributes = FileAttributes.Normal;
			}
			File.Delete(filename);
		}
	}
}
