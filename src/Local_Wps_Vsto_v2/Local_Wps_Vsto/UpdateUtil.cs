using System;
using System.IO;
using System.Net;
using System.Text;

namespace Local_Wps_Vsto
{
	internal class UpdateUtil
	{
		public static bool HttpDownloadFile(string url, string path)
		{
			bool flag = false;
			try
			{
				Stream responseStream = ((WebRequest.Create(url) as HttpWebRequest).GetResponse() as HttpWebResponse).GetResponseStream();
				Stream stream = new FileStream(path, FileMode.Create);
				byte[] array = new byte[1024];
				for (int num = responseStream.Read(array, 0, array.Length); num > 0; num = responseStream.Read(array, 0, array.Length))
				{
					stream.Write(array, 0, num);
				}
				stream.Close();
				responseStream.Close();
				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}

		public static string getHttpResponse(string strUrl)
		{
			try
			{
				byte[] bytes = new WebClient
				{
					Credentials = CredentialCache.DefaultCredentials
				}.DownloadData(strUrl);
				return Encoding.UTF8.GetString(bytes);
			}
			catch (Exception)
			{
				return "";
			}
		}

		public static bool FileIsReady(string strPath)
		{
			bool flag = false;
			try
			{
				File.Move(strPath, strPath);
				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}
	}
}
