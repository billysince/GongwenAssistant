using System;
using System.Net;

namespace Local_Wps_Vsto
{
	internal class HttpUtil
	{
		public static string getUrlResponse(string strUrl)
		{
			string result = "";
			try
			{
				result = new WebClient
				{
					Credentials = CredentialCache.DefaultCredentials
				}.DownloadString(strUrl);
				return result;
			}
			catch (Exception)
			{
				return result;
			}
		}
	}
}
