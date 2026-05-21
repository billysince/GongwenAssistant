using System;

namespace Local_Wps_Vsto
{
	internal class LocalSuCaiUtil
	{
		public static bool InsertNewSuCai(string strContent, long longUserId, string strSort)
		{
			try
			{
				SQLiteHelper sQLiteHelper = new SQLiteHelper();
				sQLiteHelper.InsertValues_WithC("wr_sucai_local", "content,creater_id,sort_name,create_time", new string[4]
				{
					strContent.ToString(),
					longUserId.ToString(),
					strSort,
					DateTime.Now.ToLocalTime().ToString("s")
				});
				sQLiteHelper.CloseConnection();
				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}
	}
}
