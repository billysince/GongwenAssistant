using System;
using System.Windows.Forms;

namespace Local_Wps_Vsto
{
	internal class MessageTipsUtil
	{
		public static void Show(string strMessage, int intShowSecond = 1000, bool boolSpeak = false)
		{
			try
			{
				MesBoxShow mesBoxShow = new MesBoxShow(strMessage, intShowSecond, boolSpeak);
				mesBoxShow.StartPosition = FormStartPosition.CenterScreen;
				mesBoxShow.TopMost = true;
				mesBoxShow.Show();
			}
			catch (Exception)
			{
			}
		}
	}
}
