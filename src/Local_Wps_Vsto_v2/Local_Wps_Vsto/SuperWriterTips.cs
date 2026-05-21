using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace Local_Wps_Vsto
{
	public class SuperWriterTips : UserControl
	{
		private MyAddin _parentRibbon;

		public string strSearchWord = "";

		public string strSearchWord_Before_Web = "";

		private RichTextBox rtbContent;

		private int flagCount;

		private string strStatus = "";

		private bool IsWebChanged = true;

		private bool IsFromWeb = true;

		public int maxResultCount = 10;

		public int gwSort;

		private string strWebResult = "";

		private BackgroundWorker All_BgWorker_Web = new BackgroundWorker();

		private List<object> All_List_Web = new List<object>();

		public MyAddin ParentRibbon
		{
			get
			{
				return _parentRibbon;
			}
			set
			{
				_parentRibbon = value;
			}
		}

		public SuperWriterTips()
		{
			InitializeComponent();
		}

		public SuperWriterTips(MyAddin ribbonDesign)
		{
			_parentRibbon = ribbonDesign;
			InitializeComponent();
		}

		public void setKeyWord(string strKeyWord, int _maxCount, bool _isFromWeb, bool boolRefresh = false)
		{
			try
			{
				maxResultCount = _maxCount;
				IsFromWeb = _isFromWeb;
				if (strKeyWord.Length <= 0)
				{
					rtbContent.Text = strKeyWord + "\r\n等待提示中……";
					return;
				}
				if (flagCount == 0)
				{
					strSearchWord = strKeyWord;
					strSearchWord_Before_Web = strSearchWord;
					flagCount++;
					IsWebChanged = true;
				}
				else
				{
					strSearchWord = strKeyWord;
					if (strSearchWord_Before_Web == strSearchWord)
					{
						IsWebChanged = false;
					}
					else
					{
						IsWebChanged = true;
						strSearchWord_Before_Web = strKeyWord;
					}
				}
				if (IsWebChanged || boolRefresh)
				{
					if (All_BgWorker_Web.IsBusy)
					{
						All_BgWorker_Web.CancelAsync();
					}
					All_BgWorker_Web = new BackgroundWorker();
					All_BgWorker_Web.WorkerReportsProgress = false;
					All_BgWorker_Web.WorkerSupportsCancellation = true;
					All_BgWorker_Web.DoWork += All_DoWork_Handler_Web;
					All_BgWorker_Web.RunWorkerCompleted += All_RunWorkerCompleted_Handler_Web;
					All_BgWorker_Web.RunWorkerAsync();
					rtbContent.Text = "搜索中……";
				}
			}
			catch (Exception)
			{
			}
		}

		private void All_DoWork_Handler_Web(object sender, DoWorkEventArgs args)
		{
			try
			{
				BackgroundWorker obj = sender as BackgroundWorker;
				strWebResult = getSearchResultList_FromLocal(strSearchWord);
				if (obj.CancellationPending)
				{
					strWebResult = "";
					args.Cancel = true;
				}
			}
			catch (Exception ex)
			{
				strStatus = ex.ToString();
			}
		}

		private void All_RunWorkerCompleted_Handler_Web(object sender, RunWorkerCompletedEventArgs args)
		{
			if (args.Cancelled)
			{
				return;
			}
			try
			{
				if (strWebResult.Length <= 0)
				{
					strWebResult = "您的素材中没有相关内容\r\n平时要注意积累哦！";
				}
				int length = strSearchWord.Length;
				for (int i = 0; i < length - 1; i++)
				{
					strSearchWord.Substring(i, 2);
				}
				strWebResult = strWebResult.Replace("\n", "\r\n");
				while (("—" + strWebResult).IndexOf("\r\r") > 0)
				{
					strWebResult = strWebResult.Replace("\r\r", "\r");
				}
				while (("—" + strWebResult).IndexOf("\r\n\r\n") > 0)
				{
					strWebResult = strWebResult.Replace("\r\n\r\n", "\r\n");
				}
				rtbContent.Text = strWebResult;
				rtbContent.ForeColor = Color.Black;
				RtfUtil.ConvertKeyToHyperLink(strSearchWord, rtbContent);
				IsWebChanged = false;
			}
			catch (Exception)
			{
			}
		}

		private string getSearchResultList_FromLocal(string strSearch)
		{
			try
			{
				string text = "";
				new List<object>();
				SQLiteHelper sQLiteHelper = new SQLiteHelper();
				string text2 = "SELECT top " + maxResultCount + " sc_id,content FROM [wr_sucai_local] where sc_id>0  ";
				int num = 0;
				string[] array = strSearch.Split(CommonConfig.charTemp);
				for (int i = 0; i < array.Length; i++)
				{
					string text3 = array[i].ToString();
					if (text3.Length > 0)
					{
						text2 = text2 + "  and content like '%" + text3 + "%' ";
					}
				}
				text2 += " order by sc_id desc  ";
				SqlCeDataReader sqlCeDataReader = sQLiteHelper.ExecuteQuery(text2);
				while (sqlCeDataReader.Read())
				{
					num++;
					text = text + sqlCeDataReader["content"].ToString() + "\r\n-----------------------------\r\n";
				}
				sQLiteHelper.CloseConnection();
				try
				{
					int num2 = maxResultCount - num;
					if (num2 > 0)
					{
						SQLiteHelper sQLiteHelper2 = new SQLiteHelper();
						text2 = "SELECT top " + num2 + " content FROM [wr_fanwen_local] where 1=1  ";
						for (int j = 0; j < array.Length; j++)
						{
							string text4 = array[j].ToString();
							if (text4.Length > 0)
							{
								text2 = text2 + " and content like '%" + text4 + "%' ";
							}
						}
						text2 += " order by fw_id desc  ";
						if (CommonConfig.IS_DEBUG)
						{
							text += text2;
						}
						SqlCeDataReader sqlCeDataReader2 = sQLiteHelper2.ExecuteQuery(text2);
						int num3 = 0;
						while (sqlCeDataReader2.Read())
						{
							num3++;
							string strContent = sqlCeDataReader2["content"].ToString();
							text = text + getSubString(strContent, strSearchWord) + "\r\n-----------------------------\r\n";
						}
						sQLiteHelper2.CloseConnection();
					}
				}
				catch (Exception ex)
				{
					if (CommonConfig.IS_DEBUG)
					{
						text += ex.ToString();
					}
				}
				return text;
			}
			catch (Exception)
			{
				return "";
			}
		}

		public string getSubString(string strContent, string strWord)
		{
			string[] array = strContent.Split(new string[1] { strWord }, StringSplitOptions.None);
			array = Regex.Split(strContent, strWord);
			if (array.Length == 1)
			{
				strContent = ((strContent.IndexOf(strWord) <= 0) ? (strWord + getRightString(strContent)) : (getLeftString(strContent) + strWord));
			}
			else
			{
				string text = "";
				for (int i = 1; i < array.Length; i++)
				{
					text = text + strWord + array[i];
				}
				strContent = getLeftString(array[0]) + getRightString(text);
			}
			return strContent;
		}

		public string getLeftString(string strContent)
		{
			int num = 60;
			if (strContent.Length > num)
			{
				strContent = "……" + strContent.Substring(strContent.Length - num, num);
			}
			return strContent;
		}

		public string getRightString(string strContent)
		{
			int num = 150;
			if (strContent.Length > num)
			{
				strContent = strContent.Substring(0, num) + "……";
			}
			return strContent;
		}

		public void setMaxResultCount(int _maxCount)
		{
			maxResultCount = _maxCount;
		}

		public void setSucaiFrom(bool _isFromWeb)
		{
			IsFromWeb = _isFromWeb;
		}

		private string getSearchResultListFromWeb(string strSort, string strSearch)
		{
			string text = "";
			try
			{
				string text2 = "";
				text2 = ((strSearch.Length <= 0) ? ("Super_Writer_Web_Sort_" + maxResultCount + "_" + strSort + "_danwei_" + gwSort + "_Key_=====") : ("Super_Writer_Web_Sort_" + maxResultCount + "_" + strSort + "_danwei_" + gwSort + "_Key_" + strSearch));
				string strKey = text2;
				string text3 = FileCacheUtil.GetCache(strKey, CommonConfig.CacheHour_SuperWriter);
				if (text3.Length <= 0)
				{
					string address = UrlUtil.strGetSuCaiUrl_Super(strSearch, strSort, "", maxResultCount, gwSort);
					text3 = new WebClient
					{
						Credentials = CredentialCache.DefaultCredentials
					}.DownloadString(address);
					FileCacheUtil.SetCache(strKey, text3);
				}
				SearchResultUtil searchResultUtil = JsonConvert.DeserializeObject<SearchResultUtil>(text3);
				if (UserUtil.IsVip())
				{
					for (int i = 0; i < searchResultUtil.result.Count; i++)
					{
						string content = searchResultUtil.result[i].Content;
						text = text + (i + 1) + "、" + content + "\r\n-----------------------------\r\n";
					}
					return text;
				}
				int num = ((searchResultUtil.result.Count > CommonConfig.FreeShowCount) ? CommonConfig.FreeShowCount : searchResultUtil.result.Count);
				for (int j = 0; j < num; j++)
				{
					string content2 = searchResultUtil.result[j].Content;
					text = text + content2 + "\r\n-----------------------------\r\n";
				}
				text = text + "\r\n------------------------\r\n由于您不是VIP用户，只能显示前" + num + "条提示\r\n";
				if (searchResultUtil.result.Count > CommonConfig.FreeShowCount)
				{
					text = text + "升级VIP即可显示剩余" + (searchResultUtil.result.Count - CommonConfig.FreeShowCount) + "+条提示\r\n";
				}
				return text + "------------------------\r\n";
			}
			catch (Exception ex)
			{
				return ex.ToString();
			}
		}

		private void SuperWriterTips_Load(object sender, EventArgs e)
		{
			try
			{
				Control.CheckForIllegalCrossThreadCalls = false;
				All_BgWorker_Web.WorkerReportsProgress = false;
				All_BgWorker_Web.WorkerSupportsCancellation = true;
				All_BgWorker_Web.DoWork += All_DoWork_Handler_Web;
				All_BgWorker_Web.RunWorkerCompleted += All_RunWorkerCompleted_Handler_Web;
			}
			catch (Exception)
			{
			}
		}

		private void SuperWriterTips_ControlRemoved(object sender, ControlEventArgs e)
		{
		}

		private void SuperWriterTips_Leave(object sender, EventArgs e)
		{
		}

		private void SuperWriterTips_VisibleChanged(object sender, EventArgs e)
		{
		}

		private void InitializeComponent()
		{
			this.rtbContent = new System.Windows.Forms.RichTextBox();
			base.SuspendLayout();
			this.rtbContent.BackColor = System.Drawing.Color.White;
			this.rtbContent.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.rtbContent.Dock = System.Windows.Forms.DockStyle.Fill;
			this.rtbContent.Font = new System.Drawing.Font("宋体", 14f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.rtbContent.Location = new System.Drawing.Point(0, 0);
			this.rtbContent.Name = "rtbContent";
			this.rtbContent.ReadOnly = true;
			this.rtbContent.Size = new System.Drawing.Size(498, 498);
			this.rtbContent.TabIndex = 2;
			this.rtbContent.Text = "";
			this.BackColor = System.Drawing.Color.White;
			base.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			base.Controls.Add(this.rtbContent);
			base.Name = "SuperWriterTips";
			base.Size = new System.Drawing.Size(498, 498);
			base.Load += new System.EventHandler(SuperWriterTips_Load);
			base.ResumeLayout(false);
		}
	}
}
