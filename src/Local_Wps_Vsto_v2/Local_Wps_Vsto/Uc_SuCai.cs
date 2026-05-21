using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Net;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace Local_Wps_Vsto
{
	public class Uc_SuCai : UserControl
	{
		private static string STR_ERROR = "网络异常！！！";

		private RichTextBox RtbWeb_All;

		private Button BtnSearch;

		private TextBox tbSearch;

		private Panel panelTop;

		private Button btnAdd;

		private TabControl tabSort;

		private TabPage TabItem_All;

		private TabPage tabPage1;

		private TabPage tabPage2;

		private TabPage tabPage3;

		private TabPage tabPage4;

		private SaveFileDialog saveFileDialog1;

		private TabPage tabPage5;

		private MyAddin _parentRibbon;

		private BackgroundWorker Local_BgWorker;

		private string strSortName = "";

		private string strStatus = "";

		private List<object> All_List_Local = new List<object>();

		private string NowSearchWord = "";

		private bool boolCache;

		private string strResult = "";

		private bool IsLocal;

		private int totalCount;

		private string strResult_Web = "";

		private Timer timerSearchChanged;

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

		public Uc_SuCai()
		{
			InitializeComponent();
		}

		public Uc_SuCai(MyAddin ribbonDesign)
		{
			_parentRibbon = ribbonDesign;
			InitializeComponent();
		}

		private void Uc_SuCai_Load(object sender, EventArgs e)
		{
			try
			{
				Control.CheckForIllegalCrossThreadCalls = false;
				BindBackgroudWorker();
				RefreshContent();
			}
			catch (Exception)
			{
			}
		}

		private void BindBackgroudWorker()
		{
			try
			{
				Local_BgWorker = new BackgroundWorker();
				Local_BgWorker.WorkerReportsProgress = false;
				Local_BgWorker.WorkerSupportsCancellation = true;
				Local_BgWorker.DoWork += Local_DoWork_Handler;
				Local_BgWorker.RunWorkerCompleted += Local_RunWorkerCompleted_Handler;
			}
			catch (Exception)
			{
			}
		}

		private void Fm_SuCai_FormClosed(object sender, FormClosedEventArgs e)
		{
		}

		private void BtnSearch_Click(object sender, EventArgs e)
		{
			RefreshContent();
		}

		public void RefreshContent()
		{
			try
			{
				NowSearchWord = tbSearch.Text.ToString();
				strSortName = CommonConfig.SortName[tabSort.SelectedIndex];
				if (Local_BgWorker.IsBusy)
				{
					Local_BgWorker.CancelAsync();
				}
				Local_BgWorker = new BackgroundWorker();
				Local_BgWorker.WorkerReportsProgress = false;
				Local_BgWorker.WorkerSupportsCancellation = true;
				Local_BgWorker.DoWork += Local_DoWork_Handler;
				Local_BgWorker.RunWorkerCompleted += Local_RunWorkerCompleted_Handler;
				Text = "素材搜索-个人素材";
				Local_BgWorker.RunWorkerAsync();
				RtbWeb_All.Text = "------------------------搜索中……------------------------";
			}
			catch (Exception)
			{
			}
		}

		private string getSearchResultList_Local(string strSort, string strSearch)
		{
			boolCache = false;
			string text = "";
			try
			{
				string text4 = "Local_sucai_sort_" + strSort + "_key_" + strSearch;
				if (text.Length > 0)
				{
					boolCache = true;
					return text;
				}
				SQLiteHelper sQLiteHelper = new SQLiteHelper();
				string text2 = "";
				text2 = ((!(strSort == "nn")) ? ("SELECT top " + CommonConfig.MAX_LOCAL + "  sc_id,content FROM [wr_sucai_local] where sort_name='" + strSort + "' ") : ("SELECT top " + CommonConfig.MAX_LOCAL + " sc_id,content FROM [wr_sucai_local] where  sort_name!='1'   "));
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
				bool flag = false;
				int num = 0;
				while (sqlCeDataReader.Read())
				{
					num++;
					if (UserUtil.IsVip())
					{
						text = text + "\r\n\r\n" + num + "、" + sqlCeDataReader["content"].ToString();
					}
					else if (num <= CommonConfig.FreeShowCount)
					{
						text = text + "\r\n\r\n" + num + "、" + sqlCeDataReader["content"].ToString();
					}
					else if (!flag)
					{
						text = text + "\r\n\r\n------------------------\r\n由于您不是VIP用户，只能显示前" + CommonConfig.FreeShowCount + "条素材\r\n";
						flag = true;
					}
				}
				sQLiteHelper.CloseConnection();
				text = "------------------------" + getShowStatus(num) + "------------------------" + text;
				if (num == 0)
				{
					return text + "\r\n\r\n本地素材库没有相关内容,平时要注意积累素材哦！";
				}
				return text;
			}
			catch (Exception)
			{
				return "\r\n本地素材库没有相关内容,平时要注意积累素材哦！";
			}
		}

		private string getShowStatus(int totalCount)
		{
			string text = "共搜到" + totalCount;
			if (totalCount >= CommonConfig.MAX_LOCAL)
			{
				return text + "+条结果";
			}
			return text + "条结果";
		}

		private void Local_DoWork_Handler(object sender, DoWorkEventArgs args)
		{
			try
			{
				strResult = getSearchResultList_Local(strSortName, NowSearchWord);
				if ((sender as BackgroundWorker).CancellationPending)
				{
					args.Cancel = true;
				}
			}
			catch (Exception)
			{
				strStatus = STR_ERROR;
			}
		}

		private void Local_RunWorkerCompleted_Handler(object sender, RunWorkerCompletedEventArgs args)
		{
			if (args.Cancelled)
			{
				return;
			}
			try
			{
				RtbWeb_All.Text = strResult;
				if (NowSearchWord.Length > 0)
				{
					string[] array = NowSearchWord.Split(CommonConfig.charTemp);
					for (int i = 0; i < array.Length; i++)
					{
						string text = array[i].ToString();
						if (text.Length > 0)
						{
							RtfUtil.ConvertKeyToHyperLink(text, RtbWeb_All);
						}
					}
				}
				RtbWeb_All.AutoScrollOffset = new Point(0, 0);
			}
			catch (Exception)
			{
			}
		}

		private void CbSort_SelectedIndexChanged(object sender, EventArgs e)
		{
			try
			{
				TabItem_All.Text = "个人全部素材";
				RefreshContent();
			}
			catch (Exception)
			{
			}
		}

		private string getSearchResultList_Web(string strSort, string strSearch)
		{
			boolCache = false;
			string text = "";
			try
			{
				string text2 = "";
				text2 = ((strSearch.Length <= 0) ? ("Web_sort_" + strSort + "_key_=====") : ("Web_sort_" + strSort + "_key_" + strSearch));
				string text3 = FileCacheUtil.GetCache(text2, CommonConfig.CacheHour_Sucai_Web);
				if (text3.Length > 0)
				{
					boolCache = true;
				}
				else
				{
					string address = UrlUtil.strGetSuCaiUrl(strSort, strSearch, "");
					text3 = new WebClient
					{
						Credentials = CredentialCache.DefaultCredentials
					}.DownloadString(address);
					FileCacheUtil.SetCache(text2, text3);
				}
				SearchResultUtil searchResultUtil = JsonConvert.DeserializeObject<SearchResultUtil>(text3);
				if (UserUtil.IsVip())
				{
					text = text + "------------------------" + getShowStatus(searchResultUtil.result.Count) + "------------------------";
					for (int i = 0; i < searchResultUtil.result.Count; i++)
					{
						text = text + "\r\n\r\n" + (i + 1) + "、" + searchResultUtil.result[i].Content;
					}
					if (searchResultUtil.result.Count == 0)
					{
						text += "\r\n\r\n云端素材库尚未收录相关内容\r\n我们会尽快完善的！";
					}
				}
				else
				{
					text = text + "------------------------" + getShowStatus(searchResultUtil.result.Count) + "------------------------";
					int num = ((searchResultUtil.result.Count > CommonConfig.FreeShowCount) ? CommonConfig.FreeShowCount : searchResultUtil.result.Count);
					for (int j = 0; j < num; j++)
					{
						string content = searchResultUtil.result[j].Content;
						text = text + "\r\n\r\n" + (j + 1) + "、" + content;
					}
					text = ((!UserUtil.HasLogin()) ? (text + "\r\n\r\n------------------------\r\n由于您尚未登录，只能显示前" + num + "条素材\r\n") : (text + "\r\n\r\n------------------------\r\n由于您不是VIP用户，只能显示前" + num + "条素材\r\n"));
					int count = searchResultUtil.result.Count;
					short freeShowCount = CommonConfig.FreeShowCount;
					text += "------------------------\r\n";
				}
				totalCount = searchResultUtil.result.Count;
				return text;
			}
			catch (Exception)
			{
				return text;
			}
		}

		private void Web_DoWork_Handler(object sender, DoWorkEventArgs args)
		{
			try
			{
				strResult_Web = getSearchResultList_Web(strSortName, NowSearchWord);
				if ((sender as BackgroundWorker).CancellationPending)
				{
					args.Cancel = true;
				}
			}
			catch (Exception)
			{
				strStatus = STR_ERROR;
			}
		}

		private void Web_RunWorkerCompleted_Handler(object sender, RunWorkerCompletedEventArgs args)
		{
			if (!args.Cancelled)
			{
				try
				{
					RtbWeb_All.Text = strResult_Web;
					if (NowSearchWord.Length > 0)
					{
						string[] array = NowSearchWord.Split(CommonConfig.charTemp);
						for (int i = 0; i < array.Length; i++)
						{
							string text = array[i].ToString();
							if (text.Length > 0)
							{
								RtfUtil.ConvertKeyToHyperLink(text, RtbWeb_All);
							}
						}
					}
					RtbWeb_All.AutoScrollOffset = new Point(0, 0);
				}
				catch (Exception)
				{
				}
			}
			MsgTips.Hide();
		}

		private void tabSort_SelectedIndexChanged(object sender, EventArgs e)
		{
			RefreshContent();
		}

		private void tbSearch_TextChanged(object sender, EventArgs e)
		{
		}

		private void Timer_SearchTxtChanged(object sender, EventArgs e)
		{
			try
			{
				RefreshContent();
				timerSearchChanged.Stop();
				timerSearchChanged = null;
			}
			catch (Exception)
			{
			}
		}

		private void btnAdd_Click(object sender, EventArgs e)
		{
			try
			{
				_parentRibbon.openNewSuCaiPanel();
			}
			catch (Exception)
			{
			}
		}

		private void InitializeComponent()
		{
			this.RtbWeb_All = new System.Windows.Forms.RichTextBox();
			this.BtnSearch = new System.Windows.Forms.Button();
			this.tbSearch = new System.Windows.Forms.TextBox();
			this.panelTop = new System.Windows.Forms.Panel();
			this.btnAdd = new System.Windows.Forms.Button();
			this.tabSort = new System.Windows.Forms.TabControl();
			this.TabItem_All = new System.Windows.Forms.TabPage();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.tabPage4 = new System.Windows.Forms.TabPage();
			this.tabPage5 = new System.Windows.Forms.TabPage();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.panelTop.SuspendLayout();
			this.tabSort.SuspendLayout();
			base.SuspendLayout();
			this.RtbWeb_All.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.RtbWeb_All.Dock = System.Windows.Forms.DockStyle.Fill;
			this.RtbWeb_All.Font = new System.Drawing.Font("宋体", 11f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.RtbWeb_All.Location = new System.Drawing.Point(0, 72);
			this.RtbWeb_All.Name = "RtbWeb_All";
			this.RtbWeb_All.Size = new System.Drawing.Size(798, 454);
			this.RtbWeb_All.TabIndex = 13;
			this.RtbWeb_All.Text = "";
			this.BtnSearch.Dock = System.Windows.Forms.DockStyle.Left;
			this.BtnSearch.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnSearch.Location = new System.Drawing.Point(438, 0);
			this.BtnSearch.Name = "BtnSearch";
			this.BtnSearch.Size = new System.Drawing.Size(93, 33);
			this.BtnSearch.TabIndex = 4;
			this.BtnSearch.Text = "搜索";
			this.BtnSearch.UseVisualStyleBackColor = true;
			this.BtnSearch.Click += new System.EventHandler(BtnSearch_Click);
			this.tbSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.tbSearch.Dock = System.Windows.Forms.DockStyle.Left;
			this.tbSearch.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tbSearch.Location = new System.Drawing.Point(0, 0);
			this.tbSearch.MaxLength = 20;
			this.tbSearch.Name = "tbSearch";
			this.tbSearch.Size = new System.Drawing.Size(438, 30);
			this.tbSearch.TabIndex = 3;
			this.panelTop.Controls.Add(this.btnAdd);
			this.panelTop.Controls.Add(this.BtnSearch);
			this.panelTop.Controls.Add(this.tbSearch);
			this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
			this.panelTop.Location = new System.Drawing.Point(0, 39);
			this.panelTop.Margin = new System.Windows.Forms.Padding(3, 10, 3, 10);
			this.panelTop.Name = "panelTop";
			this.panelTop.Size = new System.Drawing.Size(798, 33);
			this.panelTop.TabIndex = 11;
			this.btnAdd.Dock = System.Windows.Forms.DockStyle.Fill;
			this.btnAdd.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.btnAdd.Location = new System.Drawing.Point(531, 0);
			this.btnAdd.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
			this.btnAdd.Name = "btnAdd";
			this.btnAdd.Size = new System.Drawing.Size(267, 33);
			this.btnAdd.TabIndex = 5;
			this.btnAdd.Text = "添加素材";
			this.btnAdd.UseVisualStyleBackColor = true;
			this.btnAdd.Click += new System.EventHandler(btnAdd_Click);
			this.tabSort.Controls.Add(this.TabItem_All);
			this.tabSort.Controls.Add(this.tabPage1);
			this.tabSort.Controls.Add(this.tabPage2);
			this.tabSort.Controls.Add(this.tabPage3);
			this.tabSort.Controls.Add(this.tabPage4);
			this.tabSort.Controls.Add(this.tabPage5);
			this.tabSort.Dock = System.Windows.Forms.DockStyle.Top;
			this.tabSort.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tabSort.Location = new System.Drawing.Point(0, 0);
			this.tabSort.Name = "tabSort";
			this.tabSort.Padding = new System.Drawing.Point(10, 8);
			this.tabSort.SelectedIndex = 0;
			this.tabSort.Size = new System.Drawing.Size(798, 39);
			this.tabSort.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight;
			this.tabSort.TabIndex = 12;
			this.tabSort.SelectedIndexChanged += new System.EventHandler(tabSort_SelectedIndexChanged);
			this.TabItem_All.Location = new System.Drawing.Point(4, 40);
			this.TabItem_All.Name = "TabItem_All";
			this.TabItem_All.Size = new System.Drawing.Size(790, 0);
			this.TabItem_All.TabIndex = 5;
			this.TabItem_All.Text = "全部素材";
			this.TabItem_All.UseVisualStyleBackColor = true;
			this.tabPage1.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tabPage1.Location = new System.Drawing.Point(4, 40);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
			this.tabPage1.Size = new System.Drawing.Size(790, 0);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "工作常识";
			this.tabPage1.UseVisualStyleBackColor = true;
			this.tabPage2.Location = new System.Drawing.Point(4, 40);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
			this.tabPage2.Size = new System.Drawing.Size(790, 0);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "经典提纲";
			this.tabPage2.UseVisualStyleBackColor = true;
			this.tabPage3.Location = new System.Drawing.Point(4, 40);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
			this.tabPage3.Size = new System.Drawing.Size(790, 0);
			this.tabPage3.TabIndex = 2;
			this.tabPage3.Text = "句式词语";
			this.tabPage3.UseVisualStyleBackColor = true;
			this.tabPage4.Location = new System.Drawing.Point(4, 40);
			this.tabPage4.Name = "tabPage4";
			this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
			this.tabPage4.Size = new System.Drawing.Size(790, 0);
			this.tabPage4.TabIndex = 3;
			this.tabPage4.Text = "名言警句";
			this.tabPage4.UseVisualStyleBackColor = true;
			this.tabPage5.Location = new System.Drawing.Point(4, 40);
			this.tabPage5.Name = "tabPage5";
			this.tabPage5.Padding = new System.Windows.Forms.Padding(3);
			this.tabPage5.Size = new System.Drawing.Size(790, 0);
			this.tabPage5.TabIndex = 4;
			this.tabPage5.Text = "习语经典";
			this.tabPage5.UseVisualStyleBackColor = true;
			base.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			base.Controls.Add(this.RtbWeb_All);
			base.Controls.Add(this.panelTop);
			base.Controls.Add(this.tabSort);
			base.Name = "Uc_SuCai";
			base.Size = new System.Drawing.Size(798, 526);
			base.Load += new System.EventHandler(Uc_SuCai_Load);
			base.VisibleChanged += new System.EventHandler(Uc_SuCai_VisibleChanged);
			this.panelTop.ResumeLayout(false);
			this.panelTop.PerformLayout();
			this.tabSort.ResumeLayout(false);
			base.ResumeLayout(false);
		}

		private void Uc_SuCai_VisibleChanged(object sender, EventArgs e)
		{
			try
			{
				if (!base.Visible)
				{
					MyAddin.boolTbSucai = false;
					MyAddin.thisAddIn.InvalidateControl("tbSuCai");
				}
			}
			catch (Exception)
			{
			}
		}
	}
}
