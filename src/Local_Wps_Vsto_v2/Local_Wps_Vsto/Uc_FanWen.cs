using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlServerCe;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace Local_Wps_Vsto
{
	public class Uc_FanWen : UserControl
	{
		private static string STR_ERROR = "查找错误";

		private bool IsWebChanged = true;

		private Button BtnSearch;

		private TextBox tbSearch;

		private TabControl tabSort;

		private TabPage tabPage2;

		private Panel panelLocal_Right;

		private RichTextBox RtbLocal;

		private Panel panelLocal_Left;

		private ListView lvLocal;

		private ColumnHeader columnHeader1;

		private Label lbLocal;

		private Panel panelTop;

		private Button btnAdd;

		private bool IsLocalChanged = true;

		private MyAddin _parentRibbon;

		private BackgroundWorker BgWorker_Local = new BackgroundWorker();

		private List<FanWenListItem> fanWenLists_Local = new List<FanWenListItem>();

		private string strStatus = "";

		public string NowSearchWord = "";

		private int NowSortId;

		private int countLocal;

		private int countLike;

		private int countWeb;

		private int countFull;

		private bool boolQiuGaoMsg;

		private long longFwId;

		private string strLocalTitle = "";

		private BackgroundWorker BgWorker_Bind = new BackgroundWorker();

		private BackgroundWorker BgWorker_Delete = new BackgroundWorker();

		private int intFlag;

		private string strContentRtf = "";

		private string strContentTxt = "";

		private int lastSelectIndex;

		private string strWordFileName = "";

		private string strDocument = "";

		private BackgroundWorker BgWorker_ExportToWord = new BackgroundWorker();

		private bool boolSave;

		private string strTitleWeb = "";

		private string strWordUrl = "";

		private string strLocalRtf = "";

		private string strLocalWord = "";

		private int nowFlag;

		private string strFanWenTxt = "";

		private int successFlag;

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

		public Uc_FanWen()
		{
			InitializeComponent();
		}

		public Uc_FanWen(MyAddin ribbonDesign)
		{
			_parentRibbon = ribbonDesign;
			InitializeComponent();
		}

		private void Fm_FanWen_Load(object sender, EventArgs e)
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
				BgWorker_Local.WorkerReportsProgress = false;
				BgWorker_Local.WorkerSupportsCancellation = true;
				BgWorker_Local.DoWork += DoWork_Handler_Local;
				BgWorker_Local.RunWorkerCompleted += RunWorkerCompleted_Handler_Local;
				BgWorker_Delete.WorkerReportsProgress = false;
				BgWorker_Delete.WorkerSupportsCancellation = true;
				BgWorker_Delete.DoWork += DoWork_Handler_Delete;
				BgWorker_Delete.RunWorkerCompleted += RunWorkerCompleted_Handler_Delete;
			}
			catch (Exception)
			{
			}
		}

		private void DoWork_Handler_Local(object sender, DoWorkEventArgs args)
		{
			fanWenLists_Local.Clear();
			try
			{
				fanWenLists_Local = getResultList(NowSearchWord);
				if ((sender as BackgroundWorker).CancellationPending)
				{
					args.Cancel = true;
				}
			}
			catch (Exception)
			{
			}
		}

		public List<FanWenListItem> getResultList(string strSearch)
		{
			try
			{
				List<FanWenListItem> list = new List<FanWenListItem>();
				SQLiteHelper sQLiteHelper = new SQLiteHelper();
				string text = "SELECT top " + CommonConfig.MAX_LOCAL + "  fw_id,title FROM [wr_fanwen_local] where 1=1   ";
				if (text.Length > 0)
				{
					char[] separator = new char[6] { ',', '，', ' ', '\u3000', '.', '。' };
					string[] array = NowSearchWord.Split(separator);
					string text2 = "###";
					string text3 = "###";
					for (int i = 0; i < array.Length; i++)
					{
						string text4 = array[i].ToString();
						if (text4.Length > 0)
						{
							text2 = text2 + "and title like '%" + text4 + "%' ";
							text3 = text3 + "and content like '%" + text4 + "%' ";
						}
					}
					if (array.Length != 0 && text2.Length > 10)
					{
						text2 = text2.Replace("###and", "");
						text3 = text3.Replace("###and", "");
						text = text + " and ((" + text2 + ") or (" + text3 + "))";
					}
				}
				text += " order by fw_id desc  ";
				SqlCeDataReader sqlCeDataReader = sQLiteHelper.ExecuteQuery(text);
				int num = 0;
				while (sqlCeDataReader.Read())
				{
					num++;
					FanWenListItem fanWenListItem = new FanWenListItem();
					fanWenListItem.Id = num;
					fanWenListItem.FwId = int.Parse(sqlCeDataReader["fw_id"].ToString());
					fanWenListItem.Title = sqlCeDataReader["title"].ToString();
					fanWenListItem.Content = "";
					fanWenListItem.WordUrl = "";
					list.Add(fanWenListItem);
				}
				sQLiteHelper.CloseConnection();
				strStatus = list.Count.ToString();
				return list;
			}
			catch (Exception)
			{
				return null;
			}
		}

		private void RunWorkerCompleted_Handler_Local(object sender, RunWorkerCompletedEventArgs args)
		{
			try
			{
				IsLocalChanged = false;
				MsgTips.Hide();
				lvLocal.Items.Clear();
				if (args.Cancelled)
				{
					return;
				}
				lbLocal.Text = "找到" + fanWenLists_Local.Count() + "条结果";
				try
				{
					foreach (FanWenListItem item in fanWenLists_Local)
					{
						ListViewItem listViewItem = new ListViewItem(item.Title.ToString());
						listViewItem.ToolTipText = item.Title;
						lvLocal.Items.Add(listViewItem);
					}
				}
				catch (Exception)
				{
				}
			}
			catch (Exception)
			{
			}
		}

		private List<FanWenListItem> getSearchResultList(string strUrl)
		{
			try
			{
				string text = FileCacheUtil.GetCache(strUrl, CommonConfig.CacheHour_FanWen_Web);
				if (text.Length <= 0)
				{
					text = new WebClient
					{
						Credentials = CredentialCache.DefaultCredentials
					}.DownloadString(strUrl);
					FileCacheUtil.SetCache(strUrl, text);
				}
				return JsonConvert.DeserializeObject<FanWenList>(text).result;
			}
			catch (Exception)
			{
				return null;
			}
		}

		private void Fm_FanWen_FormClosed(object sender, FormClosedEventArgs e)
		{
		}

		private void BtnSearch_Click(object sender, EventArgs e)
		{
			RefreshContent();
		}

		public void RefreshContent(string strSearch = "")
		{
			try
			{
				NowSearchWord = tbSearch.Text.ToString();
				NowSortId = ((tabSort.SelectedIndex > 0) ? tabSort.SelectedIndex : 0);
				if (IsLocalChanged)
				{
					if (BgWorker_Local.IsBusy)
					{
						BgWorker_Local.CancelAsync();
					}
					BgWorker_Local = new BackgroundWorker();
					BgWorker_Local.WorkerReportsProgress = false;
					BgWorker_Local.WorkerSupportsCancellation = true;
					BgWorker_Local.DoWork += DoWork_Handler_Local;
					BgWorker_Local.RunWorkerCompleted += RunWorkerCompleted_Handler_Local;
					lbLocal.Text = "搜索中……";
					BgWorker_Local.RunWorkerAsync();
				}
			}
			catch (Exception)
			{
				MessageTipsUtil.Show("出现异常");
			}
		}

		private void DoWork_Handler_Bind(object sender, DoWorkEventArgs args)
		{
			try
			{
				SQLiteHelper sQLiteHelper = new SQLiteHelper();
				string queryString = "SELECT  content_rtf,content,creater_id FROM [wr_fanwen_local] where fw_id=" + longFwId;
				SqlCeDataReader sqlCeDataReader = sQLiteHelper.ExecuteQuery(queryString);
				if (sqlCeDataReader.Read())
				{
					intFlag = int.Parse(sqlCeDataReader["creater_id"].ToString());
					string zippedString = sqlCeDataReader["content_rtf"].ToString();
					strContentRtf = ZipHelper.GZipDecompressString(zippedString);
					strContentTxt = sqlCeDataReader["content"].ToString();
				}
				sQLiteHelper.CloseConnection();
				if (intFlag < 0 && File.Exists(strContentRtf))
				{
					RtfUtil.LoadWordTo_ClipBoard(strContentRtf);
				}
			}
			catch (Exception)
			{
			}
		}

		private void RunWorkerCompleted_Handler_Bind(object sender, RunWorkerCompletedEventArgs args)
		{
			try
			{
				RtbLocal.Text = "";
				if (intFlag > 0)
				{
					RtbLocal.Rtf = strContentRtf;
				}
				else if (intFlag == 0)
				{
					RtfUtil.setGeShi_FromTxt(RtbLocal, strContentTxt);
				}
				else if (File.Exists(strContentRtf))
				{
					RtbLocal.Text = "";
					RtbLocal.ReadOnly = false;
					RtbLocal.SelectAll();
					RtbLocal.Paste();
					RtbLocal.ReadOnly = true;
				}
				else
				{
					RtfUtil.setGeShi_FromTxt(RtbLocal, strContentTxt);
				}
			}
			catch (Exception)
			{
			}
			MsgTips.Hide();
		}

		private void DoWork_Handler_Delete(object sender, DoWorkEventArgs args)
		{
			try
			{
				SQLiteHelper sQLiteHelper = new SQLiteHelper();
				string queryString = "delete from [wr_fanwen_local] where fw_id=" + longFwId;
				sQLiteHelper.ExecuteQuery(queryString);
				sQLiteHelper.CloseConnection();
			}
			catch (Exception ex)
			{
				strStatus = ex.ToString();
			}
		}

		private void RunWorkerCompleted_Handler_Delete(object sender, RunWorkerCompletedEventArgs args)
		{
			try
			{
				IsLocalChanged = true;
				RefreshContent();
			}
			catch (Exception)
			{
				MessageTipsUtil.Show(strStatus);
			}
		}

		private void BtnExport_Click(object sender, EventArgs e)
		{
			try
			{
				string text = strLocalTitle;
				FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
				if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
				{
					MsgTips.Show("导出中……");
					string selectedPath = folderBrowserDialog.SelectedPath;
					strWordFileName = selectedPath + "\\" + text + ".docx";
					RtfUtil.SaveWord(strWordFileName, RtbLocal);
					MsgTips.Hide();
				}
			}
			catch (Exception)
			{
			}
		}

		private void DoWork_Handler_ExportToWord(object sender, DoWorkEventArgs args)
		{
			boolSave = false;
			try
			{
				RtbLocal.SaveFile(strWordFileName, RichTextBoxStreamType.RichText);
				boolSave = true;
			}
			catch (Exception)
			{
			}
		}

		private void RunWorkerCompleted_Handler_ExportToWord(object sender, RunWorkerCompletedEventArgs args)
		{
			MsgTips.Hide();
		}

		private void BtnDelete_Click(object sender, EventArgs e)
		{
			try
			{
				if (BgWorker_Delete.IsBusy)
				{
					BgWorker_Delete.CancelAsync();
				}
				BgWorker_Delete = new BackgroundWorker();
				BgWorker_Delete.WorkerReportsProgress = false;
				BgWorker_Delete.WorkerSupportsCancellation = true;
				BgWorker_Delete.DoWork += DoWork_Handler_Delete;
				BgWorker_Delete.RunWorkerCompleted += RunWorkerCompleted_Handler_Delete;
				lastSelectIndex--;
				BgWorker_Delete.RunWorkerAsync();
			}
			catch (Exception)
			{
				MessageTipsUtil.Show("删除失败");
			}
		}

		private void tbSearch_TextChanged(object sender, EventArgs e)
		{
			IsLocalChanged = true;
			IsWebChanged = true;
			RefreshContent();
		}

		private void tabSort_SelectedIndexChanged(object sender, EventArgs e)
		{
			RefreshContent();
		}

		private void DoWork_Handler_Bind_Web(object sender, DoWorkEventArgs args)
		{
			BackgroundWorker backgroundWorker = sender as BackgroundWorker;
			try
			{
				strFanWenTxt = FileCacheUtil.GetCache(strWordUrl, 240000);
				if (strFanWenTxt.Length <= 0)
				{
					strFanWenTxt = HttpUtil.getUrlResponse(strWordUrl + SecretUtil.GetMachineCode());
					FileCacheUtil.SetCache(strWordUrl, strFanWenTxt);
				}
				successFlag = 888;
				if (backgroundWorker.CancellationPending)
				{
					args.Cancel = true;
				}
			}
			catch (Exception)
			{
				successFlag = 444;
			}
		}

		private void Uc_FanWen_Load(object sender, EventArgs e)
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

		private void InitializeComponent()
		{
			this.BtnSearch = new System.Windows.Forms.Button();
			this.tbSearch = new System.Windows.Forms.TextBox();
			this.tabSort = new System.Windows.Forms.TabControl();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.panelLocal_Right = new System.Windows.Forms.Panel();
			this.RtbLocal = new System.Windows.Forms.RichTextBox();
			this.panelLocal_Left = new System.Windows.Forms.Panel();
			this.lvLocal = new System.Windows.Forms.ListView();
			this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
			this.lbLocal = new System.Windows.Forms.Label();
			this.panelTop = new System.Windows.Forms.Panel();
			this.btnAdd = new System.Windows.Forms.Button();
			this.tabSort.SuspendLayout();
			this.tabPage2.SuspendLayout();
			this.panelLocal_Right.SuspendLayout();
			this.panelLocal_Left.SuspendLayout();
			this.panelTop.SuspendLayout();
			base.SuspendLayout();
			this.BtnSearch.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnSearch.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnSearch.Location = new System.Drawing.Point(609, 0);
			this.BtnSearch.Name = "BtnSearch";
			this.BtnSearch.Size = new System.Drawing.Size(206, 35);
			this.BtnSearch.TabIndex = 4;
			this.BtnSearch.Text = "搜索";
			this.BtnSearch.UseVisualStyleBackColor = true;
			this.BtnSearch.Click += new System.EventHandler(BtnSearch_Click);
			this.tbSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.tbSearch.Dock = System.Windows.Forms.DockStyle.Left;
			this.tbSearch.Font = new System.Drawing.Font("宋体", 14f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tbSearch.Location = new System.Drawing.Point(0, 0);
			this.tbSearch.MaxLength = 20;
			this.tbSearch.Name = "tbSearch";
			this.tbSearch.Size = new System.Drawing.Size(609, 34);
			this.tbSearch.TabIndex = 3;
			this.tabSort.Controls.Add(this.tabPage2);
			this.tabSort.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tabSort.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tabSort.Location = new System.Drawing.Point(0, 35);
			this.tabSort.Name = "tabSort";
			this.tabSort.Padding = new System.Drawing.Point(10, 8);
			this.tabSort.SelectedIndex = 0;
			this.tabSort.Size = new System.Drawing.Size(1000, 665);
			this.tabSort.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight;
			this.tabSort.TabIndex = 13;
			this.tabSort.SelectedIndexChanged += new System.EventHandler(tabSort_SelectedIndexChanged);
			this.tabPage2.Controls.Add(this.panelLocal_Right);
			this.tabPage2.Controls.Add(this.panelLocal_Left);
			this.tabPage2.Location = new System.Drawing.Point(4, 40);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
			this.tabPage2.Size = new System.Drawing.Size(992, 621);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "本地范文";
			this.tabPage2.UseVisualStyleBackColor = true;
			this.panelLocal_Right.Controls.Add(this.RtbLocal);
			this.panelLocal_Right.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panelLocal_Right.Location = new System.Drawing.Point(303, 3);
			this.panelLocal_Right.Name = "panelLocal_Right";
			this.panelLocal_Right.Size = new System.Drawing.Size(686, 615);
			this.panelLocal_Right.TabIndex = 2;
			this.RtbLocal.BackColor = System.Drawing.Color.White;
			this.RtbLocal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.RtbLocal.Dock = System.Windows.Forms.DockStyle.Fill;
			this.RtbLocal.Location = new System.Drawing.Point(0, 0);
			this.RtbLocal.Name = "RtbLocal";
			this.RtbLocal.ReadOnly = true;
			this.RtbLocal.Size = new System.Drawing.Size(686, 615);
			this.RtbLocal.TabIndex = 1;
			this.RtbLocal.Text = "";
			this.panelLocal_Left.Controls.Add(this.lvLocal);
			this.panelLocal_Left.Controls.Add(this.lbLocal);
			this.panelLocal_Left.Dock = System.Windows.Forms.DockStyle.Left;
			this.panelLocal_Left.Location = new System.Drawing.Point(3, 3);
			this.panelLocal_Left.Name = "panelLocal_Left";
			this.panelLocal_Left.Size = new System.Drawing.Size(300, 615);
			this.panelLocal_Left.TabIndex = 1;
			this.lvLocal.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.lvLocal.Columns.AddRange(new System.Windows.Forms.ColumnHeader[1] { this.columnHeader1 });
			this.lvLocal.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lvLocal.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
			this.lvLocal.HideSelection = false;
			this.lvLocal.Location = new System.Drawing.Point(0, 20);
			this.lvLocal.Name = "lvLocal";
			this.lvLocal.Size = new System.Drawing.Size(300, 595);
			this.lvLocal.TabIndex = 3;
			this.lvLocal.TileSize = new System.Drawing.Size(300, 44);
			this.lvLocal.UseCompatibleStateImageBehavior = false;
			this.lvLocal.View = System.Windows.Forms.View.Tile;
			this.lvLocal.SelectedIndexChanged += new System.EventHandler(lvLocal_SelectedIndexChanged);
			this.columnHeader1.Text = "文章标题";
			this.columnHeader1.Width = 295;
			this.lbLocal.AutoSize = true;
			this.lbLocal.Dock = System.Windows.Forms.DockStyle.Top;
			this.lbLocal.Location = new System.Drawing.Point(0, 0);
			this.lbLocal.Name = "lbLocal";
			this.lbLocal.Size = new System.Drawing.Size(109, 20);
			this.lbLocal.TabIndex = 2;
			this.lbLocal.Text = "搜索中……";
			this.panelTop.Controls.Add(this.BtnSearch);
			this.panelTop.Controls.Add(this.tbSearch);
			this.panelTop.Controls.Add(this.btnAdd);
			this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
			this.panelTop.Location = new System.Drawing.Point(0, 0);
			this.panelTop.Margin = new System.Windows.Forms.Padding(3, 10, 3, 10);
			this.panelTop.Name = "panelTop";
			this.panelTop.Size = new System.Drawing.Size(1000, 35);
			this.panelTop.TabIndex = 12;
			this.btnAdd.Dock = System.Windows.Forms.DockStyle.Right;
			this.btnAdd.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.btnAdd.Location = new System.Drawing.Point(815, 0);
			this.btnAdd.Name = "btnAdd";
			this.btnAdd.Size = new System.Drawing.Size(185, 35);
			this.btnAdd.TabIndex = 4;
			this.btnAdd.Text = "添加范文";
			this.btnAdd.UseVisualStyleBackColor = true;
			this.btnAdd.Click += new System.EventHandler(btnAdd_Click);
			base.Controls.Add(this.tabSort);
			base.Controls.Add(this.panelTop);
			base.Name = "Uc_FanWen";
			base.Size = new System.Drawing.Size(1000, 700);
			base.Load += new System.EventHandler(Uc_FanWen_Load);
			this.tabSort.ResumeLayout(false);
			this.tabPage2.ResumeLayout(false);
			this.panelLocal_Right.ResumeLayout(false);
			this.panelLocal_Left.ResumeLayout(false);
			this.panelLocal_Left.PerformLayout();
			this.panelTop.ResumeLayout(false);
			this.panelTop.PerformLayout();
			base.ResumeLayout(false);
		}

		private void btnAdd_Click(object sender, EventArgs e)
		{
			_parentRibbon.openNewFanWenPanel();
		}

		private void lvLocal_SelectedIndexChanged(object sender, EventArgs e)
		{
			try
			{
				int index = (lastSelectIndex = ((lvLocal.SelectedItems[0].Index > 0) ? lvLocal.SelectedItems[0].Index : 0));
				List<FanWenListItem> range = fanWenLists_Local.GetRange(index, 1);
				FanWenListItem fanWenListItem = new FanWenListItem();
				foreach (FanWenListItem item in range)
				{
					fanWenListItem = item;
				}
				longFwId = fanWenListItem.FwId;
				strLocalTitle = fanWenListItem.Title;
				MsgTips.Show("");
				RtbLocal.Text = "加载中……";
				if (UserUtil.IsVip())
				{
					if (BgWorker_Bind.IsBusy)
					{
						BgWorker_Bind.CancelAsync();
					}
					BgWorker_Bind = new BackgroundWorker();
					BgWorker_Bind.WorkerReportsProgress = true;
					BgWorker_Bind.WorkerSupportsCancellation = true;
					BgWorker_Bind.DoWork += DoWork_Handler_Bind;
					BgWorker_Bind.RunWorkerCompleted += RunWorkerCompleted_Handler_Bind;
					BgWorker_Bind.RunWorkerAsync();
				}
				else
				{
					RtbLocal.Text = "软件尚未激活，激活后才能查看全文";
				}
			}
			catch (Exception)
			{
			}
		}
	}
}
