using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;

namespace Local_Wps_Vsto
{
	public class Uc_NewFanWen : UserControl
	{
		private MyAddin _parentRibbon;

		private BackgroundWorker BgWorker_Local = new BackgroundWorker();

		private BackgroundWorker BgWorker_Upload = new BackgroundWorker();

		private string strUp_Title = "";

		private string strUp_Content = "";

		private string strUp_Return = "";

		private string strStatus = "";

		private int totalInsert;

		private string strException = "没有发生异常";

		private List<string> list = new List<string>();

		private bool boolShare;

		private IContainer components;

		private Panel panel2;

		private Label lbFolder;

		private TextBox tbFolder;

		private Button BtnSelectFolder;

		private Panel panel1;

		private Button BtnStartFolder;

		private RichTextBox RtbFolder;

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

		public Uc_NewFanWen()
		{
			InitializeComponent();
		}

		public Uc_NewFanWen(MyAddin ribbonDesign)
		{
			_parentRibbon = ribbonDesign;
			InitializeComponent();
		}

		private void Uc_NewFanWen_Load(object sender, EventArgs e)
		{
			try
			{
				Control.CheckForIllegalCrossThreadCalls = false;
				BgWorker_Local.WorkerReportsProgress = true;
				BgWorker_Local.WorkerSupportsCancellation = true;
				BgWorker_Local.DoWork += DoWork_Handler_Local;
				BgWorker_Local.ProgressChanged += ProgressChanged_Handler_Local;
				BgWorker_Local.RunWorkerCompleted += RunWorkerCompleted_Handler_Local;
				BgWorker_Upload.WorkerReportsProgress = true;
				BgWorker_Upload.WorkerSupportsCancellation = true;
				BgWorker_Upload.DoWork += DoWork_Handler_Upload;
				BgWorker_Upload.RunWorkerCompleted += RunWorkerCompleted_Handler_Upload;
			}
			catch (Exception)
			{
			}
		}

		private void DoWork_Handler_Upload(object sender, DoWorkEventArgs args)
		{
			strUp_Return = "";
			try
			{
				string address = UrlUtil.strUploadOneArticle();
				WebClient webClient = new WebClient();
				webClient.Credentials = CredentialCache.DefaultCredentials;
				string text = "content=" + strUp_Content + "&user_id=" + 1L + "&title=" + strUp_Title;
				text = text + "&mac=" + SecretUtil.GetMachineCode();
				webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
				byte[] bytes = Encoding.UTF8.GetBytes(text);
				byte[] bytes2 = webClient.UploadData(address, "POST", bytes);
				strUp_Return = Encoding.UTF8.GetString(bytes2);
			}
			catch (Exception)
			{
			}
		}

		private void RunWorkerCompleted_Handler_Upload(object sender, RunWorkerCompletedEventArgs args)
		{
		}

		private void DoWork_Handler_Local(object sender, DoWorkEventArgs args)
		{
			try
			{
				BackgroundWorker backgroundWorker = sender as BackgroundWorker;
				strStatus = strStatus + "开始插入范文……" + list.Count;
				backgroundWorker.ReportProgress(1);
				SQLiteHelper sQLiteHelper = new SQLiteHelper();
				if (CommonConfig.IS_DEBUG)
				{
					strStatus = "开始插入范文……创建sqlite成功";
					backgroundWorker.ReportProgress(1);
				}
				int num = 1;
				Document document;
				try
				{
					document = new Document();
				}
				catch (Exception ex)
				{
					if (CommonConfig.IS_DEBUG)
					{
						strStatus = "开始插入范文……Document失败" + ex.ToString();
						backgroundWorker.ReportProgress(1);
					}
					return;
				}
				if (CommonConfig.IS_DEBUG)
				{
					strStatus = "开始插入范文……Document成功";
					backgroundWorker.ReportProgress(1);
				}
				foreach (string item in list)
				{
					if (CommonConfig.IS_DEBUG)
					{
						strStatus = "开始插入范文……Document成功" + item;
						backgroundWorker.ReportProgress(1);
					}
					if (backgroundWorker.CancellationPending)
					{
						args.Cancel = true;
						break;
					}
					long length = new FileInfo(item).Length;
					long num2 = 10485760L;
					if (length < num2)
					{
						if (CommonConfig.IS_DEBUG)
						{
							strStatus = "开始插入范文……" + item;
							backgroundWorker.ReportProgress(1);
						}
						try
						{
							string text4 = CommonConfig.strBaseFolder + "\\cache\\fw\\00003.rtf";
							string text5 = CommonConfig.strBaseFolder + "\\cache\\fw\\00003.txt";
							string text = Path.GetExtension(item).ToLower();
							string text2 = "";
							string rawString = "";
							bool flag = true;
							int num3 = 0;
							switch (text)
							{
							case ".txt":
								num3 = 0;
								text2 = File.ReadAllText(item);
								break;
							case ".rtf":
							case ".doc":
							case ".docx":
								num3 = -1;
								document.LoadFromFile(item);
								text2 = document.GetText();
								rawString = item;
								break;
							default:
								flag = false;
								break;
							}
							if (CommonConfig.IS_DEBUG)
							{
								strStatus = "开始插入范文……" + text2;
								backgroundWorker.ReportProgress(1);
							}
							if (flag)
							{
								string text3 = ZipHelper.GZipCompressString(rawString);
								string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(item);
								string md5Hash = CommonConfig.GetMd5Hash(text2);
								sQLiteHelper.InsertValues_WithC("wr_fanwen_local", "title,content,content_rtf,md5,creater_id,create_time", new string[6]
								{
									fileNameWithoutExtension,
									text2,
									text3,
									md5Hash,
									num3.ToString(),
									DateTime.Now.ToLocalTime().ToString("s")
								});
								strStatus = "插入第" + num + "篇范文\r\n" + item;
								backgroundWorker.ReportProgress(num);
							}
							else
							{
								strStatus = "插入第" + num + "篇范文失败\r\n" + item;
								backgroundWorker.ReportProgress(num);
							}
						}
						catch (Exception ex2)
						{
							strStatus = "插入第" + num + "篇范文失败\r\n" + item + ex2.ToString();
							backgroundWorker.ReportProgress(num);
						}
					}
					num++;
					totalInsert++;
				}
				sQLiteHelper.CloseConnection();
				strStatus = "已插入全部" + list.Count + "篇范文";
				document.Close();
			}
			catch (Exception ex3)
			{
				if (CommonConfig.IS_DEBUG)
				{
					strException = ex3.ToString();
				}
			}
		}

		private void RunWorkerCompleted_Handler_Local(object sender, RunWorkerCompletedEventArgs args)
		{
			lbFolder.Text = strStatus;
			if (!CommonConfig.IS_DEBUG)
			{
				MessageTipsUtil.Show("处理完成！");
				BtnStartFolder.Enabled = false;
				BtnSelectFolder.Enabled = true;
			}
		}

		private void ProgressChanged_Handler_Local(object sender, ProgressChangedEventArgs args)
		{
			int progressPercentage = args.ProgressPercentage;
			lbFolder.Text = strStatus;
		}

		private void ProgressChanged_Handler_Upload(object sender, ProgressChangedEventArgs args)
		{
			int progressPercentage = args.ProgressPercentage;
		}

		private void Fm_NewFanWen_FormClosed(object sender, FormClosedEventArgs e)
		{
		}

		private void BtnSelectFolder_Click(object sender, EventArgs e)
		{
			showSelectFolderDialog();
		}

		private void showSelectFolderDialog()
		{
			try
			{
				FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
				if (folderBrowserDialog.ShowDialog() != DialogResult.OK)
				{
					return;
				}
				tbFolder.Text = folderBrowserDialog.SelectedPath;
				this.list.Clear();
				List<string> list = FindFile(folderBrowserDialog.SelectedPath);
				string text = "";
				int num = 1;
				foreach (string item in list)
				{
					text = text + "\r\n" + num + "、" + item;
					num++;
				}
				Text = "共找到" + list.Count + "个文件";
				RtbFolder.Text = text;
				if (list.Count > 0)
				{
					BtnStartFolder.Enabled = true;
				}
			}
			catch (Exception)
			{
			}
		}

		public List<string> FindFile(string sSourcePath)
		{
			try
			{
				DirectoryInfo directoryInfo = new DirectoryInfo(sSourcePath);
				string[] array2 = new string[4] { "*.doc", "*.docx", "*.txt", "*.rtf" };
				DirectoryInfo[] directories = directoryInfo.GetDirectories();
				if (directories.Length == 0)
				{
					FileInfo[] files = directoryInfo.GetFiles("*.*", SearchOption.TopDirectoryOnly);
					foreach (FileInfo fileInfo in files)
					{
						if (fileInfo.Extension.ToLower() == ".doc" || fileInfo.Extension.ToLower() == ".docx" || fileInfo.Extension.ToLower() == ".rtf" || fileInfo.Extension.ToLower() == ".txt")
						{
							list.Add(((directoryInfo != null) ? directoryInfo.ToString() : null) + "\\" + fileInfo.ToString());
						}
					}
				}
				int num = 1;
				DirectoryInfo[] array = directories;
				foreach (DirectoryInfo directoryInfo2 in array)
				{
					FindFile(((directoryInfo != null) ? directoryInfo.ToString() : null) + "\\" + directoryInfo2.ToString());
					if (num != 1)
					{
						continue;
					}
					FileInfo[] files = directoryInfo.GetFiles("*.*", SearchOption.TopDirectoryOnly);
					foreach (FileInfo fileInfo2 in files)
					{
						if (fileInfo2.Extension == ".doc" || fileInfo2.Extension == ".docx" || fileInfo2.Extension == ".rtf" || fileInfo2.Extension == ".txt")
						{
							list.Add(((directoryInfo != null) ? directoryInfo.ToString() : null) + "\\" + fileInfo2.ToString());
						}
					}
					num++;
				}
				return list;
			}
			catch (Exception)
			{
				return null;
			}
		}

		private void tbFolder_Click(object sender, EventArgs e)
		{
			showSelectFolderDialog();
		}

		private void BtnStartFolder_Click(object sender, EventArgs e)
		{
			try
			{
				if (list.Count > 0)
				{
					MsgTips.Show("添加中……");
					BtnStartFolder.Enabled = false;
					BtnSelectFolder.Enabled = false;
					BgWorker_Local.RunWorkerAsync();
				}
			}
			catch (Exception)
			{
			}
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing && components != null)
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		private void InitializeComponent()
		{
			this.panel2 = new System.Windows.Forms.Panel();
			this.lbFolder = new System.Windows.Forms.Label();
			this.tbFolder = new System.Windows.Forms.TextBox();
			this.BtnSelectFolder = new System.Windows.Forms.Button();
			this.panel1 = new System.Windows.Forms.Panel();
			this.BtnStartFolder = new System.Windows.Forms.Button();
			this.RtbFolder = new System.Windows.Forms.RichTextBox();
			this.panel2.SuspendLayout();
			this.panel1.SuspendLayout();
			base.SuspendLayout();
			this.panel2.Controls.Add(this.lbFolder);
			this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
			this.panel2.Location = new System.Drawing.Point(0, 45);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(800, 58);
			this.panel2.TabIndex = 17;
			this.lbFolder.AutoSize = true;
			this.lbFolder.Location = new System.Drawing.Point(2, 3);
			this.lbFolder.Name = "lbFolder";
			this.lbFolder.Padding = new System.Windows.Forms.Padding(10);
			this.lbFolder.Size = new System.Drawing.Size(102, 35);
			this.lbFolder.TabIndex = 0;
			this.lbFolder.Text = "选中文件夹";
			this.lbFolder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.tbFolder.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tbFolder.Font = new System.Drawing.Font("宋体", 14f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tbFolder.Location = new System.Drawing.Point(5, 5);
			this.tbFolder.Name = "tbFolder";
			this.tbFolder.ReadOnly = true;
			this.tbFolder.Size = new System.Drawing.Size(652, 34);
			this.tbFolder.TabIndex = 1;
			this.BtnSelectFolder.Dock = System.Windows.Forms.DockStyle.Right;
			this.BtnSelectFolder.Location = new System.Drawing.Point(657, 5);
			this.BtnSelectFolder.Name = "BtnSelectFolder";
			this.BtnSelectFolder.Size = new System.Drawing.Size(138, 35);
			this.BtnSelectFolder.TabIndex = 2;
			this.BtnSelectFolder.Text = "选择文件夹……";
			this.BtnSelectFolder.UseVisualStyleBackColor = true;
			this.BtnSelectFolder.Click += new System.EventHandler(BtnSelectFolder_Click);
			this.panel1.Controls.Add(this.tbFolder);
			this.panel1.Controls.Add(this.BtnSelectFolder);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Padding = new System.Windows.Forms.Padding(5);
			this.panel1.Size = new System.Drawing.Size(800, 45);
			this.panel1.TabIndex = 16;
			this.BtnStartFolder.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.BtnStartFolder.Location = new System.Drawing.Point(0, 550);
			this.BtnStartFolder.Margin = new System.Windows.Forms.Padding(10);
			this.BtnStartFolder.Name = "BtnStartFolder";
			this.BtnStartFolder.Padding = new System.Windows.Forms.Padding(10);
			this.BtnStartFolder.Size = new System.Drawing.Size(800, 50);
			this.BtnStartFolder.TabIndex = 15;
			this.BtnStartFolder.Text = "开始添加";
			this.BtnStartFolder.UseVisualStyleBackColor = true;
			this.BtnStartFolder.Click += new System.EventHandler(BtnStartFolder_Click);
			this.RtbFolder.Location = new System.Drawing.Point(0, 109);
			this.RtbFolder.Name = "RtbFolder";
			this.RtbFolder.Size = new System.Drawing.Size(782, 428);
			this.RtbFolder.TabIndex = 13;
			this.RtbFolder.Text = "";
			base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 15f);
			base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			base.Controls.Add(this.panel2);
			base.Controls.Add(this.panel1);
			base.Controls.Add(this.BtnStartFolder);
			base.Controls.Add(this.RtbFolder);
			base.Name = "Uc_NewFanWen";
			base.Size = new System.Drawing.Size(800, 600);
			base.Load += new System.EventHandler(Uc_NewFanWen_Load);
			this.panel2.ResumeLayout(false);
			this.panel2.PerformLayout();
			this.panel1.ResumeLayout(false);
			this.panel1.PerformLayout();
			base.ResumeLayout(false);
		}
	}
}
