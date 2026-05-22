using System;
using System.ComponentModel;
using System.Data.SqlServerCe;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace Local_Wps_Vsto
{
	public class Uc_Back : UserControl
	{
		private BackgroundWorker BgWorker_Upload = new BackgroundWorker();

		private string strStatus = "";

		private string strTitle;

		private bool boolInsert = true;

		private IContainer components;

		private TextBox tbTitle;

		private Button BtnSelectFile;

		private Button BtnStart;

		private RichTextBox RtbFile;

		private Label label1;

		public Uc_Back()
		{
			InitializeComponent();
		}

		private void Uc_Back_Load(object sender, EventArgs e)
		{
			BindBackgroundWork();
		}

		private void BtnSelectFile_Click(object sender, EventArgs e)
		{
			try
			{
				FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
				if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
				{
					tbTitle.Text = folderBrowserDialog.SelectedPath;
				}
			}
			catch (Exception)
			{
			}
		}

		private void BtnStart_Click(object sender, EventArgs e)
		{
			BgWorker_Upload.RunWorkerAsync();
			tbTitle.Enabled = false;
			BtnStart.Enabled = false;
		}

		private void BindBackgroundWork()
		{
			BgWorker_Upload.WorkerReportsProgress = true;
			BgWorker_Upload.WorkerSupportsCancellation = true;
			BgWorker_Upload.DoWork += DoWork_Handler_Upload;
			BgWorker_Upload.ProgressChanged += ProgressChanged_Handler_Upload;
			BgWorker_Upload.RunWorkerCompleted += RunWorkerCompleted_Handler_Upload;
		}

		private void DoWork_Handler_Upload(object sender, DoWorkEventArgs args)
		{
			BackgroundWorker backgroundWorker = sender as BackgroundWorker;
			try
			{
				string text = DateTime.Now.ToLongDateString().ToString().Replace("/", "-");
				string path = strTitle + "\\公文助手单机版备份(" + text + ").bfdb";
				LocalZiYuan localZiYuan = new LocalZiYuan();
				StreamWriter streamWriter = new StreamWriter(path, true, Encoding.Default);
				SQLiteHelper sQLiteHelper = new SQLiteHelper();
				string queryString = "SELECT  sc_id,content,creater_id,create_time,sort_name,total_like  FROM [wr_sucai_local]  ";
				SqlCeDataReader sqlCeDataReader = sQLiteHelper.ExecuteQuery(queryString);
				strStatus = "开始读取素材……";
				backgroundWorker.ReportProgress(1);
				while (sqlCeDataReader.Read())
				{
					SucaiItem sucaiItem = new SucaiItem();
					sucaiItem.Sc_id = int.Parse(sqlCeDataReader["sc_id"].ToString());
					sucaiItem.Content = sqlCeDataReader["content"].ToString();
					sucaiItem.Creater_id = int.Parse(sqlCeDataReader["creater_id"].ToString());
					sucaiItem.Total_like = int.Parse(sqlCeDataReader["total_like"].ToString());
					sucaiItem.Sort_name = sqlCeDataReader["sort_name"].ToString();
					localZiYuan.sucaiItems.Add(sucaiItem);
					string rawString = JsonConvert.SerializeObject(sucaiItem);
					rawString = ZipHelper.GZipCompressString(rawString);
					streamWriter.WriteLine(rawString);
				}
				sQLiteHelper.CloseConnection();
				strStatus = strStatus + "\r\n素材备份完毕，共有" + localZiYuan.sucaiItems.Count + "条";
				backgroundWorker.ReportProgress(1);
				string value = ZipHelper.GZipCompressString("====分割素材和范文====");
				streamWriter.WriteLine(value);
				sQLiteHelper = new SQLiteHelper();
				queryString = "SELECT  fw_id,title,content,content_rtf,creater_id,md5   FROM [wr_fanwen_local]  ";
				sqlCeDataReader = sQLiteHelper.ExecuteQuery(queryString);
				strStatus += "\r\n开始备份范文……";
				backgroundWorker.ReportProgress(1);
				int num = 0;
				while (sqlCeDataReader.Read())
				{
					num++;
					FanWenItem fanWenItem = new FanWenItem();
					fanWenItem.FwId = int.Parse(sqlCeDataReader["fw_id"].ToString());
					fanWenItem.Title = sqlCeDataReader["title"].ToString();
					fanWenItem.Content = sqlCeDataReader["content"].ToString();
					fanWenItem.ContentRtf = sqlCeDataReader["content_rtf"].ToString();
					fanWenItem.CreaterId = int.Parse(sqlCeDataReader["creater_id"].ToString());
					fanWenItem.Md5 = sqlCeDataReader["md5"].ToString();
					fanWenItem.Createtime = "";
					localZiYuan.fanWenItems.Add(fanWenItem);
					string rawString2 = JsonConvert.SerializeObject(fanWenItem);
					rawString2 = ZipHelper.GZipCompressString(rawString2);
					streamWriter.WriteLine(rawString2);
				}
				sQLiteHelper.CloseConnection();
				strStatus = strStatus + "\r\n范文备份完毕，共有" + localZiYuan.fanWenItems.Count + "条";
				backgroundWorker.ReportProgress(1);
				JsonConvert.SerializeObject(localZiYuan);
				streamWriter.Close();
				strStatus += "\r\n备份成功";
				backgroundWorker.ReportProgress(2);
			}
			catch (Exception)
			{
				strStatus += "\r\n备份失败，请稍后再试！";
				backgroundWorker.ReportProgress(3);
			}
		}

		private void RunWorkerCompleted_Handler_Upload(object sender, RunWorkerCompletedEventArgs args)
		{
		}

		private void ProgressChanged_Handler_Upload(object sender, ProgressChangedEventArgs args)
		{
			try
			{
				RtbFile.Text = strStatus;
			}
			catch (Exception)
			{
			}
		}

		private void tbTitle_TextChanged(object sender, EventArgs e)
		{
			if (tbTitle.Text.Length > 0)
			{
				strTitle = tbTitle.Text.ToString();
				BtnStart.Enabled = true;
			}
			else
			{
				BtnStart.Enabled = false;
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
			this.tbTitle = new System.Windows.Forms.TextBox();
			this.BtnSelectFile = new System.Windows.Forms.Button();
			this.BtnStart = new System.Windows.Forms.Button();
			this.RtbFile = new System.Windows.Forms.RichTextBox();
			this.label1 = new System.Windows.Forms.Label();
			base.SuspendLayout();
			this.tbTitle.Enabled = false;
			this.tbTitle.Font = new System.Drawing.Font("宋体", 14f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tbTitle.Location = new System.Drawing.Point(3, 88);
			this.tbTitle.Name = "tbTitle";
			this.tbTitle.ReadOnly = true;
			this.tbTitle.Size = new System.Drawing.Size(478, 34);
			this.tbTitle.TabIndex = 3;
			this.tbTitle.TextChanged += new System.EventHandler(tbTitle_TextChanged);
			this.BtnSelectFile.Location = new System.Drawing.Point(476, 88);
			this.BtnSelectFile.Name = "BtnSelectFile";
			this.BtnSelectFile.Size = new System.Drawing.Size(121, 35);
			this.BtnSelectFile.TabIndex = 4;
			this.BtnSelectFile.Text = "选择位置……";
			this.BtnSelectFile.UseVisualStyleBackColor = true;
			this.BtnSelectFile.Click += new System.EventHandler(BtnSelectFile_Click);
			this.BtnStart.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.BtnStart.Enabled = false;
			this.BtnStart.Location = new System.Drawing.Point(0, 450);
			this.BtnStart.Margin = new System.Windows.Forms.Padding(10);
			this.BtnStart.Name = "BtnStart";
			this.BtnStart.Padding = new System.Windows.Forms.Padding(10);
			this.BtnStart.Size = new System.Drawing.Size(600, 50);
			this.BtnStart.TabIndex = 17;
			this.BtnStart.Text = "开始备份";
			this.BtnStart.UseVisualStyleBackColor = true;
			this.BtnStart.Click += new System.EventHandler(BtnStart_Click);
			this.RtbFile.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.RtbFile.Location = new System.Drawing.Point(0, 145);
			this.RtbFile.Name = "RtbFile";
			this.RtbFile.ReadOnly = true;
			this.RtbFile.Size = new System.Drawing.Size(600, 305);
			this.RtbFile.TabIndex = 16;
			this.RtbFile.Text = "";
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.label1.Location = new System.Drawing.Point(141, 36);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(329, 20);
			this.label1.TabIndex = 18;
			this.label1.Text = "备份资源需要较长时间，请耐心等待";
			base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 15f);
			base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			base.Controls.Add(this.label1);
			base.Controls.Add(this.RtbFile);
			base.Controls.Add(this.BtnStart);
			base.Controls.Add(this.tbTitle);
			base.Controls.Add(this.BtnSelectFile);
			base.Name = "Uc_Back";
			base.Size = new System.Drawing.Size(600, 500);
			base.Load += new System.EventHandler(Uc_Back_Load);
			base.ResumeLayout(false);
			base.PerformLayout();
		}
	}
}
