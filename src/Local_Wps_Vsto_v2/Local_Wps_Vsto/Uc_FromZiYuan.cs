using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Local_Wps_Vsto
{
	public class Uc_FromZiYuan : UserControl
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

		public Uc_FromZiYuan()
		{
			InitializeComponent();
		}

		private void Uc_FromZiYuan_Load(object sender, EventArgs e)
		{
			BindBackgroundWork();
		}

		private void BtnSelectFile_Click(object sender, EventArgs e)
		{
			try
			{
				OpenFileDialog openFileDialog = new OpenFileDialog();
				openFileDialog.Filter = "公文高手资源包(*.zydb)|*.zydb";
				if (openFileDialog.ShowDialog() == DialogResult.OK)
				{
					tbTitle.Text = openFileDialog.FileName;
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
				SQLiteHelper sQLiteHelper = new SQLiteHelper();
				strStatus = "导入资源包";
				strStatus = "\r\n开始插入素材……";
				backgroundWorker.ReportProgress(1);
				StreamReader streamReader = new StreamReader(strTitle, Encoding.Default);
				bool flag = false;
				long num = 0L;
				long num2 = 0L;
				string zippedString;
				while ((zippedString = streamReader.ReadLine()) != null)
				{
					zippedString = ZipHelper.GZipDecompressString(zippedString);
					if (zippedString.Equals("====分割素材和范文===="))
					{
						strStatus = "\r\n素材插入完毕，共有" + num + "条素材";
						backgroundWorker.ReportProgress(1);
						flag = true;
					}
					else if (flag)
					{
						try
						{
							JObject jObject = JsonConvert.DeserializeObject(zippedString) as JObject;
							sQLiteHelper.InsertValues_WithC("wr_fanwen_local", "title,content,content_rtf,creater_id,create_time,md5", new string[6]
							{
								jObject["Title"].ToString(),
								jObject["Content"].ToString(),
								jObject["ContentRtf"].ToString(),
								jObject["CreaterId"].ToString(),
								DateTime.Now.ToString("s"),
								jObject["Md5"].ToString()
							});
							num2++;
							strStatus = "\r\n插入第" + num2 + "篇范文";
							backgroundWorker.ReportProgress(1);
						}
						catch (Exception)
						{
							break;
						}
					}
					else
					{
						try
						{
							JObject jObject2 = JsonConvert.DeserializeObject(zippedString) as JObject;
							LocalSuCaiUtil.InsertNewSuCai(jObject2["Content"].ToString(), 1L, jObject2["Sort_name"].ToString());
							num++;
							strStatus = "\r\n插入第" + num + "条素材";
							backgroundWorker.ReportProgress(1);
						}
						catch (Exception)
						{
						}
					}
				}
				streamReader.Close();
				sQLiteHelper.CloseConnection();
				strStatus = "\r\n资源包插入完毕\r\n\r\n共有" + num + "条素材，" + num2 + "篇范文。";
				backgroundWorker.ReportProgress(1);
			}
			catch (Exception)
			{
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
			this.BtnSelectFile.Text = "选择资源包……";
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
			this.BtnStart.Text = "开始添加";
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
			this.label1.Size = new System.Drawing.Size(349, 20);
			this.label1.TabIndex = 18;
			this.label1.Text = "导入资源包需要较长时间，请耐心等待";
			base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 15f);
			base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			base.Controls.Add(this.label1);
			base.Controls.Add(this.RtbFile);
			base.Controls.Add(this.BtnStart);
			base.Controls.Add(this.tbTitle);
			base.Controls.Add(this.BtnSelectFile);
			base.Name = "Uc_FromZiYuan";
			base.Size = new System.Drawing.Size(600, 500);
			base.Load += new System.EventHandler(Uc_FromZiYuan_Load);
			base.ResumeLayout(false);
			base.PerformLayout();
		}
	}
}
