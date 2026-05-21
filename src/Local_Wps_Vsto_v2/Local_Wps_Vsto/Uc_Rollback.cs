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
	public class Uc_Rollback : UserControl
	{
		private BackgroundWorker BgWorker_Upload = new BackgroundWorker();

		private string strStatus = "";

		private string strTitle;

		private bool boolInsert = true;

		private IContainer components;

		private Label label1;

		private RichTextBox RtbFile;

		private Button BtnStart;

		private TextBox tbTitle;

		private Button BtnSelectFile;

		private RadioButton rbInsert;

		private RadioButton radioButton2;

		public Uc_Rollback()
		{
			InitializeComponent();
		}

		private void Uc_Rollback_Load(object sender, EventArgs e)
		{
			BindBackgroundWork();
		}

		private void BtnSelectFile_Click(object sender, EventArgs e)
		{
			try
			{
				OpenFileDialog openFileDialog = new OpenFileDialog();
				openFileDialog.Filter = "单机版备份资源(*.bfdb)|*.bfdb";
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
			BindBackgroundWork();
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
				if (boolInsert)
				{
					strStatus = "合并素材库";
					backgroundWorker.ReportProgress(1);
				}
				else
				{
					strStatus = "正在清除当前素材库……";
					backgroundWorker.ReportProgress(1);
					string queryString = "delete  FROM [wr_sucai_local]  ";
					sQLiteHelper.ExecuteQuery(queryString);
					queryString = "delete  FROM [wr_fanwen_local]  ";
					sQLiteHelper.ExecuteQuery(queryString);
				}
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

		private void radioButton1_CheckedChanged(object sender, EventArgs e)
		{
			if (rbInsert.Checked)
			{
				boolInsert = true;
			}
			else
			{
				boolInsert = false;
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
			this.label1 = new System.Windows.Forms.Label();
			this.RtbFile = new System.Windows.Forms.RichTextBox();
			this.BtnStart = new System.Windows.Forms.Button();
			this.tbTitle = new System.Windows.Forms.TextBox();
			this.BtnSelectFile = new System.Windows.Forms.Button();
			this.rbInsert = new System.Windows.Forms.RadioButton();
			this.radioButton2 = new System.Windows.Forms.RadioButton();
			base.SuspendLayout();
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.label1.Location = new System.Drawing.Point(141, 18);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(329, 20);
			this.label1.TabIndex = 23;
			this.label1.Text = "恢复备份需要较长时间，请耐心等待";
			this.RtbFile.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.RtbFile.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.RtbFile.Location = new System.Drawing.Point(0, 163);
			this.RtbFile.Name = "RtbFile";
			this.RtbFile.ReadOnly = true;
			this.RtbFile.Size = new System.Drawing.Size(600, 287);
			this.RtbFile.TabIndex = 21;
			this.RtbFile.Text = "";
			this.BtnStart.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.BtnStart.Enabled = false;
			this.BtnStart.Location = new System.Drawing.Point(0, 450);
			this.BtnStart.Margin = new System.Windows.Forms.Padding(10);
			this.BtnStart.Name = "BtnStart";
			this.BtnStart.Padding = new System.Windows.Forms.Padding(10);
			this.BtnStart.Size = new System.Drawing.Size(600, 50);
			this.BtnStart.TabIndex = 22;
			this.BtnStart.Text = "开始恢复备份";
			this.BtnStart.UseVisualStyleBackColor = true;
			this.BtnStart.Click += new System.EventHandler(BtnStart_Click);
			this.tbTitle.Enabled = false;
			this.tbTitle.Font = new System.Drawing.Font("宋体", 14f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tbTitle.Location = new System.Drawing.Point(3, 55);
			this.tbTitle.Name = "tbTitle";
			this.tbTitle.ReadOnly = true;
			this.tbTitle.Size = new System.Drawing.Size(478, 34);
			this.tbTitle.TabIndex = 19;
			this.tbTitle.TextChanged += new System.EventHandler(tbTitle_TextChanged);
			this.BtnSelectFile.Location = new System.Drawing.Point(487, 55);
			this.BtnSelectFile.Name = "BtnSelectFile";
			this.BtnSelectFile.Size = new System.Drawing.Size(98, 35);
			this.BtnSelectFile.TabIndex = 20;
			this.BtnSelectFile.Text = "选择备份…";
			this.BtnSelectFile.UseVisualStyleBackColor = true;
			this.BtnSelectFile.Click += new System.EventHandler(BtnSelectFile_Click);
			this.rbInsert.AutoSize = true;
			this.rbInsert.Checked = true;
			this.rbInsert.Location = new System.Drawing.Point(57, 110);
			this.rbInsert.Name = "rbInsert";
			this.rbInsert.Size = new System.Drawing.Size(148, 19);
			this.rbInsert.TabIndex = 24;
			this.rbInsert.TabStop = true;
			this.rbInsert.Text = "与当前资源库合并";
			this.rbInsert.UseVisualStyleBackColor = true;
			this.rbInsert.CheckedChanged += new System.EventHandler(radioButton1_CheckedChanged);
			this.radioButton2.AutoSize = true;
			this.radioButton2.Location = new System.Drawing.Point(258, 110);
			this.radioButton2.Name = "radioButton2";
			this.radioButton2.Size = new System.Drawing.Size(133, 19);
			this.radioButton2.TabIndex = 24;
			this.radioButton2.Text = "覆盖当前资源库";
			this.radioButton2.UseVisualStyleBackColor = true;
			base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 15f);
			base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			base.Controls.Add(this.radioButton2);
			base.Controls.Add(this.rbInsert);
			base.Controls.Add(this.label1);
			base.Controls.Add(this.RtbFile);
			base.Controls.Add(this.BtnStart);
			base.Controls.Add(this.tbTitle);
			base.Controls.Add(this.BtnSelectFile);
			base.Name = "Uc_Rollback";
			base.Size = new System.Drawing.Size(600, 500);
			base.ResumeLayout(false);
			base.PerformLayout();
		}
	}
}
