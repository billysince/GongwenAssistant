using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;

namespace Local_Wps_Vsto
{
	public class Uc_NewSuCai : UserControl
	{
		private MyAddin _parentRibbon;

		private BackgroundWorker BgWorker_Upload = new BackgroundWorker();

		private BackgroundWorker BgWorker_Save = new BackgroundWorker();

		private string strContent = "";

		private string strSortName = "";

		private string strReturn;

		private bool IsLotOf;

		private string strShare = "yes";

		private string strException = "插入素材没有异常";

		private int totalInsert;

		private IContainer components;

		private TextBox TbContent;

		private TabControl tabTop;

		private TabPage tabPage1;

		private TabPage tabPage2;

		private TabPage tabPage3;

		private TabPage tabPage4;

		private TabPage tabPage5;

		private FileSystemWatcher fileSystemWatcher1;

		private Panel panel1;

		private Button BtnNewSucai;

		private CheckBox CheckBox_Lot;

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

		public Uc_NewSuCai()
		{
			InitializeComponent();
		}

		public Uc_NewSuCai(MyAddin ribbonDesign)
		{
			_parentRibbon = ribbonDesign;
			InitializeComponent();
		}

		private void Uc_NewSuCai_Load(object sender, EventArgs e)
		{
			try
			{
				Control.CheckForIllegalCrossThreadCalls = false;
				CheckBox_Lot.Checked = true;
				BtnNewSucai.Enabled = false;
				TbContent.Focus();
				BindBackgroundWork();
			}
			catch (Exception)
			{
			}
		}

		private void BindBackgroundWork()
		{
			try
			{
				BgWorker_Upload.WorkerReportsProgress = true;
				BgWorker_Upload.WorkerSupportsCancellation = true;
				BgWorker_Upload.DoWork += DoWork_Handler_Upload;
				BgWorker_Upload.RunWorkerCompleted += RunWorkerCompleted_Handler_Upload;
				BgWorker_Save.WorkerReportsProgress = false;
				BgWorker_Save.WorkerSupportsCancellation = true;
				BgWorker_Save.DoWork += DoWork_Handler_Save;
				BgWorker_Save.RunWorkerCompleted += RunWorkerCompleted_Handler_Save;
			}
			catch (Exception)
			{
			}
		}

		private void DoWork_Handler_Upload(object sender, DoWorkEventArgs args)
		{
			try
			{
				string address = UrlUtil.strUploadOneSucai();
				WebClient webClient = new WebClient();
				webClient.Credentials = CredentialCache.DefaultCredentials;
				string text = "content=" + strContent + "&user_id=" + 1L + "&sort_name=" + strSortName;
				text = text + "&mac=" + SecretUtil.GetMachineCode();
				text = text + "&share=" + strShare;
				webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
				byte[] bytes = Encoding.UTF8.GetBytes(text);
				byte[] bytes2 = webClient.UploadData(address, "POST", bytes);
				strReturn = Encoding.UTF8.GetString(bytes2);
			}
			catch (Exception)
			{
			}
		}

		private void RunWorkerCompleted_Handler_Upload(object sender, RunWorkerCompletedEventArgs args)
		{
		}

		private void DoWork_Handler_Save(object sender, DoWorkEventArgs args)
		{
			try
			{
				if (IsLotOf)
				{
					string[] source = strContent.Split('\r', '\n');
					source = source.Where((string s) => !string.IsNullOrEmpty(s)).ToArray();
					SQLiteHelper sQLiteHelper = new SQLiteHelper();
					foreach (string text in source)
					{
						sQLiteHelper.InsertValues_WithC("wr_sucai_local", "content,creater_id,sort_name,create_time", new string[4]
						{
							text,
							1L.ToString(),
							strSortName,
							DateTime.Now.ToLocalTime().ToString("s")
						});
					}
				}
				else
				{
					LocalSuCaiUtil.InsertNewSuCai(strContent, 1L, strSortName);
				}
			}
			catch (Exception ex)
			{
				strException = ex.ToString();
			}
		}

		private void RunWorkerCompleted_Handler_Save(object sender, RunWorkerCompletedEventArgs args)
		{
			try
			{
				totalInsert++;
				TbContent.Text = "";
				TbContent.Enabled = true;
				BtnNewSucai.Enabled = false;
				MessageTipsUtil.Show("添加素材成功");
				if (CommonConfig.IS_DEBUG)
				{
					MessageBox.Show(strException);
				}
			}
			catch (Exception)
			{
			}
		}

		private void BtnNewSucai_Click(object sender, EventArgs e)
		{
			try
			{
				strShare = "no";
				if (TbContent.Text.Length > 0)
				{
					int num = ((tabTop.SelectedIndex > 0) ? tabTop.SelectedIndex : 0);
					strContent = TbContent.Text.ToString();
					if (CheckBox_Lot.Checked)
					{
						IsLotOf = true;
						strShare = "yes";
					}
					else
					{
						IsLotOf = false;
						strShare = "no";
					}
					strSortName = CommonConfig.SortName[num + 1];
					BgWorker_Save = new BackgroundWorker();
					BgWorker_Save.WorkerReportsProgress = true;
					BgWorker_Save.WorkerSupportsCancellation = true;
					BgWorker_Save.DoWork += DoWork_Handler_Save;
					BgWorker_Save.RunWorkerCompleted += RunWorkerCompleted_Handler_Save;
					BgWorker_Save.RunWorkerAsync();
					TbContent.Enabled = false;
					BtnNewSucai.Enabled = false;
				}
			}
			catch (Exception)
			{
			}
		}

		private void tabTop_SelectedIndexChanged(object sender, EventArgs e)
		{
			CheckBox_Lot.Checked = true;
			if (tabTop.SelectedIndex == 1)
			{
				CheckBox_Lot.Checked = false;
			}
		}

		private void TbContent_TextChanged(object sender, EventArgs e)
		{
			if (TbContent.Text.Length > 0)
			{
				BtnNewSucai.Enabled = true;
			}
			else
			{
				BtnNewSucai.Enabled = false;
			}
		}

		private void Fm_NewSuCai_FormClosed(object sender, FormClosedEventArgs e)
		{
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
			this.TbContent = new System.Windows.Forms.TextBox();
			this.tabTop = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.tabPage4 = new System.Windows.Forms.TabPage();
			this.tabPage5 = new System.Windows.Forms.TabPage();
			this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
			this.panel1 = new System.Windows.Forms.Panel();
			this.BtnNewSucai = new System.Windows.Forms.Button();
			this.CheckBox_Lot = new System.Windows.Forms.CheckBox();
			this.tabTop.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)this.fileSystemWatcher1).BeginInit();
			this.panel1.SuspendLayout();
			base.SuspendLayout();
			this.TbContent.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.TbContent.Dock = System.Windows.Forms.DockStyle.Top;
			this.TbContent.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.TbContent.Location = new System.Drawing.Point(0, 39);
			this.TbContent.Multiline = true;
			this.TbContent.Name = "TbContent";
			this.TbContent.Size = new System.Drawing.Size(798, 466);
			this.TbContent.TabIndex = 7;
			this.TbContent.TextChanged += new System.EventHandler(TbContent_TextChanged);
			this.tabTop.Controls.Add(this.tabPage1);
			this.tabTop.Controls.Add(this.tabPage2);
			this.tabTop.Controls.Add(this.tabPage3);
			this.tabTop.Controls.Add(this.tabPage4);
			this.tabTop.Controls.Add(this.tabPage5);
			this.tabTop.Dock = System.Windows.Forms.DockStyle.Top;
			this.tabTop.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tabTop.Location = new System.Drawing.Point(0, 0);
			this.tabTop.Name = "tabTop";
			this.tabTop.Padding = new System.Drawing.Point(10, 8);
			this.tabTop.SelectedIndex = 0;
			this.tabTop.Size = new System.Drawing.Size(798, 39);
			this.tabTop.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
			this.tabTop.TabIndex = 6;
			this.tabTop.SelectedIndexChanged += new System.EventHandler(tabTop_SelectedIndexChanged);
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
			this.fileSystemWatcher1.EnableRaisingEvents = true;
			this.fileSystemWatcher1.SynchronizingObject = this;
			this.panel1.Controls.Add(this.BtnNewSucai);
			this.panel1.Controls.Add(this.CheckBox_Lot);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.panel1.Location = new System.Drawing.Point(0, 492);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(798, 106);
			this.panel1.TabIndex = 8;
			this.BtnNewSucai.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnNewSucai.Location = new System.Drawing.Point(295, 54);
			this.BtnNewSucai.Margin = new System.Windows.Forms.Padding(10);
			this.BtnNewSucai.Name = "BtnNewSucai";
			this.BtnNewSucai.Padding = new System.Windows.Forms.Padding(1);
			this.BtnNewSucai.Size = new System.Drawing.Size(149, 35);
			this.BtnNewSucai.TabIndex = 2;
			this.BtnNewSucai.Text = "添加素材";
			this.BtnNewSucai.UseVisualStyleBackColor = true;
			this.BtnNewSucai.Click += new System.EventHandler(BtnNewSucai_Click);
			this.CheckBox_Lot.AutoSize = true;
			this.CheckBox_Lot.Location = new System.Drawing.Point(30, 30);
			this.CheckBox_Lot.Name = "CheckBox_Lot";
			this.CheckBox_Lot.Size = new System.Drawing.Size(239, 19);
			this.CheckBox_Lot.TabIndex = 1;
			this.CheckBox_Lot.Text = "批量添加，每段作为一个素材！";
			this.CheckBox_Lot.UseVisualStyleBackColor = true;
			base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 15f);
			base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			base.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			base.Controls.Add(this.TbContent);
			base.Controls.Add(this.tabTop);
			base.Controls.Add(this.panel1);
			base.Name = "Uc_NewSuCai";
			base.Size = new System.Drawing.Size(798, 598);
			base.Load += new System.EventHandler(Uc_NewSuCai_Load);
			this.tabTop.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)this.fileSystemWatcher1).EndInit();
			this.panel1.ResumeLayout(false);
			this.panel1.PerformLayout();
			base.ResumeLayout(false);
			base.PerformLayout();
		}
	}
}
