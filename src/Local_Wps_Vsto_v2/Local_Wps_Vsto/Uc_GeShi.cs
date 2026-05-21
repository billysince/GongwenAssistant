using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel;
using Word;

namespace Local_Wps_Vsto
{
	public class Uc_GeShi : UserControl
	{
		private GroupBox groupBox3;

		private TableLayoutPanel tableLayoutPanel3;

		private Button BtnXY_Tg;

		private Button BtnXY_Ml;

		private Button BtnXY_Cz;

		private Button BtnXY_Zm;

		private Button BtnXY_Gg;

		private Button BtnXY_Han;

		private Button BtnXY_Tz;

		private Button BtnXY_Hyjy;

		private Button BtnSX_4;

		private Button BtnSX_3;

		private Button BtnSX_2;

		private Button BtnSX_1;

		private TableLayoutPanel tableLayoutPanel2;

		private GroupBox groupBox2;

		private Button BtnXX_4;

		private Button BtnXX_3;

		private Button BtnXX_2;

		public Button BtnXX_1;

		private TableLayoutPanel tableLayoutPanel1;

		private Button BtnPT_A5_1;

		private Button BtnPT_A4_3;

		private Button BtnPT_A4_1;

		private TableLayoutPanel tableLayoutPanel5;

		private Button BtnPT_A4_2;

		private GroupBox groupBox5;

		private Button BtnHW_Zxpw;

		private Button BtnHW_Jczc;

		private Button BtnHW_Hwsc;

		private TableLayoutPanel tableLayoutPanel4;

		private Button BtnHW_Hyzc;

		private GroupBox groupBox4;

		private GroupBox groupBox1;

		private Word.Application wordApp;

		private Excel.Application excelApp = (Excel.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00024500-0000-0000-C000-000000000046")));

		private string objExcelFullName;

		private string strFileName = "";

		public Uc_GeShi()
		{
			InitializeComponent();
		}

		private void Uc_GeShi_Load(object sender, EventArgs e)
		{
		}

		private void Excel_DoWork_Handler_Web(object sender, DoWorkEventArgs args)
		{
			BackgroundWorker backgroundWorker = sender as BackgroundWorker;
			try
			{
				if (excelApp == null)
				{
					excelApp = (Excel.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00024500-0000-0000-C000-000000000046")));
				}
				Path.GetFileName(strFileName);
				object template = CommonConfig.strBaseFolder + strFileName;
				excelApp.Workbooks.Add(template);
				if (backgroundWorker.CancellationPending)
				{
					args.Cancel = true;
				}
			}
			catch (Exception)
			{
			}
		}

		private void Excel_RunWorkerCompleted_Handler_Web(object sender, RunWorkerCompletedEventArgs args)
		{
			if (!args.Cancelled)
			{
				try
				{
					excelApp.Visible = true;
				}
				catch (Exception)
				{
					MessageTipsUtil.Show("打开Excel文档出错");
				}
			}
			MsgTips.Hide();
		}

		private void Word_DoWork_Handler_Web(object sender, DoWorkEventArgs args)
		{
			BackgroundWorker backgroundWorker = sender as BackgroundWorker;
			try
			{
				wordApp = MyAddin.wordApp;
				object Template = CommonConfig.strBaseFolder + strFileName;
				Documents documents = wordApp.Documents;
				object NewTemplate = Type.Missing;
				object DocumentType = Type.Missing;
				object Visible = Type.Missing;
				documents.Add(ref Template, ref NewTemplate, ref DocumentType, ref Visible);
				if (backgroundWorker.CancellationPending)
				{
					args.Cancel = true;
				}
			}
			catch (Exception)
			{
			}
		}

		private void Word_RunWorkerCompleted_Handler_Web(object sender, RunWorkerCompletedEventArgs args)
		{
			if (!args.Cancelled)
			{
				try
				{
					wordApp.Visible = true;
					MsgTips.Hide();
				}
				catch (Exception)
				{
					MessageTipsUtil.Show("打开Word文档出错");
				}
			}
			MsgTips.Hide();
		}

		private void BtnXX_1_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\xx_1.docx");
		}

		private void OpenNewWordTemplate(string _strFileName)
		{
			strFileName = _strFileName;
			if (UserUtil.HasLogin())
			{
				if (UserUtil.IsVip())
				{
					try
					{
						string text = CommonConfig.strBaseFolder + strFileName;
						Documents documents = MyAddin.wordApp.Documents;
						object Template = text;
						object NewTemplate = Type.Missing;
						object DocumentType = Type.Missing;
						object Visible = Type.Missing;
						documents.Add(ref Template, ref NewTemplate, ref DocumentType, ref Visible);
						try
						{
							MyAddin.geshiPanel.Visible = false;
							return;
						}
						catch (Exception)
						{
							return;
						}
					}
					catch (Exception)
					{
						MessageTipsUtil.Show("打开Word文档出错");
						return;
					}
				}
				MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
			}
			else
			{
				MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
			}
		}

		private void OpenNewExcelTemplate(string _strFileName, string _NewFileName)
		{
			strFileName = _strFileName;
			if (UserUtil.HasLogin())
			{
				if (UserUtil.IsVip())
				{
					try
					{
						string sourceFileName = CommonConfig.strBaseFolder + strFileName;
						string text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + _NewFileName + "-" + DateTime.Now.ToLongDateString().ToString();
						text = ((!File.Exists(text + ".xls")) ? (text + ".xls") : (text + "-" + new Random().Next(10, 99) + ".xls"));
						File.Copy(sourceFileName, text, true);
						Process.Start(text);
						try
						{
							MyAddin.geshiPanel.Visible = false;
							return;
						}
						catch (Exception)
						{
							return;
						}
					}
					catch (Exception)
					{
						MessageTipsUtil.Show("打开Excel表格出错");
						return;
					}
				}
				MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
			}
			else
			{
				MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
			}
		}

		private void BtnXX_2_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\xx_2.docx");
		}

		private void BtnXX_3_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\xx_3.docx");
		}

		private void BtnXX_4_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\xx_4.docx");
		}

		private void BtnSX_1_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\sx_1.docx");
		}

		private void BtnSX_2_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\sx_2.docx");
		}

		private void BtnSX_3_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\sx_3.docx");
		}

		private void BtnSX_4_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\sx_4.docx");
		}

		private void BtnXY_Hyjy_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\hyjy_1.docx");
		}

		private void BtnXY_Tz_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\tz_1.docx");
		}

		private void BtnXY_Han_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\han_1.docx");
		}

		private void BtnXY_Gg_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\gg_1.docx");
		}

		private void BtnXY_Tg_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\tg_1.docx");
		}

		private void BtnXY_Ml_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\ml_1.docx");
		}

		private void BtnXY_Cz_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\cz_1.docx");
		}

		private void BtnXY_Zm_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\zm_1.docx");
		}

		private void BtnHW_Hwsc_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewExcelTemplate("\\template\\hylcqd.xls", "会议流程清单");
		}

		private void BtnHW_Hyzc_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewExcelTemplate("\\template\\hyzc.xls", "会议座次");
		}

		private void BtnHW_Jczc_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewExcelTemplate("\\template\\jczc.xls", "就餐座次");
		}

		private void BtnHW_Zxpw_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewExcelTemplate("\\template\\zxpw.xls", "照相排位");
		}

		private void BtnPT_A4_1_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\A4_1.docx");
		}

		private void BtnPT_A4_2_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\A4_2.docx");
		}

		private void BtnPT_A4_3_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\A4_3.docx");
		}

		private void BtnPT_A5_1_Click(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\A5_3.docx");
		}

		private void BtnXY_Tz_Click_1(object sender, EventArgs e)
		{
		}

		private void BtnSX_1_Click_1(object sender, EventArgs e)
		{
			MyAddin.OpenNewWordTemplate("\\template\\xx_1.docx");
		}

		private void InitializeComponent()
		{
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
			this.BtnXY_Tg = new System.Windows.Forms.Button();
			this.BtnXY_Ml = new System.Windows.Forms.Button();
			this.BtnXY_Cz = new System.Windows.Forms.Button();
			this.BtnXY_Zm = new System.Windows.Forms.Button();
			this.BtnXY_Gg = new System.Windows.Forms.Button();
			this.BtnXY_Han = new System.Windows.Forms.Button();
			this.BtnXY_Tz = new System.Windows.Forms.Button();
			this.BtnXY_Hyjy = new System.Windows.Forms.Button();
			this.BtnSX_4 = new System.Windows.Forms.Button();
			this.BtnSX_3 = new System.Windows.Forms.Button();
			this.BtnSX_2 = new System.Windows.Forms.Button();
			this.BtnSX_1 = new System.Windows.Forms.Button();
			this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.BtnXX_4 = new System.Windows.Forms.Button();
			this.BtnXX_3 = new System.Windows.Forms.Button();
			this.BtnXX_2 = new System.Windows.Forms.Button();
			this.BtnXX_1 = new System.Windows.Forms.Button();
			this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
			this.BtnPT_A5_1 = new System.Windows.Forms.Button();
			this.BtnPT_A4_3 = new System.Windows.Forms.Button();
			this.BtnPT_A4_1 = new System.Windows.Forms.Button();
			this.tableLayoutPanel5 = new System.Windows.Forms.TableLayoutPanel();
			this.BtnPT_A4_2 = new System.Windows.Forms.Button();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.BtnHW_Zxpw = new System.Windows.Forms.Button();
			this.BtnHW_Jczc = new System.Windows.Forms.Button();
			this.BtnHW_Hwsc = new System.Windows.Forms.Button();
			this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
			this.BtnHW_Hyzc = new System.Windows.Forms.Button();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox3.SuspendLayout();
			this.tableLayoutPanel3.SuspendLayout();
			this.tableLayoutPanel2.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.tableLayoutPanel1.SuspendLayout();
			this.tableLayoutPanel5.SuspendLayout();
			this.groupBox5.SuspendLayout();
			this.tableLayoutPanel4.SuspendLayout();
			this.groupBox4.SuspendLayout();
			this.groupBox1.SuspendLayout();
			base.SuspendLayout();
			this.groupBox3.Controls.Add(this.tableLayoutPanel3);
			this.groupBox3.Dock = System.Windows.Forms.DockStyle.Top;
			this.groupBox3.Font = new System.Drawing.Font("宋体", 13f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.groupBox3.Location = new System.Drawing.Point(0, 440);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1000, 187);
			this.groupBox3.TabIndex = 12;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "常用公文";
			this.tableLayoutPanel3.ColumnCount = 4;
			this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel3.Controls.Add(this.BtnXY_Tg, 0, 1);
			this.tableLayoutPanel3.Controls.Add(this.BtnXY_Ml, 0, 1);
			this.tableLayoutPanel3.Controls.Add(this.BtnXY_Cz, 0, 1);
			this.tableLayoutPanel3.Controls.Add(this.BtnXY_Zm, 0, 1);
			this.tableLayoutPanel3.Controls.Add(this.BtnXY_Gg, 3, 0);
			this.tableLayoutPanel3.Controls.Add(this.BtnXY_Han, 2, 0);
			this.tableLayoutPanel3.Controls.Add(this.BtnXY_Tz, 1, 0);
			this.tableLayoutPanel3.Controls.Add(this.BtnXY_Hyjy, 0, 0);
			this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel3.Location = new System.Drawing.Point(3, 28);
			this.tableLayoutPanel3.Name = "tableLayoutPanel3";
			this.tableLayoutPanel3.RowCount = 2;
			this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
			this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
			this.tableLayoutPanel3.Size = new System.Drawing.Size(994, 156);
			this.tableLayoutPanel3.TabIndex = 2;
			this.BtnXY_Tg.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnXY_Tg.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnXY_Tg.Location = new System.Drawing.Point(20, 93);
			this.BtnXY_Tg.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnXY_Tg.Name = "BtnXY_Tg";
			this.BtnXY_Tg.Size = new System.Drawing.Size(208, 48);
			this.BtnXY_Tg.TabIndex = 7;
			this.BtnXY_Tg.Text = "通告";
			this.BtnXY_Tg.UseVisualStyleBackColor = true;
			this.BtnXY_Tg.Click += new System.EventHandler(BtnXY_Tg_Click);
			this.BtnXY_Ml.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnXY_Ml.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnXY_Ml.Location = new System.Drawing.Point(268, 93);
			this.BtnXY_Ml.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnXY_Ml.Name = "BtnXY_Ml";
			this.BtnXY_Ml.Size = new System.Drawing.Size(208, 48);
			this.BtnXY_Ml.TabIndex = 6;
			this.BtnXY_Ml.Text = "命令";
			this.BtnXY_Ml.UseVisualStyleBackColor = true;
			this.BtnXY_Ml.Click += new System.EventHandler(BtnXY_Ml_Click);
			this.BtnXY_Cz.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnXY_Cz.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnXY_Cz.Location = new System.Drawing.Point(516, 93);
			this.BtnXY_Cz.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnXY_Cz.Name = "BtnXY_Cz";
			this.BtnXY_Cz.Size = new System.Drawing.Size(208, 48);
			this.BtnXY_Cz.TabIndex = 5;
			this.BtnXY_Cz.Text = "传真";
			this.BtnXY_Cz.UseVisualStyleBackColor = true;
			this.BtnXY_Cz.Click += new System.EventHandler(BtnXY_Cz_Click);
			this.BtnXY_Zm.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnXY_Zm.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnXY_Zm.Location = new System.Drawing.Point(764, 93);
			this.BtnXY_Zm.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnXY_Zm.Name = "BtnXY_Zm";
			this.BtnXY_Zm.Size = new System.Drawing.Size(210, 48);
			this.BtnXY_Zm.TabIndex = 4;
			this.BtnXY_Zm.Text = "证明";
			this.BtnXY_Zm.UseVisualStyleBackColor = true;
			this.BtnXY_Zm.Click += new System.EventHandler(BtnXY_Zm_Click);
			this.BtnXY_Gg.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnXY_Gg.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnXY_Gg.Location = new System.Drawing.Point(764, 15);
			this.BtnXY_Gg.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnXY_Gg.Name = "BtnXY_Gg";
			this.BtnXY_Gg.Size = new System.Drawing.Size(210, 48);
			this.BtnXY_Gg.TabIndex = 3;
			this.BtnXY_Gg.Text = "公告";
			this.BtnXY_Gg.UseVisualStyleBackColor = true;
			this.BtnXY_Gg.Click += new System.EventHandler(BtnXY_Gg_Click);
			this.BtnXY_Han.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnXY_Han.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnXY_Han.Location = new System.Drawing.Point(516, 15);
			this.BtnXY_Han.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnXY_Han.Name = "BtnXY_Han";
			this.BtnXY_Han.Size = new System.Drawing.Size(208, 48);
			this.BtnXY_Han.TabIndex = 2;
			this.BtnXY_Han.Text = "函";
			this.BtnXY_Han.UseVisualStyleBackColor = true;
			this.BtnXY_Han.Click += new System.EventHandler(BtnXY_Han_Click);
			this.BtnXY_Tz.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnXY_Tz.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnXY_Tz.Location = new System.Drawing.Point(268, 15);
			this.BtnXY_Tz.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnXY_Tz.Name = "BtnXY_Tz";
			this.BtnXY_Tz.Size = new System.Drawing.Size(208, 48);
			this.BtnXY_Tz.TabIndex = 1;
			this.BtnXY_Tz.Text = "通知";
			this.BtnXY_Tz.UseVisualStyleBackColor = true;
			this.BtnXY_Tz.Click += new System.EventHandler(BtnXY_Tz_Click);
			this.BtnXY_Hyjy.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnXY_Hyjy.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnXY_Hyjy.Location = new System.Drawing.Point(20, 15);
			this.BtnXY_Hyjy.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnXY_Hyjy.Name = "BtnXY_Hyjy";
			this.BtnXY_Hyjy.Size = new System.Drawing.Size(208, 48);
			this.BtnXY_Hyjy.TabIndex = 0;
			this.BtnXY_Hyjy.Text = "会议纪要";
			this.BtnXY_Hyjy.UseVisualStyleBackColor = true;
			this.BtnXY_Hyjy.Click += new System.EventHandler(BtnXY_Hyjy_Click);
			this.BtnSX_4.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnSX_4.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnSX_4.Location = new System.Drawing.Point(764, 15);
			this.BtnSX_4.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnSX_4.Name = "BtnSX_4";
			this.BtnSX_4.Size = new System.Drawing.Size(210, 49);
			this.BtnSX_4.TabIndex = 3;
			this.BtnSX_4.Text = "4单位联合发文";
			this.BtnSX_4.UseVisualStyleBackColor = true;
			this.BtnSX_4.Click += new System.EventHandler(BtnSX_4_Click);
			this.BtnSX_3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnSX_3.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnSX_3.Location = new System.Drawing.Point(516, 15);
			this.BtnSX_3.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnSX_3.Name = "BtnSX_3";
			this.BtnSX_3.Size = new System.Drawing.Size(208, 49);
			this.BtnSX_3.TabIndex = 2;
			this.BtnSX_3.Text = "3单位联合发文";
			this.BtnSX_3.UseVisualStyleBackColor = true;
			this.BtnSX_3.Click += new System.EventHandler(BtnSX_3_Click);
			this.BtnSX_2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnSX_2.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnSX_2.Location = new System.Drawing.Point(268, 15);
			this.BtnSX_2.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnSX_2.Name = "BtnSX_2";
			this.BtnSX_2.Size = new System.Drawing.Size(208, 49);
			this.BtnSX_2.TabIndex = 1;
			this.BtnSX_2.Text = "2单位联合发文";
			this.BtnSX_2.UseVisualStyleBackColor = true;
			this.BtnSX_2.Click += new System.EventHandler(BtnSX_2_Click);
			this.BtnSX_1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnSX_1.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnSX_1.Location = new System.Drawing.Point(20, 15);
			this.BtnSX_1.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnSX_1.Name = "BtnSX_1";
			this.BtnSX_1.Size = new System.Drawing.Size(208, 49);
			this.BtnSX_1.TabIndex = 0;
			this.BtnSX_1.Text = "单一机关发文";
			this.BtnSX_1.UseVisualStyleBackColor = true;
			this.BtnSX_1.Click += new System.EventHandler(BtnSX_1_Click);
			this.tableLayoutPanel2.ColumnCount = 4;
			this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel2.Controls.Add(this.BtnSX_4, 3, 0);
			this.tableLayoutPanel2.Controls.Add(this.BtnSX_3, 2, 0);
			this.tableLayoutPanel2.Controls.Add(this.BtnSX_2, 1, 0);
			this.tableLayoutPanel2.Controls.Add(this.BtnSX_1, 0, 0);
			this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 28);
			this.tableLayoutPanel2.Name = "tableLayoutPanel2";
			this.tableLayoutPanel2.RowCount = 1;
			this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
			this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
			this.tableLayoutPanel2.Size = new System.Drawing.Size(994, 79);
			this.tableLayoutPanel2.TabIndex = 1;
			this.groupBox2.Controls.Add(this.tableLayoutPanel2);
			this.groupBox2.Dock = System.Windows.Forms.DockStyle.Top;
			this.groupBox2.Font = new System.Drawing.Font("宋体", 13f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.groupBox2.Location = new System.Drawing.Point(0, 330);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1000, 110);
			this.groupBox2.TabIndex = 11;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "上行文（红头正式文件）";
			this.BtnXX_4.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnXX_4.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnXX_4.Location = new System.Drawing.Point(764, 15);
			this.BtnXX_4.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnXX_4.Name = "BtnXX_4";
			this.BtnXX_4.Size = new System.Drawing.Size(210, 49);
			this.BtnXX_4.TabIndex = 3;
			this.BtnXX_4.Text = "4单位联合发文";
			this.BtnXX_4.UseVisualStyleBackColor = true;
			this.BtnXX_4.Click += new System.EventHandler(BtnXX_4_Click);
			this.BtnXX_3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnXX_3.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnXX_3.Location = new System.Drawing.Point(516, 15);
			this.BtnXX_3.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnXX_3.Name = "BtnXX_3";
			this.BtnXX_3.Size = new System.Drawing.Size(208, 49);
			this.BtnXX_3.TabIndex = 2;
			this.BtnXX_3.Text = "3单位联合发文";
			this.BtnXX_3.UseVisualStyleBackColor = true;
			this.BtnXX_3.Click += new System.EventHandler(BtnXX_3_Click);
			this.BtnXX_2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnXX_2.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnXX_2.Location = new System.Drawing.Point(268, 15);
			this.BtnXX_2.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnXX_2.Name = "BtnXX_2";
			this.BtnXX_2.Size = new System.Drawing.Size(208, 49);
			this.BtnXX_2.TabIndex = 1;
			this.BtnXX_2.Text = "2单位联合发文";
			this.BtnXX_2.UseVisualStyleBackColor = true;
			this.BtnXX_2.Click += new System.EventHandler(BtnXX_2_Click);
			this.BtnXX_1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnXX_1.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnXX_1.Location = new System.Drawing.Point(20, 15);
			this.BtnXX_1.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnXX_1.Name = "BtnXX_1";
			this.BtnXX_1.Size = new System.Drawing.Size(208, 49);
			this.BtnXX_1.TabIndex = 0;
			this.BtnXX_1.Text = "单一机关发文";
			this.BtnXX_1.UseVisualStyleBackColor = true;
			this.BtnXX_1.Click += new System.EventHandler(BtnXX_1_Click);
			this.tableLayoutPanel1.ColumnCount = 4;
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel1.Controls.Add(this.BtnXX_4, 3, 0);
			this.tableLayoutPanel1.Controls.Add(this.BtnXX_3, 2, 0);
			this.tableLayoutPanel1.Controls.Add(this.BtnXX_2, 1, 0);
			this.tableLayoutPanel1.Controls.Add(this.BtnXX_1, 0, 0);
			this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 28);
			this.tableLayoutPanel1.Name = "tableLayoutPanel1";
			this.tableLayoutPanel1.RowCount = 1;
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
			this.tableLayoutPanel1.Size = new System.Drawing.Size(994, 79);
			this.tableLayoutPanel1.TabIndex = 0;
			this.BtnPT_A5_1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnPT_A5_1.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnPT_A5_1.Location = new System.Drawing.Point(764, 15);
			this.BtnPT_A5_1.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnPT_A5_1.Name = "BtnPT_A5_1";
			this.BtnPT_A5_1.Size = new System.Drawing.Size(210, 49);
			this.BtnPT_A5_1.TabIndex = 3;
			this.BtnPT_A5_1.Text = "A5折页手册";
			this.BtnPT_A5_1.UseVisualStyleBackColor = true;
			this.BtnPT_A5_1.Click += new System.EventHandler(BtnPT_A5_1_Click);
			this.BtnPT_A4_3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnPT_A4_3.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnPT_A4_3.Location = new System.Drawing.Point(516, 15);
			this.BtnPT_A4_3.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnPT_A4_3.Name = "BtnPT_A4_3";
			this.BtnPT_A4_3.Size = new System.Drawing.Size(208, 49);
			this.BtnPT_A4_3.TabIndex = 2;
			this.BtnPT_A4_3.Text = "A4折页手册";
			this.BtnPT_A4_3.UseVisualStyleBackColor = true;
			this.BtnPT_A4_3.Click += new System.EventHandler(BtnPT_A4_3_Click);
			this.BtnPT_A4_1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnPT_A4_1.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnPT_A4_1.Location = new System.Drawing.Point(20, 15);
			this.BtnPT_A4_1.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnPT_A4_1.Name = "BtnPT_A4_1";
			this.BtnPT_A4_1.Size = new System.Drawing.Size(208, 49);
			this.BtnPT_A4_1.TabIndex = 0;
			this.BtnPT_A4_1.Text = "A4单面";
			this.BtnPT_A4_1.UseVisualStyleBackColor = true;
			this.BtnPT_A4_1.Click += new System.EventHandler(BtnPT_A4_1_Click);
			this.tableLayoutPanel5.ColumnCount = 4;
			this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel5.Controls.Add(this.BtnPT_A5_1, 3, 0);
			this.tableLayoutPanel5.Controls.Add(this.BtnPT_A4_3, 2, 0);
			this.tableLayoutPanel5.Controls.Add(this.BtnPT_A4_2, 1, 0);
			this.tableLayoutPanel5.Controls.Add(this.BtnPT_A4_1, 0, 0);
			this.tableLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel5.Location = new System.Drawing.Point(3, 28);
			this.tableLayoutPanel5.Name = "tableLayoutPanel5";
			this.tableLayoutPanel5.RowCount = 1;
			this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
			this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
			this.tableLayoutPanel5.Size = new System.Drawing.Size(994, 79);
			this.tableLayoutPanel5.TabIndex = 2;
			this.BtnPT_A4_2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnPT_A4_2.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnPT_A4_2.Location = new System.Drawing.Point(268, 15);
			this.BtnPT_A4_2.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnPT_A4_2.Name = "BtnPT_A4_2";
			this.BtnPT_A4_2.Size = new System.Drawing.Size(208, 49);
			this.BtnPT_A4_2.TabIndex = 1;
			this.BtnPT_A4_2.Text = "A4双面";
			this.BtnPT_A4_2.UseVisualStyleBackColor = true;
			this.BtnPT_A4_2.Click += new System.EventHandler(BtnPT_A4_2_Click);
			this.groupBox5.Controls.Add(this.tableLayoutPanel5);
			this.groupBox5.Dock = System.Windows.Forms.DockStyle.Top;
			this.groupBox5.Font = new System.Drawing.Font("宋体", 13f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.groupBox5.Location = new System.Drawing.Point(0, 220);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(1000, 110);
			this.groupBox5.TabIndex = 14;
			this.groupBox5.TabStop = false;
			this.groupBox5.Text = "普通排版";
			this.BtnHW_Zxpw.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnHW_Zxpw.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnHW_Zxpw.Location = new System.Drawing.Point(764, 15);
			this.BtnHW_Zxpw.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnHW_Zxpw.Name = "BtnHW_Zxpw";
			this.BtnHW_Zxpw.Size = new System.Drawing.Size(210, 49);
			this.BtnHW_Zxpw.TabIndex = 3;
			this.BtnHW_Zxpw.Text = "照相排位";
			this.BtnHW_Zxpw.UseVisualStyleBackColor = true;
			this.BtnHW_Zxpw.Click += new System.EventHandler(BtnHW_Zxpw_Click);
			this.BtnHW_Jczc.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnHW_Jczc.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnHW_Jczc.Location = new System.Drawing.Point(516, 15);
			this.BtnHW_Jczc.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnHW_Jczc.Name = "BtnHW_Jczc";
			this.BtnHW_Jczc.Size = new System.Drawing.Size(208, 49);
			this.BtnHW_Jczc.TabIndex = 2;
			this.BtnHW_Jczc.Text = "就餐座次";
			this.BtnHW_Jczc.UseVisualStyleBackColor = true;
			this.BtnHW_Jczc.Click += new System.EventHandler(BtnHW_Jczc_Click);
			this.BtnHW_Hwsc.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnHW_Hwsc.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnHW_Hwsc.Location = new System.Drawing.Point(20, 15);
			this.BtnHW_Hwsc.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnHW_Hwsc.Name = "BtnHW_Hwsc";
			this.BtnHW_Hwsc.Size = new System.Drawing.Size(208, 49);
			this.BtnHW_Hwsc.TabIndex = 0;
			this.BtnHW_Hwsc.Text = "会务流程清单";
			this.BtnHW_Hwsc.UseVisualStyleBackColor = true;
			this.BtnHW_Hwsc.Click += new System.EventHandler(BtnHW_Hwsc_Click);
			this.tableLayoutPanel4.ColumnCount = 4;
			this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25f));
			this.tableLayoutPanel4.Controls.Add(this.BtnHW_Zxpw, 3, 0);
			this.tableLayoutPanel4.Controls.Add(this.BtnHW_Jczc, 2, 0);
			this.tableLayoutPanel4.Controls.Add(this.BtnHW_Hyzc, 1, 0);
			this.tableLayoutPanel4.Controls.Add(this.BtnHW_Hwsc, 0, 0);
			this.tableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel4.Location = new System.Drawing.Point(3, 28);
			this.tableLayoutPanel4.Name = "tableLayoutPanel4";
			this.tableLayoutPanel4.RowCount = 1;
			this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
			this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50f));
			this.tableLayoutPanel4.Size = new System.Drawing.Size(994, 79);
			this.tableLayoutPanel4.TabIndex = 2;
			this.BtnHW_Hyzc.Dock = System.Windows.Forms.DockStyle.Fill;
			this.BtnHW_Hyzc.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.BtnHW_Hyzc.Location = new System.Drawing.Point(268, 15);
			this.BtnHW_Hyzc.Margin = new System.Windows.Forms.Padding(20, 15, 20, 15);
			this.BtnHW_Hyzc.Name = "BtnHW_Hyzc";
			this.BtnHW_Hyzc.Size = new System.Drawing.Size(208, 49);
			this.BtnHW_Hyzc.TabIndex = 1;
			this.BtnHW_Hyzc.Text = "会议座次";
			this.BtnHW_Hyzc.UseVisualStyleBackColor = true;
			this.BtnHW_Hyzc.Click += new System.EventHandler(BtnHW_Hyzc_Click);
			this.groupBox4.Controls.Add(this.tableLayoutPanel4);
			this.groupBox4.Dock = System.Windows.Forms.DockStyle.Top;
			this.groupBox4.Font = new System.Drawing.Font("宋体", 13f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.groupBox4.Location = new System.Drawing.Point(0, 110);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1000, 110);
			this.groupBox4.TabIndex = 13;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "会务接待（一般用于大型会议）";
			this.groupBox1.Controls.Add(this.tableLayoutPanel1);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
			this.groupBox1.Font = new System.Drawing.Font("宋体", 13f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Margin = new System.Windows.Forms.Padding(5, 15, 5, 15);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1000, 110);
			this.groupBox1.TabIndex = 10;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "下行文（红头正式文件）";
			base.Controls.Add(this.groupBox3);
			base.Controls.Add(this.groupBox2);
			base.Controls.Add(this.groupBox5);
			base.Controls.Add(this.groupBox4);
			base.Controls.Add(this.groupBox1);
			base.Name = "Uc_GeShi";
			base.Size = new System.Drawing.Size(1000, 633);
			this.groupBox3.ResumeLayout(false);
			this.tableLayoutPanel3.ResumeLayout(false);
			this.tableLayoutPanel2.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.tableLayoutPanel1.ResumeLayout(false);
			this.tableLayoutPanel5.ResumeLayout(false);
			this.groupBox5.ResumeLayout(false);
			this.tableLayoutPanel4.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			base.ResumeLayout(false);
		}
	}
}
