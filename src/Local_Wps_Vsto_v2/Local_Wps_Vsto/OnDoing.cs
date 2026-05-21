using System;
using System.ComponentModel;
using System.Drawing;
using System.Threading;
using System.Timers;
using System.Windows.Forms;

namespace Local_Wps_Vsto
{
	public class OnDoing : Form
	{
		private System.Timers.Timer timerWriterTips;

		public Label txtMsg;

		private TableLayoutPanel tableLayoutPanel1;

		private Label lb_8;

		private Label lb_7;

		private Label lb_6;

		private Label lb_5;

		private Label lb_4;

		private Label lb_3;

		private Label lb_2;

		private Label lb_1;

		private BackgroundWorker BgWorker_Upload = new BackgroundWorker();

		private System.Threading.Timer threadTimer;

		private int nowCount;

		public OnDoing()
		{
			InitializeComponent();
			Control.CheckForIllegalCrossThreadCalls = false;
		}

		private void OnDoing_Load(object sender, EventArgs e)
		{
			try
			{
				base.TransparencyKey = Color.White;
				setCircle();
				BgWorker_Upload.WorkerReportsProgress = true;
				BgWorker_Upload.WorkerSupportsCancellation = true;
				BgWorker_Upload.DoWork += DoWork_Handler_Upload;
				BgWorker_Upload.ProgressChanged += ProgressChanged_Handler_Upload;
				BgWorker_Upload.RunWorkerCompleted += RunWorkerCompleted_Handler_Upload;
				BgWorker_Upload.RunWorkerAsync();
			}
			catch (Exception)
			{
			}
		}

		private void timer_Elapsed(object sender, ElapsedEventArgs e)
		{
			Console.WriteLine(e.SignalTime.ToString());
			setCircle();
		}

		private void timersTimer_Elapsed(object sender, ElapsedEventArgs e)
		{
			setCircle();
		}

		private void DoWork_Handler_Upload(object sender, DoWorkEventArgs args)
		{
			BackgroundWorker backgroundWorker = sender as BackgroundWorker;
			try
			{
				while (true)
				{
					backgroundWorker.ReportProgress(2);
					Thread.Sleep(80);
				}
			}
			catch (Exception)
			{
			}
		}

		private void ProgressChanged_Handler_Upload(object sender, ProgressChangedEventArgs args)
		{
			setCircle();
		}

		private void RunWorkerCompleted_Handler_Upload(object sender, RunWorkerCompletedEventArgs args)
		{
		}

		private void Timer_TipsChanged(object sender, EventArgs e)
		{
			setCircle();
		}

		private void setCircle()
		{
			try
			{
				switch (nowCount % 9)
				{
				case 1:
					lb_1.Text = "●";
					lb_2.Text = "";
					lb_3.Text = "";
					lb_4.Text = "";
					lb_5.Text = "";
					lb_6.Text = "";
					lb_7.Text = "";
					lb_8.Text = "";
					break;
				case 2:
					lb_1.Text = "●";
					lb_2.Text = "●";
					lb_3.Text = "";
					lb_4.Text = "";
					lb_5.Text = "";
					lb_6.Text = "";
					lb_7.Text = "";
					lb_8.Text = "";
					break;
				case 3:
					lb_1.Text = "●";
					lb_2.Text = "●";
					lb_3.Text = "●";
					lb_4.Text = "";
					lb_5.Text = "";
					lb_6.Text = "";
					lb_7.Text = "";
					lb_8.Text = "";
					break;
				case 4:
					lb_1.Text = "●";
					lb_2.Text = "●";
					lb_3.Text = "●";
					lb_4.Text = "●";
					lb_5.Text = "";
					lb_6.Text = "";
					lb_7.Text = "";
					lb_8.Text = "";
					break;
				case 5:
					lb_1.Text = "●";
					lb_2.Text = "●";
					lb_3.Text = "●";
					lb_4.Text = "●";
					lb_5.Text = "●";
					lb_6.Text = "";
					lb_7.Text = "";
					lb_8.Text = "";
					break;
				case 6:
					lb_1.Text = "●";
					lb_2.Text = "●";
					lb_3.Text = "●";
					lb_4.Text = "●";
					lb_5.Text = "●";
					lb_6.Text = "●";
					lb_7.Text = "";
					lb_8.Text = "";
					break;
				case 7:
					lb_1.Text = "●";
					lb_2.Text = "●";
					lb_3.Text = "●";
					lb_4.Text = "●";
					lb_5.Text = "●";
					lb_6.Text = "●";
					lb_7.Text = "●";
					lb_8.Text = "";
					break;
				case 8:
					lb_1.Text = "●";
					lb_2.Text = "●";
					lb_3.Text = "●";
					lb_4.Text = "●";
					lb_5.Text = "●";
					lb_6.Text = "●";
					lb_7.Text = "●";
					lb_8.Text = "●";
					break;
				default:
					lb_1.Text = "";
					lb_2.Text = "";
					lb_3.Text = "";
					lb_4.Text = "";
					lb_5.Text = "";
					lb_6.Text = "";
					lb_7.Text = "";
					lb_8.Text = "";
					break;
				}
				nowCount++;
			}
			catch (Exception)
			{
			}
		}

		private void InitializeComponent()
		{
			this.txtMsg = new System.Windows.Forms.Label();
			this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
			this.lb_8 = new System.Windows.Forms.Label();
			this.lb_7 = new System.Windows.Forms.Label();
			this.lb_6 = new System.Windows.Forms.Label();
			this.lb_5 = new System.Windows.Forms.Label();
			this.lb_4 = new System.Windows.Forms.Label();
			this.lb_3 = new System.Windows.Forms.Label();
			this.lb_2 = new System.Windows.Forms.Label();
			this.lb_1 = new System.Windows.Forms.Label();
			this.tableLayoutPanel1.SuspendLayout();
			base.SuspendLayout();
			this.txtMsg.AutoSize = true;
			this.tableLayoutPanel1.SetColumnSpan(this.txtMsg, 8);
			this.txtMsg.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtMsg.Font = new System.Drawing.Font("宋体", 14f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 134);
			this.txtMsg.ForeColor = System.Drawing.Color.Green;
			this.txtMsg.Location = new System.Drawing.Point(3, 62);
			this.txtMsg.Name = "txtMsg";
			this.txtMsg.Size = new System.Drawing.Size(276, 91);
			this.txtMsg.TabIndex = 0;
			this.txtMsg.Text = "提示信息";
			this.txtMsg.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.tableLayoutPanel1.ColumnCount = 8;
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5f));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5f));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5f));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5f));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5f));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5f));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5f));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5f));
			this.tableLayoutPanel1.Controls.Add(this.lb_8, 7, 1);
			this.tableLayoutPanel1.Controls.Add(this.lb_7, 6, 1);
			this.tableLayoutPanel1.Controls.Add(this.lb_6, 5, 1);
			this.tableLayoutPanel1.Controls.Add(this.lb_5, 4, 1);
			this.tableLayoutPanel1.Controls.Add(this.lb_4, 3, 1);
			this.tableLayoutPanel1.Controls.Add(this.lb_3, 2, 1);
			this.tableLayoutPanel1.Controls.Add(this.lb_2, 1, 1);
			this.tableLayoutPanel1.Controls.Add(this.txtMsg, 0, 2);
			this.tableLayoutPanel1.Controls.Add(this.lb_1, 0, 1);
			this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
			this.tableLayoutPanel1.Name = "tableLayoutPanel1";
			this.tableLayoutPanel1.RowCount = 3;
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 28.75817f));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 71.24183f));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 90f));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20f));
			this.tableLayoutPanel1.Size = new System.Drawing.Size(282, 153);
			this.tableLayoutPanel1.TabIndex = 2;
			this.lb_8.AutoSize = true;
			this.lb_8.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lb_8.Font = new System.Drawing.Font("宋体", 15f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.lb_8.ForeColor = System.Drawing.Color.Green;
			this.lb_8.Location = new System.Drawing.Point(248, 18);
			this.lb_8.Name = "lb_8";
			this.lb_8.Size = new System.Drawing.Size(31, 44);
			this.lb_8.TabIndex = 8;
			this.lb_8.Text = "●";
			this.lb_8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lb_7.AutoSize = true;
			this.lb_7.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lb_7.Font = new System.Drawing.Font("宋体", 15f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.lb_7.ForeColor = System.Drawing.Color.Green;
			this.lb_7.Location = new System.Drawing.Point(213, 18);
			this.lb_7.Name = "lb_7";
			this.lb_7.Size = new System.Drawing.Size(29, 44);
			this.lb_7.TabIndex = 7;
			this.lb_7.Text = "●";
			this.lb_7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lb_6.AutoSize = true;
			this.lb_6.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lb_6.Font = new System.Drawing.Font("宋体", 15f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.lb_6.ForeColor = System.Drawing.Color.Green;
			this.lb_6.Location = new System.Drawing.Point(178, 18);
			this.lb_6.Name = "lb_6";
			this.lb_6.Size = new System.Drawing.Size(29, 44);
			this.lb_6.TabIndex = 6;
			this.lb_6.Text = "●";
			this.lb_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lb_5.AutoSize = true;
			this.lb_5.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lb_5.Font = new System.Drawing.Font("宋体", 15f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.lb_5.ForeColor = System.Drawing.Color.Green;
			this.lb_5.Location = new System.Drawing.Point(143, 18);
			this.lb_5.Name = "lb_5";
			this.lb_5.Size = new System.Drawing.Size(29, 44);
			this.lb_5.TabIndex = 5;
			this.lb_5.Text = "●";
			this.lb_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lb_4.AutoSize = true;
			this.lb_4.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lb_4.Font = new System.Drawing.Font("宋体", 15f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.lb_4.ForeColor = System.Drawing.Color.Green;
			this.lb_4.Location = new System.Drawing.Point(108, 18);
			this.lb_4.Name = "lb_4";
			this.lb_4.Size = new System.Drawing.Size(29, 44);
			this.lb_4.TabIndex = 4;
			this.lb_4.Text = "●";
			this.lb_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lb_3.AutoSize = true;
			this.lb_3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lb_3.Font = new System.Drawing.Font("宋体", 15f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.lb_3.ForeColor = System.Drawing.Color.Green;
			this.lb_3.Location = new System.Drawing.Point(73, 18);
			this.lb_3.Name = "lb_3";
			this.lb_3.Size = new System.Drawing.Size(29, 44);
			this.lb_3.TabIndex = 3;
			this.lb_3.Text = "●";
			this.lb_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lb_2.AutoSize = true;
			this.lb_2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lb_2.Font = new System.Drawing.Font("宋体", 15f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.lb_2.ForeColor = System.Drawing.Color.Green;
			this.lb_2.Location = new System.Drawing.Point(38, 18);
			this.lb_2.Name = "lb_2";
			this.lb_2.Size = new System.Drawing.Size(29, 44);
			this.lb_2.TabIndex = 2;
			this.lb_2.Text = "●";
			this.lb_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lb_1.AutoSize = true;
			this.lb_1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lb_1.Font = new System.Drawing.Font("宋体", 15f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.lb_1.ForeColor = System.Drawing.Color.Green;
			this.lb_1.Location = new System.Drawing.Point(3, 18);
			this.lb_1.Name = "lb_1";
			this.lb_1.Size = new System.Drawing.Size(29, 44);
			this.lb_1.TabIndex = 1;
			this.lb_1.Text = "●";
			this.lb_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.BackColor = System.Drawing.Color.White;
			base.ClientSize = new System.Drawing.Size(282, 153);
			base.Controls.Add(this.tableLayoutPanel1);
			base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			base.Name = "OnDoing";
			this.tableLayoutPanel1.ResumeLayout(false);
			this.tableLayoutPanel1.PerformLayout();
			base.ResumeLayout(false);
		}
	}
}
