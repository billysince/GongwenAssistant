using System;
using System.ComponentModel;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace Local_Wps_Vsto
{
	public class MesBoxShow : Form
	{
		private string _strMessage = "";

		private int _intShowSecond = 1000;

		private bool _boolSpeak;

		private Label TbMsg;

		private System.Windows.Forms.Timer timerSearchChanged;

		private BackgroundWorker BgWorker_Upload = new BackgroundWorker();

		private BackgroundWorker BgWorker_Speaker = new BackgroundWorker();

		private int _sleepSecond = 1000;

		private dynamic spVoice;

		public MesBoxShow()
		{
			InitializeComponent();
		}

		private void MesBoxShow_Load(object sender, EventArgs e)
		{
			base.ShowInTaskbar = false;
			Control.CheckForIllegalCrossThreadCalls = false;
			TbMsg.Text = _strMessage;
			_sleepSecond = _intShowSecond;
			timerSearchChanged = new System.Windows.Forms.Timer();
			timerSearchChanged.Enabled = true;
			timerSearchChanged.Interval = _intShowSecond;
			timerSearchChanged.Tick += Timer_SearchTxtChanged;
			timerSearchChanged.Start();
		}

		private void Timer_SearchTxtChanged(object sender, EventArgs e)
		{
			Close();
		}

		public MesBoxShow(string strMessage, int intShowSecond = 1000, bool boolSpeak = false)
		{
			try
			{
				InitializeComponent();
				_strMessage = strMessage;
				_intShowSecond = intShowSecond;
				_boolSpeak = boolSpeak;
			}
			catch (Exception)
			{
			}
		}

		private void DoWork_Handler_Speaker(object sender, DoWorkEventArgs args)
		{
			try
			{
				Type typeFromProgID = Type.GetTypeFromProgID("SAPI.SpVoice");
				spVoice = Activator.CreateInstance(typeFromProgID);
				spVoice.Speak(CommonConfig.STR_TIP_NEED_LOGIN);
			}
			catch (Exception)
			{
			}
		}

		private void DoWork_Handler_Upload(object sender, DoWorkEventArgs args)
		{
			try
			{
				Thread.Sleep(_sleepSecond);
			}
			catch (Exception)
			{
			}
		}

		private void ProgressChanged_Handler_Upload(object sender, ProgressChangedEventArgs args)
		{
			switch (args.ProgressPercentage)
			{
			case 1:
				base.Opacity += 0.1;
				break;
			case 2:
				base.Opacity -= 0.1;
				break;
			}
		}

		private void RunWorkerCompleted_Handler_Upload(object sender, RunWorkerCompletedEventArgs args)
		{
			try
			{
				Close();
			}
			catch (Exception)
			{
			}
		}

		private void MesBoxShow_Deactivate(object sender, EventArgs e)
		{
			try
			{
				Close();
			}
			catch (Exception)
			{
			}
		}

		private void MesBoxShow_FormClosing(object sender, FormClosingEventArgs e)
		{
			try
			{
				e.Cancel = true;
				Hide();
			}
			catch (Exception)
			{
			}
		}

		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Local_Wps_Vsto.MesBoxShow));
			this.TbMsg = new System.Windows.Forms.Label();
			base.SuspendLayout();
			this.TbMsg.BackColor = System.Drawing.Color.ForestGreen;
			this.TbMsg.Dock = System.Windows.Forms.DockStyle.Fill;
			this.TbMsg.Font = new System.Drawing.Font("宋体", 16f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.TbMsg.ForeColor = System.Drawing.Color.White;
			this.TbMsg.Location = new System.Drawing.Point(0, 0);
			this.TbMsg.Name = "TbMsg";
			this.TbMsg.Size = new System.Drawing.Size(350, 250);
			this.TbMsg.TabIndex = 1;
			this.TbMsg.Text = "提示";
			this.TbMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.TbMsg.Click += new System.EventHandler(TbMsg_Click);
			this.BackColor = System.Drawing.Color.Green;
			base.ClientSize = new System.Drawing.Size(350, 250);
			base.Controls.Add(this.TbMsg);
			base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			base.Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
			base.Name = "MesBoxShow";
			base.ShowInTaskbar = false;
			base.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			base.Deactivate += new System.EventHandler(MesBoxShow_Deactivate);
			base.Load += new System.EventHandler(MesBoxShow_Load);
			base.ResumeLayout(false);
		}

		private void TbMsg_Click(object sender, EventArgs e)
		{
		}
	}
}
