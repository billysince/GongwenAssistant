using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace Local_Wps_Vsto
{
	public class Uc_Update : UserControl
	{
		private EventLog eventLog1;

		private Panel panel1;

		private Button btnUpdate;

		private RichTextBox rtbDes;

		public UpdateInfo _updateInfo;

		private MyAddin _parentRibbon;

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

		public Uc_Update()
		{
			InitializeComponent();
		}

		public Uc_Update(UpdateInfo updateInfo, MyAddin ribbonDesign)
		{
			try
			{
				_updateInfo = updateInfo;
				_parentRibbon = ribbonDesign;
				InitializeComponent();
			}
			catch (Exception)
			{
			}
		}

		public Uc_Update(MyAddin ribbonDesign)
		{
			_parentRibbon = ribbonDesign;
			InitializeComponent();
		}

		private void btnUpdate_Click(object sender, EventArgs e)
		{
			try
			{
				Process.Start(_updateInfo.DownUrl);
				_parentRibbon.updatePanel.Visible = false;
				_parentRibbon.updatePanel.Delete();
			}
			catch (Exception)
			{
			}
		}

		private void btnIngore_Click(object sender, EventArgs e)
		{
		}

		private void Uc_Update_Load(object sender, EventArgs e)
		{
			Control.CheckForIllegalCrossThreadCalls = false;
		}

		public void UpdateUrlInfo()
		{
			rtbDes.Text = "发现新版本：" + _updateInfo.AppVersion + "\r\n更新日期：" + _updateInfo.CreateTime + "\r\n版本简介：" + _updateInfo.Describe;
		}

		private void InitializeComponent()
		{
			this.eventLog1 = new System.Diagnostics.EventLog();
			this.btnUpdate = new System.Windows.Forms.Button();
			this.panel1 = new System.Windows.Forms.Panel();
			this.rtbDes = new System.Windows.Forms.RichTextBox();
			((System.ComponentModel.ISupportInitialize)this.eventLog1).BeginInit();
			this.panel1.SuspendLayout();
			base.SuspendLayout();
			this.eventLog1.SynchronizingObject = this;
			this.btnUpdate.Dock = System.Windows.Forms.DockStyle.Fill;
			this.btnUpdate.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.btnUpdate.Font = new System.Drawing.Font("宋体", 14f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.btnUpdate.Location = new System.Drawing.Point(0, 0);
			this.btnUpdate.Margin = new System.Windows.Forms.Padding(10);
			this.btnUpdate.Name = "btnUpdate";
			this.btnUpdate.Size = new System.Drawing.Size(598, 48);
			this.btnUpdate.TabIndex = 0;
			this.btnUpdate.Text = "下载新版本";
			this.btnUpdate.UseVisualStyleBackColor = true;
			this.btnUpdate.Click += new System.EventHandler(btnUpdate_Click);
			this.panel1.Controls.Add(this.btnUpdate);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel1.Location = new System.Drawing.Point(0, 350);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(598, 48);
			this.panel1.TabIndex = 4;
			this.rtbDes.BackColor = System.Drawing.Color.White;
			this.rtbDes.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.rtbDes.Dock = System.Windows.Forms.DockStyle.Top;
			this.rtbDes.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.rtbDes.Location = new System.Drawing.Point(0, 0);
			this.rtbDes.Name = "rtbDes";
			this.rtbDes.ReadOnly = true;
			this.rtbDes.Size = new System.Drawing.Size(598, 350);
			this.rtbDes.TabIndex = 3;
			this.rtbDes.Text = "";
			base.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			base.Controls.Add(this.panel1);
			base.Controls.Add(this.rtbDes);
			base.Name = "Uc_Update";
			base.Size = new System.Drawing.Size(598, 398);
			((System.ComponentModel.ISupportInitialize)this.eventLog1).EndInit();
			this.panel1.ResumeLayout(false);
			base.ResumeLayout(false);
		}
	}
}
