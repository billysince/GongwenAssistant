using System;
using System.Drawing;
using System.Windows.Forms;

namespace Local_Wps_Vsto
{
	public class Uc_UserInfo : UserControl
	{
		private Button btnImprove;

		private Label lbEmail;

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

		public Uc_UserInfo()
		{
			InitializeComponent();
		}

		public Uc_UserInfo(MyAddin ribbonDesign)
		{
			_parentRibbon = ribbonDesign;
			InitializeComponent();
		}

		private void Uc_UserInfo_Load(object sender, EventArgs e)
		{
			try
			{
				if (UserUtil.IsVip())
				{
					lbEmail.Text = "插件已激活";
				}
				else
				{
					lbEmail.Text = "插件尚未激活";
				}
			}
			catch (Exception)
			{
			}
		}

		private void btnImprove_Click(object sender, EventArgs e)
		{
			try
			{
				_parentRibbon.btnImprove_Click();
			}
			catch (Exception)
			{
			}
		}

		private void InitializeComponent()
		{
			this.btnImprove = new System.Windows.Forms.Button();
			this.lbEmail = new System.Windows.Forms.Label();
			base.SuspendLayout();
			this.btnImprove.Font = new System.Drawing.Font("宋体", 16f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.btnImprove.Location = new System.Drawing.Point(187, 356);
			this.btnImprove.Name = "btnImprove";
			this.btnImprove.Size = new System.Drawing.Size(200, 50);
			this.btnImprove.TabIndex = 11;
			this.btnImprove.Text = "立即激活";
			this.btnImprove.UseVisualStyleBackColor = true;
			this.btnImprove.Click += new System.EventHandler(btnImprove_Click);
			this.lbEmail.AutoSize = true;
			this.lbEmail.Font = new System.Drawing.Font("宋体", 16f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.lbEmail.Location = new System.Drawing.Point(182, 185);
			this.lbEmail.Name = "lbEmail";
			this.lbEmail.Size = new System.Drawing.Size(120, 27);
			this.lbEmail.TabIndex = 8;
			this.lbEmail.Text = "尚未激活";
			base.Controls.Add(this.btnImprove);
			base.Controls.Add(this.lbEmail);
			base.Name = "Uc_UserInfo";
			base.Size = new System.Drawing.Size(600, 500);
			base.Load += new System.EventHandler(Uc_UserInfo_Load);
			base.ResumeLayout(false);
			base.PerformLayout();
		}
	}
}
