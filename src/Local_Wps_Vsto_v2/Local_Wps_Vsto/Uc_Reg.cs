using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace Local_Wps_Vsto
{
	public class Uc_Reg : UserControl
	{
		private MyAddin _parentRibbon;

		private string strEmail = "";

		private string strPwd = "";

		private BackgroundWorker BgWorker_GetWebInfo = new BackgroundWorker();

		private string strReturn = "";

		private UserLoginItem userLoginItem = new UserLoginItem();

		private string strStatus = "";

		private IContainer components;

		private TextBox tbPwd;

		private TextBox tbEmail;

		private Label label2;

		private Label label1;

		private Label lbTips;

		private Button btLogin;

		private Button btFindPwd;

		private Button btRegister;

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

		public Uc_Reg()
		{
			InitializeComponent();
		}

		public Uc_Reg(MyAddin ribbonDesign)
		{
			_parentRibbon = ribbonDesign;
			InitializeComponent();
		}

		private void btRegister_Click(object sender, EventArgs e)
		{
			try
			{
				strEmail = tbEmail.Text.ToString();
				strPwd = tbPwd.Text.ToString();
				if (isRightEmail(strEmail))
				{
					if (isRightPwd(strPwd))
					{
						MessageTipsUtil.Show("注册中……", 5000);
						btLogin.Enabled = false;
						if (BgWorker_GetWebInfo.IsBusy)
						{
							BgWorker_GetWebInfo.CancelAsync();
						}
						BgWorker_GetWebInfo = new BackgroundWorker();
						BgWorker_GetWebInfo.WorkerReportsProgress = true;
						BgWorker_GetWebInfo.WorkerSupportsCancellation = true;
						BgWorker_GetWebInfo.DoWork += DoWork_Handler_GetWebInfo;
						BgWorker_GetWebInfo.RunWorkerCompleted += RunWorkerCompleted_GetWebInfo;
						BgWorker_GetWebInfo.RunWorkerAsync();
					}
					else
					{
						tbPwd.Focus();
						MessageTipsUtil.Show("密码格式错误");
					}
				}
				else
				{
					tbEmail.Focus();
					MessageTipsUtil.Show("邮箱格式错误");
				}
			}
			catch (Exception)
			{
			}
		}

		private void Uc_Reg_Load(object sender, EventArgs e)
		{
			try
			{
				BgWorker_GetWebInfo.WorkerReportsProgress = true;
				BgWorker_GetWebInfo.WorkerSupportsCancellation = true;
				BgWorker_GetWebInfo.DoWork += DoWork_Handler_GetWebInfo;
				BgWorker_GetWebInfo.RunWorkerCompleted += RunWorkerCompleted_GetWebInfo;
				tbEmail.Focus();
			}
			catch (Exception)
			{
			}
		}

		private void DoWork_Handler_GetWebInfo(object sender, DoWorkEventArgs args)
		{
			try
			{
				string address = UrlUtil.strRegUrl(strEmail, strPwd);
				byte[] bytes = new WebClient
				{
					Credentials = CredentialCache.DefaultCredentials
				}.DownloadData(address);
				strReturn = Encoding.UTF8.GetString(bytes);
				userLoginItem = JsonConvert.DeserializeObject<UserLoginItem>(strReturn);
			}
			catch (Exception)
			{
			}
		}

		private void RunWorkerCompleted_GetWebInfo(object sender, RunWorkerCompletedEventArgs args)
		{
			try
			{
				MsgTips.Hide();
				if (userLoginItem.ReturnCode == "ok")
				{
					try
					{
						MyAddin.CheckImproveInfo_Local();
						MyAddin.refreshLoginPanel();
					}
					catch (Exception)
					{
					}
					MessageTipsUtil.Show("注册成功");
					MyAddin.regPanel.Visible = false;
				}
				else
				{
					MessageTipsUtil.Show(userLoginItem.ReturnCode);
				}
				btLogin.Enabled = true;
			}
			catch (Exception)
			{
			}
		}

		private void btFindPwd_Click(object sender, EventArgs e)
		{
			try
			{
				Process.Start("http://www.6dgww.com/findpwd.php");
			}
			catch (Exception)
			{
			}
		}

		private void tbEmail_TextChanged(object sender, EventArgs e)
		{
			if (tbEmail.Text.Length >= 6 && tbPwd.Text.Length >= 6)
			{
				btRegister.Enabled = true;
			}
			else
			{
				btRegister.Enabled = false;
			}
		}

		private void tbEmail_LostFocus(object sender, EventArgs e)
		{
			isRightEmail(tbEmail.Text.ToString());
		}

		private bool isRightEmail(string strEmail)
		{
			try
			{
				if (new Regex("^\\s*([A-Za-z0-9_-]+(\\.\\w+)*@(\\w+\\.)+\\w{2,5})\\s*$").IsMatch(strEmail))
				{
					return true;
				}
				return false;
			}
			catch (Exception)
			{
				return false;
			}
		}

		private bool isRightPwd(string strPwd)
		{
			try
			{
				if (new Regex("^[A-Za-z0-9]+$").IsMatch(strPwd))
				{
					return true;
				}
				return false;
			}
			catch (Exception)
			{
				return false;
			}
		}

		private void tbPwd_PasswordChanged(object sender, EventArgs e)
		{
			if (tbEmail.Text.Length >= 6 && tbPwd.Text.Length >= 6)
			{
				btRegister.Enabled = true;
			}
			else
			{
				btRegister.Enabled = false;
			}
		}

		private void btLogin_Click(object sender, EventArgs e)
		{
			try
			{
				MyAddin.btnLogin_Click();
				MyAddin.regPanel.Visible = false;
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
			this.tbPwd = new System.Windows.Forms.TextBox();
			this.tbEmail = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lbTips = new System.Windows.Forms.Label();
			this.btLogin = new System.Windows.Forms.Button();
			this.btFindPwd = new System.Windows.Forms.Button();
			this.btRegister = new System.Windows.Forms.Button();
			base.SuspendLayout();
			this.tbPwd.Font = new System.Drawing.Font("宋体", 14f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tbPwd.Location = new System.Drawing.Point(235, 221);
			this.tbPwd.MaxLength = 30;
			this.tbPwd.Name = "tbPwd";
			this.tbPwd.Size = new System.Drawing.Size(250, 34);
			this.tbPwd.TabIndex = 19;
			this.tbPwd.TextChanged += new System.EventHandler(tbPwd_PasswordChanged);
			this.tbEmail.Font = new System.Drawing.Font("宋体", 14f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tbEmail.Location = new System.Drawing.Point(235, 158);
			this.tbEmail.Name = "tbEmail";
			this.tbEmail.Size = new System.Drawing.Size(250, 34);
			this.tbEmail.TabIndex = 18;
			this.tbEmail.TextChanged += new System.EventHandler(tbEmail_TextChanged);
			this.tbEmail.Leave += new System.EventHandler(tbEmail_LostFocus);
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.label2.Location = new System.Drawing.Point(143, 227);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(69, 20);
			this.label2.TabIndex = 24;
			this.label2.Text = "密码：";
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.label1.Location = new System.Drawing.Point(147, 164);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(69, 20);
			this.label1.TabIndex = 25;
			this.label1.Text = "邮箱：";
			this.lbTips.AutoSize = true;
			this.lbTips.Font = new System.Drawing.Font("宋体", 22f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.lbTips.Location = new System.Drawing.Point(210, 71);
			this.lbTips.Name = "lbTips";
			this.lbTips.Size = new System.Drawing.Size(165, 37);
			this.lbTips.TabIndex = 23;
			this.lbTips.Text = "欢迎注册";
			this.lbTips.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.btLogin.Font = new System.Drawing.Font("宋体", 10f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.btLogin.Location = new System.Drawing.Point(308, 399);
			this.btLogin.Name = "btLogin";
			this.btLogin.Size = new System.Drawing.Size(250, 30);
			this.btLogin.TabIndex = 22;
			this.btLogin.Text = "登录已有账号";
			this.btLogin.UseVisualStyleBackColor = true;
			this.btLogin.Click += new System.EventHandler(btLogin_Click);
			this.btFindPwd.Font = new System.Drawing.Font("宋体", 10f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.btFindPwd.Location = new System.Drawing.Point(42, 399);
			this.btFindPwd.Name = "btFindPwd";
			this.btFindPwd.Size = new System.Drawing.Size(150, 30);
			this.btFindPwd.TabIndex = 21;
			this.btFindPwd.Text = "找回密码";
			this.btFindPwd.UseVisualStyleBackColor = true;
			this.btFindPwd.Click += new System.EventHandler(btFindPwd_Click);
			this.btRegister.Enabled = false;
			this.btRegister.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.btRegister.Location = new System.Drawing.Point(246, 305);
			this.btRegister.Name = "btRegister";
			this.btRegister.Size = new System.Drawing.Size(150, 40);
			this.btRegister.TabIndex = 20;
			this.btRegister.Text = "立即注册";
			this.btRegister.UseVisualStyleBackColor = true;
			this.btRegister.Click += new System.EventHandler(btRegister_Click);
			base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 15f);
			base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			base.Controls.Add(this.tbPwd);
			base.Controls.Add(this.tbEmail);
			base.Controls.Add(this.label2);
			base.Controls.Add(this.label1);
			base.Controls.Add(this.lbTips);
			base.Controls.Add(this.btLogin);
			base.Controls.Add(this.btFindPwd);
			base.Controls.Add(this.btRegister);
			base.Name = "Uc_Reg";
			base.Size = new System.Drawing.Size(600, 500);
			base.Load += new System.EventHandler(Uc_Reg_Load);
			base.ResumeLayout(false);
			base.PerformLayout();
		}
	}
}
