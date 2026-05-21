using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Local_Wps_Vsto
{
	public class Uc_Improve : UserControl
	{
		private Button btnImprove;

		private Label lbEmail;

		private MyAddin _parentRibbon;

		private string strLocalNeedCode = "";

		private IContainer components;

		private TabPage tpCode;

		private Button btCode;

		private TextBox tbCode;

		private TabControl tabControl1;

		private TabPage tpFree;

		private RichTextBox rtbShow;

		private Label LbMachineCode;

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

		public Uc_Improve()
		{
			InitializeComponent();
		}

		public Uc_Improve(MyAddin ribbonDesign)
		{
			_parentRibbon = ribbonDesign;
			InitializeComponent();
		}

		private void Uc_Improve_Load(object sender, EventArgs e)
		{
			try
			{
				string machineCode = SecretUtil.GetMachineCode();
				strLocalNeedCode = SecretUtil.getNeedActiviteCode();
				LbMachineCode.Text = "机器识别码：" + machineCode;
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
			try
			{
				string rtf = File.ReadAllText(CommonConfig.strBaseFolder + "\\template\\h_i.rtf");
				rtbShow.Rtf = rtf;
			}
			catch (Exception)
			{
			}
		}

		private void btCode_Click(object sender, EventArgs e)
		{
			try
			{
				if (tbCode.Text.Length <= 1)
				{
					MessageTipsUtil.Show("请输入正确的激活码");
					return;
				}
				string text = tbCode.Text.ToString().ToUpper();
				if (strLocalNeedCode.Equals(text.ToUpper()))
				{
					MessageTipsUtil.Show("激活成功！");
					UserUtil.SaveActiviteCode(strLocalNeedCode);
					MyAddin.CheckImproveInfo_Local();
					try
					{
						MyAddin.improvePanel.Visible = false;
						MyAddin.improvePanel.Delete();
						return;
					}
					catch (Exception)
					{
						return;
					}
				}
				if (text.ToUpper().Equals("AABBCCDDEE"))
				{
					MessageTipsUtil.Show("该激活码仅适用于在线版软件", 3000);
				}
				else
				{
					MessageTipsUtil.Show("激活码错误");
				}
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
			this.tpCode = new System.Windows.Forms.TabPage();
			this.LbMachineCode = new System.Windows.Forms.Label();
			this.btCode = new System.Windows.Forms.Button();
			this.tbCode = new System.Windows.Forms.TextBox();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tpFree = new System.Windows.Forms.TabPage();
			this.rtbShow = new System.Windows.Forms.RichTextBox();
			this.tpCode.SuspendLayout();
			this.tabControl1.SuspendLayout();
			this.tpFree.SuspendLayout();
			base.SuspendLayout();
			this.tpCode.AutoScroll = true;
			this.tpCode.BackColor = System.Drawing.Color.White;
			this.tpCode.Controls.Add(this.LbMachineCode);
			this.tpCode.Controls.Add(this.btCode);
			this.tpCode.Controls.Add(this.tbCode);
			this.tpCode.Font = new System.Drawing.Font("宋体", 14f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tpCode.Location = new System.Drawing.Point(4, 44);
			this.tpCode.Name = "tpCode";
			this.tpCode.Padding = new System.Windows.Forms.Padding(10);
			this.tpCode.Size = new System.Drawing.Size(792, 652);
			this.tpCode.TabIndex = 1;
			this.tpCode.Text = "激活码升级";
			this.LbMachineCode.AutoSize = true;
			this.LbMachineCode.Location = new System.Drawing.Point(232, 110);
			this.LbMachineCode.Name = "LbMachineCode";
			this.LbMachineCode.Size = new System.Drawing.Size(130, 24);
			this.LbMachineCode.TabIndex = 2;
			this.LbMachineCode.Text = "机器识别码";
			this.LbMachineCode.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.btCode.Location = new System.Drawing.Point(236, 327);
			this.btCode.Margin = new System.Windows.Forms.Padding(3, 3, 3, 200);
			this.btCode.Name = "btCode";
			this.btCode.Size = new System.Drawing.Size(181, 45);
			this.btCode.TabIndex = 1;
			this.btCode.Text = "验证激活码";
			this.btCode.UseVisualStyleBackColor = true;
			this.btCode.Click += new System.EventHandler(btCode_Click);
			this.tbCode.Font = new System.Drawing.Font("宋体", 16f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 134);
			this.tbCode.Location = new System.Drawing.Point(210, 208);
			this.tbCode.MaxLength = 20;
			this.tbCode.Name = "tbCode";
			this.tbCode.Size = new System.Drawing.Size(250, 38);
			this.tbCode.TabIndex = 0;
			this.tbCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.tabControl1.Controls.Add(this.tpCode);
			this.tabControl1.Controls.Add(this.tpFree);
			this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tabControl1.Font = new System.Drawing.Font("宋体", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.tabControl1.Location = new System.Drawing.Point(0, 0);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.Padding = new System.Drawing.Point(26, 10);
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(800, 700);
			this.tabControl1.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight;
			this.tabControl1.TabIndex = 1;
			this.tpFree.Controls.Add(this.rtbShow);
			this.tpFree.Location = new System.Drawing.Point(4, 44);
			this.tpFree.Name = "tpFree";
			this.tpFree.Size = new System.Drawing.Size(692, 552);
			this.tpFree.TabIndex = 2;
			this.tpFree.Text = "如何获取激活码";
			this.tpFree.UseVisualStyleBackColor = true;
			this.rtbShow.BackColor = System.Drawing.Color.White;
			this.rtbShow.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.rtbShow.Dock = System.Windows.Forms.DockStyle.Fill;
			this.rtbShow.Font = new System.Drawing.Font("宋体", 14f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.rtbShow.Location = new System.Drawing.Point(0, 0);
			this.rtbShow.Name = "rtbShow";
			this.rtbShow.ReadOnly = true;
			this.rtbShow.Size = new System.Drawing.Size(692, 552);
			this.rtbShow.TabIndex = 0;
			this.rtbShow.Text = "";
			base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 15f);
			base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			base.Controls.Add(this.tabControl1);
			base.Name = "Uc_Improve";
			base.Size = new System.Drawing.Size(800, 700);
			base.Load += new System.EventHandler(Uc_Improve_Load);
			this.tpCode.ResumeLayout(false);
			this.tpCode.PerformLayout();
			this.tabControl1.ResumeLayout(false);
			this.tpFree.ResumeLayout(false);
			base.ResumeLayout(false);
		}
	}
}
