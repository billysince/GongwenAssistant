using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Local_Wps_Vsto.Properties;

namespace Local_Wps_Vsto
{
	public class Uc_Help_KeFu : UserControl
	{
		private IContainer components;

		private Label label1;

		private PictureBox pictureBox1;

		public Uc_Help_KeFu()
		{
			InitializeComponent();
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
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			((System.ComponentModel.ISupportInitialize)this.pictureBox1).BeginInit();
			base.SuspendLayout();
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("黑体", 18f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.label1.Location = new System.Drawing.Point(147, 39);
			this.label1.Name = "label1";
			this.label1.Padding = new System.Windows.Forms.Padding(0, 10, 0, 10);
			this.label1.Size = new System.Drawing.Size(283, 50);
			this.label1.TabIndex = 4;
			this.label1.Text = "微信扫码，联系客服";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.pictureBox1.Image = Local_Wps_Vsto.Properties.Resources.weixin_kefu;
			this.pictureBox1.Location = new System.Drawing.Point(129, 132);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(301, 300);
			this.pictureBox1.TabIndex = 5;
			this.pictureBox1.TabStop = false;
			base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 15f);
			base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			base.Controls.Add(this.pictureBox1);
			base.Controls.Add(this.label1);
			base.Name = "Uc_Help_KeFu";
			base.Size = new System.Drawing.Size(600, 500);
			((System.ComponentModel.ISupportInitialize)this.pictureBox1).EndInit();
			base.ResumeLayout(false);
			base.PerformLayout();
		}
	}
}
