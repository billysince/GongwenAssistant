using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;
using Word;

namespace Local_Wps_Vsto
{
	public class WordCorrect : UserControl
	{
		private MyAddin _parentRibbon;

		private string _strContent = "";

		private Range _rangeSelectContent;

		private List<Range> _sentencesList = new List<Range>();

		private BackgroundWorker BgWorker_Correct = new BackgroundWorker();

		private List<SentenceUtil> sentenceUtils = new List<SentenceUtil>();

		private string strCorrectResult = "";

		private Label lbTips;

		private TableLayoutPanel tableLayoutPanel1;

		private Label lbCorrect;

		private Button btnCorrect;

		private Button btnNext;

		private Button btnIgnore;

		private Label lbError;

		private Button btnPre;

		private TextBox tbResult;

		private Sentences pSentences;

		private List<SentenceShowUtil> sentenceShowUtil_List = new List<SentenceShowUtil>();

		private int intPianYi;

		private int nowCorrectPos;

		private Range rangeNowSentence;

		private int _nowSentencePos;

		private int _start;

		private int _end;

		private string _strErrorWord;

		public WordCorrect()
		{
			InitializeComponent();
		}

		public void getCorrectResult(Range rangeSelectContent, MyAddin ribbonDesign)
		{
			try
			{
				nowCorrectPos = 0;
				_nowSentencePos = 0;
				intPianYi = 0;
				_sentencesList.Clear();
				sentenceUtils.Clear();
				strCorrectResult = "";
				sentenceShowUtil_List.Clear();
				MessageTipsUtil.Show("正在自动校稿\r\n请耐心等待……", 100000);
				_strContent = "";
				_rangeSelectContent = rangeSelectContent;
				pSentences = rangeSelectContent.Sentences;
				if (pSentences.Count > 0)
				{
					_parentRibbon = ribbonDesign;
					if (BgWorker_Correct.IsBusy)
					{
						BgWorker_Correct.CancelAsync();
					}
					BgWorker_Correct = new BackgroundWorker();
					BgWorker_Correct.WorkerReportsProgress = false;
					BgWorker_Correct.WorkerSupportsCancellation = true;
					BgWorker_Correct.DoWork += DoWork_Handler_Correct;
					BgWorker_Correct.RunWorkerCompleted += RunWorkerCompleted_Handler_Correct;
					BgWorker_Correct.RunWorkerAsync();
					lbTips.Text = "正在联网校对……";
					btnPre.Enabled = false;
					btnNext.Enabled = false;
					btnIgnore.Enabled = false;
					btnCorrect.Enabled = false;
					tbResult.Text = "正在联网校对\r\n请耐心等待……";
					lbError.Text = "";
					lbCorrect.Text = "";
				}
				else
				{
					MessageTipsUtil.Show("没有发现可校对文字");
				}
			}
			catch (Exception)
			{
			}
		}

		private void DoWork_Handler_Correct(object sender, DoWorkEventArgs args)
		{
			try
			{
				for (int i = 1; i <= pSentences.Count; i++)
				{
					pSentences[i].Text.Replace(Convert.ToChar(10).ToString(), "").Replace(Convert.ToChar(13).ToString(), "");
					_strContent = _strContent + pSentences[i].Text + "#";
					_sentencesList.Add(pSentences[i]);
				}
				string address = UrlUtil.strGetWordCorrectUrl();
				WebClient obj = new WebClient
				{
					Credentials = CredentialCache.DefaultCredentials
				};
				string text = "content=" + _strContent + "&user_id=" + 1L;
				text = text + "&mac=" + SecretUtil.GetMachineCode();
				obj.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
				byte[] bytes = Encoding.UTF8.GetBytes(text);
				byte[] bytes2 = obj.UploadData(address, "POST", bytes);
				CorrectUtil correctUtil = JsonConvert.DeserializeObject<CorrectUtil>(strCorrectResult = Encoding.UTF8.GetString(bytes2));
				sentenceUtils = correctUtil.result;
				if ((sender as BackgroundWorker).CancellationPending)
				{
					args.Cancel = true;
				}
			}
			catch (Exception ex)
			{
				strCorrectResult += ex.ToString();
			}
		}

		private void RunWorkerCompleted_Handler_Correct(object sender, RunWorkerCompletedEventArgs args)
		{
			try
			{
				if (!args.Cancelled)
				{
					tbResult.Text = "";
					sentenceShowUtil_List.Clear();
					int num = 0;
					foreach (SentenceUtil sentenceUtil in sentenceUtils)
					{
						if (!sentenceUtil.sentence.Equals(sentenceUtil.sen_correct))
						{
							foreach (ErrorUtil item in sentenceUtil.error_list)
							{
								SentenceShowUtil sentenceShowUtil = new SentenceShowUtil();
								sentenceShowUtil.senPos = num;
								sentenceShowUtil.sentence = sentenceUtil.sentence;
								sentenceShowUtil.sen_correct = sentenceUtil.sen_correct;
								sentenceShowUtil.error_word = item.error_word;
								sentenceShowUtil.correct_word = item.correct_word;
								sentenceShowUtil.start = item.start;
								sentenceShowUtil.end = item.end;
								sentenceShowUtil_List.Add(sentenceShowUtil);
							}
						}
						num++;
					}
					intPianYi = 0;
					try
					{
						int num2 = 0;
						string text = "";
						bool flag = false;
						foreach (Range sentences in _sentencesList)
						{
							if (flag)
							{
								break;
							}
							if (sentences != null)
							{
								text = sentences.Text;
								text = text.Replace(Convert.ToChar(10).ToString(), "");
								text = text.Replace(Convert.ToChar(13).ToString(), "");
								if (text.Length > 0)
								{
									flag = true;
									break;
								}
							}
							num2++;
						}
						int num3 = 0;
						flag = false;
						using (List<SentenceUtil>.Enumerator enumerator = sentenceUtils.GetEnumerator())
						{
							while (enumerator.MoveNext() && !(enumerator.Current.sentence == text))
							{
								num3++;
							}
						}
						intPianYi = num3 - num2;
					}
					catch (Exception)
					{
					}
					if (sentenceShowUtil_List.Count <= 0)
					{
						lbTips.Text = "恭喜您，没有发现错误";
						MessageTipsUtil.Show(lbTips.Text);
					}
					else
					{
						lbError.Text = "共发现" + sentenceShowUtil_List.Count + "处错误";
						_rangeSelectContent.Paragraphs[1].Range.Select();
						nowCorrectPos = 0;
						setNowCorrectWord();
						MessageTipsUtil.Show("校对完成\r\n共发现" + sentenceShowUtil_List.Count + "处错误");
					}
				}
				MyAddin.boolBtnJiaoGao = true;
				MyAddin.thisAddIn.InvalidateControl("btnJiaoGao");
			}
			catch (Exception)
			{
			}
		}

		public void setNowCorrectWord()
		{
			try
			{
				int count = sentenceShowUtil_List.Count;
				if (nowCorrectPos >= count)
				{
					nowCorrectPos = count - 1;
				}
				if (nowCorrectPos < 0)
				{
					nowCorrectPos = 0;
				}
				if (count <= 0)
				{
					lbTips.Text = "恭喜您，错误全部修正";
					MessageTipsUtil.Show(lbTips.Text);
					btnPre.Enabled = false;
					btnNext.Enabled = false;
					btnIgnore.Enabled = false;
					btnCorrect.Enabled = false;
					tbResult.Text = lbTips.Text;
					lbError.Text = "";
					lbCorrect.Text = "";
					MessageTipsUtil.Show(lbTips.Text);
					return;
				}
				SentenceShowUtil sentenceShowUtil = sentenceShowUtil_List[nowCorrectPos];
				lbTips.Text = "剩余" + count + "条错误";
				btnIgnore.Enabled = true;
				btnCorrect.Enabled = true;
				switch (count)
				{
				case 1:
					btnPre.Enabled = false;
					btnNext.Enabled = false;
					break;
				case 2:
					if (nowCorrectPos == 0)
					{
						btnPre.Enabled = false;
						btnNext.Enabled = true;
					}
					else if (nowCorrectPos == count - 1)
					{
						btnPre.Enabled = true;
						btnNext.Enabled = false;
					}
					break;
				default:
					if (nowCorrectPos == 0)
					{
						btnPre.Enabled = false;
						btnNext.Enabled = true;
					}
					else if (nowCorrectPos == count - 1)
					{
						btnPre.Enabled = true;
						btnNext.Enabled = false;
					}
					else
					{
						btnPre.Enabled = true;
						btnNext.Enabled = true;
					}
					break;
				}
				tbResult.Text = sentenceShowUtil.sentence;
				lbError.Text = sentenceShowUtil.error_word + "->";
				lbCorrect.Text = sentenceShowUtil.correct_word;
				setBackgroudColor(sentenceShowUtil.senPos, sentenceShowUtil.start, sentenceShowUtil.end, sentenceShowUtil.error_word);
			}
			catch (Exception)
			{
			}
		}

		public void setBackgroudColor(int nowSentencePos, int start, int end, string strErrorWord)
		{
			_nowSentencePos = nowSentencePos;
			_start = start;
			_end = end;
			_strErrorWord = strErrorWord;
			try
			{
				Range range = _sentencesList[nowSentencePos + intPianYi];
				Range range2 = _sentencesList[nowSentencePos + intPianYi];
				int start2 = range2.Start;
				int end2 = range2.End;
				range2.SetRange(range2.Start + start, range2.Start + end);
				TextBox textBox = tbResult;
				textBox.Text = textBox.Text + "\r\nstart:" + start + "--end:" + end;
				TextBox textBox2 = tbResult;
				textBox2.Text = textBox2.Text + "\r\nreturn_errorword:" + strErrorWord;
				TextBox textBox3 = tbResult;
				textBox3.Text = textBox3.Text + "\r\nlocal_get:" + range2.Text;
				range2.HighlightColorIndex = WdColorIndex.wdYellow;
				range2.SetRange(start2, end2);
				range.Select();
			}
			catch (Exception)
			{
			}
		}

		private void resetBackgroundColor()
		{
			try
			{
				_sentencesList[_nowSentencePos + intPianYi].HighlightColorIndex = WdColorIndex.wdAuto;
			}
			catch (Exception)
			{
			}
		}

		private void btnNext_Click(object sender, EventArgs e)
		{
			try
			{
				resetBackgroundColor();
				nowCorrectPos++;
				setNowCorrectWord();
			}
			catch (Exception)
			{
			}
		}

		private void btnPre_Click(object sender, EventArgs e)
		{
			try
			{
				resetBackgroundColor();
				nowCorrectPos--;
				setNowCorrectWord();
			}
			catch (Exception)
			{
			}
		}

		private void btnIgnore_Click(object sender, EventArgs e)
		{
			try
			{
				resetBackgroundColor();
				sentenceShowUtil_List.RemoveAt(nowCorrectPos);
				setNowCorrectWord();
			}
			catch (Exception)
			{
			}
		}

		private void btnCorrect_Click(object sender, EventArgs e)
		{
			try
			{
				SentenceShowUtil sentenceShowUtil = sentenceShowUtil_List[nowCorrectPos];
				correctErrorWord(sentenceShowUtil.senPos, sentenceShowUtil.start, sentenceShowUtil.end, sentenceShowUtil.error_word, sentenceShowUtil.correct_word);
				resetBackgroundColor();
				sentenceShowUtil_List.RemoveAt(nowCorrectPos);
				setNowCorrectWord();
			}
			catch (Exception)
			{
			}
		}

		public void correctErrorWord(int nowSentencePos, int start, int end, string strErrorWord, string strRightWord)
		{
			_nowSentencePos = nowSentencePos;
			_start = start;
			_end = end;
			_strErrorWord = strErrorWord;
			try
			{
				Range range = _sentencesList[nowSentencePos];
				Range range2 = _sentencesList[nowSentencePos];
				int start2 = range2.Start;
				int end2 = range2.End;
				range2.SetRange(range2.Start + start, range2.Start + end);
				TextBox textBox = tbResult;
				textBox.Text = textBox.Text + "\r\nstart:" + start + "--end:" + end;
				TextBox textBox2 = tbResult;
				textBox2.Text = textBox2.Text + "\r\nreturn_errorword:" + strErrorWord;
				TextBox textBox3 = tbResult;
				textBox3.Text = textBox3.Text + "\r\nlocal_get:" + range2.Text;
				range2.Text = strRightWord;
				range2.HighlightColorIndex = WdColorIndex.wdAuto;
				range2.SetRange(start2, end2);
				range.Select();
			}
			catch (Exception)
			{
			}
		}

		private void WordCorrect_Load(object sender, EventArgs e)
		{
			lbError.Text = "";
			lbCorrect.Text = "";
		}

		private void InitializeComponent()
		{
			this.lbTips = new System.Windows.Forms.Label();
			this.lbCorrect = new System.Windows.Forms.Label();
			this.btnCorrect = new System.Windows.Forms.Button();
			this.btnNext = new System.Windows.Forms.Button();
			this.btnIgnore = new System.Windows.Forms.Button();
			this.lbError = new System.Windows.Forms.Label();
			this.btnPre = new System.Windows.Forms.Button();
			this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
			this.tbResult = new System.Windows.Forms.TextBox();
			this.tableLayoutPanel1.SuspendLayout();
			base.SuspendLayout();
			this.lbTips.AutoSize = true;
			this.tableLayoutPanel1.SetColumnSpan(this.lbTips, 2);
			this.lbTips.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lbTips.Location = new System.Drawing.Point(3, 0);
			this.lbTips.Name = "lbTips";
			this.lbTips.Size = new System.Drawing.Size(594, 50);
			this.lbTips.TabIndex = 3;
			this.lbTips.Text = "正在联网校对中……";
			this.lbTips.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lbCorrect.AutoSize = true;
			this.lbCorrect.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lbCorrect.Font = new System.Drawing.Font("宋体", 18f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 134);
			this.lbCorrect.ForeColor = System.Drawing.Color.Red;
			this.lbCorrect.Location = new System.Drawing.Point(220, 100);
			this.lbCorrect.Name = "lbCorrect";
			this.lbCorrect.Size = new System.Drawing.Size(377, 100);
			this.lbCorrect.TabIndex = 7;
			this.lbCorrect.Text = "right";
			this.lbCorrect.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCorrect.Dock = System.Windows.Forms.DockStyle.Fill;
			this.btnCorrect.Location = new System.Drawing.Point(227, 205);
			this.btnCorrect.Margin = new System.Windows.Forms.Padding(10, 5, 10, 5);
			this.btnCorrect.Name = "btnCorrect";
			this.btnCorrect.Size = new System.Drawing.Size(363, 45);
			this.btnCorrect.TabIndex = 2;
			this.btnCorrect.Text = "立即修正错误";
			this.btnCorrect.UseVisualStyleBackColor = true;
			this.btnCorrect.Click += new System.EventHandler(btnCorrect_Click);
			this.btnNext.Dock = System.Windows.Forms.DockStyle.Fill;
			this.btnNext.Location = new System.Drawing.Point(227, 55);
			this.btnNext.Margin = new System.Windows.Forms.Padding(10, 5, 10, 5);
			this.btnNext.Name = "btnNext";
			this.btnNext.Size = new System.Drawing.Size(363, 40);
			this.btnNext.TabIndex = 5;
			this.btnNext.Text = "下一条";
			this.btnNext.UseVisualStyleBackColor = true;
			this.btnNext.Click += new System.EventHandler(btnNext_Click);
			this.btnIgnore.Dock = System.Windows.Forms.DockStyle.Fill;
			this.btnIgnore.Location = new System.Drawing.Point(10, 205);
			this.btnIgnore.Margin = new System.Windows.Forms.Padding(10, 5, 10, 5);
			this.btnIgnore.Name = "btnIgnore";
			this.btnIgnore.Size = new System.Drawing.Size(197, 45);
			this.btnIgnore.TabIndex = 1;
			this.btnIgnore.Text = "忽略此条";
			this.btnIgnore.UseVisualStyleBackColor = true;
			this.btnIgnore.Click += new System.EventHandler(btnIgnore_Click);
			this.lbError.AutoSize = true;
			this.lbError.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lbError.Font = new System.Drawing.Font("宋体", 16f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
			this.lbError.Location = new System.Drawing.Point(3, 100);
			this.lbError.Name = "lbError";
			this.lbError.Size = new System.Drawing.Size(211, 100);
			this.lbError.TabIndex = 6;
			this.lbError.Text = "error";
			this.lbError.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnPre.Dock = System.Windows.Forms.DockStyle.Fill;
			this.btnPre.Location = new System.Drawing.Point(10, 55);
			this.btnPre.Margin = new System.Windows.Forms.Padding(10, 5, 10, 5);
			this.btnPre.Name = "btnPre";
			this.btnPre.Size = new System.Drawing.Size(197, 40);
			this.btnPre.TabIndex = 4;
			this.btnPre.Text = "上一条";
			this.btnPre.UseVisualStyleBackColor = true;
			this.btnPre.Click += new System.EventHandler(btnPre_Click);
			this.tableLayoutPanel1.ColumnCount = 2;
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 36.33333f));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 63.66667f));
			this.tableLayoutPanel1.Controls.Add(this.lbTips, 0, 0);
			this.tableLayoutPanel1.Controls.Add(this.lbCorrect, 1, 2);
			this.tableLayoutPanel1.Controls.Add(this.btnCorrect, 1, 3);
			this.tableLayoutPanel1.Controls.Add(this.btnNext, 1, 1);
			this.tableLayoutPanel1.Controls.Add(this.btnIgnore, 0, 3);
			this.tableLayoutPanel1.Controls.Add(this.lbError, 0, 2);
			this.tableLayoutPanel1.Controls.Add(this.btnPre, 0, 1);
			this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top;
			this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
			this.tableLayoutPanel1.Name = "tableLayoutPanel1";
			this.tableLayoutPanel1.RowCount = 4;
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50f));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50f));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 100f));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50f));
			this.tableLayoutPanel1.Size = new System.Drawing.Size(600, 255);
			this.tableLayoutPanel1.TabIndex = 10;
			this.tbResult.Location = new System.Drawing.Point(80, 449);
			this.tbResult.Multiline = true;
			this.tbResult.Name = "tbResult";
			this.tbResult.Size = new System.Drawing.Size(190, 42);
			this.tbResult.TabIndex = 9;
			this.tbResult.Text = "正在联网校对中……";
			this.tbResult.Visible = false;
			base.Controls.Add(this.tableLayoutPanel1);
			base.Controls.Add(this.tbResult);
			base.Name = "WordCorrect";
			base.Size = new System.Drawing.Size(600, 600);
			base.Load += new System.EventHandler(WordCorrect_Load);
			this.tableLayoutPanel1.ResumeLayout(false);
			this.tableLayoutPanel1.PerformLayout();
			base.ResumeLayout(false);
			base.PerformLayout();
		}
	}
}
