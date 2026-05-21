using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Word;

namespace Local_Wps_Vsto
{
	public class Uc_TigangCorrect : UserControl
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

		private List<TiGangShowUtil> tiGangShowUtil_List = new List<TiGangShowUtil>();

		private int intPianYi;

		private int nowCorrectPos;

		private Range rangeNowSentence;

		private int _nowSentencePos;

		private int _start;

		private int _end;

		private string _strErrorWord;

		private Range nowRangeSentence;

		public Uc_TigangCorrect()
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
				tiGangShowUtil_List.Clear();
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
					lbTips.Text = "正在校对……";
					btnPre.Enabled = false;
					btnNext.Enabled = false;
					btnIgnore.Enabled = false;
					btnCorrect.Enabled = false;
					tbResult.Text = "正在校对\r\n请耐心等待……";
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
				List<Range> list = new List<Range>();
				for (int i = 1; i <= pSentences.Count; i++)
				{
					string strTitle = pSentences[i].Text;
					if (WpsHelper.Is_BiaoTi_One(strTitle) || WpsHelper.Is_BiaoTi_Two(strTitle) || WpsHelper.Is_BiaoTi_Three(strTitle))
					{
						list.Add(pSentences[i]);
					}
				}
				string[] array = new string[109]
				{
					"一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、", "十、",
					"十一、", "十二、", "十三、", "十四、", "十五、", "十六、", "十七、", "十八、", "十九、", "二十、",
					"二十一、", "二十二、", "二十三、", "二十四、", "二十五、", "二十六、", "二十七、", "二十八、", "二十九、", "三十、",
					"三十一、", "三十二、", "三十三、", "三十四、", "三十五、", "三十六、", "三十七、", "三十八、", "三十九、", "三十、",
					"三十一、", "三十二、", "三十三、", "三十四、", "三十五、", "三十六、", "三十七、", "三十八、", "三十九、", "四十、",
					"四十一、", "四十二、", "四十三、", "四十四、", "四十五、", "四十六、", "四十七、", "四十八、", "四十九、", "五十、",
					"五十一、", "五十二、", "五十三、", "五十四、", "五十五、", "五十六、", "五十七、", "五十八、", "五十九、", "六十、",
					"六十一、", "六十二、", "六十三、", "六十四、", "六十五、", "六十六、", "六十七、", "六十八、", "六十九、", "七十、",
					"七十一、", "七十二、", "七十三、", "七十四、", "七十五、", "七十六、", "七十七、", "七十八、", "七十九、", "八十、",
					"八十一、", "八十二、", "八十三、", "八十四、", "八十五、", "八十六、", "八十七、", "八十八、", "八十九、", "九十、",
					"九十一、", "九十二、", "九十三、", "九十四、", "九十五、", "九十六、", "九十七、", "九十八、", "九十九、"
				};
				string[] array2 = new string[109]
				{
					"（一）", "（二）", "（三）", "（四）", "（五）", "（六）", "（七）", "（八）", "（九）", "（十）",
					"（十一）", "（十二）", "（十三）", "（十四）", "（十五）", "（十六）", "（十七）", "（十八）", "（十九）", "（二十）",
					"（二十一）", "（二十二）", "（二十三）", "（二十四）", "（二十五）", "（二十六）", "（二十七）", "（二十八）", "（二十九）", "（三十）",
					"（三十一）", "（三十二）", "（三十三）", "（三十四）", "（三十五）", "（三十六）", "（三十七）", "（三十八）", "（三十九）", "（三十）",
					"（三十一）", "（三十二）", "（三十三）", "（三十四）", "（三十五）", "（三十六）", "（三十七）", "（三十八）", "（三十九）", "（四十）",
					"（四十一）", "（四十二）", "（四十三）", "（四十四）", "（四十五）", "（四十六）", "（四十七）", "（四十八）", "（四十九）", "（五十）",
					"（五十一）", "（五十二）", "（五十三）", "（五十四）", "（五十五）", "（五十六）", "（五十七）", "（五十八）", "（五十九）", "（六十）",
					"（六十一）", "（六十二）", "（六十三）", "（六十四）", "（六十五）", "（六十六）", "（六十七）", "（六十八）", "（六十九）", "（七十）",
					"（七十一）", "（七十二）", "（七十三）", "（七十四）", "（七十五）", "（七十六）", "（七十七）", "（七十八）", "（七十九）", "（八十）",
					"（八十一）", "（八十二）", "（八十三）", "（八十四）", "（八十五）", "（八十六）", "（八十七）", "（八十八）", "（八十九）", "（九十）",
					"（九十一）", "（九十二）", "（九十三）", "（九十四）", "（九十五）", "（九十六）", "（九十七）", "（九十八）", "（九十九）"
				};
				string[] array3 = new string[109]
				{
					"一是", "二是", "三是", "四是", "五是", "六是", "七是", "八是", "九是", "十是",
					"十一是", "十二是", "十三是", "十四是", "十五是", "十六是", "十七是", "十八是", "十九是", "二十是",
					"二十一是", "二十二是", "二十三是", "二十四是", "二十五是", "二十六是", "二十七是", "二十八是", "二十九是", "三十是",
					"三十一是", "三十二是", "三十三是", "三十四是", "三十五是", "三十六是", "三十七是", "三十八是", "三十九是", "三十是",
					"三十一是", "三十二是", "三十三是", "三十四是", "三十五是", "三十六是", "三十七是", "三十八是", "三十九是", "四十是",
					"四十一是", "四十二是", "四十三是", "四十四是", "四十五是", "四十六是", "四十七是", "四十八是", "四十九是", "五十是",
					"五十一是", "五十二是", "五十三是", "五十四是", "五十五是", "五十六是", "五十七是", "五十八是", "五十九是", "六十是",
					"六十一是", "六十二是", "六十三是", "六十四是", "六十五是", "六十六是", "六十七是", "六十八是", "六十九是", "七十是",
					"七十一是", "七十二是", "七十三是", "七十四是", "七十五是", "七十六是", "七十七是", "七十八是", "七十九是", "八十是",
					"八十一是", "八十二是", "八十三是", "八十四是", "八十五是", "八十六是", "八十七是", "八十八是", "八十九是", "九十是",
					"九十一是", "九十二是", "九十三是", "九十四是", "九十五是", "九十六是", "九十七是", "九十八是", "九十九是"
				};
				string[] array4 = new string[109]
				{
					"一要", "二要", "三要", "四要", "五要", "六要", "七要", "八要", "九要", "十要",
					"十一要", "十二要", "十三要", "十四要", "十五要", "十六要", "十七要", "十八要", "十九要", "二十要",
					"二十一要", "二十二要", "二十三要", "二十四要", "二十五要", "二十六要", "二十七要", "二十八要", "二十九要", "三十要",
					"三十一要", "三十二要", "三十三要", "三十四要", "三十五要", "三十六要", "三十七要", "三十八要", "三十九要", "三十要",
					"三十一要", "三十二要", "三十三要", "三十四要", "三十五要", "三十六要", "三十七要", "三十八要", "三十九要", "四十要",
					"四十一要", "四十二要", "四十三要", "四十四要", "四十五要", "四十六要", "四十七要", "四十八要", "四十九要", "五十要",
					"五十一要", "五十二要", "五十三要", "五十四要", "五十五要", "五十六要", "五十七要", "五十八要", "五十九要", "六十要",
					"六十一要", "六十二要", "六十三要", "六十四要", "六十五要", "六十六要", "六十七要", "六十八要", "六十九要", "七十要",
					"七十一要", "七十二要", "七十三要", "七十四要", "七十五要", "七十六要", "七十七要", "七十八要", "七十九要", "八十要",
					"八十一要", "八十二要", "八十三要", "八十四要", "八十五要", "八十六要", "八十七要", "八十八要", "八十九要", "九十要",
					"九十一要", "九十二要", "九十三要", "九十四要", "九十五要", "九十六要", "九十七要", "九十八要", "九十九要"
				};
				int num = 0;
				int num2 = 0;
				int num3 = 0;
				for (int j = 0; j < list.Count; j++)
				{
					string text = list[j].Text;
					if (WpsHelper.Is_BiaoTi_One(text))
					{
						if (!text.Substring(0, array[num].Length).Equals(array[num]))
						{
							TiGangShowUtil tiGangShowUtil = new TiGangShowUtil();
							tiGangShowUtil.rangeSentece = list[j];
							tiGangShowUtil.error_word = text.Substring(0, array[num].Length);
							tiGangShowUtil.correct_word = array[num];
							tiGangShowUtil.start = 0;
							tiGangShowUtil.end = array[num].Length;
							tiGangShowUtil_List.Add(tiGangShowUtil);
						}
						num++;
						num2 = 0;
						num3 = 0;
					}
					else if (WpsHelper.Is_BiaoTi_Two(text))
					{
						if (!text.Substring(0, array2[num2].Length).Equals(array2[num2]))
						{
							TiGangShowUtil tiGangShowUtil2 = new TiGangShowUtil();
							tiGangShowUtil2.rangeSentece = list[j];
							tiGangShowUtil2.error_word = text.Substring(0, array2[num2].Length);
							tiGangShowUtil2.correct_word = array2[num2];
							tiGangShowUtil2.start = 0;
							tiGangShowUtil2.end = array2[num2].Length;
							tiGangShowUtil_List.Add(tiGangShowUtil2);
						}
						num2++;
						num3 = 0;
					}
					else if (WpsHelper.Is_BiaoTi_Three(text))
					{
						if (!text.Substring(0, array3[num3].Length).Equals(array3[num3]))
						{
							TiGangShowUtil tiGangShowUtil3 = new TiGangShowUtil();
							tiGangShowUtil3.rangeSentece = list[j];
							tiGangShowUtil3.error_word = text.Substring(0, array3[num3].Length);
							tiGangShowUtil3.correct_word = array3[num3];
							tiGangShowUtil3.start = 0;
							tiGangShowUtil3.end = array3[num3].Length;
							tiGangShowUtil_List.Add(tiGangShowUtil3);
						}
						num3++;
					}
				}
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
					if (tiGangShowUtil_List.Count <= 0)
					{
						lbTips.Text = "恭喜您，没有发现错误";
						MessageTipsUtil.Show(lbTips.Text);
					}
					else
					{
						lbError.Text = "共发现" + tiGangShowUtil_List.Count + "处错误";
						_rangeSelectContent.Paragraphs[1].Range.Select();
						nowCorrectPos = 0;
						setNowCorrectWord();
						MessageTipsUtil.Show("校对完成\r\n共发现" + tiGangShowUtil_List.Count + "处错误");
					}
				}
				MyAddin.boolBtnTigang = true;
				MyAddin.thisAddIn.InvalidateControl("btnTiGang");
			}
			catch (Exception)
			{
			}
		}

		public void setNowCorrectWord()
		{
			try
			{
				int count = tiGangShowUtil_List.Count;
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
				TiGangShowUtil tiGangShowUtil = tiGangShowUtil_List[nowCorrectPos];
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
				tbResult.Text = tiGangShowUtil.rangeSentece.Text;
				lbError.Text = tiGangShowUtil.error_word + "->";
				lbCorrect.Text = tiGangShowUtil.correct_word;
				setBackgroudColor(tiGangShowUtil.rangeSentece, tiGangShowUtil.start, tiGangShowUtil.end, tiGangShowUtil.error_word);
			}
			catch (Exception)
			{
			}
		}

		public void setBackgroudColor(Range rangeSentence, int start, int end, string strErrorWord)
		{
			_start = start;
			_end = end;
			_strErrorWord = strErrorWord;
			try
			{
				nowRangeSentence = rangeSentence;
				int start2 = rangeSentence.Start;
				int end2 = rangeSentence.End;
				rangeSentence.SetRange(rangeSentence.Start + start, rangeSentence.Start + end);
				TextBox textBox = tbResult;
				textBox.Text = textBox.Text + "\r\nstart:" + start + "--end:" + end;
				TextBox textBox2 = tbResult;
				textBox2.Text = textBox2.Text + "\r\nreturn_errorword:" + strErrorWord;
				TextBox textBox3 = tbResult;
				textBox3.Text = textBox3.Text + "\r\nlocal_get:" + rangeSentence.Text;
				rangeSentence.HighlightColorIndex = WdColorIndex.wdYellow;
				rangeSentence.SetRange(start2, end2);
				rangeSentence.Select();
			}
			catch (Exception)
			{
			}
		}

		private void resetBackgroundColor()
		{
			try
			{
				nowRangeSentence.HighlightColorIndex = WdColorIndex.wdAuto;
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
				tiGangShowUtil_List.RemoveAt(nowCorrectPos);
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
				TiGangShowUtil tiGangShowUtil = tiGangShowUtil_List[nowCorrectPos];
				correctErrorWord(tiGangShowUtil.rangeSentece, tiGangShowUtil.start, tiGangShowUtil.end, tiGangShowUtil.error_word, tiGangShowUtil.correct_word);
				resetBackgroundColor();
				tiGangShowUtil_List.RemoveAt(nowCorrectPos);
				setNowCorrectWord();
			}
			catch (Exception)
			{
			}
		}

		public void correctErrorWord(Range rangeSentence, int start, int end, string strErrorWord, string strRightWord)
		{
			_start = start;
			_end = end;
			_strErrorWord = strErrorWord;
			try
			{
				int start2 = rangeSentence.Start;
				int end2 = rangeSentence.End;
				rangeSentence.SetRange(rangeSentence.Start + start, rangeSentence.Start + end);
				TextBox textBox = tbResult;
				textBox.Text = textBox.Text + "\r\nstart:" + start + "--end:" + end;
				TextBox textBox2 = tbResult;
				textBox2.Text = textBox2.Text + "\r\nreturn_errorword:" + strErrorWord;
				TextBox textBox3 = tbResult;
				textBox3.Text = textBox3.Text + "\r\nlocal_get:" + rangeSentence.Text;
				rangeSentence.Text = strRightWord;
				rangeSentence.HighlightColorIndex = WdColorIndex.wdAuto;
				rangeSentence.SetRange(start2, end2);
				rangeSentence.Select();
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
			this.tableLayoutPanel1.TabIndex = 12;
			this.tbResult.Location = new System.Drawing.Point(80, 528);
			this.tbResult.Multiline = true;
			this.tbResult.Name = "tbResult";
			this.tbResult.Size = new System.Drawing.Size(190, 42);
			this.tbResult.TabIndex = 11;
			this.tbResult.Text = "正在联网校对中……";
			this.tbResult.Visible = false;
			base.Controls.Add(this.tableLayoutPanel1);
			base.Controls.Add(this.tbResult);
			base.Name = "Uc_TigangCorrect";
			base.Size = new System.Drawing.Size(600, 633);
			base.Load += new System.EventHandler(WordCorrect_Load);
			this.tableLayoutPanel1.ResumeLayout(false);
			this.tableLayoutPanel1.PerformLayout();
			base.ResumeLayout(false);
			base.PerformLayout();
		}
	}
}
