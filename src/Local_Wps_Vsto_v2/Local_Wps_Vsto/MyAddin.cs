using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Speech.Synthesis;
using System.Windows.Forms;
using System.Windows.Threading;
using AddInDesignerObjects;
using ICSharpCode.SharpZipLib.Zip;
using Office;
using Word;

namespace Local_Wps_Vsto
{
	public class MyAddin : IDTExtensibility2, _IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
	{
		public static Word.Application wordApp = null;

		public static object wpp;

		private bool boolBtnPlay = true;

		public CustomTaskPane CTP;

		private int intNo;

		public static ICTPFactory CustomTaskPaneCreater;

		private static string _label = "登录";

		private static string strLbLogin = "登录";

		private static bool boolLbLogin = true;

		private static string strLbReg = "注册";

		private static bool boolLbReg = true;

		private static string strLbUserMsg = "用户";

		private static bool boolLbUserMsg = false;

		private static string strLbImprove = "升级";

		private static bool boolLbImprove = false;

		public static bool boolBtnTigang = true;

		public static bool boolBtnJiaoGao = true;

		public static bool boolTbSucai = false;

		public static bool boolTbFanWen = false;

		public static bool boolTbSuperWriter = false;

		private bool boolShow = true;

		public static IRibbonUI thisAddIn;

		private BackgroundWorker BgWorker_Speaker = new BackgroundWorker();

		private BackgroundWorker BgWorker_PaiBan = new BackgroundWorker();

		public SpeechSynthesizer voiceSpeaker = new SpeechSynthesizer();

		private Range rangePre;

		private Range rangeNow;

		private Range rangeVoice;

		private bool boolPlay = true;

		private int intSpeakerPosition;

		private BackgroundWorker BgWorker_SetPage;

		private CustomTaskPane kefuPanel;

		private string strException = "";

		public CustomTaskPane updatePanel;

		public static CustomTaskPane loginPanel;

		public static CustomTaskPane regPanel;

		public static CustomTaskPane improvePanel;

		public static object missing = Type.Missing;

		private Timer timerWriterTips;

		private SuperWriterTips superWriterTips;

		private CustomTaskPane superWriterPanel;

		public int CurserPositionX;

		public int CurserPositionY;

		public int rangewidth;

		public int rangeheight;

		private int maxLength = 3;

		private bool boolWindowsLoaded;

		private int maxResultCount = 10;

		private int intLengthIndex;

		private int intCountIndex;

		private int intFromIndex = 1;

		private bool IsFromWeb = true;

		public static CustomTaskPane geshiPanel;

		public static CustomTaskPane fromZiYuanPanel;

		public static CustomTaskPane backPanel;

		public static CustomTaskPane rollBackPanel;

		private CustomTaskPane sucaiPanel;

		private CustomTaskPane newsucaiPanel;

		public CustomTaskPane fanwenPanel;

		public CustomTaskPane newfanwenPanel;

		private CustomTaskPane correctPanel;

		private Dispatcher _dispatcher;

		private CustomTaskPane tigangPanel;

		private string strFileName = "";

		public Dispatcher Dispatcher
		{
			get
			{
				return _dispatcher;
			}
		}

		public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
		{
			try
			{
				wpp = Application;
				wordApp = wpp as Word.Application;
			}
			catch (Exception)
			{
			}
		}

		public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
		{
			throw new NotImplementedException();
		}

		public void OnAddInsUpdate(ref Array custom)
		{
			throw new NotImplementedException();
		}

		public void OnStartupComplete(ref Array custom)
		{
			throw new NotImplementedException();
		}

		public void OnBeginShutdown(ref Array custom)
		{
			throw new NotImplementedException();
		}

		public string GetCustomUI(string RibbonID)
		{
			return Resource1.MyRibbon;
		}

		void ICustomTaskPaneConsumer.CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
		{
			try
			{
				CustomTaskPaneCreater = CTPFactoryInst;
			}
			catch (Exception)
			{
			}
		}

		public Bitmap GetRibbonImage(IRibbonControl ctrl)
		{
			try
			{
				switch (ctrl.Id)
				{
				case "btnLogin":
					return Resource1.u_login;
				case "btnReg":
					return Resource1.u_reg;
				case "btnUserMsg":
					if (UserUtil.IsVip())
					{
						return Resource1.u_forever;
					}
					return Resource1.u_free;
				case "btnImprove":
					if (UserUtil.IsVip())
					{
						return Resource1._continue;
					}
					return Resource1.improve;
				case "btnHongTou":
					return Resource1.hongtou;
				case "tbSuCai":
					return Resource1.search;
				case "tbFanWen":
					return Resource1.search;
				case "tbSuperWriter":
					return Resource1.writer;
				case "btnA4_New":
					return Resource1.page_a4;
				case "btnWps":
					return Resource1.paiban;
				case "menuWps":
					return Resource1.gaoji;
				case "btnInsert_One":
					return Resource1.yiji;
				case "btnInsert_Two":
					return Resource1.erji;
				case "btnInsert_Three":
					return Resource1.sanji;
				case "gFuHao":
					return Resource1.fuhao;
				case "menuPageNo":
					return Resource1.yema;
				case "btnAlignLeft":
					return Resource1.yema;
				case "btnAlignRight":
					return Resource1.yema;
				case "btnAlilgnCenter":
					return Resource1.yema;
				case "btnInsert_Date":
					return Resource1.riqi;
				case "btnInsert_HengYe":
					return Resource1.hengye;
				case "btnPlay":
					if (boolBtnPlay)
					{
						return Resource1.play;
					}
					return Resource1.pause;
				case "btnTiGang":
					return Resource1.tigang;
				case "btnJiaoGao":
					return Resource1.wenzhi;
				case "btnSave":
					return Resource1.save;
				case "btnSaveAsPdf":
					return Resource1.pdf;
				case "btnHelp":
					return Resource1.help;
				case "btnQQ":
					return Resource1.qq;
				case "btnKeFu":
					return Resource1.weichat_tans;
				}
			}
			catch (Exception)
			{
			}
			return null;
		}

		// Ribbon UI 回调：返回每个控件的显示文本。
		// "gwgs" 是 Ribbon XML 中定义的 tab id（见 Resource1.resx 中嵌入的 ribbon.xml）。
		// 这里把 tab 标题改成 "公文助手 <版本号>"，保证用户在 WPS 顶栏看到的是新品牌。
		// 其它 case 对应右上角"激活/登录"区域的按钮文案（已在 refreshLoginPanel 中统一）。
		public string GetLabel(IRibbonControl control)
		{
			try
			{
				switch (control.Id)
				{
				case "btnLogin":
					return strLbLogin;
				case "btnReg":
					return strLbReg;
				case "btnUserMsg":
					return strLbUserMsg;
				case "btnImprove":
					return strLbImprove;
				case "gwgs":
					return "公文助手 " + CommonConfig.strVersionCode;
				}
			}
			catch (Exception)
			{
			}
			return "按钮";
		}

		public bool GetVisible(IRibbonControl control)
		{
			try
			{
				switch (control.Id)
				{
				case "btnLogin":
					return boolLbLogin;
				case "btnReg":
					return boolLbReg;
				case "btnUserMsg":
					return boolLbUserMsg;
				case "btnImprove":
					return boolLbImprove;
				case "tbSuCai":
					return boolTbSucai;
				case "tbFanWen":
					return boolTbFanWen;
				case "tbSuperWriter":
					return boolTbSuperWriter;
				case "btnTiGang":
					return boolBtnTigang;
				case "btnJiaoGao":
					return boolBtnJiaoGao;
				}
			}
			catch (Exception)
			{
			}
			return true;
		}

		public static void refreshLoginPanel()
		{
			try
			{
				thisAddIn.InvalidateControl("btnUserMsg");
			}
			catch (Exception)
			{
			}
		}

		public void Click(IRibbonControl control)
		{
			try
			{
				_label = "登录窗口";
				thisAddIn.InvalidateControl("btnLogin");
				boolShow = false;
				thisAddIn.InvalidateControl("btnImprove");
			}
			catch (Exception ex)
			{
				if (CommonConfig.IS_DEBUG)
				{
					MessageBox.Show(ex.ToString());
				}
			}
		}

		public bool onLoginVisible(IRibbonControl control)
		{
			return boolShow;
		}

		public void MyAddInInitialize(IRibbonUI Ribbon)
		{
			thisAddIn = Ribbon;
			RibbonDesign_Load();
		}

		public void RibbonDesign_Load()
		{
			try
			{
				BindBackgroudWorker();
				checkFolder();
				checkDataBase();
				CheckImproveInfo_Local();
				loadConfigInfo();
				boolWindowsLoaded = true;
			}
			catch (Exception ex)
			{
				if (CommonConfig.IS_DEBUG)
				{
					MessageBox.Show(ex.ToString());
				}
			}
		}

		public void DoWork_Handler_Speaker(object sender, DoWorkEventArgs args)
		{
			try
			{
				rangePre = null;
				rangeNow = null;
				Sentences sentences = rangeVoice.Sentences;
				for (int i = 1; i <= sentences.Count; i++)
				{
					rangePre = rangeNow;
					if (rangePre != null)
					{
						rangePre.HighlightColorIndex = WdColorIndex.wdAuto;
					}
					rangeNow = sentences[i];
					rangeNow.HighlightColorIndex = WdColorIndex.wdYellow;
					wordApp.Selection.Start = rangeNow.Start;
					wordApp.Selection.End = rangeNow.Start;
					string text = sentences[i].Text;
					voiceSpeaker.Speak(text);
				}
			}
			catch (Exception)
			{
			}
		}

		public void RunWorkerCompleted_Handler_Speaker(object sender, RunWorkerCompletedEventArgs args)
		{
			try
			{
				if (rangeNow != null)
				{
					rangeNow.HighlightColorIndex = WdColorIndex.wdAuto;
				}
				boolBtnPlay = true;
				thisAddIn.InvalidateControl("btnPlay");
				boolPlay = true;
			}
			catch (Exception)
			{
			}
		}

		public void gFuHao_Click(IRibbonControl ctrl, string selectId, int selectIndex)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				string text = "";
				switch (selectId)
				{
				case "__id10":
					text = "〔〕";
					break;
				case "__id11":
					text = "《》";
					break;
				case "__id12":
					text = "〈〉";
					break;
				case "__id13":
					text = "·";
					break;
				case "__id14":
					text = "☆";
					break;
				case "__id15":
					text = "＋";
					break;
				case "__id16":
					text = "－";
					break;
				case "__id17":
					text = "×";
					break;
				case "__id18":
					text = "÷";
					break;
				case "__id19":
					text = "≈";
					break;
				case "__id20":
					text = "①";
					break;
				case "__id21":
					text = "②";
					break;
				case "__id22":
					text = "③";
					break;
				case "__id23":
					text = "④";
					break;
				case "__id24":
					text = "⑤";
					break;
				case "__id25":
					text = "⑥";
					break;
				case "__id26":
					text = "⑦";
					break;
				case "__id27":
					text = "⑧";
					break;
				case "__id28":
					text = "⑨";
					break;
				case "__id29":
					text = "⑩";
					break;
				case "__id30":
					text = "★";
					break;
				case "__id31":
					text = "℃";
					break;
				case "__id32":
					text = "‰";
					break;
				case "__id33":
					text = "㎡";
					break;
				case "__id34":
					text = "m³";
					break;
				}
				wordApp.Selection.InsertAfter(text);
			}
			catch (Exception)
			{
			}
		}

		public void btnPlay_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				rangeVoice = null;
				if (boolPlay)
				{
					Document activeDocument = wordApp.ActiveDocument;
					try
					{
						if (wordApp.Selection.End == wordApp.Selection.Start)
						{
							Selection selection = wordApp.Selection;
							int end = wordApp.Selection.End;
							object Start = Type.Missing;
							object End = Type.Missing;
							int end2 = activeDocument.Range(ref Start, ref End).End;
							End = end;
							Start = Type.Missing;
							rangeVoice = activeDocument.Range(ref End, ref Start);
						}
						else
						{
							rangeVoice = wordApp.Selection.Range;
						}
					}
					catch (Exception)
					{
						object Start = Type.Missing;
						object End = Type.Missing;
						rangeVoice = activeDocument.Range(ref Start, ref End);
					}
					if (rangeVoice.Text.Replace(Convert.ToChar(10).ToString(), "").Replace(Convert.ToChar(13).ToString(), "").Replace(" ", "")
						.Replace(" ", "")
						.Length == 0)
					{
						MessageTipsUtil.Show("没有发现可朗读文字");
						return;
					}
					boolBtnPlay = false;
					thisAddIn.InvalidateControl("btnPlay");
					BgWorker_Speaker.RunWorkerAsync();
					boolPlay = !boolPlay;
				}
				else
				{
					voiceSpeaker.SpeakAsyncCancelAll();
					if (rangeNow != null)
					{
						rangeNow.HighlightColorIndex = WdColorIndex.wdAuto;
					}
					boolBtnPlay = true;
					thisAddIn.InvalidateControl("btnPlay");
					boolPlay = !boolPlay;
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnStop_Click(IRibbonControl ctrl)
		{
			try
			{
				boolPlay = false;
				boolBtnPlay = true;
				thisAddIn.InvalidateControl("btnPlay");
				intSpeakerPosition = 0;
			}
			catch (Exception)
			{
			}
		}

		public void btnWps_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				MessageTipsUtil.Show("自动排版中……", 100000);
				if (BgWorker_PaiBan.IsBusy)
				{
					BgWorker_PaiBan.CancelAsync();
				}
				BgWorker_PaiBan = new BackgroundWorker();
				BgWorker_PaiBan.WorkerReportsProgress = false;
				BgWorker_PaiBan.WorkerSupportsCancellation = true;
				BgWorker_PaiBan.DoWork += DoWork_Handler_PaiBan;
				BgWorker_PaiBan.RunWorkerCompleted += RunWorkerCompleted_Handler_PaiBan;
				BgWorker_PaiBan.RunWorkerAsync();
			}
			catch (Exception)
			{
			}
		}

		public void DoWork_Handler_PaiBan(object sender, DoWorkEventArgs args)
		{
			try
			{
				DocWpsUtil.AutoPaiBan();
			}
			catch (Exception)
			{
			}
		}

		public void RunWorkerCompleted_Handler_PaiBan(object sender, RunWorkerCompletedEventArgs args)
		{
			MessageTipsUtil.Show("自动排版完成", 1500);
		}

		public void wordProcess()
		{
			try
			{
				Word.Application application = wordApp;
				application.ActiveDocument.PageSetup.LineNumbering.Active = 0;
				application.ActiveDocument.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
				application.ActiveDocument.PageSetup.TopMargin = application.CentimetersToPoints(float.Parse("2.54"));
				application.ActiveDocument.PageSetup.BottomMargin = application.CentimetersToPoints(float.Parse("2.54"));
				application.ActiveDocument.PageSetup.LeftMargin = application.CentimetersToPoints(float.Parse("3.17"));
				application.ActiveDocument.PageSetup.RightMargin = application.CentimetersToPoints(float.Parse("3.17"));
				application.ActiveDocument.PageSetup.Gutter = application.CentimetersToPoints(float.Parse("0"));
				application.ActiveDocument.PageSetup.HeaderDistance = application.CentimetersToPoints(float.Parse("1.5"));
				application.ActiveDocument.PageSetup.FooterDistance = application.CentimetersToPoints(float.Parse("1.75"));
				application.ActiveDocument.PageSetup.PageWidth = application.CentimetersToPoints(float.Parse("21"));
				application.ActiveDocument.PageSetup.PageHeight = application.CentimetersToPoints(float.Parse("29.7"));
				application.ActiveDocument.PageSetup.FirstPageTray = WdPaperTray.wdPrinterDefaultBin;
				application.ActiveDocument.PageSetup.OtherPagesTray = WdPaperTray.wdPrinterDefaultBin;
				application.ActiveDocument.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage;
				application.ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter = 0;
				application.ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = 0;
				application.ActiveDocument.PageSetup.VerticalAlignment = WdVerticalAlignment.wdAlignVerticalTop;
				application.ActiveDocument.PageSetup.SuppressEndnotes = 0;
				application.ActiveDocument.PageSetup.MirrorMargins = 0;
				application.ActiveDocument.PageSetup.TwoPagesOnOne = false;
				application.ActiveDocument.PageSetup.BookFoldPrinting = false;
				application.ActiveDocument.PageSetup.BookFoldRevPrinting = false;
				application.ActiveDocument.PageSetup.BookFoldPrintingSheets = 1;
				application.ActiveDocument.PageSetup.GutterPos = WdGutterStyle.wdGutterPosLeft;
				application.ActiveDocument.PageSetup.LinesPage = 40f;
				application.ActiveDocument.PageSetup.LayoutMode = WdLayoutMode.wdLayoutModeLineGrid;
			}
			catch (Exception)
			{
			}
		}

		public void btnA4_Single_Click(IRibbonControl ctrl)
		{
			try
			{
				DocWpsUtil.setPage_A4_Single();
			}
			catch (Exception)
			{
			}
		}

		public void InsertFooter(string footer)
		{
			try
			{
				Word.Application application = wordApp;
				if (application.ActiveWindow.ActivePane.View.Type == WdViewType.wdNormalView || application.ActiveWindow.ActivePane.View.Type == WdViewType.wdOutlineView)
				{
					application.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView;
				}
				application.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
				application.Application.Selection.HeaderFooter.LinkToPrevious = false;
				application.Application.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				application.ActiveWindow.ActivePane.Selection.InsertAfter(footer);
				application.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
			}
			catch (Exception)
			{
			}
		}

		public void btnDelete_HeaderLine_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				wordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
				wordApp.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
				wordApp.ActiveWindow.ActivePane.Selection.Borders[WdBorderType.wdBorderBottom].Visible = false;
				wordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument;
			}
			catch (Exception)
			{
			}
		}

		public void InsertPageNumber(string strType, bool bHeader)
		{
			try
			{
				object PageNumberAlignment = WdPageNumberAlignment.wdAlignPageNumberCenter;
				object FirstPage = bHeader;
				WdHeaderFooterIndex index = WdHeaderFooterIndex.wdHeaderFooterPrimary;
				switch (strType.ToLower().ToString())
				{
				case "Center":
					PageNumberAlignment = WdPageNumberAlignment.wdAlignPageNumberCenter;
					break;
				case "Right":
					PageNumberAlignment = WdPageNumberAlignment.wdAlignPageNumberRight;
					break;
				case "Left":
					PageNumberAlignment = WdPageNumberAlignment.wdAlignPageNumberLeft;
					break;
				}
				wordApp.ActiveDocument.Sections[1].Footers[index].PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic;
				wordApp.ActiveDocument.Sections[1].Footers[index].PageNumbers.Add(ref PageNumberAlignment, ref FirstPage);
			}
			catch (Exception)
			{
			}
		}

		public void InsertPageFooterNumber()
		{
			object Text = Missing.Value;
			try
			{
				wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
				wordApp.Selection.WholeStory();
				wordApp.Selection.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
				wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
				wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
				wordApp.Selection.TypeText("第");
				object Type = WdFieldType.wdFieldPage;
				wordApp.Selection.Fields.Add(wordApp.Selection.Range, ref Type, ref Text, ref Text);
				wordApp.Selection.TypeText("页/共");
				object Type2 = WdFieldType.wdFieldNumPages;
				wordApp.Selection.Fields.Add(wordApp.Selection.Range, ref Type2, ref Text, ref Text);
				wordApp.Selection.TypeText("页");
				wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
			}
			catch (Exception)
			{
			}
		}

		public void insertNewPageNumber_Double()
		{
			try
			{
				Word.Application application = wordApp;
				if (application.ActiveWindow.ActivePane.View.Type == WdViewType.wdNormalView || application.ActiveWindow.ActivePane.View.Type == WdViewType.wdOutlineView)
				{
					application.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView;
				}
				application.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
				application.Application.Selection.HeaderFooter.LinkToPrevious = false;
				Range range = application.Application.Selection.HeaderFooter.Range;
				object Unit = System.Type.Missing;
				object Count = System.Type.Missing;
				range.Delete(ref Unit, ref Count);
				application.ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter = -1;
				application.ActiveWindow.ActivePane.Selection.Font.Size = 14f;
				application.ActiveWindow.ActivePane.Selection.Font.Name = "宋体";
				object Text = Missing.Value;
				object Type = WdFieldType.wdFieldPage;
				application.ActiveWindow.View.SeekView = WdSeekView.wdSeekEvenPagesFooter;
				application.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
				application.ActiveWindow.ActivePane.Selection.Font.Size = 14f;
				application.ActiveWindow.ActivePane.Selection.Font.Name = "宋体";
				application.Application.Selection.TypeText("—");
				application.Application.Selection.Fields.Add(application.Selection.Range, ref Type, ref Text, ref Text);
				application.Application.Selection.TypeText("—");
				application.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
				application.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryFooter;
				application.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
				application.ActiveWindow.ActivePane.Selection.Font.Size = 14f;
				application.ActiveWindow.ActivePane.Selection.Font.Name = "宋体";
				application.Application.Selection.TypeText("—");
				application.Application.Selection.Fields.Add(application.Selection.Range, ref Type, ref Text, ref Text);
				application.Application.Selection.TypeText("—");
				application.ActiveDocument.PageSetup.HeaderDistance = application.CentimetersToPoints(float.Parse("1.5"));
				application.ActiveDocument.PageSetup.FooterDistance = application.CentimetersToPoints(float.Parse("2.5"));
				application.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
				wordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
				application.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
				application.ActiveWindow.ActivePane.Selection.Borders[WdBorderType.wdBorderBottom].Visible = false;
				application.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
			}
			catch (Exception)
			{
			}
		}

		public void btnA4_Double_Click(IRibbonControl ctrl)
		{
			try
			{
				insertNewPageNumber_Double();
			}
			catch (Exception)
			{
			}
		}

		public void DoWork_Handler_SetPage(object sender, DoWorkEventArgs args)
		{
			try
			{
				DocWpsUtil.setPage_A4_Single();
			}
			catch (Exception)
			{
			}
		}

		public void RunWorkerCompleted_SetPage(object sender, RunWorkerCompletedEventArgs args)
		{
			MessageTipsUtil.Show("设置成功");
		}

		public void btnA4_New_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				DocWpsUtil.setPage_A4_Single();
				MessageTipsUtil.Show("设置成功");
			}
			catch (Exception)
			{
			}
		}

		public void btnMw_D_kongge_Click(IRibbonControl ctrl)
		{
			if (!UserUtil.HasLogin())
			{
				MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				return;
			}
			if (!UserUtil.IsVip())
			{
				MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				return;
			}
			try
			{
				Document activeDocument = wordApp.ActiveDocument;
				object Start = Type.Missing;
				object End = Type.Missing;
				Range selectionContent = activeDocument.Range(ref Start, ref End);
				DocWpsUtil.FindAndReplace(selectionContent, " ", "");
				DocWpsUtil.FindAndReplace(selectionContent, "\u2003", "");
				DocWpsUtil.FindAndReplace(selectionContent, "\u3000\u3000", "");
			}
			catch (Exception)
			{
			}
		}

		public void btnMw_D_kongge_S_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				Range range = wordApp.Application.Selection.Range;
				Words words = wordApp.Selection.Words;
				DocWpsUtil.FindAndReplace(words.Application.Selection, " ", "");
				DocWpsUtil.FindAndReplace(words.Application.Selection, "\u2003", "");
				DocWpsUtil.FindAndReplace(words.Application.Selection, "\u3000\u3000", "");
			}
			catch (Exception)
			{
			}
		}

		public void btnMw_D_konghang_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				Document activeDocument = wordApp.ActiveDocument;
				object Start = Type.Missing;
				object End = Type.Missing;
				DocWpsUtil.DeleteBlankLine(activeDocument.Range(ref Start, ref End), wordApp);
			}
			catch (Exception)
			{
			}
		}

		public void btnMw_D_konghang_Select_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					DocWpsUtil.DeleteBlankLine(wordApp.Selection.Words.Application.Selection.Range, wordApp);
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnMw_D_konghang_FuHao_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				Document activeDocument = wordApp.ActiveDocument;
				object Start = Type.Missing;
				object End = Type.Missing;
				Range selectionContent = activeDocument.Range(ref Start, ref End);
				DocWpsUtil.FindAndReplace(selectionContent, ",", "，");
				DocWpsUtil.FindAndReplace(selectionContent, "(", "（");
				DocWpsUtil.FindAndReplace(selectionContent, ")", "）");
				DocWpsUtil.FindAndReplace(selectionContent, "!", "！");
				DocWpsUtil.FindAndReplace(selectionContent, ":", "：");
				DocWpsUtil.FindAndReplace(selectionContent, ";", "；");
			}
			catch (Exception)
			{
			}
		}

		public void btnWps_YiJi_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					WpsHelper.setBiaoTi_One_JuZhi(wordApp.Selection.Words.Application.Selection.Range);
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnWps_ErJi_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					WpsHelper.setBiaoTi_Two_JuZhi(wordApp.Selection.Words.Application.Selection.Range);
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnWps_SanJi_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					WpsHelper.setBiaoTi_Three_JuZhi(wordApp.Selection.Words.Application.Selection.Range);
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnWps_NeiRong_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					WpsHelper.setBiaoTi_Content_JuZhi(wordApp.Selection.Words.Application.Selection.Range);
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnWps_Top_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					WpsHelper.setBiaoTi_TopTitle_JuZhi(wordApp.Selection.Words.Application.Selection.Range);
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnWps_Time_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					WpsHelper.setBiaoTi_TopAuthor_JuZhi(wordApp.Selection.Words.Application.Selection.Range);
				}
			}
			catch (Exception)
			{
			}
		}

		public void gFuHao_ButtonClick(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
			}
			catch (Exception)
			{
			}
			Document activeDocument = wordApp.ActiveDocument;
			object Start = Type.Missing;
			object End = Type.Missing;
			Range selectionContent = activeDocument.Range(ref Start, ref End);
			try
			{
				switch (ctrl.Id)
				{
				case "btnReplace_1":
					DocWpsUtil.FindAndReplace(selectionContent, "【", "〔");
					DocWpsUtil.FindAndReplace(selectionContent, "】", "〕");
					break;
				case "btnReplace_2":
					DocWpsUtil.FindAndReplace(selectionContent, "[", "〔");
					DocWpsUtil.FindAndReplace(selectionContent, "]", "〕");
					break;
				case "btnReplace_3":
					DocWpsUtil.FindAndReplace(selectionContent, "<<", "《");
					DocWpsUtil.FindAndReplace(selectionContent, ">>", "》");
					break;
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnKeFu_Click(IRibbonControl ctrl)
		{
			try
			{
				kefuPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_Help_KeFu", "联系客服", missing);
				kefuPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
				kefuPanel.Width = 600;
				kefuPanel.Height = 500;
				kefuPanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
				kefuPanel.Visible = true;
			}
			catch (Exception)
			{
			}
		}

		public void loadConfigInfo()
		{
			try
			{
				ConfigLoad configLoad = new ConfigLoad(CommonConfig.CONFIG_FILE_INI);
				intLengthIndex = configLoad.GetIntValue("word_length");
				intCountIndex = configLoad.GetIntValue("word_count");
				intFromIndex = configLoad.GetIntValue("word_from");
				maxResultCount = (intCountIndex + 1) * 5;
				maxLength = (intLengthIndex + 1) * 3;
				if (intFromIndex == 0)
				{
					IsFromWeb = true;
				}
				else
				{
					IsFromWeb = false;
				}
				thisAddIn.InvalidateControl("dropLength");
				thisAddIn.InvalidateControl("dropCount");
				thisAddIn.InvalidateControl("dropFrom");
			}
			catch (Exception)
			{
			}
		}

		public void checkDataBase()
		{
			try
			{
				string text = CommonConfig.strBaseFolder_Common + "\\gwgs.sdf";
				string text2 = CommonConfig.strBaseFolder + "\\conf\\gwgs.sdf";
				string text3 = "C:\\公文高手WPS插件\\gwgs.sdf";
				if (!File.Exists(text))
				{
					if (File.Exists(text3))
					{
						new FileInfo(text3).MoveTo(text);
					}
					else if (File.Exists(text2))
					{
						new FileInfo(text2).CopyTo(text);
					}
				}
				if (!File.Exists(text))
				{
					MessageBox.Show("数据库损坏，重新安装插件！");
				}
			}
			catch (Exception ex)
			{
				if (CommonConfig.IS_DEBUG)
				{
					MessageBox.Show(ex.ToString());
				}
			}
		}

		public void checkDataBase_Old()
		{
			try
			{
				string text = CommonConfig.strBaseFolder + "\\gwgs.sdf";
				string text2 = CommonConfig.strBaseFolder + "\\conf\\gwgs.sdf";
				if (File.Exists(SQLiteHelper.strGwgsDbFile) || File.Exists(text))
				{
					return;
				}
				if (!File.Exists(text2))
				{
					string path = CommonConfig.strBaseFolder + CommonConfig.Init_Table;
					if (File.Exists(path))
					{
						string queryString = File.ReadAllText(path);
						SQLiteHelper sQLiteHelper = new SQLiteHelper();
						sQLiteHelper.ExecuteQuery(queryString);
						sQLiteHelper.CloseConnection();
						string path2 = CommonConfig.strBaseFolder + CommonConfig.Init_Data;
						if (File.Exists(path2))
						{
							string queryString2 = File.ReadAllText(path2);
							SQLiteHelper sQLiteHelper2 = new SQLiteHelper();
							sQLiteHelper2.ExecuteQuery(queryString2);
							sQLiteHelper2.CloseConnection();
						}
					}
					else
					{
						MessageBox.Show("数据库损坏，请到网站下载最新版本！");
					}
				}
				else
				{
					new FileInfo(text2).CopyTo(text);
				}
			}
			catch (Exception)
			{
			}
		}

		public void checkFolder()
		{
			try
			{
				string[] array = new string[7]
				{
					CommonConfig.strBaseFolder + "\\cache",
					CommonConfig.strBaseFolder + "\\cache\\fw",
					CommonConfig.strBaseFolder + "\\cache\\localrtf",
					CommonConfig.strBaseFolder + "\\cache\\rtf",
					CommonConfig.strBaseFolder + "\\cache\\word",
					CommonConfig.strBaseFolder + "\\update",
					CommonConfig.strBaseFolder_Common
				};
				foreach (string path in array)
				{
					if (!Directory.Exists(path))
					{
						Directory.CreateDirectory(path);
					}
				}
			}
			catch (Exception)
			{
			}
		}

		public void CheckLoginInfo()
		{
		}

		public static void CheckImproveInfo_Local()
		{
			try
			{
				UserUtil.readActiviedCode();
				boolLbImprove = true;
				boolLbUserMsg = true;
				// 原版分支：if (IsVip) 显示用户名; else 显示"尚未激活"。
				// 本版本：UserUtil.IsVip() 已 patch 为恒 true，分支结果实际相同；
				// 两个 strLbUserMsg 都赋为 "公文助手"，保证 UI 完全一致，并避免
				// 任何未来 patch 移除后仍触发"尚未激活"提示。
				if (UserUtil.IsVip())
				{
					strLbUserMsg = "公文助手";
				}
				else
				{
					strLbUserMsg = "公文助手";
				}
				refreshLoginPanel();
			}
			catch (Exception ex)
			{
				if (CommonConfig.IS_DEBUG)
				{
					MessageBox.Show(ex.ToString());
				}
				MessageBox.Show(ex.ToString());
			}
		}

		public void BindBackgroudWorker()
		{
			try
			{
				BgWorker_Speaker.WorkerReportsProgress = false;
				BgWorker_Speaker.WorkerSupportsCancellation = true;
				BgWorker_Speaker.DoWork += DoWork_Handler_Speaker;
				BgWorker_Speaker.RunWorkerCompleted += RunWorkerCompleted_Handler_Speaker;
				BgWorker_PaiBan.WorkerReportsProgress = false;
				BgWorker_PaiBan.WorkerSupportsCancellation = true;
				BgWorker_PaiBan.DoWork += DoWork_Handler_PaiBan;
				BgWorker_PaiBan.RunWorkerCompleted += RunWorkerCompleted_Handler_PaiBan;
			}
			catch (Exception)
			{
			}
		}

		public static void DeleteDir(string strFolder)
		{
			try
			{
				new DirectoryInfo(strFolder).Attributes = (FileAttributes)0;
				File.SetAttributes(strFolder, FileAttributes.Normal);
				if (!Directory.Exists(strFolder))
				{
					return;
				}
				string[] fileSystemEntries = Directory.GetFileSystemEntries(strFolder);
				foreach (string text in fileSystemEntries)
				{
					if (File.Exists(text))
					{
						File.Delete(text);
						Console.WriteLine(text);
					}
					else
					{
						DeleteDir(text);
					}
				}
				Directory.Delete(strFolder);
			}
			catch (Exception)
			{
			}
		}

		private static void UnZipFile(string zipFilePath, string targetDir)
		{
			try
			{
				new FastZip(new FastZipEvents()).ExtractZip(zipFilePath, targetDir, "");
			}
			catch (Exception)
			{
			}
		}

		public void btnQQ_Click(IRibbonControl ctrl)
		{
			try
			{
				Process.Start(UrlUtil.strGoToQqUrl());
			}
			catch (Exception)
			{
			}
		}

		public void btnHelp_Click(IRibbonControl ctrl)
		{
			try
			{
				Process.Start(UrlUtil.strGoToHelpUrl());
			}
			catch (Exception)
			{
			}
		}

		public void btnLogin_Click(IRibbonControl ctrl)
		{
			btnLogin_Click();
		}

		public static void btnLogin_Click()
		{
			try
			{
				loginPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_Login", "欢迎登陆", missing);
				loginPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
				loginPanel.Width = 600;
				loginPanel.Height = 500;
				loginPanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
				loginPanel.Visible = true;
			}
			catch (Exception ex)
			{
				if (CommonConfig.IS_DEBUG)
				{
					MessageBox.Show(ex.ToString());
				}
			}
		}

		public void btnReg_Click(IRibbonControl ctrl)
		{
			btnReg_Click();
		}

		public static void btnReg_Click()
		{
			try
			{
				try
				{
					regPanel.Visible = false;
				}
				catch (Exception)
				{
				}
				regPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_Reg", "欢迎注册", missing);
				regPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
				regPanel.Width = 600;
				regPanel.Height = 500;
				regPanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
				regPanel.Visible = true;
			}
			catch (Exception ex2)
			{
				if (CommonConfig.IS_DEBUG)
				{
					MessageBox.Show(ex2.ToString());
				}
			}
		}

		public void btnUserMsg_Click(IRibbonControl ctrl)
		{
			if (UserUtil.IsVip())
			{
				MessageTipsUtil.Show("软件已激活");
				return;
			}
			try
			{
				improvePanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_Improve", "软件激活", missing);
				improvePanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
				improvePanel.Width = 710;
				improvePanel.Height = 610;
				improvePanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
				improvePanel.Visible = true;
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}

		public void btnImprove_Click(IRibbonControl ctrl)
		{
			btnImprove_Click();
		}

		public void btnImprove_Click()
		{
		}

		public void btnSuCai_New_Click(IRibbonControl ctrl)
		{
		}

		public void tbSuperWriter_Click(IRibbonControl ctrl, bool pressed)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					boolTbSuperWriter = false;
					thisAddIn.InvalidateControl("tbSuperWriter");
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					boolTbSuperWriter = false;
					thisAddIn.InvalidateControl("tbSuperWriter");
					return;
				}
			}
			catch (Exception)
			{
			}
			int width = Screen.PrimaryScreen.Bounds.Width;
			int height = Screen.PrimaryScreen.Bounds.Height;
			Selection selection = wordApp.Selection;
			int start = selection.Start;
			int end = selection.End;
			try
			{
				if (pressed)
				{
					boolTbSuperWriter = true;
					superWriterPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.SuperWriterTips", "写作提词器", missing);
					superWriterTips = (SuperWriterTips)(dynamic)superWriterPanel.ContentControl;
					superWriterTips.ParentRibbon = this;
					superWriterPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
					superWriterPanel.Width = 510;
					superWriterPanel.Height = 510;
					superWriterPanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
					new ComAwareEventInfo(typeof(_CustomTaskPaneEvents_Event), "VisibleStateChange").AddEventHandler(superWriterPanel, new _CustomTaskPaneEvents_VisibleStateChangeEventHandler(SuperWriterPanel_VisibleStateChange));
					superWriterPanel.Visible = true;
					if (timerWriterTips == null)
					{
						timerWriterTips = new Timer();
						timerWriterTips.Enabled = true;
						timerWriterTips.Interval = CommonConfig.TIME_DELAY;
						timerWriterTips.Tick += Timer_TipsChanged;
					}
					else
					{
						timerWriterTips.Stop();
					}
					timerWriterTips.Start();
					return;
				}
				boolTbSuperWriter = false;
				try
				{
					superWriterPanel.Visible = false;
				}
				catch (Exception)
				{
				}
				try
				{
					timerWriterTips.Stop();
				}
				catch (Exception)
				{
				}
			}
			catch (Exception ex4)
			{
				if (CommonConfig.IS_DEBUG)
				{
					MessageBox.Show(ex4.ToString());
				}
			}
		}

		private void SuperWriterPanel_VisibleStateChange(CustomTaskPane CustomTaskPaneInst)
		{
			try
			{
				if (!CustomTaskPaneInst.Visible)
				{
					boolTbSuperWriter = false;
					thisAddIn.InvalidateControl("tbSuperWriter");
				}
			}
			catch (Exception)
			{
			}
		}

		public void Timer_TipsChanged(object sender, EventArgs e)
		{
			try
			{
				DocumentContextOnUpdate_New();
			}
			catch (Exception)
			{
			}
		}

		public void DocumentContextOnUpdate_New(bool boolRefresh = false)
		{
			try
			{
				Range range = wordApp.Selection.Range;
				wordApp.ActiveDocument.ActiveWindow.GetPoint(out CurserPositionX, out CurserPositionY, out rangewidth, out rangeheight, range);
				int curserPositionX = CurserPositionX;
				int rangewidth2 = rangewidth;
				int curserPositionY = CurserPositionY;
				int rangeheight2 = rangeheight;
				string text = range.Text;
				string text2 = "";
				string text3 = "";
				if (wordApp.Selection.End == wordApp.Selection.Start)
				{
					text2 = range.Sentences[1].Text.ToString();
					text2 = range.Characters[1].Text.ToString();
					object obj = 0;
					object obj2 = wordApp.Selection.End;
					Document activeDocument = wordApp.ActiveDocument;
					object Start = obj;
					object End = obj2;
					text2 = activeDocument.Range(ref Start, ref End).Text;
					if (text2 == null)
					{
						text2 = "";
					}
					char[] separator = new char[17]
					{
						',', '，', '、', '；', ';', '.', '。', '!', '！', ':',
						'：', '\n', '\r', '\n', ' ', ' ', '…'
					};
					string[] array = text2.Split(separator);
					int num = array.Length;
					if (num > 0)
					{
						text3 = array[num - 1];
						while (text3.Length == 0 && num > 0)
						{
							text3 = array[num - 1];
							num--;
						}
						if (!string.IsNullOrEmpty(text3) && text3.Length > 0)
						{
							text3 = getLastSubString(text3, maxLength);
							superWriterTips.setKeyWord(text3, maxResultCount, IsFromWeb, boolRefresh);
						}
					}
				}
				else
				{
					text3 = range.Text.ToString();
					if (text3.Length > 0)
					{
						superWriterTips.setKeyWord(text3, maxResultCount, IsFromWeb, boolRefresh);
					}
				}
				if (text == null)
				{
					text = "";
				}
			}
			catch (Exception)
			{
			}
		}

		public string getLastSubString(string strWord, int _maxLength)
		{
			if (strWord == null)
			{
				return "";
			}
			if (strWord.Length <= _maxLength)
			{
				return strWord;
			}
			return strWord.Substring(strWord.Length - _maxLength);
		}

		public void dropCount_SelectionChanged(IRibbonControl ctrl, string selectId, int intSelectIndex)
		{
			if (!boolWindowsLoaded)
			{
				return;
			}
			try
			{
				maxResultCount = (intSelectIndex + 1) * 5;
				if (boolTbSuperWriter)
				{
					DocumentContextOnUpdate_New(true);
				}
			}
			catch (Exception)
			{
			}
			try
			{
				if (boolWindowsLoaded)
				{
					ConfigSet configSet = new ConfigSet(CommonConfig.CONFIG_FILE_INI);
					configSet.SetConfigValue("word_count", intSelectIndex.ToString());
					configSet.WriteConfigToFile();
				}
			}
			catch (Exception)
			{
			}
		}

		public void rxitemddSelectSheet_getItemId(IRibbonControl ctrl, int index)
		{
		}

		public int dorpLength_GetSelectedItemIndex(IRibbonControl ctrl)
		{
			int result = 0;
			try
			{
				switch (ctrl.Id)
				{
				case "dropLength":
					result = intLengthIndex;
					return result;
				case "dropCount":
					result = intCountIndex;
					return result;
				case "dropFrom":
					result = intFromIndex;
					return result;
				default:
					return result;
				}
			}
			catch (Exception)
			{
				return result;
			}
		}

		public void dropLength_SelectionChanged(IRibbonControl ctrl, string selectId, int intSelectIndex)
		{
			if (!boolWindowsLoaded)
			{
				return;
			}
			try
			{
				maxLength = (intSelectIndex + 1) * 3;
				if (boolTbSuperWriter)
				{
					DocumentContextOnUpdate_New(true);
				}
			}
			catch (Exception)
			{
			}
			try
			{
				if (boolWindowsLoaded)
				{
					ConfigSet configSet = new ConfigSet(CommonConfig.CONFIG_FILE_INI);
					configSet.SetConfigValue("word_length", intSelectIndex.ToString());
					configSet.WriteConfigToFile();
				}
			}
			catch (Exception)
			{
			}
		}

		public void dropFrom_SelectionChanged(IRibbonControl ctrl, string selectId, int intSelectIndex)
		{
			if (!boolWindowsLoaded)
			{
				return;
			}
			try
			{
				if (intSelectIndex == 0)
				{
					IsFromWeb = true;
				}
				else
				{
					IsFromWeb = false;
				}
				if (boolTbSuperWriter)
				{
					DocumentContextOnUpdate_New(true);
				}
			}
			catch (Exception)
			{
			}
			try
			{
				if (boolWindowsLoaded)
				{
					ConfigSet configSet = new ConfigSet(CommonConfig.CONFIG_FILE_INI);
					configSet.SetConfigValue("word_from", intSelectIndex.ToString());
					configSet.WriteConfigToFile();
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnSave_Click(IRibbonControl ctrl)
		{
			try
			{
				wordApp.ActiveDocument.Save();
				MessageTipsUtil.Show("保存成功", 500);
			}
			catch (Exception)
			{
			}
		}

		public void btnSaveAsPdf_Click(IRibbonControl ctrl)
		{
			try
			{
				Word.Application wordApp2 = wordApp;
				SaveFileDialog saveFileDialog = new SaveFileDialog();
				saveFileDialog.Filter = "PDF文件(*.pdf)|*.pdf";
				saveFileDialog.InitialDirectory = wordApp.ActiveDocument.Path;
				saveFileDialog.FileName = Path.GetFileNameWithoutExtension(wordApp.ActiveDocument.Name);
				saveFileDialog.DefaultExt = "pdf";
				saveFileDialog.AddExtension = true;
				if (saveFileDialog.ShowDialog() == DialogResult.OK)
				{
					Document activeDocument = wordApp.ActiveDocument;
					string fileName = saveFileDialog.FileName;
					object FixedFormatExtClassPtr = Type.Missing;
					activeDocument.ExportAsFixedFormat(fileName, WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument, 1, 1, WdExportItem.wdExportDocumentContent, false, true, WdExportCreateBookmarks.wdExportCreateNoBookmarks, true, true, false, ref FixedFormatExtClassPtr);
					MessageTipsUtil.Show("保存Pdf成功");
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnHongTou_Click(IRibbonControl ctrl)
		{
			try
			{
				geshiPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_GeShi", "红头模板", missing);
				geshiPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
				geshiPanel.Width = 1000;
				geshiPanel.Height = 655;
				geshiPanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
				geshiPanel.Visible = true;
			}
			catch (Exception)
			{
			}
		}

		public void btnFromZiYuan_Click(IRibbonControl ctrl)
		{
			try
			{
				fromZiYuanPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_FromZiYuan", "导入资源包", missing);
				fromZiYuanPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
				fromZiYuanPanel.Width = 610;
				fromZiYuanPanel.Height = 510;
				fromZiYuanPanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
				fromZiYuanPanel.Visible = true;
			}
			catch (Exception)
			{
			}
		}

		public void btnBack_Click(IRibbonControl ctrl)
		{
			try
			{
				backPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_Back", "备份资源", missing);
				backPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
				backPanel.Width = 610;
				backPanel.Height = 510;
				backPanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
				backPanel.Visible = true;
			}
			catch (Exception)
			{
			}
		}

		public void btnRollback_Click(IRibbonControl ctrl)
		{
			try
			{
				rollBackPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_Rollback", "恢复备份", missing);
				rollBackPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
				rollBackPanel.Width = 610;
				rollBackPanel.Height = 510;
				rollBackPanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
				rollBackPanel.Visible = true;
			}
			catch (Exception)
			{
			}
		}

		private void BtnXX_1_Click(object sender, EventArgs e)
		{
			OpenNewWordTemplate("\\template\\xx_1.docx");
		}

		public void tbSuCai_Click(IRibbonControl ctrl, bool pressed)
		{
			try
			{
				if (pressed)
				{
					sucaiPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_SuCai", "素材搜索", missing);
					((Uc_SuCai)(dynamic)sucaiPanel.ContentControl).ParentRibbon = this;
					sucaiPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
					sucaiPanel.Width = 800;
					sucaiPanel.Height = 528;
					sucaiPanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
					new ComAwareEventInfo(typeof(_CustomTaskPaneEvents_Event), "VisibleStateChange").AddEventHandler(sucaiPanel, new _CustomTaskPaneEvents_VisibleStateChangeEventHandler(SucaiPanel_VisibleStateChange));
					sucaiPanel.Visible = true;
					return;
				}
				try
				{
					sucaiPanel.Visible = false;
				}
				catch (Exception)
				{
				}
			}
			catch (Exception)
			{
			}
		}

		private void SucaiPanel_VisibleStateChange(CustomTaskPane CustomTaskPaneInst)
		{
			try
			{
				if (!CustomTaskPaneInst.Visible)
				{
					boolTbSucai = false;
					thisAddIn.InvalidateControl("tbSuCai");
				}
			}
			catch (Exception)
			{
			}
		}

		public void openNewSuCaiPanel()
		{
			try
			{
				try
				{
					sucaiPanel.Visible = false;
				}
				catch (Exception)
				{
				}
				boolTbSucai = false;
				thisAddIn.InvalidateControl("tbSuCai");
				newsucaiPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_NewSuCai", "添加素材", missing);
				((Uc_NewSuCai)(dynamic)newsucaiPanel.ContentControl).ParentRibbon = this;
				newsucaiPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
				newsucaiPanel.Width = 800;
				newsucaiPanel.Height = 628;
				newsucaiPanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
				newsucaiPanel.Visible = true;
			}
			catch (Exception)
			{
			}
		}

		public void SucaiPanel_VisibleChanged(object sender, EventArgs e)
		{
		}

		public void testTaskPanel()
		{
		}

		public void tbSucai_TaskPane_VisibleChanged(object sender, EventArgs e)
		{
		}

		public void tbFanWen_Click(IRibbonControl ctrl, bool pressed)
		{
			if (pressed)
			{
				try
				{
					fanwenPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_FanWen", "范文搜索", missing);
					((Uc_FanWen)(dynamic)fanwenPanel.ContentControl).ParentRibbon = this;
					fanwenPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
					fanwenPanel.Width = 1000;
					fanwenPanel.Height = 725;
					fanwenPanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
					new ComAwareEventInfo(typeof(_CustomTaskPaneEvents_Event), "VisibleStateChange").AddEventHandler(fanwenPanel, new _CustomTaskPaneEvents_VisibleStateChangeEventHandler(FanwenPanel_VisibleStateChange));
					fanwenPanel.Visible = true;
					return;
				}
				catch (Exception ex)
				{
					if (CommonConfig.IS_DEBUG)
					{
						MessageBox.Show(ex.ToString());
					}
					return;
				}
			}
			try
			{
				fanwenPanel.Visible = false;
			}
			catch (Exception)
			{
			}
		}

		private void FanwenPanel_VisibleStateChange(CustomTaskPane CustomTaskPaneInst)
		{
			try
			{
				if (!CustomTaskPaneInst.Visible)
				{
					boolTbFanWen = false;
					thisAddIn.InvalidateControl("tbFanWen");
				}
			}
			catch (Exception)
			{
			}
		}

		public void openNewFanWenPanel()
		{
			try
			{
				try
				{
					fanwenPanel.Visible = false;
				}
				catch (Exception)
				{
				}
				boolTbSucai = false;
				thisAddIn.InvalidateControl("tbFanWen");
				newfanwenPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_NewFanWen", "添加范文", missing);
				((Uc_NewFanWen)(dynamic)newfanwenPanel.ContentControl).ParentRibbon = this;
				newfanwenPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
				newfanwenPanel.Width = 800;
				newfanwenPanel.Height = 628;
				newfanwenPanel.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
				newfanwenPanel.Visible = true;
			}
			catch (Exception)
			{
			}
		}

		public void tb_NewSuCai_Click(IRibbonControl ctrl)
		{
		}

		public void tb_NewFanWen_Click(IRibbonControl ctrl)
		{
		}

		public void btnJiaoGao_Click(IRibbonControl ctrl)
		{
			if (!UserUtil.HasLogin())
			{
				MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				return;
			}
			if (!UserUtil.IsVip())
			{
				MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				return;
			}
			int width2 = Screen.PrimaryScreen.Bounds.Width;
			Range rangeSelectContent = null;
			try
			{
				Document activeDocument = wordApp.ActiveDocument;
				Range range = wordApp.Selection.Range;
				object End;
				object Start;
				try
				{
					if (range.Text.Length > 0)
					{
						rangeSelectContent = wordApp.Selection.Range;
					}
				}
				catch (Exception)
				{
					Start = Type.Missing;
					End = Type.Missing;
					rangeSelectContent = activeDocument.Range(ref Start, ref End);
				}
				End = Type.Missing;
				Start = Type.Missing;
				rangeSelectContent = activeDocument.Range(ref End, ref Start);
			}
			catch (Exception)
			{
			}
			boolBtnJiaoGao = false;
			thisAddIn.InvalidateControl("btnJiaoGao");
			try
			{
				try
				{
					correctPanel.Delete();
				}
				catch (Exception)
				{
				}
				correctPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.WordCorrect", "文字校对", missing);
				WordCorrect obj = (WordCorrect)(dynamic)correctPanel.ContentControl;
				correctPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
				int width = 500;
				correctPanel.Width = width;
				tigangPanel.Height = 400;
				correctPanel.Visible = true;
				obj.getCorrectResult(rangeSelectContent, this);
			}
			catch (Exception)
			{
			}
		}

		public void btnInsert_Date_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				string text = DateTime.Now.ToLongDateString();
				text = text.Replace("星期一", "");
				text = text.Replace("星期二", "");
				text = text.Replace("星期三", "");
				text = text.Replace("星期四", "");
				text = text.Replace("星期五", "");
				text = text.Replace("星期六", "");
				text = text.Replace("星期日", "");
				text = text.Replace("星期天", "");
				text = text.Replace(" ", "");
				wordApp.Selection.InsertAfter(text);
			}
			catch (Exception)
			{
			}
		}

		public void btnInsert_One_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				string text = "";
				Document activeDocument = wordApp.ActiveDocument;
				object Start = 0;
				object End = wordApp.Selection.Range.Start;
				Range range = activeDocument.Range(ref Start, ref End);
				int num = 0;
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
				foreach (Paragraph paragraph in range.Paragraphs)
				{
					string text2 = paragraph.Range.Text;
					int num2 = array.Length;
					for (int i = 0; i < num2; i++)
					{
						if (text2.IndexOf(array[i]) == 0)
						{
							if (num <= i)
							{
								num = i + 1;
							}
							break;
						}
					}
				}
				text = array[num];
				int end = wordApp.Selection.End;
				wordApp.Selection.InsertAfter(text);
				int end2 = wordApp.Selection.End;
				Document activeDocument2 = wordApp.ActiveDocument;
				End = end;
				Start = end2;
				Range biaoTi_One_JuZhi = activeDocument2.Range(ref End, ref Start);
				wordApp.Selection.InsertParagraphAfter();
				WpsHelper.setBiaoTi_One_JuZhi(biaoTi_One_JuZhi);
				Document activeDocument3 = wordApp.ActiveDocument;
				Start = end2;
				End = end2;
				activeDocument3.Range(ref Start, ref End).Select();
			}
			catch (Exception)
			{
			}
		}

		public void btnInsert_Two_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				string text = "";
				Document activeDocument = wordApp.ActiveDocument;
				object Start = 0;
				object End = wordApp.Selection.Range.Start;
				Range range = activeDocument.Range(ref Start, ref End);
				int num = 0;
				string[] array = new string[109]
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
				Sentences sentences = range.Sentences;
				int count = sentences.Count;
				count = sentences.Count;
				while (count > 0 && !WpsHelper.Is_BiaoTi_One(sentences[count].Text))
				{
					count--;
				}
				if (count == 1)
				{
					foreach (Paragraph paragraph in range.Paragraphs)
					{
						string text2 = paragraph.Range.Text;
						int num2 = array.Length;
						for (int i = 0; i < num2; i++)
						{
							if (text2.IndexOf(array[i]) == 0)
							{
								if (num <= i)
								{
									num = i + 1;
								}
								break;
							}
						}
					}
				}
				else
				{
					for (; count <= sentences.Count; count++)
					{
						string text3 = sentences[count].Text;
						int num3 = array.Length;
						for (int j = 0; j < num3; j++)
						{
							if (text3.IndexOf(array[j]) == 0)
							{
								if (num <= j)
								{
									num = j + 1;
								}
								break;
							}
						}
					}
				}
				text = array[num];
				int end = wordApp.Selection.End;
				wordApp.Selection.InsertAfter(text);
				int end2 = wordApp.Selection.End;
				Document activeDocument2 = wordApp.ActiveDocument;
				End = end;
				Start = end2;
				Range biaoTi_Two_JuZhi = activeDocument2.Range(ref End, ref Start);
				wordApp.Selection.InsertParagraphAfter();
				WpsHelper.setBiaoTi_Two_JuZhi(biaoTi_Two_JuZhi);
				Document activeDocument3 = wordApp.ActiveDocument;
				Start = end2;
				End = end2;
				activeDocument3.Range(ref Start, ref End).Select();
			}
			catch (Exception)
			{
			}
		}

		public void btnInsert_Three_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				string text = "";
				Document activeDocument = wordApp.ActiveDocument;
				object Start = 0;
				object End = wordApp.Selection.Range.Start;
				Range range = activeDocument.Range(ref Start, ref End);
				int num = 0;
				string[] array = new string[109]
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
				Sentences sentences = range.Sentences;
				int count = sentences.Count;
				count = sentences.Count;
				while (count > 0 && !WpsHelper.Is_BiaoTi_Two(sentences[count].Text))
				{
					count--;
				}
				if (count == 1)
				{
					count = sentences.Count;
					while (count > 0 && !WpsHelper.Is_BiaoTi_Two(sentences[count].Text))
					{
						count--;
					}
					if (count == 1)
					{
						foreach (Paragraph paragraph in range.Paragraphs)
						{
							string text2 = paragraph.Range.Text;
							int num2 = array.Length;
							for (int i = 0; i < num2; i++)
							{
								if (text2.IndexOf(array[i]) == 0)
								{
									if (num <= i)
									{
										num = i + 1;
									}
									break;
								}
							}
						}
					}
					else
					{
						for (; count <= sentences.Count; count++)
						{
							string text3 = sentences[count].Text;
							int num3 = array.Length;
							for (int j = 0; j < num3; j++)
							{
								if (text3.IndexOf(array[j]) == 0)
								{
									if (num <= j)
									{
										num = j + 1;
									}
									break;
								}
							}
						}
					}
				}
				else
				{
					for (; count <= sentences.Count; count++)
					{
						string text4 = sentences[count].Text;
						int num4 = array.Length;
						for (int k = 0; k < num4; k++)
						{
							if (text4.IndexOf(array[k]) == 0)
							{
								if (num <= k)
								{
									num = k + 1;
								}
								break;
							}
						}
					}
				}
				text = array[num];
				int end = wordApp.Selection.End;
				wordApp.Selection.InsertAfter(text);
				int end2 = wordApp.Selection.End;
				Document activeDocument2 = wordApp.ActiveDocument;
				End = end;
				Start = end2;
				Range biaoTi_Three_JuZhi = activeDocument2.Range(ref End, ref Start);
				wordApp.Selection.InsertParagraphAfter();
				WpsHelper.setBiaoTi_Three_JuZhi(biaoTi_Three_JuZhi);
				Document activeDocument3 = wordApp.ActiveDocument;
				Start = end2;
				End = end2;
				activeDocument3.Range(ref Start, ref End).Select();
			}
			catch (Exception)
			{
			}
		}

		public void btnAlignLeft_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					DocWpsUtil.insertPageNumber_Bottom_Left();
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnAlignRight_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					DocWpsUtil.insertPageNumber_Bottom_Right();
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnAlilgnCenter_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					DocWpsUtil.insertPageNumber_Bottom_Center();
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnWps_Send_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					WpsHelper.setBiaoTi_People_JuZhi(wordApp.Selection.Words.Application.Selection.Range);
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnWps_Fujian_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					WpsHelper.setBiaoTi_FuJian_JuZhi(wordApp.Selection.Words.Application.Selection.Range);
				}
			}
			catch (Exception)
			{
			}
		}

		public void btnWps_Date_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				}
				else if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				}
				else
				{
					WpsHelper.setBiaoTi_LastAuthor_JuZhi(wordApp.Selection.Words.Application.Selection.Range);
				}
			}
			catch (Exception)
			{
			}
		}

		public void group4_DialogLauncherClick(IRibbonControl ctrl)
		{
		}

		public void group5_DialogLauncherClick(IRibbonControl ctrl)
		{
		}

		public void groupCorrect_DialogLauncherClick(IRibbonControl ctrl)
		{
		}

		public void btnTiGang_Click(IRibbonControl ctrl)
		{
			if (!UserUtil.HasLogin())
			{
				MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
				return;
			}
			if (!UserUtil.IsVip())
			{
				MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
				return;
			}
			int width = Screen.PrimaryScreen.Bounds.Width;
			try
			{
				boolBtnTigang = false;
				thisAddIn.InvalidateControl("btnTiGang");
			}
			catch (Exception)
			{
			}
			Document activeDocument = wordApp.ActiveDocument;
			object Start = Type.Missing;
			object End = Type.Missing;
			Range rangeSelectContent = activeDocument.Range(ref Start, ref End);
			try
			{
				try
				{
					tigangPanel.Delete();
				}
				catch (Exception)
				{
				}
				tigangPanel = CustomTaskPaneCreater.CreateCTP("Local_Wps_Vsto.Uc_TigangCorrect", "检查提纲", missing);
				Uc_TigangCorrect obj = (Uc_TigangCorrect)(dynamic)tigangPanel.ContentControl;
				tigangPanel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
				int num = 400;
				if (width <= 1024)
				{
					num = 400;
				}
				else if (width <= 1680)
				{
					num = 500;
				}
				else
				{
					num = 600;
				}
				num = 500;
				tigangPanel.Width = num;
				tigangPanel.Height = 400;
				tigangPanel.Visible = true;
				obj.getCorrectResult(rangeSelectContent, this);
			}
			catch (Exception)
			{
			}
		}

		public void btnInsert_HengYe_Click(IRibbonControl ctrl)
		{
			try
			{
				if (!UserUtil.HasLogin())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_LOGIN);
					return;
				}
				if (!UserUtil.IsVip())
				{
					MessageTipsUtil.Show(CommonConfig.STR_TIP_NEED_VIP);
					return;
				}
				DateTime.Now.ToLongDateString();
				Selection selection = wordApp.Selection;
				object Type = WdBreakType.wdSectionBreakNextPage;
				selection.InsertBreak(ref Type);
				Range range = wordApp.Selection.Range;
				wordApp.Selection.Sections[1].PageSetup.Orientation = WdOrientation.wdOrientLandscape;
				Selection selection2 = wordApp.Selection;
				Type = WdBreakType.wdSectionBreakContinuous;
				selection2.InsertBreak(ref Type);
				wordApp.Selection.Sections[1].PageSetup.Orientation = WdOrientation.wdOrientPortrait;
				range.Select();
			}
			catch (Exception)
			{
			}
		}

		public static void OpenNewWordTemplate(string _strFileName)
		{
			if (UserUtil.HasLogin())
			{
				if (UserUtil.IsVip())
				{
					try
					{
						string text = CommonConfig.strBaseFolder + _strFileName;
						Documents documents = wordApp.Documents;
						object Template = text;
						object NewTemplate = Type.Missing;
						object DocumentType = Type.Missing;
						object Visible = Type.Missing;
						documents.Add(ref Template, ref NewTemplate, ref DocumentType, ref Visible);
						try
						{
							geshiPanel.Visible = false;
							geshiPanel.Delete();
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

		public static void OpenNewExcelTemplate(string _strFileName, string _NewFileName)
		{
			if (UserUtil.HasLogin())
			{
				if (UserUtil.IsVip())
				{
					try
					{
						string sourceFileName = CommonConfig.strBaseFolder + _strFileName;
						string text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + _NewFileName + "-" + DateTime.Now.ToLongDateString().ToString();
						text = ((!File.Exists(text + ".xls")) ? (text + ".xls") : (text + "-" + new Random().Next(10, 99) + ".xls"));
						File.Copy(sourceFileName, text, true);
						Process.Start(text);
						try
						{
							geshiPanel.Visible = false;
							geshiPanel.Delete();
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
	}
}
