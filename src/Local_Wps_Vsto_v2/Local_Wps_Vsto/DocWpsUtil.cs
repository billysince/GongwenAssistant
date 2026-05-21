using System;
using System.Reflection;
using System.Windows.Forms;
using Word;

namespace Local_Wps_Vsto
{
	internal class DocWpsUtil
	{
		public static Word.Application wordApp = MyAddin.wordApp;

		public static void AutoPaiBan()
		{
			try
			{
				Document activeDocument = wordApp.ActiveDocument;
				Range range = null;
				int start = wordApp.Selection.Start;
				int end = wordApp.Selection.End;
				Range range2 = wordApp.Selection.Range;
				try
				{
					if (range2.Text.Length > 0)
					{
						range = wordApp.Selection.Range;
					}
				}
				catch (Exception)
				{
					object Start = Type.Missing;
					object End = Type.Missing;
					range = activeDocument.Range(ref Start, ref End);
				}
				range.Font.Size = 16f;
				range.Font.Name = "仿宋_GB2312";
				range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
				range.ParagraphFormat.LineSpacing = 28.95f;
				range.ParagraphFormat.WidowControl = 0;
				int num = 0;
				Range range3 = null;
				for (int num2 = range.Paragraphs.Count; num2 > 0; num2--)
				{
					if (range.Paragraphs[num2].Range.ListFormat.ListValue > 0)
					{
						ListFormat listFormat = range.Paragraphs[num2].Range.ListFormat;
						object End = Type.Missing;
						listFormat.ConvertNumbersToText(ref End);
					}
				}
				Range selectionContent = range;
				FindAndReplace(selectionContent, " ", "");
				FindAndReplace(selectionContent, "\u2003", "");
				FindAndReplace(selectionContent, "\t", "");
				FindAndReplace(selectionContent, "\u3000\u3000", "");
				FindAndReplace(selectionContent, ",", "，");
				FindAndReplace(selectionContent, "(", "（");
				FindAndReplace(selectionContent, ")", "）");
				FindAndReplace(selectionContent, "!", "！");
				FindAndReplace(selectionContent, ":", "：");
				FindAndReplace(selectionContent, ";", "；");
				FindAndReplace(selectionContent, "【", "〔");
				FindAndReplace(selectionContent, "】", "〕");
				FindAndReplace(selectionContent, "[", "〔");
				FindAndReplace(selectionContent, "]", "〕");
				FindAndReplace(selectionContent, "<<", "《");
				FindAndReplace(selectionContent, ">>", "》");
				foreach (Paragraph paragraph in range.Paragraphs)
				{
					string text = paragraph.Range.Text;
					if (num == 0)
					{
						range3 = paragraph.Range;
					}
					paragraph.Range.Select();
					if (text.IndexOf("。") > 0)
					{
						WpsHelper.setBiaoTi_Content(paragraph);
						int count = paragraph.Range.Sentences.Count;
						try
						{
							Sentences sentences = paragraph.Range.Sentences;
							for (int i = 1; i <= count; i++)
							{
								string strTitle = sentences[i].Text.ToString();
								if (WpsHelper.Is_BiaoTi_One(strTitle))
								{
									WpsHelper.setBiaoTi_One_JuZhi(sentences[i]);
								}
								else if (WpsHelper.Is_BiaoTi_Two(strTitle))
								{
									WpsHelper.setBiaoTi_Two_JuZhi(sentences[i]);
								}
								else if (WpsHelper.Is_BiaoTi_Three(strTitle))
								{
									WpsHelper.setBiaoTi_Three_JuZhi(sentences[i]);
								}
							}
						}
						catch (Exception)
						{
						}
					}
					else if (text.IndexOf("年") > 0 && text.IndexOf("月") > 0)
					{
						if (num <= 5)
						{
							WpsHelper.setBiaoTi_TopAuthor(paragraph);
						}
						else
						{
							WpsHelper.setBiaoTi_LastAuthor(paragraph);
						}
					}
					else if (text.IndexOf("：") > 0 || text.IndexOf(":") > 0)
					{
						if (WpsHelper.Is_BiaoTi_People(text, num))
						{
							WpsHelper.setBiaoTi_People(paragraph);
						}
						else
						{
							WpsHelper.setBiaoTi_Content(paragraph);
						}
					}
					else if (WpsHelper.Is_BiaoTi_TopTitle(text, num))
					{
						WpsHelper.setBiaoTi_TopTitle(paragraph);
					}
					else if (WpsHelper.Is_BiaoTi_Time(text, num))
					{
						WpsHelper.setBiaoTi_TopAuthor(paragraph);
					}
					else if (WpsHelper.Is_BiaoTi_One(text))
					{
						WpsHelper.setBiaoTi_One_Paragraph(paragraph);
					}
					else if (WpsHelper.Is_BiaoTi_Two(text))
					{
						WpsHelper.setBiaoTi_Two_Paragraph(paragraph);
					}
					else if (WpsHelper.Is_BiaoTi_Three(text))
					{
						WpsHelper.setBiaoTi_Three_Paragraph(paragraph);
					}
					else
					{
						WpsHelper.setBiaoTi_Content(paragraph);
					}
					int count2 = paragraph.Range.Words.Count;
					num++;
				}
				range3.Select();
				range.Select();
			}
			catch (Exception)
			{
			}
		}

		public static void setPage_A4_Single()
		{
			try
			{
				wordApp.ActiveDocument.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
				wordApp.ActiveDocument.PageSetup.TopMargin = wordApp.CentimetersToPoints(float.Parse("3.7"));
				wordApp.ActiveDocument.PageSetup.BottomMargin = wordApp.CentimetersToPoints(float.Parse("3.5"));
				wordApp.ActiveDocument.PageSetup.LeftMargin = wordApp.CentimetersToPoints(float.Parse("2.8"));
				wordApp.ActiveDocument.PageSetup.RightMargin = wordApp.CentimetersToPoints(float.Parse("2.6"));
				wordApp.ActiveDocument.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
				insertPageNumber_Bottom_Right();
				wordApp.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
				wordApp.ActiveWindow.ActivePane.Selection.Borders[WdBorderType.wdBorderBottom].Visible = false;
				wordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument;
			}
			catch (Exception)
			{
			}
		}

		public static void insertPageNumber_Bottom_Right()
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
				application.Application.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
				application.ActiveWindow.ActivePane.Selection.Font.Size = 14f;
				application.ActiveWindow.ActivePane.Selection.Font.Name = "宋体";
				application.ActiveWindow.ActivePane.Selection.ParagraphFormat.CharacterUnitRightIndent = 1f;
				object Text = Missing.Value;
				application.Application.Selection.TypeText("—");
				object Type = WdFieldType.wdFieldPage;
				application.Application.Selection.Fields.Add(application.Application.Selection.Range, ref Type, ref Text, ref Text);
				application.Application.Selection.TypeText("—");
				application.ActiveDocument.PageSetup.HeaderDistance = application.CentimetersToPoints(float.Parse("1.5"));
				application.ActiveDocument.PageSetup.FooterDistance = application.CentimetersToPoints(float.Parse("2.5"));
				application.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
				wordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
				application.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
				application.ActiveWindow.ActivePane.Selection.Borders[WdBorderType.wdBorderBottom].Visible = false;
				application.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
			}
			catch (Exception ex)
			{
				if (CommonConfig.IS_DEBUG)
				{
					MessageBox.Show(ex.ToString());
				}
			}
		}

		public static void insertPageNumber_Bottom_Left()
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
				application.Application.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				application.ActiveWindow.ActivePane.Selection.Font.Size = 14f;
				application.ActiveWindow.ActivePane.Selection.Font.Name = "宋体";
				application.ActiveWindow.ActivePane.Selection.ParagraphFormat.CharacterUnitLeftIndent = 1f;
				object Text = Missing.Value;
				application.Application.Selection.TypeText("—");
				object Type = WdFieldType.wdFieldPage;
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

		public static void insertPageNumber_Bottom_Center()
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
				application.Application.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				application.ActiveWindow.ActivePane.Selection.Font.Size = 14f;
				application.ActiveWindow.ActivePane.Selection.Font.Name = "宋体";
				object Text = Missing.Value;
				application.Application.Selection.TypeText("—");
				object Type = WdFieldType.wdFieldPage;
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

		public static void DeleteBlankLine(Range selectionContent, Word.Application wordApp)
		{
			try
			{
				int num = 1;
				foreach (Paragraph paragraph in selectionContent.Paragraphs)
				{
					try
					{
						if (paragraph.Range.Text.Length <= 1)
						{
							paragraph.Range.Select();
							Selection selection = wordApp.Selection;
							object Unit = Type.Missing;
							object Count = Type.Missing;
							selection.Delete(ref Unit, ref Count);
						}
					}
					catch (Exception)
					{
					}
					num++;
				}
			}
			catch (Exception)
			{
			}
		}

		public static void FindAndReplace(Selection selectionContent, string strOld, string strNew)
		{
			try
			{
				object FindText = Type.Missing;
				selectionContent.Find.ClearFormatting();
				selectionContent.Find.Replacement.ClearFormatting();
				selectionContent.Find.Text = strOld;
				selectionContent.Find.Replacement.Text = strNew;
				object Replace = WdReplace.wdReplaceAll;
				selectionContent.Find.Execute(ref FindText, ref FindText, ref FindText, ref FindText, ref FindText, ref FindText, ref FindText, ref FindText, ref FindText, ref FindText, ref Replace, ref FindText, ref FindText, ref FindText, ref FindText);
			}
			catch (Exception)
			{
			}
		}

		public static void FindAndReplace(Range selectionContent, string strOld, string strNew)
		{
			try
			{
				object FindText = Type.Missing;
				selectionContent.Find.ClearFormatting();
				selectionContent.Find.Replacement.ClearFormatting();
				selectionContent.Find.Text = strOld;
				selectionContent.Find.Replacement.Text = strNew;
				object Replace = WdReplace.wdReplaceAll;
				selectionContent.Find.Execute(ref FindText, ref FindText, ref FindText, ref FindText, ref FindText, ref FindText, ref FindText, ref FindText, ref FindText, ref FindText, ref Replace, ref FindText, ref FindText, ref FindText, ref FindText);
			}
			catch (Exception)
			{
			}
		}
	}
}
