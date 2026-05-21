using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word;

namespace Local_Wps_Vsto
{
	internal class RtfUtil
	{
		public static void ConvertKeyToHyperLink(string hilightString, RichTextBox control)
		{
			try
			{
				int selectionStart = control.SelectionStart;
				int selectionLength = control.SelectionLength;
				control.SelectAll();
				int num;
				for (num = 0; num < control.Text.Length; num++)
				{
					num = control.Find(hilightString, num, RichTextBoxFinds.None);
					if (num < 0)
					{
						break;
					}
					control.Select(num, hilightString.Length);
					control.SelectionColor = Color.Red;
				}
				control.Select(selectionStart, selectionLength);
			}
			catch (Exception)
			{
			}
		}

		public static void setGeShi_FromTxt(RichTextBox RtbContent, string strContextTxt)
		{
			RtbContent.Text = strContextTxt;
		}

		public static void LoadWordTo_ClipBoard(string strFileName)
		{
			try
			{
				Word.Application application = (Word.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
				Document document = null;
				object ConfirmConversions = Missing.Value;
				object FileName = strFileName;
				object ReadOnly = false;
				object Visible = false;
				try
				{
					document = application.Documents.Open(ref FileName, ref ConfirmConversions, ref ReadOnly, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref Visible, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
					document.ActiveWindow.Selection.WholeStory();
					document.ActiveWindow.Selection.Copy();
				}
				finally
				{
					if (document != null)
					{
						document.Close(ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
						document = null;
					}
					if (application != null)
					{
						application.Quit(ref ConfirmConversions, ref ConfirmConversions, ref ConfirmConversions);
						application = null;
					}
				}
			}
			catch (Exception)
			{
			}
		}

		public static void SaveWord(string strFileName, RichTextBox RtbTemp)
		{
			object Template = CommonConfig.strBaseFolder + "\\template\\A4_1.docx";
			Word.Application application = (Word.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
			Document document = null;
			object FileFormat = Missing.Value;
			object FileName = strFileName;
			try
			{
				Documents documents = application.Documents;
				object NewTemplate = Type.Missing;
				object DocumentType = Type.Missing;
				object Visible = Type.Missing;
				document = documents.Add(ref Template, ref NewTemplate, ref DocumentType, ref Visible);
				RtbTemp.SelectAll();
				RtbTemp.Copy();
				document.ActiveWindow.Selection.WholeStory();
				document.ActiveWindow.Selection.Paste();
				document.SaveAs(ref FileName, ref FileFormat, ref FileFormat, ref FileFormat, ref FileFormat, ref FileFormat, ref FileFormat, ref FileFormat, ref FileFormat, ref FileFormat, ref FileFormat, ref FileFormat, ref FileFormat, ref FileFormat, ref FileFormat, ref FileFormat);
			}
			finally
			{
				if (document != null)
				{
					document.Close(ref FileFormat, ref FileFormat, ref FileFormat);
					document = null;
				}
				if (application != null)
				{
					application.Quit(ref FileFormat, ref FileFormat, ref FileFormat);
					application = null;
				}
			}
		}
	}
}
