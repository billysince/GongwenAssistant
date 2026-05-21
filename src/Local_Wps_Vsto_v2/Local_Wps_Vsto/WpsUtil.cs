namespace Local_Wps_Vsto
{
	internal class WpsUtil
	{
		private string _strXmlContent = "";

		public void StartDocument()
		{
			_strXmlContent += "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><w:document xmlns:wpc='http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' xmlns:v='urn:schemas-microsoft-com:vml' xmlns:wp14='http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing' xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml' xmlns:w10='urn:schemas-microsoft-com:office:word' xmlns:w15='http://schemas.microsoft.com/office/word/2012/wordml' xmlns:wpg='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup' xmlns:wpi='http://schemas.microsoft.com/office/word/2010/wordprocessingInk' xmlns:wne='http://schemas.microsoft.com/office/word/2006/wordml' xmlns:wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape' xmlns:wpsCustomData='http://www.wps.cn/officeDocument/2013/wpsCustomData' mc:Ignorable='w14 w15 wp14'><w:body>' ";
		}

		public void CloseDocument()
		{
			_strXmlContent += "<w:sectPr><w:footerReference r:id='rId3' w:type='default'/><w:pgSz w:w='11906' w:h='16838'/><w:pgMar w:top='2098' w:right='1474' w:bottom='1984' w:left='1587' w:header='851' w:footer='1134' w:gutter='0'/><w:cols w:space='425' w:num='1'/><w:docGrid w:type='lines' w:linePitch='312' w:charSpace='0'/></w:sectPr></w:body></w:document>";
		}

		public void Add_Top(string strTxt)
		{
			_strXmlContent += "<w:p><w:pPr><w:ind w:firstLine='0' w:firstLineChars='0'/><w:jc w:val='center'/><w:rPr><w:rFonts w:ascii='方正小标宋简体' w:hAnsi='方正小标宋简体' w:eastAsia='方正小标宋简体' w:cs='方正小标宋简体'/><w:sz w:val='44'/><w:szCs w:val='36'/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint='eastAsia' w:ascii='方正小标宋简体' w:hAnsi='方正小标宋简体' w:eastAsia='方正小标宋简体' w:cs='方正小标宋简体'/><w:sz w:val='44'/><w:szCs w:val='36'/></w:rPr><w:t>";
			_strXmlContent += strTxt;
			_strXmlContent += "</w:t></w:r></w:p>";
		}

		public void Add_Time(string strTxt)
		{
			_strXmlContent += "<w:p><w:pPr><w:ind w:firstLine='0' w:firstLineChars='0'/><w:jc w:val='center'/><w:rPr><w:rFonts w:hint='eastAsia' w:ascii='楷体_GB2312' w:eastAsia='楷体_GB2312'/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint='eastAsia' w:ascii='楷体_GB2312' w:eastAsia='楷体_GB2312'/> </w:rPr><w:t>";
			_strXmlContent += strTxt;
			_strXmlContent += "</w:t></w:r></w:p>";
		}

		public void Add_People(string strTxt)
		{
			_strXmlContent += "<w:p><w:pPr><w:ind w:left='0' w:leftChars='0' w:firstLine='0' w:firstLineChars='0' /><w:rPr><w:rFonts w:hint='eastAsia' w:eastAsia='仿宋_GB2312' /><w:lang w:eastAsia='zh-CN' /></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint='eastAsia' /><w:lang w:eastAsia='zh-CN' /></w:rPr><w:t> ";
			_strXmlContent += strTxt;
			_strXmlContent += "</w:t></w:r><w:bookmarkStart w:id='0' w:name='_GoBack' /><w:bookmarkEnd w:id='0' /></w:p> ";
		}

		public void Add_One(string strTxt)
		{
			_strXmlContent += "<w:p><w:pPr><w:ind w:firstLine='640'/><w:rPr><w:rFonts w:ascii='黑体' w:hAnsi='黑体' w:eastAsia='黑体' w:cs='黑体'/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint='eastAsia' w:ascii='黑体' w:hAnsi='黑体' w:eastAsia='黑体' w:cs='黑体'/></w:rPr><w:t>";
			_strXmlContent += strTxt;
			_strXmlContent += "</w:t></w:r></w:p>";
		}

		public void Add_Two(string strTxt)
		{
			_strXmlContent += "<w:p><w:pPr><w:ind w:firstLine='643'/><w:rPr><w:rFonts w:hint='eastAsia' w:ascii='楷体_GB2312' w:hAnsi='楷体_GB2312' w:eastAsia='楷体_GB2312' w:cs='楷体_GB2312'/><w:b/><w:bCs/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint='eastAsia' w:ascii='楷体_GB2312' w:hAnsi='楷体_GB2312' w:eastAsia='楷体_GB2312' w:cs='楷体_GB2312'/><w:b/><w:bCs/></w:rPr><w:t>";
			_strXmlContent += strTxt;
			_strXmlContent += "</w:t></w:r></w:p>";
		}

		public void Add_Three(string strTxt)
		{
			_strXmlContent += "<w:p><w:pPr><w:ind w:firstLine='643'/><w:rPr><w:rFonts w:hint='eastAsia'/><w:b/><w:bCs/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint='eastAsia'/><w:b/><w:bCs/></w:rPr><w:t>";
			_strXmlContent += strTxt;
			_strXmlContent += "</w:t></w:r></w:p>";
		}

		public void Start_P()
		{
			_strXmlContent += "<w:p><w:pPr><w:ind w:firstLine='643'/></w:pPr>";
		}

		public void Close_P()
		{
			_strXmlContent += "</w:p>";
		}

		public void Add_One_DuanNei(string strTxt)
		{
			_strXmlContent += "<w:pPr><w:ind w:firstLine='640'/></w:pPr><w:r><w:rPr><w:rFonts w:hint='eastAsia' w:ascii='黑体' w:hAnsi='黑体' w:eastAsia='黑体' w:cs='黑体'/><w:lang w:eastAsia='zh-CN'/></w:rPr><w:t>";
			_strXmlContent += strTxt;
			_strXmlContent += "</w:t></w:r>";
		}

		public void Add_Two_DuanNei(string strTxt)
		{
			_strXmlContent += "<w:r><w:rPr><w:rFonts w:hint='eastAsia' w:ascii='楷体_GB2312' w:hAnsi='楷体_GB2312' w:eastAsia='楷体_GB2312' w:cs='楷体_GB2312'/><w:b/><w:bCs/></w:rPr><w:t>";
			_strXmlContent += strTxt;
			_strXmlContent += "</w:t></w:r>";
		}

		public void Add_Three_DuanNei(string strTxt)
		{
			_strXmlContent += "<w:r><w:rPr><w:rFonts w:hint='eastAsia'/><w:b/><w:bCs/></w:rPr><w:t>";
			_strXmlContent += strTxt;
			_strXmlContent += "</w:t></w:r>";
		}

		public void Add_NeiRong_DuanNei(string strTxt)
		{
			_strXmlContent += "<w:r><w:rPr><w:rFonts w:hint='eastAsia'/></w:rPr><w:t>";
			_strXmlContent += strTxt;
			_strXmlContent += "</w:t></w:r>";
		}

		public void Add_Last_Time(string strTxt)
		{
			_strXmlContent += "<w:p><w:pPr><w:wordWrap w:val='0' /><w:ind w:left='0' w:leftChars='0' w:firstLine='0' w:firstLineChars='0' /><w:jc w:val='right' /><w:rPr><w:rFonts w:hint='default' w:eastAsia='仿宋_GB2312' /><w:lang w:val='en-US' w:eastAsia='zh-CN' /> </w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint='eastAsia' /><w:lang w:val='en-US' w:eastAsia='zh-CN' />   </w:rPr><w:t xml:space='preserve'> ";
			_strXmlContent += strTxt;
			_strXmlContent += "    </w:t></w:r></w:p>";
		}

		public string getContent()
		{
			return _strXmlContent;
		}
	}
}
