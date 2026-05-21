using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[TypeIdentifier]
	[CompilerGenerated]
	[Guid("00020971-0000-0000-C000-000000000046")]
	public interface PageSetup
	{
		[DispId(100)]
		float TopMargin
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(100)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(100)]
			[param: In]
			set;
		}

		[DispId(101)]
		float BottomMargin
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(101)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(101)]
			[param: In]
			set;
		}

		[DispId(102)]
		float LeftMargin
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(102)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(102)]
			[param: In]
			set;
		}

		[DispId(103)]
		float RightMargin
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(103)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(103)]
			[param: In]
			set;
		}

		[DispId(104)]
		float Gutter
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(104)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(104)]
			[param: In]
			set;
		}

		[DispId(105)]
		float PageWidth
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(105)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(105)]
			[param: In]
			set;
		}

		[DispId(106)]
		float PageHeight
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(106)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(106)]
			[param: In]
			set;
		}

		[DispId(107)]
		WdOrientation Orientation
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(107)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(107)]
			[param: In]
			set;
		}

		[DispId(108)]
		WdPaperTray FirstPageTray
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(108)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(108)]
			[param: In]
			set;
		}

		[DispId(109)]
		WdPaperTray OtherPagesTray
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(109)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(109)]
			[param: In]
			set;
		}

		[DispId(110)]
		WdVerticalAlignment VerticalAlignment
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(110)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(110)]
			[param: In]
			set;
		}

		[DispId(111)]
		int MirrorMargins
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(111)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(111)]
			[param: In]
			set;
		}

		[DispId(112)]
		float HeaderDistance
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(112)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(112)]
			[param: In]
			set;
		}

		[DispId(113)]
		float FooterDistance
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(113)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(113)]
			[param: In]
			set;
		}

		[DispId(114)]
		WdSectionStart SectionStart
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(114)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(114)]
			[param: In]
			set;
		}

		[DispId(115)]
		int OddAndEvenPagesHeaderFooter
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(115)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(115)]
			[param: In]
			set;
		}

		[DispId(116)]
		int DifferentFirstPageHeaderFooter
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(116)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(116)]
			[param: In]
			set;
		}

		[DispId(117)]
		int SuppressEndnotes
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(117)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(117)]
			[param: In]
			set;
		}

		[DispId(118)]
		LineNumbering LineNumbering
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(118)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(118)]
			[param: In]
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(120)]
		WdPaperSize PaperSize
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(120)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(120)]
			[param: In]
			set;
		}

		[DispId(121)]
		bool TwoPagesOnOne
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(121)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(121)]
			[param: In]
			set;
		}

		[DispId(124)]
		float LinesPage
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(124)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(124)]
			[param: In]
			set;
		}

		[DispId(131)]
		WdLayoutMode LayoutMode
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(131)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(131)]
			[param: In]
			set;
		}

		[DispId(1222)]
		WdGutterStyle GutterPos
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1222)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1222)]
			[param: In]
			set;
		}

		[DispId(1223)]
		bool BookFoldPrinting
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1223)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1223)]
			[param: In]
			set;
		}

		[DispId(1224)]
		bool BookFoldRevPrinting
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1224)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1224)]
			[param: In]
			set;
		}

		[DispId(1225)]
		int BookFoldPrintingSheets
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1225)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1225)]
			[param: In]
			set;
		}

		void _VtblGap1_3();

		void _VtblGap2_2();

		void _VtblGap3_4();

		void _VtblGap4_8();
	}
}
