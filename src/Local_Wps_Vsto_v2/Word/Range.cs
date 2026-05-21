using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[Guid("0002095E-0000-0000-C000-000000000046")]
	[DefaultMember("Text")]
	[CompilerGenerated]
	[TypeIdentifier]
	public interface Range
	{
		[DispId(0)]
		string Text
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[param: In]
			[param: MarshalAs(UnmanagedType.BStr)]
			set;
		}

		[DispId(3)]
		int Start
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(3)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(3)]
			[param: In]
			set;
		}

		[DispId(4)]
		int End
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(4)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(4)]
			[param: In]
			set;
		}

		[DispId(5)]
		Font Font
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(5)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(5)]
			[param: In]
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(51)]
		Words Words
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(51)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(52)]
		Sentences Sentences
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(52)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(53)]
		Characters Characters
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(53)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(59)]
		Paragraphs Paragraphs
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(59)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(1102)]
		ParagraphFormat ParagraphFormat
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1102)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1102)]
			[param: In]
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(68)]
		ListFormat ListFormat
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(68)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(130)]
		int Bold
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(130)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(130)]
			[param: In]
			set;
		}

		[DispId(139)]
		WdUnderline Underline
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(139)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(139)]
			[param: In]
			set;
		}

		[DispId(301)]
		WdColorIndex HighlightColorIndex
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(301)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(301)]
			[param: In]
			set;
		}

		[DispId(262)]
		Find Find
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(262)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		void _VtblGap1_2();

		void _VtblGap2_3();

		void _VtblGap3_5();

		void _VtblGap4_8();

		void _VtblGap5_4();

		void _VtblGap6_2();

		void _VtblGap7_18();

		void _VtblGap8_7();

		void _VtblGap9_17();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(65535)]
		void Select();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(100)]
		void SetRange([In] int Start, [In] int End);

		void _VtblGap10_23();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(127)]
		int Delete([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Unit, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Count);
	}
}
