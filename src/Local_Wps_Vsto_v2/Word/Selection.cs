using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[Guid("00020975-0000-0000-C000-000000000046")]
	[DefaultMember("Text")]
	[TypeIdentifier]
	[CompilerGenerated]
	public interface Selection
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

		[DispId(58)]
		Sections Sections
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(58)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(1100)]
		Borders Borders
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1100)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1100)]
			[param: In]
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(64)]
		Fields Fields
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(64)]
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

		[DispId(306)]
		HeaderFooter HeaderFooter
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(306)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(262)]
		Find Find
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(262)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(400)]
		Range Range
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(400)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		void _VtblGap1_2();

		void _VtblGap2_5();

		void _VtblGap3_6();

		void _VtblGap4_1();

		void _VtblGap5_1();

		void _VtblGap6_2();

		void _VtblGap7_13();

		void _VtblGap8_3();

		void _VtblGap9_23();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(104)]
		void InsertAfter([In][MarshalAs(UnmanagedType.BStr)] string Text);

		void _VtblGap10_14();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(120)]
		void Copy();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(121)]
		void Paste();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(122)]
		void InsertBreak([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Type);

		void _VtblGap11_3();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(127)]
		int Delete([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Unit, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Count);

		void _VtblGap12_2();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(161)]
		void InsertParagraphAfter();

		void _VtblGap13_28();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(507)]
		void TypeText([In][MarshalAs(UnmanagedType.BStr)] string Text);

		void _VtblGap14_14();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(524)]
		void WholeStory();

		void _VtblGap15_39();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(1013)]
		void PasteAndFormat([In] WdRecoveryType Type);
	}
}
