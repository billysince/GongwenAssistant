using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[Guid("00020957-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	[DefaultMember("Range")]
	[CompilerGenerated]
	public interface Paragraph
	{
		[DispId(0)]
		Range Range
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(1102)]
		ParagraphFormat Format
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

		[DispId(101)]
		WdParagraphAlignment Alignment
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(101)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(101)]
			[param: In]
			set;
		}

		[DispId(106)]
		float RightIndent
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
		float LeftIndent
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
		float FirstLineIndent
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
		float LineSpacing
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
		WdLineSpacing LineSpacingRule
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(110)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(110)]
			[param: In]
			set;
		}

		[DispId(114)]
		int WidowControl
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(114)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(114)]
			[param: In]
			set;
		}

		[DispId(125)]
		int DisableLineHeightGrid
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(125)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(125)]
			[param: In]
			set;
		}

		[DispId(128)]
		float CharacterUnitFirstLineIndent
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(128)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(128)]
			[param: In]
			set;
		}

		[DispId(129)]
		float LineUnitBefore
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(129)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(129)]
			[param: In]
			set;
		}

		[DispId(130)]
		float LineUnitAfter
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(130)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(130)]
			[param: In]
			set;
		}

		void _VtblGap1_3();

		void _VtblGap2_7();

		void _VtblGap3_8();

		void _VtblGap4_6();

		void _VtblGap5_17();

		void _VtblGap6_24();
	}
}
