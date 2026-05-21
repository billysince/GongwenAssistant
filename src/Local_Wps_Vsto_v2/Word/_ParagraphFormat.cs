using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[TypeIdentifier]
	[Guid("00020953-0000-0000-C000-000000000046")]
	[CompilerGenerated]
	public interface _ParagraphFormat
	{
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

		[DispId(126)]
		float CharacterUnitRightIndent
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(126)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(126)]
			[param: In]
			set;
		}

		[DispId(127)]
		float CharacterUnitLeftIndent
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(127)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(127)]
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

		void _VtblGap1_6();

		void _VtblGap2_8();

		void _VtblGap3_6();

		void _VtblGap4_20();

		void _VtblGap5_14();
	}
}
