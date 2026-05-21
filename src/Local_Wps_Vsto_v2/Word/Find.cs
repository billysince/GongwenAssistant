using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[CompilerGenerated]
	[Guid("000209B0-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	public interface Find
	{
		[DispId(22)]
		string Text
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(22)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(22)]
			[param: In]
			[param: MarshalAs(UnmanagedType.BStr)]
			set;
		}

		[DispId(25)]
		Replacement Replacement
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(25)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		void _VtblGap1_26();

		void _VtblGap2_4();

		void _VtblGap3_12();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(31)]
		void ClearFormatting();

		void _VtblGap4_2();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(444)]
		bool Execute([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FindText, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchCase, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchWholeWord, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchWildcards, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchSoundsLike, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchAllWordForms, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Forward, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Wrap, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Format, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object ReplaceWith, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Replace, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchKashida, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchDiacritics, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchAlefHamza, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchControl);
	}
}
