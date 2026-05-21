using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[Guid("0002096B-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	[DefaultMember("Name")]
	[CompilerGenerated]
	public interface _Document
	{
		[DispId(0)]
		string Name
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}

		[DispId(3)]
		string Path
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(3)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}

		[DispId(15)]
		Sections Sections
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(15)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(16)]
		Paragraphs Paragraphs
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(16)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(1101)]
		PageSetup PageSetup
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1101)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1101)]
			[param: In]
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(42)]
		Window ActiveWindow
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(42)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		void _VtblGap1_5();

		void _VtblGap2_14();

		void _VtblGap3_15();

		void _VtblGap4_10();

		void _VtblGap5_109();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(1105)]
		void Close([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveChanges, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object OriginalFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object RouteDocument);

		void _VtblGap6_7();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(108)]
		void Save();

		void _VtblGap7_2();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(2000)]
		[return: MarshalAs(UnmanagedType.Interface)]
		Range Range([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Start, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object End);

		void _VtblGap8_118();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(376)]
		void SaveAs([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FileName, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FileFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object LockComments, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Password, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AddToRecentFiles, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object WritePassword, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object ReadOnlyRecommended, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object EmbedTrueTypeFonts, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveNativePictureFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveFormsData, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveAsAOCELetter, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Encoding, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object InsertLineBreaks, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AllowSubstitutions, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object LineEnding, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AddBiDiMarks);

		void _VtblGap9_119();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(552)]
		void ExportAsFixedFormat([In][MarshalAs(UnmanagedType.BStr)] string OutputFileName, [In] WdExportFormat ExportFormat, [In] bool OpenAfterExport, [In] WdExportOptimizeFor OptimizeFor, [In] WdExportRange Range, [In] int From, [In] int To, [In] WdExportItem Item, [In] bool IncludeDocProps, [In] bool KeepIRM, [In] WdExportCreateBookmarks CreateBookmarks, [In] bool DocStructureTags, [In] bool BitmapMissingFonts, [In] bool UseISO19005_1, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FixedFormatExtClassPtr);
	}
}
