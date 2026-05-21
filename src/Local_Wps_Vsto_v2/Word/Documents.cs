using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[CompilerGenerated]
	[Guid("0002096C-0000-0000-C000-000000000046")]
	[DefaultMember("Item")]
	[TypeIdentifier]
	public interface Documents : IEnumerable
	{
		void _VtblGap1_10();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(14)]
		[return: MarshalAs(UnmanagedType.Interface)]
		Document Add([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Template, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object NewTemplate, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object DocumentType, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Visible);

		void _VtblGap2_4();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(19)]
		[return: MarshalAs(UnmanagedType.Interface)]
		Document Open([In][MarshalAs(UnmanagedType.Struct)] ref object FileName, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object ConfirmConversions, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object ReadOnly, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AddToRecentFiles, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object PasswordDocument, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object PasswordTemplate, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Revert, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object WritePasswordDocument, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object WritePasswordTemplate, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Format, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Encoding, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Visible, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object OpenAndRepair, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object DocumentDirection, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object NoEncodingDialog, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object XMLTransform);
	}
}
