using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Office
{
	[ComImport]
	[CompilerGenerated]
	[Guid("000C033D-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	public interface ICTPFactory
	{
		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(1)]
		[return: MarshalAs(UnmanagedType.Interface)]
		CustomTaskPane CreateCTP([In][MarshalAs(UnmanagedType.BStr)] string CTPAxID, [In][MarshalAs(UnmanagedType.BStr)] string CTPTitle, [Optional][In][MarshalAs(UnmanagedType.Struct)] object CTPParentWindow);
	}
}
