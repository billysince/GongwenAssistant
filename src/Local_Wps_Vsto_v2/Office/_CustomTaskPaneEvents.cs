using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Office
{
	[ComImport]
	[TypeIdentifier]
	[InterfaceType(2)]
	[Guid("000C033C-0000-0000-C000-000000000046")]
	[CompilerGenerated]
	public interface _CustomTaskPaneEvents
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
		[DispId(1)]
		void VisibleStateChange([In][MarshalAs(UnmanagedType.Interface)] CustomTaskPane CustomTaskPaneInst);
	}
}
