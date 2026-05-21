using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Office
{
	[ComImport]
	[Guid("000C0395-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	[CompilerGenerated]
	public interface IRibbonControl
	{
		[DispId(1)]
		string Id
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}
	}
}
