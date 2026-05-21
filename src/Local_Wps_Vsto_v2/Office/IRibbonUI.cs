using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Office
{
	[ComImport]
	[TypeIdentifier]
	[Guid("000C03A7-0000-0000-C000-000000000046")]
	[CompilerGenerated]
	public interface IRibbonUI
	{
		void _VtblGap1_1();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(2)]
		void InvalidateControl([In][MarshalAs(UnmanagedType.BStr)] string ControlID);
	}
}
