using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[Guid("00020984-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	[CompilerGenerated]
	public interface HeadersFooters : IEnumerable
	{
		[DispId(0)]
		HeaderFooter this[[In] WdHeaderFooterIndex Index]
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		void _VtblGap1_5();
	}
}
