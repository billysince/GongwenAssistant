using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[TypeIdentifier]
	[CompilerGenerated]
	[Guid("0002095A-0000-0000-C000-000000000046")]
	public interface Sections : IEnumerable
	{
		[DispId(0)]
		Section this[[In] int Index]
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		void _VtblGap1_9();
	}
}
