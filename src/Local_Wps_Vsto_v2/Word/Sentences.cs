using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[CompilerGenerated]
	[TypeIdentifier]
	[Guid("0002095B-0000-0000-C000-000000000046")]
	public interface Sentences : IEnumerable
	{
		[DispId(2)]
		int Count
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(2)]
			get;
		}

		[DispId(0)]
		Range this[[In] int Index]
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		void _VtblGap1_1();

		void _VtblGap2_5();
	}
}
