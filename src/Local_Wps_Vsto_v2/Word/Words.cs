using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[TypeIdentifier]
	[CompilerGenerated]
	[DefaultMember("Item")]
	[Guid("0002095C-0000-0000-C000-000000000046")]
	public interface Words : IEnumerable
	{
		[DispId(2)]
		int Count
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(2)]
			get;
		}

		[DispId(1000)]
		Application Application
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1000)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		void _VtblGap1_1();

		void _VtblGap2_2();
	}
}
