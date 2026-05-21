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
	[Guid("00020930-0000-0000-C000-000000000046")]
	public interface Fields : IEnumerable
	{
		void _VtblGap1_12();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(105)]
		[return: MarshalAs(UnmanagedType.Interface)]
		Field Add([In][MarshalAs(UnmanagedType.Interface)] Range Range, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Type, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Text, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object PreserveFormatting);
	}
}
