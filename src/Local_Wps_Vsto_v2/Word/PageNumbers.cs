using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[TypeIdentifier]
	[CompilerGenerated]
	[Guid("00020986-0000-0000-C000-000000000046")]
	[DefaultMember("Item")]
	public interface PageNumbers : IEnumerable
	{
		[DispId(2)]
		WdPageNumberStyle NumberStyle
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(2)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(2)]
			[param: In]
			set;
		}

		void _VtblGap1_5();

		void _VtblGap2_13();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(101)]
		[return: MarshalAs(UnmanagedType.Interface)]
		PageNumber Add([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object PageNumberAlignment, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FirstPage);
	}
}
