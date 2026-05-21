using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[Guid("0002092F-0000-0000-C000-000000000046")]
	[DefaultMember("Code")]
	[TypeIdentifier]
	[CompilerGenerated]
	public interface Field
	{
		[DispId(0)]
		Range Code
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[param: In]
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		void _VtblGap1_3();
	}
}
