using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[TypeIdentifier]
	[Guid("00020960-0000-0000-C000-000000000046")]
	[CompilerGenerated]
	public interface Pane
	{
		[DispId(3)]
		Selection Selection
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(3)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(10)]
		View View
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(10)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		void _VtblGap1_4();

		void _VtblGap2_6();
	}
}
