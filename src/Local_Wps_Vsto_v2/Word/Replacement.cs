using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[CompilerGenerated]
	[Guid("000209B1-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	public interface Replacement
	{
		[DispId(15)]
		string Text
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(15)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(15)]
			[param: In]
			[param: MarshalAs(UnmanagedType.BStr)]
			set;
		}

		void _VtblGap1_9();

		void _VtblGap2_7();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(20)]
		void ClearFormatting();
	}
}
