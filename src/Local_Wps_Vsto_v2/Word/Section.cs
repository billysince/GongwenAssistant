using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[Guid("00020959-0000-0000-C000-000000000046")]
	[DefaultMember("Range")]
	[TypeIdentifier]
	[CompilerGenerated]
	public interface Section
	{
		[DispId(0)]
		Range Range
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(1101)]
		PageSetup PageSetup
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1101)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1101)]
			[param: In]
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(122)]
		HeadersFooters Footers
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(122)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		void _VtblGap1_3();

		void _VtblGap2_1();
	}
}
