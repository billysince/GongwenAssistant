using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[DefaultMember("Range")]
	[TypeIdentifier]
	[CompilerGenerated]
	[Guid("00020985-0000-0000-C000-000000000046")]
	public interface HeaderFooter
	{
		[DispId(0)]
		Range Range
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(5)]
		PageNumbers PageNumbers
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(5)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(6)]
		bool LinkToPrevious
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(6)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(6)]
			[param: In]
			set;
		}

		void _VtblGap1_3();

		void _VtblGap2_4();
	}
}
