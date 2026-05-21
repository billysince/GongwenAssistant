using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[Guid("00020952-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	[CompilerGenerated]
	public interface _Font
	{
		[DispId(130)]
		int Bold
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(130)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(130)]
			[param: In]
			set;
		}

		[DispId(131)]
		int Italic
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(131)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(131)]
			[param: In]
			set;
		}

		[DispId(141)]
		float Size
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(141)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(141)]
			[param: In]
			set;
		}

		[DispId(142)]
		string Name
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(142)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(142)]
			[param: In]
			[param: MarshalAs(UnmanagedType.BStr)]
			set;
		}

		[DispId(159)]
		WdColor Color
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(159)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(159)]
			[param: In]
			set;
		}

		void _VtblGap1_4();

		void _VtblGap2_18();

		void _VtblGap3_35();
	}
}
