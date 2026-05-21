using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[TypeIdentifier]
	[Guid("00020970-0000-0000-C000-000000000046")]
	[CompilerGenerated]
	[DefaultMember("Name")]
	public interface _Application
	{
		[DispId(1000)]
		Application Application
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1000)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(0)]
		string Name
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}

		[DispId(6)]
		Documents Documents
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(6)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(3)]
		Document ActiveDocument
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(3)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(4)]
		Window ActiveWindow
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(4)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(5)]
		Selection Selection
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(5)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(23)]
		bool Visible
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(23)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(23)]
			[param: In]
			set;
		}

		void _VtblGap1_2();

		void _VtblGap2_1();

		void _VtblGap3_17();

		void _VtblGap4_85();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(1105)]
		void Quit([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveChanges, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object OriginalFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object RouteDocument);

		void _VtblGap5_41();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(371)]
		float CentimetersToPoints([In] float Centimeters);
	}
}
