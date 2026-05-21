using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Office
{
	[ComImport]
	[DefaultMember("Title")]
	[Guid("000C033B-0000-0000-C000-000000000046")]
	[CompilerGenerated]
	[TypeIdentifier]
	public interface _CustomTaskPane
	{
		[DispId(0)]
		string Title
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}

		[DispId(3)]
		bool Visible
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(3)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(3)]
			[param: In]
			set;
		}

		[DispId(4)]
		object ContentControl
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(4)]
			[return: MarshalAs(UnmanagedType.IDispatch)]
			get;
		}

		[DispId(5)]
		int Height
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(5)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(5)]
			[param: In]
			set;
		}

		[DispId(6)]
		int Width
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(6)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(6)]
			[param: In]
			set;
		}

		[DispId(7)]
		MsoCTPDockPosition DockPosition
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(7)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(7)]
			[param: In]
			set;
		}

		[DispId(8)]
		MsoCTPDockPositionRestrict DockPositionRestrict
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(8)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(8)]
			[param: In]
			set;
		}

		void _VtblGap1_2();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(9)]
		void Delete();
	}
}
