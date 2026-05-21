using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[CompilerGenerated]
	[DefaultMember("Caption")]
	[Guid("00020962-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	public interface Window
	{
		[DispId(1)]
		Pane ActivePane
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(1)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(4)]
		Selection Selection
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(4)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(0)]
		string Caption
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[param: In]
			[param: MarshalAs(UnmanagedType.BStr)]
			set;
		}

		[DispId(14)]
		View View
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(14)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(36)]
		bool Visible
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(36)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(36)]
			[param: In]
			set;
		}

		void _VtblGap1_3();

		void _VtblGap2_2();

		void _VtblGap3_12();

		void _VtblGap4_6();

		void _VtblGap5_34();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(112)]
		void GetPoint(out int ScreenPixelsLeft, out int ScreenPixelsTop, out int ScreenPixelsWidth, out int ScreenPixelsHeight, [In][MarshalAs(UnmanagedType.IDispatch)] object obj);

		void _VtblGap6_9();
	}
}
