using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace AddInDesignerObjects
{
	[ComImport]
	[TypeIdentifier]
	[CompilerGenerated]
	[Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
	public interface _IDTExtensibility2
	{
		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(1)]
		void OnConnection([In][MarshalAs(UnmanagedType.IDispatch)] object Application, [In] ext_ConnectMode ConnectMode, [In][MarshalAs(UnmanagedType.IDispatch)] object AddInInst, [In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(2)]
		void OnDisconnection([In] ext_DisconnectMode RemoveMode, [In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(3)]
		void OnAddInsUpdate([In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(4)]
		void OnStartupComplete([In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(5)]
		void OnBeginShutdown([In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
	}
}
