using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Office
{
	[ComImport]
	[Guid("000C033E-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	[CompilerGenerated]
	public interface ICustomTaskPaneConsumer
	{
		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(1)]
		void CTPFactoryAvailable([In][MarshalAs(UnmanagedType.Interface)] ICTPFactory CTPFactoryInst);
	}
}
