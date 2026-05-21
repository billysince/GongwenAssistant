using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Excel
{
	[ComImport]
	[CompilerGenerated]
	[Guid("000208DB-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	[DefaultMember("_Default")]
	public interface Workbooks : IEnumerable
	{
		void _VtblGap1_3();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[LCIDConversion(1)]
		[DispId(181)]
		[return: MarshalAs(UnmanagedType.Interface)]
		Workbook Add([Optional][In][MarshalAs(UnmanagedType.Struct)] object Template);
	}
}
