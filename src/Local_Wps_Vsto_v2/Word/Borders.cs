using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[Guid("0002093C-0000-0000-C000-000000000046")]
	[CompilerGenerated]
	[TypeIdentifier]
	public interface Borders : IEnumerable
	{
		[DispId(0)]
		Border this[[In] WdBorderType Index]
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		void _VtblGap1_45();
	}
}
