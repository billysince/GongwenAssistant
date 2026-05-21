using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[TypeIdentifier]
	[Guid("00020972-0000-0000-C000-000000000046")]
	[CompilerGenerated]
	public interface LineNumbering
	{
		[DispId(104)]
		int Active
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(104)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(104)]
			[param: In]
			set;
		}

		void _VtblGap1_11();
	}
}
