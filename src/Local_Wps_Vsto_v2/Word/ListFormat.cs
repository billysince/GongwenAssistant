using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[Guid("000209C0-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	[CompilerGenerated]
	public interface ListFormat
	{
		[DispId(71)]
		int ListValue
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(71)]
			get;
		}

		void _VtblGap1_4();

		void _VtblGap2_9();

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(186)]
		void ConvertNumbersToText([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object NumberType);
	}
}
