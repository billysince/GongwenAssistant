using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[CompilerGenerated]
	[Guid("00020958-0000-0000-C000-000000000046")]
	[TypeIdentifier]
	public interface Paragraphs : IEnumerable
	{
		[DispId(2)]
		int Count
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(2)]
			get;
		}

		[DispId(0)]
		Paragraph this[[In] int Index]
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[MethodImpl(MethodImplOptions.InternalCall)]
		[DispId(-4)]
		[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
		new IEnumerator GetEnumerator();

		void _VtblGap1_62();
	}
}
