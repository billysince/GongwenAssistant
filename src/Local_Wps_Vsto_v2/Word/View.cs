using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[DefaultMember("Type")]
	[Guid("000209A5-0000-0000-C000-000000000046")]
	[CompilerGenerated]
	[TypeIdentifier]
	public interface View
	{
		[DispId(0)]
		WdViewType Type
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[param: In]
			set;
		}

		[DispId(28)]
		WdSeekView SeekView
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(28)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(28)]
			[param: In]
			set;
		}

		void _VtblGap1_3();

		void _VtblGap2_51();
	}
}
