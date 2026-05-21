using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Word
{
	[ComImport]
	[Guid("0002093B-0000-0000-C000-000000000046")]
	[CompilerGenerated]
	[DefaultMember("Visible")]
	[TypeIdentifier]
	public interface Border
	{
		[DispId(0)]
		bool Visible
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(0)]
			[param: In]
			set;
		}

		[DispId(3)]
		WdLineStyle LineStyle
		{
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(3)]
			get;
			[MethodImpl(MethodImplOptions.InternalCall)]
			[DispId(3)]
			[param: In]
			set;
		}

		void _VtblGap1_3();

		void _VtblGap2_3();
	}
}
