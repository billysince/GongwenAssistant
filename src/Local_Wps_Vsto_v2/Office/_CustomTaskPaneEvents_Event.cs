using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Office
{
	[ComImport]
	[Guid("BC2E1B7B-6A56-4D2C-8D5F-91D3DD7E0A01")]
	[ComEventInterface(typeof(_CustomTaskPaneEvents), typeof(_CustomTaskPaneEvents))]
	[TypeIdentifier("2df8d04c-5bfa-101b-bde5-00aa0044de52", "Office._CustomTaskPaneEvents_Event")]
	[CompilerGenerated]
	public interface _CustomTaskPaneEvents_Event
	{
		event _CustomTaskPaneEvents_VisibleStateChangeEventHandler VisibleStateChange;
	}
}
