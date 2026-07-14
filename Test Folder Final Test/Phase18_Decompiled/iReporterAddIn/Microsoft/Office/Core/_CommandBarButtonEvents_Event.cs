using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[ComEventInterface(typeof(_CommandBarButtonEvents), typeof(_CommandBarButtonEvents))]
[TypeIdentifier("2df8d04c-5bfa-101b-bde5-00aa0044de52", "Microsoft.Office.Core._CommandBarButtonEvents_Event")]
public interface _CommandBarButtonEvents_Event
{
	event _CommandBarButtonEvents_ClickEventHandler Click;
}
