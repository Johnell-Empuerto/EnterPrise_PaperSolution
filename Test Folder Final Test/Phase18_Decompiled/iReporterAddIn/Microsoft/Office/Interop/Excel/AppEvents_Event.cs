using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[ComEventInterface(typeof(AppEvents), typeof(AppEvents))]
[TypeIdentifier("00020813-0000-0000-c000-000000000046", "Microsoft.Office.Interop.Excel.AppEvents_Event")]
public interface AppEvents_Event
{
	void _VtblGap1_2();

	event AppEvents_SheetSelectionChangeEventHandler SheetSelectionChange;

	void _VtblGap2_4();

	event AppEvents_SheetActivateEventHandler SheetActivate;

	void _VtblGap3_4();

	event AppEvents_SheetChangeEventHandler SheetChange;
}
