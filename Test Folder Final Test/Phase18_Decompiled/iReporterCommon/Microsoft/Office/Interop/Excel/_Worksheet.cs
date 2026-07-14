using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[Guid("000208D8-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface _Worksheet
{
	void _VtblGap1_49();

	[DispId(1069)]
	Range CircularReference
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1069)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
