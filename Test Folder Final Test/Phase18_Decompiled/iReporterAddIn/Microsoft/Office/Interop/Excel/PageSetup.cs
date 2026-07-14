using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[InterfaceType(2)]
[Guid("000208B4-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface PageSetup
{
	void _VtblGap1_41();

	[DispId(1019)]
	string PrintArea
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1019)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1019)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}
}
