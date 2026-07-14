using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[InterfaceType(2)]
[Guid("00024413-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface AppEvents
{
	void _VtblGap1_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1558)]
	void SheetSelectionChange([In][MarshalAs(UnmanagedType.IDispatch)] object Sh, [In][MarshalAs(UnmanagedType.Interface)] Range Target);

	void _VtblGap2_2();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1561)]
	void SheetActivate([In][MarshalAs(UnmanagedType.IDispatch)] object Sh);

	void _VtblGap3_2();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1564)]
	void SheetChange([In][MarshalAs(UnmanagedType.IDispatch)] object Sh, [In][MarshalAs(UnmanagedType.Interface)] Range Target);
}
