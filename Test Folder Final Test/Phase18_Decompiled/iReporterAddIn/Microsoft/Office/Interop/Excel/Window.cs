using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[InterfaceType(2)]
[Guid("00020893-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface Window
{
	void _VtblGap1_9();

	[DispId(307)]
	object ActiveSheet
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(307)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap2_47();

	[DispId(656)]
	Sheets SelectedSheets
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(656)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
