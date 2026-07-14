using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[InterfaceType(2)]
[Guid("00024427-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface Comment
{
	void _VtblGap1_2();

	[DispId(150)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(150)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap2_4();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(138)]
	[return: MarshalAs(UnmanagedType.BStr)]
	string Text([Optional][In][MarshalAs(UnmanagedType.Struct)] object Text, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Start, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Overwrite);
}
