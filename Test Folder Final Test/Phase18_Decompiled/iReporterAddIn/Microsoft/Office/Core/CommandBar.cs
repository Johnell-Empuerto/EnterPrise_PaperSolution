using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[Guid("000C0304-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface CommandBar : _IMsoOleAccDispObj
{
	void _VtblGap1_26();

	[DispId(1610874883)]
	CommandBarControls Controls
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610874883)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_10();

	[DispId(1610874894)]
	string Name
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610874894)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610874894)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}
}
