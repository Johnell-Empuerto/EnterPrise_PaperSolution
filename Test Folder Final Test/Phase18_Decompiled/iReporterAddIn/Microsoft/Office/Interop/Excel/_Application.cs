using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[DefaultMember("_Default")]
[Guid("000208D5-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface _Application
{
	void _VtblGap1_9();

	[DispId(307)]
	object ActiveSheet
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(307)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[DispId(759)]
	Window ActiveWindow
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(759)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(308)]
	Workbook ActiveWorkbook
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(308)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_3();

	[DispId(238)]
	Range Cells
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(238)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_2();

	[DispId(1439)]
	CommandBars CommandBars
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1439)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_10();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(766)]
	[LCIDConversion(30)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Range Intersect([In][MarshalAs(UnmanagedType.Interface)] Range Arg1, [In][MarshalAs(UnmanagedType.Interface)] Range Arg2, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg3, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg4, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg5, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg6, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg7, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg8, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg9, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg10, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg11, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg12, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg13, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg14, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg15, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg16, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg17, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg18, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg19, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg20, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg21, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg22, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg23, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg24, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg25, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg26, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg27, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg28, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg29, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg30);

	void _VtblGap5_7();

	[DispId(147)]
	object Selection
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(147)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap6_6();

	[DispId(430)]
	Windows Windows
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(430)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap7_61();

	[DispId(0)]
	string _Default
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	void _VtblGap8_237();

	[DispId(1950)]
	int Hwnd
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1950)]
		get;
	}
}
