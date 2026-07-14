using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[Guid("000C0306-0000-0000-C000-000000000046")]
[DefaultMember("Item")]
[TypeIdentifier]
public interface CommandBarControls : _IMsoDispObj, IEnumerable
{
	void _VtblGap1_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809344)]
	[return: MarshalAs(UnmanagedType.Interface)]
	CommandBarControl Add([Optional][In][MarshalAs(UnmanagedType.Struct)] object Type, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Id, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Parameter, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Before, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Temporary);
}
