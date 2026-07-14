
//  Microsoft (R) .NET Framework IL Disassembler.  Version 4.8.3928.0
//  Copyright (c) Microsoft Corporation.  All rights reserved.



// Metadata version: v4.0.30319
.module extern user32.dll
.assembly extern mscorlib
{
  .publickeytoken = (B7 7A 5C 56 19 34 E0 89 )                         // .z\V.4..
  .ver 4:0:0:0
}
.assembly extern System.Windows.Forms
{
  .publickeytoken = (B7 7A 5C 56 19 34 E0 89 )                         // .z\V.4..
  .ver 4:0:0:0
}
.assembly extern System
{
  .publickeytoken = (B7 7A 5C 56 19 34 E0 89 )                         // .z\V.4..
  .ver 4:0:0:0
}
.assembly extern 'Microsoft.Office.Tools.Common.v4.0.Utilities'
{
  .publickeytoken = (B0 3F 5F 7F 11 D5 0A 3A )                         // .?_....:
  .ver 10:0:0:0
}
.assembly extern Microsoft.Office.Tools.Common
{
  .publickeytoken = (B0 3F 5F 7F 11 D5 0A 3A )                         // .?_....:
  .ver 10:0:0:0
}
.assembly extern Microsoft.VisualStudio.Tools.Applications.Runtime
{
  .publickeytoken = (B0 3F 5F 7F 11 D5 0A 3A )                         // .?_....:
  .ver 10:0:0:0
}
.assembly extern iReporterExcelAddInCommon
{
  .ver 8:2:26060:1
}
.assembly extern Cimtops.R2Cluster
{
  .ver 8:2:26060:1
}
.assembly extern System.Drawing
{
  .publickeytoken = (B0 3F 5F 7F 11 D5 0A 3A )                         // .?_....:
  .ver 4:0:0:0
}
.assembly extern PresentationFramework
{
  .publickeytoken = (31 BF 38 56 AD 36 4E 35 )                         // 1.8V.6N5
  .ver 4:0:0:0
}
.assembly extern Microsoft.Office.Tools.Excel
{
  .publickeytoken = (B0 3F 5F 7F 11 D5 0A 3A )                         // .?_....:
  .ver 10:0:0:0
}
.assembly extern System.Core
{
  .publickeytoken = (B7 7A 5C 56 19 34 E0 89 )                         // .z\V.4..
  .ver 4:0:0:0
}
.assembly extern Cimtops.Excel
{
  .ver 8:2:26060:1
}
.assembly extern Microsoft.CSharp
{
  .publickeytoken = (B0 3F 5F 7F 11 D5 0A 3A )                         // .?_....:
  .ver 4:0:0:0
}
.assembly extern Microsoft.Office.Tools
{
  .publickeytoken = (B0 3F 5F 7F 11 D5 0A 3A )                         // .?_....:
  .ver 10:0:0:0
}
.assembly iReporterExcelAddIn
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilationRelaxationsAttribute::.ctor(int32) = ( 01 00 08 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.CompilerServices.RuntimeCompatibilityAttribute::.ctor() = ( 01 00 01 00 54 02 16 57 72 61 70 4E 6F 6E 45 78   // ....T..WrapNonEx
                                                                                                             63 65 70 74 69 6F 6E 54 68 72 6F 77 73 01 )       // ceptionThrows.

  // --- The following custom attribute is added automatically, do not uncomment -------
  //  .custom instance void [mscorlib]System.Diagnostics.DebuggableAttribute::.ctor(valuetype [mscorlib]System.Diagnostics.DebuggableAttribute/DebuggingModes) = ( 01 00 02 00 00 00 00 00 ) 

  .custom instance void [mscorlib]System.Reflection.AssemblyTitleAttribute::.ctor(string) = ( 01 00 13 69 52 65 70 6F 72 74 65 72 45 78 63 65   // ...iReporterExce
                                                                                              6C 41 64 64 49 6E 00 00 )                         // lAddIn..
  .custom instance void [mscorlib]System.Reflection.AssemblyDescriptionAttribute::.ctor(string) = ( 01 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Reflection.AssemblyConfigurationAttribute::.ctor(string) = ( 01 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Reflection.AssemblyCompanyAttribute::.ctor(string) = ( 01 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Reflection.AssemblyProductAttribute::.ctor(string) = ( 01 00 13 69 52 65 70 6F 72 74 65 72 45 78 63 65   // ...iReporterExce
                                                                                                6C 41 64 64 49 6E 00 00 )                         // lAddIn..
  .custom instance void [mscorlib]System.Reflection.AssemblyCopyrightAttribute::.ctor(string) = ( 01 00 24 43 6F 70 79 72 69 67 68 74 20 63 20 32   // ..$Copyright c 2
                                                                                                  30 31 36 20 43 49 4D 54 4F 50 53 20 43 4F 52 50   // 016 CIMTOPS CORP
                                                                                                  4F 52 41 54 49 4F 4E 00 00 )                      // ORATION..
  .custom instance void [mscorlib]System.Reflection.AssemblyTrademarkAttribute::.ctor(string) = ( 01 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComVisibleAttribute::.ctor(bool) = ( 01 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 32 30 65 37 32 61 39 63 2D 37 31 61 38   // ..$20e72a9c-71a8
                                                                                                  2D 34 38 32 65 2D 39 61 36 39 2D 66 62 32 33 36   // -482e-9a69-fb236
                                                                                                  31 62 61 37 37 63 38 00 00 )                      // 1ba77c8..
  .custom instance void [mscorlib]System.Reflection.AssemblyFileVersionAttribute::.ctor(string) = ( 01 00 0B 38 2E 32 2E 32 36 30 36 30 2E 31 00 00 ) // ...8.2.26060.1..
  .custom instance void [mscorlib]System.Runtime.Versioning.TargetFrameworkAttribute::.ctor(string) = ( 01 00 1A 2E 4E 45 54 46 72 61 6D 65 77 6F 72 6B   // ....NETFramework
                                                                                                        2C 56 65 72 73 69 6F 6E 3D 76 34 2E 35 01 00 54   // ,Version=v4.5..T
                                                                                                        0E 14 46 72 61 6D 65 77 6F 72 6B 44 69 73 70 6C   // ..FrameworkDispl
                                                                                                        61 79 4E 61 6D 65 12 2E 4E 45 54 20 46 72 61 6D   // ayName..NET Fram
                                                                                                        65 77 6F 72 6B 20 34 2E 35 )                      // ework 4.5
  .hash algorithm 0x00008004
  .ver 8:2:26060:1
}
.mresource public iReporterExcelAddIn2019.Properties.Resources.resources
{
  // Offset: 0x00000000 Length: 0x00000815
  // WARNING: managed resource file iReporterExcelAddIn2019.Properties.Resources.resources created
}
.mresource public iReporterExcelAddIn2019.Ribbon1.resources
{
  // Offset: 0x00000820 Length: 0x00002A1C
  // WARNING: managed resource file iReporterExcelAddIn2019.Ribbon1.resources created
}
.mresource public iReporterExcelAddIn2019.WaitDialog.resources
{
  // Offset: 0x00003240 Length: 0x00000F83
  // WARNING: managed resource file iReporterExcelAddIn2019.WaitDialog.resources created
}
.module iReporterExcelAddIn.dll
// MVID: {B9A234C8-C2C3-4272-A26F-1BF91665F2BD}
.imagebase 0x10000000
.file alignment 0x00000200
.stackreserve 0x00100000
.subsystem 0x0003       // WINDOWS_CUI
.corflags 0x00000001    //  ILONLY
// Image base: 0x058F0000


// =============== CLASS MEMBERS DECLARATION ===================

.class public sequential ansi sealed beforefieldinit RECT
       extends [mscorlib]System.ValueType
{
  .field public int32 Left
  .field public int32 Top
  .field public int32 Right
  .field public int32 Bottom
} // end of class RECT

.class public auto ansi beforefieldinit iReporterExcelAddIn2019.WaitDialog
       extends [System.Windows.Forms]System.Windows.Forms.Form
{
  .field public class [System.Windows.Forms]System.Windows.Forms.Label labelMainMsg
  .field private class [System.Windows.Forms]System.Windows.Forms.Button button1
  .field private class [System]System.ComponentModel.Container components
  .field private initonly class [mscorlib]System.Threading.CancellationTokenSource '<Src>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .method private hidebysig specialname instance class [mscorlib]System.Threading.CancellationTokenSource 
          get_Src() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [mscorlib]System.Threading.CancellationTokenSource iReporterExcelAddIn2019.WaitDialog::'<Src>k__BackingField'
    IL_0006:  ret
  } // end of method WaitDialog::get_Src

  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(class [mscorlib]System.Threading.CancellationTokenSource src) cil managed
  {
    // Code size       20 (0x14)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [System.Windows.Forms]System.Windows.Forms.Form::.ctor()
    IL_0006:  ldarg.0
    IL_0007:  call       instance void iReporterExcelAddIn2019.WaitDialog::InitializeComponent()
    IL_000c:  ldarg.0
    IL_000d:  ldarg.1
    IL_000e:  stfld      class [mscorlib]System.Threading.CancellationTokenSource iReporterExcelAddIn2019.WaitDialog::'<Src>k__BackingField'
    IL_0013:  ret
  } // end of method WaitDialog::.ctor

  .method family hidebysig virtual instance void 
          Dispose(bool disposing) cil managed
  {
    // Code size       30 (0x1e)
    .maxstack  8
    IL_0000:  ldarg.1
    IL_0001:  brfalse.s  IL_0016

    IL_0003:  ldarg.0
    IL_0004:  ldfld      class [System]System.ComponentModel.Container iReporterExcelAddIn2019.WaitDialog::components
    IL_0009:  brfalse.s  IL_0016

    IL_000b:  ldarg.0
    IL_000c:  ldfld      class [System]System.ComponentModel.Container iReporterExcelAddIn2019.WaitDialog::components
    IL_0011:  callvirt   instance void [System]System.ComponentModel.Container::Dispose()
    IL_0016:  ldarg.0
    IL_0017:  ldarg.1
    IL_0018:  call       instance void [System.Windows.Forms]System.Windows.Forms.Form::Dispose(bool)
    IL_001d:  ret
  } // end of method WaitDialog::Dispose

  .method private hidebysig instance void 
          InitializeComponent() cil managed
  {
    // Code size       243 (0xf3)
    .maxstack  4
    IL_0000:  ldtoken    iReporterExcelAddIn2019.WaitDialog
    IL_0005:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_000a:  newobj     instance void [System]System.ComponentModel.ComponentResourceManager::.ctor(class [mscorlib]System.Type)
    IL_000f:  ldarg.0
    IL_0010:  newobj     instance void [System.Windows.Forms]System.Windows.Forms.Label::.ctor()
    IL_0015:  stfld      class [System.Windows.Forms]System.Windows.Forms.Label iReporterExcelAddIn2019.WaitDialog::labelMainMsg
    IL_001a:  ldarg.0
    IL_001b:  newobj     instance void [System.Windows.Forms]System.Windows.Forms.Button::.ctor()
    IL_0020:  stfld      class [System.Windows.Forms]System.Windows.Forms.Button iReporterExcelAddIn2019.WaitDialog::button1
    IL_0025:  ldarg.0
    IL_0026:  call       instance void [System.Windows.Forms]System.Windows.Forms.Control::SuspendLayout()
    IL_002b:  dup
    IL_002c:  ldarg.0
    IL_002d:  ldfld      class [System.Windows.Forms]System.Windows.Forms.Label iReporterExcelAddIn2019.WaitDialog::labelMainMsg
    IL_0032:  ldstr      "labelMainMsg"
    IL_0037:  callvirt   instance void [System]System.ComponentModel.ComponentResourceManager::ApplyResources(object,
                                                                                                              string)
    IL_003c:  ldarg.0
    IL_003d:  ldfld      class [System.Windows.Forms]System.Windows.Forms.Label iReporterExcelAddIn2019.WaitDialog::labelMainMsg
    IL_0042:  ldstr      "labelMainMsg"
    IL_0047:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.Control::set_Name(string)
    IL_004c:  dup
    IL_004d:  ldarg.0
    IL_004e:  ldfld      class [System.Windows.Forms]System.Windows.Forms.Button iReporterExcelAddIn2019.WaitDialog::button1
    IL_0053:  ldstr      "button1"
    IL_0058:  callvirt   instance void [System]System.ComponentModel.ComponentResourceManager::ApplyResources(object,
                                                                                                              string)
    IL_005d:  ldarg.0
    IL_005e:  ldfld      class [System.Windows.Forms]System.Windows.Forms.Button iReporterExcelAddIn2019.WaitDialog::button1
    IL_0063:  ldstr      "button1"
    IL_0068:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.Control::set_Name(string)
    IL_006d:  ldarg.0
    IL_006e:  ldfld      class [System.Windows.Forms]System.Windows.Forms.Button iReporterExcelAddIn2019.WaitDialog::button1
    IL_0073:  ldc.i4.1
    IL_0074:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.ButtonBase::set_UseVisualStyleBackColor(bool)
    IL_0079:  ldarg.0
    IL_007a:  ldfld      class [System.Windows.Forms]System.Windows.Forms.Button iReporterExcelAddIn2019.WaitDialog::button1
    IL_007f:  ldarg.0
    IL_0080:  ldftn      instance void iReporterExcelAddIn2019.WaitDialog::button1_Click(object,
                                                                                         class [mscorlib]System.EventArgs)
    IL_0086:  newobj     instance void [mscorlib]System.EventHandler::.ctor(object,
                                                                            native int)
    IL_008b:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.Control::add_Click(class [mscorlib]System.EventHandler)
    IL_0090:  ldarg.0
    IL_0091:  ldstr      "$this"
    IL_0096:  callvirt   instance void [System]System.ComponentModel.ComponentResourceManager::ApplyResources(object,
                                                                                                              string)
    IL_009b:  ldarg.0
    IL_009c:  ldc.i4.0
    IL_009d:  call       instance void [System.Windows.Forms]System.Windows.Forms.Form::set_ControlBox(bool)
    IL_00a2:  ldarg.0
    IL_00a3:  call       instance class [System.Windows.Forms]System.Windows.Forms.Control/ControlCollection [System.Windows.Forms]System.Windows.Forms.Control::get_Controls()
    IL_00a8:  ldarg.0
    IL_00a9:  ldfld      class [System.Windows.Forms]System.Windows.Forms.Button iReporterExcelAddIn2019.WaitDialog::button1
    IL_00ae:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.Control/ControlCollection::Add(class [System.Windows.Forms]System.Windows.Forms.Control)
    IL_00b3:  ldarg.0
    IL_00b4:  call       instance class [System.Windows.Forms]System.Windows.Forms.Control/ControlCollection [System.Windows.Forms]System.Windows.Forms.Control::get_Controls()
    IL_00b9:  ldarg.0
    IL_00ba:  ldfld      class [System.Windows.Forms]System.Windows.Forms.Label iReporterExcelAddIn2019.WaitDialog::labelMainMsg
    IL_00bf:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.Control/ControlCollection::Add(class [System.Windows.Forms]System.Windows.Forms.Control)
    IL_00c4:  ldarg.0
    IL_00c5:  ldc.i4.3
    IL_00c6:  call       instance void [System.Windows.Forms]System.Windows.Forms.Form::set_FormBorderStyle(valuetype [System.Windows.Forms]System.Windows.Forms.FormBorderStyle)
    IL_00cb:  ldarg.0
    IL_00cc:  ldc.i4.0
    IL_00cd:  call       instance void [System.Windows.Forms]System.Windows.Forms.Form::set_MaximizeBox(bool)
    IL_00d2:  ldarg.0
    IL_00d3:  ldc.i4.0
    IL_00d4:  call       instance void [System.Windows.Forms]System.Windows.Forms.Form::set_MinimizeBox(bool)
    IL_00d9:  ldarg.0
    IL_00da:  ldstr      "WaitDialog"
    IL_00df:  call       instance void [System.Windows.Forms]System.Windows.Forms.Control::set_Name(string)
    IL_00e4:  ldarg.0
    IL_00e5:  ldc.i4.0
    IL_00e6:  call       instance void [System.Windows.Forms]System.Windows.Forms.Form::set_ShowInTaskbar(bool)
    IL_00eb:  ldarg.0
    IL_00ec:  ldc.i4.0
    IL_00ed:  call       instance void [System.Windows.Forms]System.Windows.Forms.Control::ResumeLayout(bool)
    IL_00f2:  ret
  } // end of method WaitDialog::InitializeComponent

  .method private hidebysig instance void 
          button1_Click(object sender,
                        class [mscorlib]System.EventArgs e) cil managed
  {
    // Code size       18 (0x12)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance class [mscorlib]System.Threading.CancellationTokenSource iReporterExcelAddIn2019.WaitDialog::get_Src()
    IL_0006:  callvirt   instance void [mscorlib]System.Threading.CancellationTokenSource::Cancel()
    IL_000b:  ldarg.0
    IL_000c:  call       instance void [System.Windows.Forms]System.Windows.Forms.Form::Close()
    IL_0011:  ret
  } // end of method WaitDialog::button1_Click

  .property instance class [mscorlib]System.Threading.CancellationTokenSource
          Src()
  {
    .get instance class [mscorlib]System.Threading.CancellationTokenSource iReporterExcelAddIn2019.WaitDialog::get_Src()
  } // end of property WaitDialog::Src
} // end of class iReporterExcelAddIn2019.WaitDialog

.class public auto ansi beforefieldinit iReporterExcelAddIn2019.Ribbon1
       extends ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase
{
  .field private class [System]System.ComponentModel.IContainer components
  .field assembly class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab tab1
  .field assembly class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup group1
  .field assembly class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton autoJudgeButton
  .field assembly class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton clusterSettingButton
  .field assembly class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton colorSettingButton
  .field assembly class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton tableSettingButton
  .field assembly class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton checkSheetNameButton
  .method private hidebysig instance void 
          Ribbon1_Load(object sender,
                       class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs e) cil managed
  {
    // Code size       1 (0x1)
    .maxstack  8
    IL_0000:  ret
  } // end of method Ribbon1::Ribbon1_Load

  .method private hidebysig instance void 
          autoJudgeButton_Click(object sender,
                                class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e) cil managed
  {
    // Code size       12 (0xc)
    .maxstack  8
    IL_0000:  call       class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.Globals::get_ThisAddIn()
    IL_0005:  ldc.i4.1
    IL_0006:  callvirt   instance void iReporterExcelAddIn2019.ThisAddIn::AutoJudge(bool)
    IL_000b:  ret
  } // end of method Ribbon1::autoJudgeButton_Click

  .method private hidebysig instance void 
          clusterSettingButton_Click(object sender,
                                     class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e) cil managed
  {
    // Code size       11 (0xb)
    .maxstack  8
    IL_0000:  call       class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.Globals::get_ThisAddIn()
    IL_0005:  callvirt   instance void iReporterExcelAddIn2019.ThisAddIn::ManualCluster()
    IL_000a:  ret
  } // end of method Ribbon1::clusterSettingButton_Click

  .method private hidebysig instance void 
          colorSettingButton_Click(object sender,
                                   class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e) cil managed
  {
    // Code size       11 (0xb)
    .maxstack  8
    IL_0000:  call       class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.Globals::get_ThisAddIn()
    IL_0005:  callvirt   instance void iReporterExcelAddIn2019.ThisAddIn::SetClusterColors()
    IL_000a:  ret
  } // end of method Ribbon1::colorSettingButton_Click

  .method private hidebysig instance void 
          tableSettingButton_Click(object sender,
                                   class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e) cil managed
  {
    // Code size       11 (0xb)
    .maxstack  8
    IL_0000:  call       class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.Globals::get_ThisAddIn()
    IL_0005:  callvirt   instance void iReporterExcelAddIn2019.ThisAddIn::ShowTableForm()
    IL_000a:  ret
  } // end of method Ribbon1::tableSettingButton_Click

  .method private hidebysig instance void 
          checkSheetNameButton_Click(object sender,
                                     class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e) cil managed
  {
    // Code size       11 (0xb)
    .maxstack  8
    IL_0000:  call       class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.Globals::get_ThisAddIn()
    IL_0005:  callvirt   instance void iReporterExcelAddIn2019.ThisAddIn::CheckAllSheetName()
    IL_000a:  ret
  } // end of method Ribbon1::checkSheetNameButton_Click

  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       23 (0x17)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory iReporterExcelAddIn2019.Globals::get_Factory()
    IL_0006:  callvirt   instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Factory::GetRibbonFactory()
    IL_000b:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::.ctor(class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory)
    IL_0010:  ldarg.0
    IL_0011:  call       instance void iReporterExcelAddIn2019.Ribbon1::InitializeComponent()
    IL_0016:  ret
  } // end of method Ribbon1::.ctor

  .method family hidebysig virtual instance void 
          Dispose(bool disposing) cil managed
  {
    // Code size       30 (0x1e)
    .maxstack  8
    IL_0000:  ldarg.1
    IL_0001:  brfalse.s  IL_0016

    IL_0003:  ldarg.0
    IL_0004:  ldfld      class [System]System.ComponentModel.IContainer iReporterExcelAddIn2019.Ribbon1::components
    IL_0009:  brfalse.s  IL_0016

    IL_000b:  ldarg.0
    IL_000c:  ldfld      class [System]System.ComponentModel.IContainer iReporterExcelAddIn2019.Ribbon1::components
    IL_0011:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
    IL_0016:  ldarg.0
    IL_0017:  ldarg.1
    IL_0018:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::Dispose(bool)
    IL_001d:  ret
  } // end of method Ribbon1::Dispose

  .method private hidebysig instance void 
          InitializeComponent() cil managed
  {
    // Code size       887 (0x377)
    .maxstack  4
    IL_0000:  ldtoken    iReporterExcelAddIn2019.Ribbon1
    IL_0005:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_000a:  newobj     instance void [System]System.ComponentModel.ComponentResourceManager::.ctor(class [mscorlib]System.Type)
    IL_000f:  ldarg.0
    IL_0010:  ldarg.0
    IL_0011:  call       instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::get_Factory()
    IL_0016:  callvirt   instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory::CreateRibbonTab()
    IL_001b:  stfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab iReporterExcelAddIn2019.Ribbon1::tab1
    IL_0020:  ldarg.0
    IL_0021:  ldarg.0
    IL_0022:  call       instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::get_Factory()
    IL_0027:  callvirt   instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory::CreateRibbonGroup()
    IL_002c:  stfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup iReporterExcelAddIn2019.Ribbon1::group1
    IL_0031:  ldarg.0
    IL_0032:  ldarg.0
    IL_0033:  call       instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::get_Factory()
    IL_0038:  callvirt   instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory::CreateRibbonButton()
    IL_003d:  stfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::autoJudgeButton
    IL_0042:  ldarg.0
    IL_0043:  ldarg.0
    IL_0044:  call       instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::get_Factory()
    IL_0049:  callvirt   instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory::CreateRibbonButton()
    IL_004e:  stfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::clusterSettingButton
    IL_0053:  ldarg.0
    IL_0054:  ldarg.0
    IL_0055:  call       instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::get_Factory()
    IL_005a:  callvirt   instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory::CreateRibbonButton()
    IL_005f:  stfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::colorSettingButton
    IL_0064:  ldarg.0
    IL_0065:  ldarg.0
    IL_0066:  call       instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::get_Factory()
    IL_006b:  callvirt   instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory::CreateRibbonButton()
    IL_0070:  stfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::tableSettingButton
    IL_0075:  ldarg.0
    IL_0076:  ldarg.0
    IL_0077:  call       instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::get_Factory()
    IL_007c:  callvirt   instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory::CreateRibbonButton()
    IL_0081:  stfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::checkSheetNameButton
    IL_0086:  ldarg.0
    IL_0087:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab iReporterExcelAddIn2019.Ribbon1::tab1
    IL_008c:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::SuspendLayout()
    IL_0091:  ldarg.0
    IL_0092:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup iReporterExcelAddIn2019.Ribbon1::group1
    IL_0097:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::SuspendLayout()
    IL_009c:  ldarg.0
    IL_009d:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::SuspendLayout()
    IL_00a2:  ldarg.0
    IL_00a3:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab iReporterExcelAddIn2019.Ribbon1::tab1
    IL_00a8:  callvirt   instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlId [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab::get_ControlId()
    IL_00ad:  ldc.i4.1
    IL_00ae:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlId::set_ControlIdType(valuetype [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlIdType)
    IL_00b3:  ldarg.0
    IL_00b4:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab iReporterExcelAddIn2019.Ribbon1::tab1
    IL_00b9:  callvirt   instance class [mscorlib]System.Collections.Generic.IList`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup> [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab::get_Groups()
    IL_00be:  ldarg.0
    IL_00bf:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup iReporterExcelAddIn2019.Ribbon1::group1
    IL_00c4:  callvirt   instance void class [mscorlib]System.Collections.Generic.ICollection`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup>::Add(!0)
    IL_00c9:  dup
    IL_00ca:  ldarg.0
    IL_00cb:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab iReporterExcelAddIn2019.Ribbon1::tab1
    IL_00d0:  ldstr      "tab1"
    IL_00d5:  callvirt   instance void [System]System.ComponentModel.ComponentResourceManager::ApplyResources(object,
                                                                                                              string)
    IL_00da:  ldarg.0
    IL_00db:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab iReporterExcelAddIn2019.Ribbon1::tab1
    IL_00e0:  ldstr      "tab1"
    IL_00e5:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::set_Name(string)
    IL_00ea:  ldarg.0
    IL_00eb:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup iReporterExcelAddIn2019.Ribbon1::group1
    IL_00f0:  callvirt   instance class [mscorlib]System.Collections.Generic.IList`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControl> [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup::get_Items()
    IL_00f5:  ldarg.0
    IL_00f6:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::autoJudgeButton
    IL_00fb:  callvirt   instance void class [mscorlib]System.Collections.Generic.ICollection`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControl>::Add(!0)
    IL_0100:  ldarg.0
    IL_0101:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup iReporterExcelAddIn2019.Ribbon1::group1
    IL_0106:  callvirt   instance class [mscorlib]System.Collections.Generic.IList`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControl> [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup::get_Items()
    IL_010b:  ldarg.0
    IL_010c:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::clusterSettingButton
    IL_0111:  callvirt   instance void class [mscorlib]System.Collections.Generic.ICollection`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControl>::Add(!0)
    IL_0116:  ldarg.0
    IL_0117:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup iReporterExcelAddIn2019.Ribbon1::group1
    IL_011c:  callvirt   instance class [mscorlib]System.Collections.Generic.IList`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControl> [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup::get_Items()
    IL_0121:  ldarg.0
    IL_0122:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::colorSettingButton
    IL_0127:  callvirt   instance void class [mscorlib]System.Collections.Generic.ICollection`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControl>::Add(!0)
    IL_012c:  ldarg.0
    IL_012d:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup iReporterExcelAddIn2019.Ribbon1::group1
    IL_0132:  callvirt   instance class [mscorlib]System.Collections.Generic.IList`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControl> [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup::get_Items()
    IL_0137:  ldarg.0
    IL_0138:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::tableSettingButton
    IL_013d:  callvirt   instance void class [mscorlib]System.Collections.Generic.ICollection`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControl>::Add(!0)
    IL_0142:  ldarg.0
    IL_0143:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup iReporterExcelAddIn2019.Ribbon1::group1
    IL_0148:  callvirt   instance class [mscorlib]System.Collections.Generic.IList`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControl> [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup::get_Items()
    IL_014d:  ldarg.0
    IL_014e:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::checkSheetNameButton
    IL_0153:  callvirt   instance void class [mscorlib]System.Collections.Generic.ICollection`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControl>::Add(!0)
    IL_0158:  dup
    IL_0159:  ldarg.0
    IL_015a:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup iReporterExcelAddIn2019.Ribbon1::group1
    IL_015f:  ldstr      "group1"
    IL_0164:  callvirt   instance void [System]System.ComponentModel.ComponentResourceManager::ApplyResources(object,
                                                                                                              string)
    IL_0169:  ldarg.0
    IL_016a:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup iReporterExcelAddIn2019.Ribbon1::group1
    IL_016f:  ldstr      "group1"
    IL_0174:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::set_Name(string)
    IL_0179:  ldarg.0
    IL_017a:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::autoJudgeButton
    IL_017f:  ldc.i4.1
    IL_0180:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_ControlSize(valuetype Microsoft.Office.Core.RibbonControlSize)
    IL_0185:  dup
    IL_0186:  ldarg.0
    IL_0187:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::autoJudgeButton
    IL_018c:  ldstr      "autoJudgeButton"
    IL_0191:  callvirt   instance void [System]System.ComponentModel.ComponentResourceManager::ApplyResources(object,
                                                                                                              string)
    IL_0196:  ldarg.0
    IL_0197:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::autoJudgeButton
    IL_019c:  ldstr      "autoJudgeButton"
    IL_01a1:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::set_Name(string)
    IL_01a6:  ldarg.0
    IL_01a7:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::autoJudgeButton
    IL_01ac:  ldc.i4.1
    IL_01ad:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_ShowImage(bool)
    IL_01b2:  ldarg.0
    IL_01b3:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::autoJudgeButton
    IL_01b8:  ldarg.0
    IL_01b9:  ldftn      instance void iReporterExcelAddIn2019.Ribbon1::autoJudgeButton_Click(object,
                                                                                              class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
    IL_01bf:  newobj     instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler::.ctor(object,
                                                                                                                                     native int)
    IL_01c4:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::add_Click(class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler)
    IL_01c9:  ldarg.0
    IL_01ca:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::clusterSettingButton
    IL_01cf:  ldc.i4.1
    IL_01d0:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_ControlSize(valuetype Microsoft.Office.Core.RibbonControlSize)
    IL_01d5:  dup
    IL_01d6:  ldarg.0
    IL_01d7:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::clusterSettingButton
    IL_01dc:  ldstr      "clusterSettingButton"
    IL_01e1:  callvirt   instance void [System]System.ComponentModel.ComponentResourceManager::ApplyResources(object,
                                                                                                              string)
    IL_01e6:  ldarg.0
    IL_01e7:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::clusterSettingButton
    IL_01ec:  ldstr      "clusterSettingButton"
    IL_01f1:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::set_Name(string)
    IL_01f6:  ldarg.0
    IL_01f7:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::clusterSettingButton
    IL_01fc:  ldc.i4.1
    IL_01fd:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_ShowImage(bool)
    IL_0202:  ldarg.0
    IL_0203:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::clusterSettingButton
    IL_0208:  ldarg.0
    IL_0209:  ldftn      instance void iReporterExcelAddIn2019.Ribbon1::clusterSettingButton_Click(object,
                                                                                                   class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
    IL_020f:  newobj     instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler::.ctor(object,
                                                                                                                                     native int)
    IL_0214:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::add_Click(class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler)
    IL_0219:  ldarg.0
    IL_021a:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::colorSettingButton
    IL_021f:  ldc.i4.1
    IL_0220:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_ControlSize(valuetype Microsoft.Office.Core.RibbonControlSize)
    IL_0225:  dup
    IL_0226:  ldarg.0
    IL_0227:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::colorSettingButton
    IL_022c:  ldstr      "colorSettingButton"
    IL_0231:  callvirt   instance void [System]System.ComponentModel.ComponentResourceManager::ApplyResources(object,
                                                                                                              string)
    IL_0236:  ldarg.0
    IL_0237:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::colorSettingButton
    IL_023c:  ldstr      "colorSettingButton"
    IL_0241:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::set_Name(string)
    IL_0246:  ldarg.0
    IL_0247:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::colorSettingButton
    IL_024c:  ldc.i4.1
    IL_024d:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_ShowImage(bool)
    IL_0252:  ldarg.0
    IL_0253:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::colorSettingButton
    IL_0258:  ldarg.0
    IL_0259:  ldftn      instance void iReporterExcelAddIn2019.Ribbon1::colorSettingButton_Click(object,
                                                                                                 class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
    IL_025f:  newobj     instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler::.ctor(object,
                                                                                                                                     native int)
    IL_0264:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::add_Click(class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler)
    IL_0269:  ldarg.0
    IL_026a:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::tableSettingButton
    IL_026f:  ldc.i4.1
    IL_0270:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_ControlSize(valuetype Microsoft.Office.Core.RibbonControlSize)
    IL_0275:  dup
    IL_0276:  ldarg.0
    IL_0277:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::tableSettingButton
    IL_027c:  ldstr      "tableSettingButton"
    IL_0281:  callvirt   instance void [System]System.ComponentModel.ComponentResourceManager::ApplyResources(object,
                                                                                                              string)
    IL_0286:  ldarg.0
    IL_0287:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::tableSettingButton
    IL_028c:  ldstr      "tableSettingButton"
    IL_0291:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::set_Name(string)
    IL_0296:  ldarg.0
    IL_0297:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::tableSettingButton
    IL_029c:  ldc.i4.1
    IL_029d:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_ShowImage(bool)
    IL_02a2:  ldarg.0
    IL_02a3:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::tableSettingButton
    IL_02a8:  ldarg.0
    IL_02a9:  ldftn      instance void iReporterExcelAddIn2019.Ribbon1::tableSettingButton_Click(object,
                                                                                                 class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
    IL_02af:  newobj     instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler::.ctor(object,
                                                                                                                                     native int)
    IL_02b4:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::add_Click(class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler)
    IL_02b9:  ldarg.0
    IL_02ba:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::checkSheetNameButton
    IL_02bf:  ldc.i4.1
    IL_02c0:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_ControlSize(valuetype Microsoft.Office.Core.RibbonControlSize)
    IL_02c5:  ldarg.0
    IL_02c6:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::checkSheetNameButton
    IL_02cb:  ldstr      "checkSheetNameButton"
    IL_02d0:  callvirt   instance void [System]System.ComponentModel.ComponentResourceManager::ApplyResources(object,
                                                                                                              string)
    IL_02d5:  ldarg.0
    IL_02d6:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::checkSheetNameButton
    IL_02db:  ldstr      "checkSheetNameButton"
    IL_02e0:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::set_Name(string)
    IL_02e5:  ldarg.0
    IL_02e6:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::checkSheetNameButton
    IL_02eb:  ldc.i4.1
    IL_02ec:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_ShowImage(bool)
    IL_02f1:  ldarg.0
    IL_02f2:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::checkSheetNameButton
    IL_02f7:  ldarg.0
    IL_02f8:  ldftn      instance void iReporterExcelAddIn2019.Ribbon1::checkSheetNameButton_Click(object,
                                                                                                   class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
    IL_02fe:  newobj     instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler::.ctor(object,
                                                                                                                                     native int)
    IL_0303:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::add_Click(class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler)
    IL_0308:  ldarg.0
    IL_0309:  ldstr      "Ribbon1"
    IL_030e:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::set_Name(string)
    IL_0313:  ldarg.0
    IL_0314:  ldstr      "Microsoft.Excel.Workbook"
    IL_0319:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::set_RibbonType(string)
    IL_031e:  ldarg.0
    IL_031f:  call       instance class [mscorlib]System.Collections.Generic.IList`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab> ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::get_Tabs()
    IL_0324:  ldarg.0
    IL_0325:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab iReporterExcelAddIn2019.Ribbon1::tab1
    IL_032a:  callvirt   instance void class [mscorlib]System.Collections.Generic.ICollection`1<class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab>::Add(!0)
    IL_032f:  ldarg.0
    IL_0330:  ldarg.0
    IL_0331:  ldftn      instance void iReporterExcelAddIn2019.Ribbon1::Ribbon1_Load(object,
                                                                                     class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs)
    IL_0337:  newobj     instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler::.ctor(object,
                                                                                                                                native int)
    IL_033c:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::add_Load(class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler)
    IL_0341:  ldarg.0
    IL_0342:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab iReporterExcelAddIn2019.Ribbon1::tab1
    IL_0347:  ldc.i4.0
    IL_0348:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::ResumeLayout(bool)
    IL_034d:  ldarg.0
    IL_034e:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonTab iReporterExcelAddIn2019.Ribbon1::tab1
    IL_0353:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::PerformLayout()
    IL_0358:  ldarg.0
    IL_0359:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup iReporterExcelAddIn2019.Ribbon1::group1
    IL_035e:  ldc.i4.0
    IL_035f:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::ResumeLayout(bool)
    IL_0364:  ldarg.0
    IL_0365:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonGroup iReporterExcelAddIn2019.Ribbon1::group1
    IL_036a:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonComponent::PerformLayout()
    IL_036f:  ldarg.0
    IL_0370:  ldc.i4.0
    IL_0371:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonBase::ResumeLayout(bool)
    IL_0376:  ret
  } // end of method Ribbon1::InitializeComponent

} // end of class iReporterExcelAddIn2019.Ribbon1

.class private auto ansi sealed beforefieldinit iReporterExcelAddIn2019.ThisRibbonCollection
       extends ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonCollectionBase
{
  .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                      string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                  75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                  4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                  6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                  2E 30 2E 30 00 00 )                               // .0.0..
  .method assembly hidebysig specialname 
          instance class iReporterExcelAddIn2019.Ribbon1 
          get_Ribbon1() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance !!0 ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonCollectionBase::GetRibbon<class iReporterExcelAddIn2019.Ribbon1>()
    IL_0006:  ret
  } // end of method ThisRibbonCollection::get_Ribbon1

  .method assembly hidebysig specialname rtspecialname 
          instance void  .ctor(class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory factory) cil managed
  {
    // Code size       8 (0x8)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.Ribbon.RibbonCollectionBase::.ctor(class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory)
    IL_0007:  ret
  } // end of method ThisRibbonCollection::.ctor

  .property instance class iReporterExcelAddIn2019.Ribbon1
          Ribbon1()
  {
    .get instance class iReporterExcelAddIn2019.Ribbon1 iReporterExcelAddIn2019.ThisRibbonCollection::get_Ribbon1()
  } // end of property ThisRibbonCollection::Ribbon1
} // end of class iReporterExcelAddIn2019.ThisRibbonCollection

.class public auto ansi sealed beforefieldinit iReporterExcelAddIn2019.ThisAddIn
       extends ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase
{
  .custom instance void [Microsoft.VisualStudio.Tools.Applications.Runtime]Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
  .permissionset demand
             = {[mscorlib]System.Security.Permissions.PermissionSetAttribute = {property string 'Name' = string('FullTrust')}}
  .class auto ansi nested private beforefieldinit Win32Window
         extends [mscorlib]System.Object
         implements [System.Windows.Forms]System.Windows.Forms.IWin32Window
  {
    .field public native int _handle
    .method public hidebysig newslot specialname virtual final 
            instance native int  get_Handle() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  ldfld      native int iReporterExcelAddIn2019.ThisAddIn/Win32Window::_handle
      IL_0006:  ret
    } // end of method Win32Window::get_Handle

    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
      IL_0006:  ret
    } // end of method Win32Window::.ctor

    .property instance native int Handle()
    {
      .get instance native int iReporterExcelAddIn2019.ThisAddIn/Win32Window::get_Handle()
    } // end of property Win32Window::Handle
  } // end of class Win32Window

  .class auto ansi nested private beforefieldinit Judged
         extends [mscorlib]System.Object
  {
    .field public class [mscorlib]System.Exception ex
    .field public class [mscorlib]System.Collections.Generic.List`1<string> msgs
    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       18 (0x12)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  newobj     instance void class [mscorlib]System.Collections.Generic.List`1<string>::.ctor()
      IL_0006:  stfld      class [mscorlib]System.Collections.Generic.List`1<string> iReporterExcelAddIn2019.ThisAddIn/Judged::msgs
      IL_000b:  ldarg.0
      IL_000c:  call       instance void [mscorlib]System.Object::.ctor()
      IL_0011:  ret
    } // end of method Judged::.ctor

  } // end of class Judged

  .class abstract auto ansi sealed nested private beforefieldinit '<>o__25'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> '<>p__0'
  } // end of class '<>o__25'

  .class auto ansi serializable sealed nested private beforefieldinit '<>c'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static initonly class iReporterExcelAddIn2019.ThisAddIn/'<>c' '<>9'
    .field public static class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>> '<>9__25_0'
    .field public static class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>> '<>9__25_1'
    .field public static class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`3<string,string,string>> '<>9__25_2'
    .field public static class [mscorlib]System.Func`2<class [mscorlib]System.Tuple`2<int32,string>,int32> '<>9__36_0'
    .field public static class [mscorlib]System.Func`2<int32,bool> '<>9__36_1'
    .field public static class [mscorlib]System.Comparison`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table> '<>9__45_0'
    .field public static class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Models.SheetNameInvalidCharacter,class [mscorlib]System.Collections.Generic.List`1<string>> '<>9__48_0'
    .field public static class [mscorlib]System.Func`2<class Microsoft.Office.Interop.Excel.Worksheet,int32> '<>9__51_0'
    .field public static class [mscorlib]System.Func`2<class Microsoft.Office.Interop.Excel.Worksheet,string> '<>9__52_0'
    .method private hidebysig specialname rtspecialname static 
            void  .cctor() cil managed
    {
      // Code size       11 (0xb)
      .maxstack  8
      IL_0000:  newobj     instance void iReporterExcelAddIn2019.ThisAddIn/'<>c'::.ctor()
      IL_0005:  stsfld     class iReporterExcelAddIn2019.ThisAddIn/'<>c' iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9'
      IL_000a:  ret
    } // end of method '<>c'::.cctor

    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
      IL_0006:  ret
    } // end of method '<>c'::.ctor

    .method assembly hidebysig instance class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color> 
            '<AutoJudgeAsync>b__25_0'(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple t) cil managed
    {
      // Code size       28 (0x1c)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  ldfld      string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple::Item1
      IL_0006:  call       valuetype [System.Drawing]System.Drawing.Color [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Util::ToColor(string)
      IL_000b:  ldarg.1
      IL_000c:  ldfld      string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple::Item2
      IL_0011:  call       valuetype [System.Drawing]System.Drawing.Color [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Util::ToColor(string)
      IL_0016:  call       class [mscorlib]System.Tuple`2<!!0,!!1> [mscorlib]System.Tuple::Create<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>(!!0,
                                                                                                                                                                                                 !!1)
      IL_001b:  ret
    } // end of method '<>c'::'<AutoJudgeAsync>b__25_0'

    .method assembly hidebysig instance class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color> 
            '<AutoJudgeAsync>b__25_1'(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple t) cil managed
    {
      // Code size       28 (0x1c)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  ldfld      string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple::Item1
      IL_0006:  call       valuetype [System.Drawing]System.Drawing.Color [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Util::ToColor(string)
      IL_000b:  ldarg.1
      IL_000c:  ldfld      string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple::Item2
      IL_0011:  call       valuetype [System.Drawing]System.Drawing.Color [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Util::ToColor(string)
      IL_0016:  call       class [mscorlib]System.Tuple`2<!!0,!!1> [mscorlib]System.Tuple::Create<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>(!!0,
                                                                                                                                                                                                 !!1)
      IL_001b:  ret
    } // end of method '<>c'::'<AutoJudgeAsync>b__25_1'

    .method assembly hidebysig instance class [mscorlib]System.Tuple`3<string,string,string> 
            '<AutoJudgeAsync>b__25_2'(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple t) cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance class [mscorlib]System.Tuple`3<string,string,string> [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple::ToTuple()
      IL_0006:  ret
    } // end of method '<>c'::'<AutoJudgeAsync>b__25_2'

    .method assembly hidebysig instance int32 
            '<WriteComment>b__36_0'(class [mscorlib]System.Tuple`2<int32,string> t) cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance !0 class [mscorlib]System.Tuple`2<int32,string>::get_Item1()
      IL_0006:  ret
    } // end of method '<>c'::'<WriteComment>b__36_0'

    .method assembly hidebysig instance bool 
            '<WriteComment>b__36_1'(int32 n) cil managed
    {
      // Code size       14 (0xe)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  brfalse.s  IL_000c

      IL_0003:  ldarg.1
      IL_0004:  ldc.i4.1
      IL_0005:  beq.s      IL_000c

      IL_0007:  ldarg.1
      IL_0008:  ldc.i4.2
      IL_0009:  ceq
      IL_000b:  ret

      IL_000c:  ldc.i4.1
      IL_000d:  ret
    } // end of method '<>c'::'<WriteComment>b__36_1'

    .method assembly hidebysig instance int32 
            '<ShowTableForm>b__45_0'(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table lhs,
                                     class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table rhs) cil managed
    {
      // Code size       64 (0x40)
      .maxstack  2
      .locals init (valuetype [mscorlib]System.Nullable`1<int32> V_0)
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::get_No()
      IL_0006:  call       valuetype [mscorlib]System.Nullable`1<int32> [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Util::TryParseInt(string)
      IL_000b:  stloc.0
      IL_000c:  ldloca.s   V_0
      IL_000e:  call       instance bool valuetype [mscorlib]System.Nullable`1<int32>::get_HasValue()
      IL_0013:  brtrue.s   IL_0018

      IL_0015:  ldc.i4.0
      IL_0016:  br.s       IL_001f

      IL_0018:  ldloca.s   V_0
      IL_001a:  call       instance !0 valuetype [mscorlib]System.Nullable`1<int32>::GetValueOrDefault()
      IL_001f:  ldarg.2
      IL_0020:  callvirt   instance string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::get_No()
      IL_0025:  call       valuetype [mscorlib]System.Nullable`1<int32> [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Util::TryParseInt(string)
      IL_002a:  stloc.0
      IL_002b:  ldloca.s   V_0
      IL_002d:  call       instance bool valuetype [mscorlib]System.Nullable`1<int32>::get_HasValue()
      IL_0032:  brtrue.s   IL_0037

      IL_0034:  ldc.i4.0
      IL_0035:  br.s       IL_003e

      IL_0037:  ldloca.s   V_0
      IL_0039:  call       instance !0 valuetype [mscorlib]System.Nullable`1<int32>::GetValueOrDefault()
      IL_003e:  sub
      IL_003f:  ret
    } // end of method '<>c'::'<ShowTableForm>b__45_0'

    .method assembly hidebysig instance class [mscorlib]System.Collections.Generic.List`1<string> 
            '<CheckAllSheetName>b__48_0'(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Models.SheetNameInvalidCharacter x) cil managed
    {
      // Code size       30 (0x1e)
      .maxstack  8
      IL_0000:  newobj     instance void class [mscorlib]System.Collections.Generic.List`1<string>::.ctor()
      IL_0005:  dup
      IL_0006:  ldarg.1
      IL_0007:  callvirt   instance string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Models.SheetNameInvalidCharacter::get_Symbol()
      IL_000c:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<string>::Add(!0)
      IL_0011:  dup
      IL_0012:  ldarg.1
      IL_0013:  callvirt   instance string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Models.SheetNameInvalidCharacter::get_Name()
      IL_0018:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<string>::Add(!0)
      IL_001d:  ret
    } // end of method '<>c'::'<CheckAllSheetName>b__48_0'

    .method assembly hidebysig instance int32 
            '<GetSheetIndexes>b__51_0'(class Microsoft.Office.Interop.Excel.Worksheet ws) cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance int32 Microsoft.Office.Interop.Excel._Worksheet::get_Index()
      IL_0006:  ret
    } // end of method '<>c'::'<GetSheetIndexes>b__51_0'

    .method assembly hidebysig instance string 
            '<GetSheetNames>b__52_0'(class Microsoft.Office.Interop.Excel.Worksheet ws) cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
      IL_0006:  ret
    } // end of method '<>c'::'<GetSheetNames>b__52_0'

  } // end of class '<>c'

  .class auto ansi sealed nested private beforefieldinit '<AutoJudgeAsync>d__25'
         extends [mscorlib]System.ValueType
         implements [mscorlib]System.Runtime.CompilerServices.IAsyncStateMachine
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public int32 '<>1__state'
    .field public valuetype [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder '<>t__builder'
    .field public class [System.Windows.Forms]System.Windows.Forms.Form wait
    .field public class iReporterExcelAddIn2019.ThisAddIn '<>4__this'
    .field public valuetype [mscorlib]System.Threading.CancellationToken token
    .field public bool openDlg
    .method private hidebysig newslot virtual final 
            instance void  MoveNext() cil managed
    {
      .override [mscorlib]System.Runtime.CompilerServices.IAsyncStateMachine::MoveNext
      // Code size       1280 (0x500)
      .maxstack  5
      .locals init (int32 V_0,
               class iReporterExcelAddIn2019.ThisAddIn V_1,
               class [System.Windows.Forms]System.Windows.Forms.Form V_2,
               class iReporterExcelAddIn2019.ThisAddIn/Judged V_3,
               class Microsoft.Office.Interop.Excel.Workbook V_4,
               class [Cimtops.Excel]Cimtops.Excel.Book V_5,
               class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.UserConfig V_6,
               class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>[] V_7,
               class [mscorlib]System.Tuple`3<string,string,string>[] V_8,
               class [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder/AutoJudgeResult V_9,
               class [mscorlib]System.Collections.Generic.IEnumerator`1<class Microsoft.Office.Interop.Excel.Window> V_10,
               class Microsoft.Office.Interop.Excel.Sheets V_11,
               class Microsoft.Office.Interop.Excel.Worksheet V_12,
               class [mscorlib]System.Collections.IEnumerator V_13,
               class Microsoft.Office.Interop.Excel.Worksheet V_14,
               class [mscorlib]System.Collections.Generic.Dictionary`2<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify> V_15,
               valuetype [mscorlib]System.Collections.Generic.Dictionary`2/Enumerator<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify> V_16,
               valuetype [mscorlib]System.Collections.Generic.KeyValuePair`2<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify> V_17,
               class Microsoft.Office.Interop.Excel.Range V_18,
               valuetype [System.Drawing]System.Drawing.Point V_19,
               valuetype [System.Drawing]System.Drawing.Color V_20,
               class [mscorlib]System.IDisposable V_21,
               class [mscorlib]System.Exception V_22,
               valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<string> V_23,
               class [mscorlib]System.Exception V_24)
      IL_0000:  ldarg.0
      IL_0001:  ldfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::'<>1__state'
      IL_0006:  stloc.0
      IL_0007:  ldarg.0
      IL_0008:  ldfld      class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::'<>4__this'
      IL_000d:  stloc.1
      .try
      {
        .try
        {
          IL_000e:  ldarg.0
          IL_000f:  ldfld      class [System.Windows.Forms]System.Windows.Forms.Form iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::wait
          IL_0014:  stloc.2
          .try
          {
            IL_0015:  ldarg.0
            IL_0016:  ldfld      class [System.Windows.Forms]System.Windows.Forms.Form iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::wait
            IL_001b:  ldloc.1
            IL_001c:  call       instance class [System.Windows.Forms]System.Windows.Forms.IWin32Window iReporterExcelAddIn2019.ThisAddIn::CreateWinrowHandle()
            IL_0021:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.Form::Show(class [System.Windows.Forms]System.Windows.Forms.IWin32Window)
            IL_0026:  newobj     instance void iReporterExcelAddIn2019.ThisAddIn/Judged::.ctor()
            IL_002b:  stloc.3
            IL_002c:  call       class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.Globals::get_ThisAddIn()
            IL_0031:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
            IL_0036:  callvirt   instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
            IL_003b:  stloc.s    V_4
            IL_003d:  ldloc.s    V_4
            IL_003f:  ldloc.3
            IL_0040:  ldflda     class [mscorlib]System.Exception iReporterExcelAddIn2019.ThisAddIn/Judged::ex
            IL_0045:  call       class [Cimtops.Excel]Cimtops.Excel.Book [Cimtops.Excel]Cimtops.Excel.BookFactory::Create(class Microsoft.Office.Interop.Excel.Workbook,
                                                                                                                          class [mscorlib]System.Exception&)
            IL_004a:  stloc.s    V_5
            IL_004c:  ldloc.s    V_5
            IL_004e:  brfalse    IL_0413

            IL_0053:  ldc.i4.0
            IL_0054:  call       class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.UserConfig [iReporterExcelAddInCommon]iReporterExcelAddInCommon.UserConfig::GetInstance(valuetype [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ConfigKind)
            IL_0059:  stloc.s    V_6
            IL_005b:  ldloc.s    V_6
            IL_005d:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.PairList [iReporterExcelAddInCommon]iReporterExcelAddInCommon.UserConfig::ClusterColors
            IL_0062:  ldsfld     class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__25_0'
            IL_0067:  dup
            IL_0068:  brtrue.s   IL_0081

            IL_006a:  pop
            IL_006b:  ldsfld     class iReporterExcelAddIn2019.ThisAddIn/'<>c' iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9'
            IL_0070:  ldftn      instance class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<AutoJudgeAsync>b__25_0'(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple)
            IL_0076:  newobj     instance void class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>>::.ctor(object,
                                                                                                                                                                                                                                                                                       native int)
            IL_007b:  dup
            IL_007c:  stsfld     class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__25_0'
            IL_0081:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!1> [System.Core]System.Linq.Enumerable::Select<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                                                                                                                                                                                                                              class [mscorlib]System.Func`2<!!0,!!1>)
            IL_0086:  call       !!0[] [System.Core]System.Linq.Enumerable::ToArray<class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
            IL_008b:  ldloc.s    V_6
            IL_008d:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.PairList [iReporterExcelAddInCommon]iReporterExcelAddInCommon.UserConfig::NotClusterColors
            IL_0092:  ldsfld     class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__25_1'
            IL_0097:  dup
            IL_0098:  brtrue.s   IL_00b1

            IL_009a:  pop
            IL_009b:  ldsfld     class iReporterExcelAddIn2019.ThisAddIn/'<>c' iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9'
            IL_00a0:  ldftn      instance class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<AutoJudgeAsync>b__25_1'(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple)
            IL_00a6:  newobj     instance void class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>>::.ctor(object,
                                                                                                                                                                                                                                                                                       native int)
            IL_00ab:  dup
            IL_00ac:  stsfld     class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__25_1'
            IL_00b1:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!1> [System.Core]System.Linq.Enumerable::Select<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                                                                                                                                                                                                                              class [mscorlib]System.Func`2<!!0,!!1>)
            IL_00b6:  call       !!0[] [System.Core]System.Linq.Enumerable::ToArray<class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
            IL_00bb:  stloc.s    V_7
            IL_00bd:  ldloc.s    V_6
            IL_00bf:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.PairList [iReporterExcelAddInCommon]iReporterExcelAddInCommon.UserConfig::CapAndKinds
            IL_00c4:  ldsfld     class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`3<string,string,string>> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__25_2'
            IL_00c9:  dup
            IL_00ca:  brtrue.s   IL_00e3

            IL_00cc:  pop
            IL_00cd:  ldsfld     class iReporterExcelAddIn2019.ThisAddIn/'<>c' iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9'
            IL_00d2:  ldftn      instance class [mscorlib]System.Tuple`3<string,string,string> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<AutoJudgeAsync>b__25_2'(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple)
            IL_00d8:  newobj     instance void class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`3<string,string,string>>::.ctor(object,
                                                                                                                                                                                                              native int)
            IL_00dd:  dup
            IL_00de:  stsfld     class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`3<string,string,string>> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__25_2'
            IL_00e3:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!1> [System.Core]System.Linq.Enumerable::Select<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.CimTuple,class [mscorlib]System.Tuple`3<string,string,string>>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                                                                                                                                                     class [mscorlib]System.Func`2<!!0,!!1>)
            IL_00e8:  call       !!0[] [System.Core]System.Linq.Enumerable::ToArray<class [mscorlib]System.Tuple`3<string,string,string>>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
            IL_00ed:  stloc.s    V_8
            IL_00ef:  ldloc.s    V_7
            IL_00f1:  ldloc.s    V_8
            IL_00f3:  call       void [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder::CreateColorDic(class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>[],
                                                                                                   class [mscorlib]System.Tuple`2<valuetype [System.Drawing]System.Drawing.Color,valuetype [System.Drawing]System.Drawing.Color>[],
                                                                                                   class [mscorlib]System.Tuple`3<string,string,string>[])
            IL_00f8:  ldloc.s    V_5
            IL_00fa:  ldarg.0
            IL_00fb:  ldfld      valuetype [mscorlib]System.Threading.CancellationToken iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::token
            IL_0100:  ldloc.s    V_6
            IL_0102:  callvirt   instance bool [iReporterExcelAddInCommon]iReporterExcelAddInCommon.UserConfig::get_UseAI()
            IL_0107:  brfalse.s  IL_0117

            IL_0109:  ldloc.s    V_6
            IL_010b:  callvirt   instance class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Language [iReporterExcelAddInCommon]iReporterExcelAddInCommon.UserConfig::get_Language()
            IL_0110:  callvirt   instance bool [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Language::get_CanUseAI()
            IL_0115:  br.s       IL_0118

            IL_0117:  ldc.i4.0
            IL_0118:  ldloc.s    V_6
            IL_011a:  ldfld      valuetype [Cimtops.Excel]Cimtops.Excel.DeviceType [iReporterExcelAddInCommon]iReporterExcelAddInCommon.UserConfig::SelectedDeviceType
            IL_011f:  call       class [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder/AutoJudgeResult [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder::DoAutomaticJudgement(class [Cimtops.Excel]Cimtops.Excel.Book,
                                                                                                                                                                       valuetype [mscorlib]System.Threading.CancellationToken,
                                                                                                                                                                       bool,
                                                                                                                                                                       valuetype [Cimtops.Excel]Cimtops.Excel.DeviceType)
            IL_0124:  stloc.s    V_9
            IL_0126:  ldarg.0
            IL_0127:  ldflda     valuetype [mscorlib]System.Threading.CancellationToken iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::token
            IL_012c:  call       instance bool [mscorlib]System.Threading.CancellationToken::get_IsCancellationRequested()
            IL_0131:  brfalse.s  IL_0138

            IL_0133:  leave      IL_04ec

            IL_0138:  ldloc.s    V_9
            IL_013a:  callvirt   instance string [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder/AutoJudgeResult::get_Empties()
            IL_013f:  brfalse.s  IL_0167

            IL_0141:  ldloc.3
            IL_0142:  ldfld      class [mscorlib]System.Collections.Generic.List`1<string> iReporterExcelAddIn2019.ThisAddIn/Judged::msgs
            IL_0147:  call       string iReporterExcelAddIn2019.Properties.Resources::get_EMPTY_SHEET()
            IL_014c:  call       string [mscorlib]System.Environment::get_NewLine()
            IL_0151:  call       string iReporterExcelAddIn2019.Properties.Resources::get_SHEET_NAME()
            IL_0156:  ldloc.s    V_9
            IL_0158:  callvirt   instance string [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder/AutoJudgeResult::get_Empties()
            IL_015d:  call       string [mscorlib]System.String::Concat(string,
                                                                        string,
                                                                        string,
                                                                        string)
            IL_0162:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<string>::Add(!0)
            IL_0167:  ldloc.s    V_9
            IL_0169:  callvirt   instance string [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder/AutoJudgeResult::get_Multis()
            IL_016e:  brfalse.s  IL_0196

            IL_0170:  ldloc.3
            IL_0171:  ldfld      class [mscorlib]System.Collections.Generic.List`1<string> iReporterExcelAddIn2019.ThisAddIn/Judged::msgs
            IL_0176:  call       string iReporterExcelAddIn2019.Properties.Resources::get_MULTI_PRINT()
            IL_017b:  call       string [mscorlib]System.Environment::get_NewLine()
            IL_0180:  call       string iReporterExcelAddIn2019.Properties.Resources::get_SHEET_NAME()
            IL_0185:  ldloc.s    V_9
            IL_0187:  callvirt   instance string [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder/AutoJudgeResult::get_Multis()
            IL_018c:  call       string [mscorlib]System.String::Concat(string,
                                                                        string,
                                                                        string,
                                                                        string)
            IL_0191:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<string>::Add(!0)
            IL_0196:  ldloc.s    V_9
            IL_0198:  callvirt   instance bool [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder/AutoJudgeResult::get_IsAPIError()
            IL_019d:  brfalse.s  IL_01af

            IL_019f:  ldloc.3
            IL_01a0:  ldfld      class [mscorlib]System.Collections.Generic.List`1<string> iReporterExcelAddIn2019.ThisAddIn/Judged::msgs
            IL_01a5:  call       string iReporterExcelAddIn2019.Properties.Resources::get_API_ERROR()
            IL_01aa:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<string>::Add(!0)
            IL_01af:  nop
            .try
            {
              IL_01b0:  ldloc.1
              IL_01b1:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
              IL_01b6:  callvirt   instance class Microsoft.Office.Interop.Excel.Windows Microsoft.Office.Interop.Excel._Application::get_Windows()
              IL_01bb:  dup
              IL_01bc:  brtrue.s   IL_01c2

              IL_01be:  pop
              IL_01bf:  ldnull
              IL_01c0:  br.s       IL_01c7

              IL_01c2:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0> [System.Core]System.Linq.Enumerable::OfType<class Microsoft.Office.Interop.Excel.Window>(class [mscorlib]System.Collections.IEnumerable)
              IL_01c7:  dup
              IL_01c8:  brtrue.s   IL_01d0

              IL_01ca:  pop
              IL_01cb:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0> [System.Core]System.Linq.Enumerable::Empty<class Microsoft.Office.Interop.Excel.Window>()
              IL_01d0:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerator`1<!0> class [mscorlib]System.Collections.Generic.IEnumerable`1<class Microsoft.Office.Interop.Excel.Window>::GetEnumerator()
              IL_01d5:  stloc.s    V_10
              .try
              {
                IL_01d7:  br.s       IL_0218

                IL_01d9:  ldloc.s    V_10
                IL_01db:  callvirt   instance !0 class [mscorlib]System.Collections.Generic.IEnumerator`1<class Microsoft.Office.Interop.Excel.Window>::get_Current()
                IL_01e0:  callvirt   instance class Microsoft.Office.Interop.Excel.Sheets Microsoft.Office.Interop.Excel.Window::get_SelectedSheets()
                IL_01e5:  stloc.s    V_11
                IL_01e7:  ldloc.s    V_11
                IL_01e9:  brfalse.s  IL_0218

                IL_01eb:  ldloc.s    V_11
                IL_01ed:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Sheets::get_Count()
                IL_01f2:  ldc.i4.1
                IL_01f3:  ble.s      IL_0218

                IL_01f5:  ldloc.s    V_11
                IL_01f7:  ldc.i4.1
                IL_01f8:  box        [mscorlib]System.Int32
                IL_01fd:  callvirt   instance object Microsoft.Office.Interop.Excel.Sheets::get__Default(object)
                IL_0202:  isinst     Microsoft.Office.Interop.Excel.Worksheet
                IL_0207:  dup
                IL_0208:  stloc.s    V_12
                IL_020a:  brfalse.s  IL_0218

                IL_020c:  ldloc.s    V_12
                IL_020e:  ldsfld     object [mscorlib]System.Type::Missing
                IL_0213:  callvirt   instance void Microsoft.Office.Interop.Excel._Worksheet::Select(object)
                IL_0218:  ldloc.s    V_10
                IL_021a:  callvirt   instance bool [mscorlib]System.Collections.IEnumerator::MoveNext()
                IL_021f:  brtrue.s   IL_01d9

                IL_0221:  leave.s    IL_0233

              }  // end .try
              finally
              {
                IL_0223:  ldloc.0
                IL_0224:  ldc.i4.0
                IL_0225:  bge.s      IL_0232

                IL_0227:  ldloc.s    V_10
                IL_0229:  brfalse.s  IL_0232

                IL_022b:  ldloc.s    V_10
                IL_022d:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
                IL_0232:  endfinally
              }  // end handler
              IL_0233:  ldloc.s    V_4
              IL_0235:  callvirt   instance class Microsoft.Office.Interop.Excel.Sheets Microsoft.Office.Interop.Excel._Workbook::get_Sheets()
              IL_023a:  callvirt   instance class [mscorlib]System.Collections.IEnumerator Microsoft.Office.Interop.Excel.Sheets::GetEnumerator()
              IL_023f:  stloc.s    V_13
              .try
              {
                IL_0241:  br         IL_03de

                IL_0246:  ldloc.s    V_13
                IL_0248:  callvirt   instance object [mscorlib]System.Collections.IEnumerator::get_Current()
                IL_024d:  isinst     Microsoft.Office.Interop.Excel.Worksheet
                IL_0252:  dup
                IL_0253:  stloc.s    V_14
                IL_0255:  brfalse    IL_03de

                IL_025a:  ldloc.s    V_9
                IL_025c:  callvirt   instance class [mscorlib]System.Collections.Generic.Dictionary`2<string,class [mscorlib]System.Collections.Generic.Dictionary`2<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify>> [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder/AutoJudgeResult::get_Modify()
                IL_0261:  ldloc.s    V_14
                IL_0263:  callvirt   instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
                IL_0268:  ldloca.s   V_15
                IL_026a:  callvirt   instance bool class [mscorlib]System.Collections.Generic.Dictionary`2<string,class [mscorlib]System.Collections.Generic.Dictionary`2<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify>>::TryGetValue(!0,
                                                                                                                                                                                                                                                                                      !1&)
                IL_026f:  brfalse    IL_03de

                IL_0274:  ldloc.s    V_15
                IL_0276:  callvirt   instance valuetype [mscorlib]System.Collections.Generic.Dictionary`2/Enumerator<!0,!1> class [mscorlib]System.Collections.Generic.Dictionary`2<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify>::GetEnumerator()
                IL_027b:  stloc.s    V_16
                .try
                {
                  IL_027d:  br         IL_03be

                  IL_0282:  ldloca.s   V_16
                  IL_0284:  call       instance valuetype [mscorlib]System.Collections.Generic.KeyValuePair`2<!0,!1> valuetype [mscorlib]System.Collections.Generic.Dictionary`2/Enumerator<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify>::get_Current()
                  IL_0289:  stloc.s    V_17
                  IL_028b:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> iReporterExcelAddIn2019.ThisAddIn/'<>o__25'::'<>p__0'
                  IL_0290:  brtrue.s   IL_02b6

                  IL_0292:  ldc.i4.0
                  IL_0293:  ldtoken    Microsoft.Office.Interop.Excel.Range
                  IL_0298:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
                  IL_029d:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
                  IL_02a2:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
                  IL_02a7:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::Convert(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                          class [mscorlib]System.Type,
                                                                                                                                                                          class [mscorlib]System.Type)
                  IL_02ac:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
                  IL_02b1:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> iReporterExcelAddIn2019.ThisAddIn/'<>o__25'::'<>p__0'
                  IL_02b6:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> iReporterExcelAddIn2019.ThisAddIn/'<>o__25'::'<>p__0'
                  IL_02bb:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>>::Target
                  IL_02c0:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> iReporterExcelAddIn2019.ThisAddIn/'<>o__25'::'<>p__0'
                  IL_02c5:  ldloc.s    V_14
                  IL_02c7:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Cells()
                  IL_02cc:  ldloca.s   V_17
                  IL_02ce:  call       instance !0 valuetype [mscorlib]System.Collections.Generic.KeyValuePair`2<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify>::get_Key()
                  IL_02d3:  stloc.s    V_19
                  IL_02d5:  ldloca.s   V_19
                  IL_02d7:  call       instance int32 [System.Drawing]System.Drawing.Point::get_Y()
                  IL_02dc:  box        [mscorlib]System.Int32
                  IL_02e1:  ldloca.s   V_17
                  IL_02e3:  call       instance !0 valuetype [mscorlib]System.Collections.Generic.KeyValuePair`2<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify>::get_Key()
                  IL_02e8:  stloc.s    V_19
                  IL_02ea:  ldloca.s   V_19
                  IL_02ec:  call       instance int32 [System.Drawing]System.Drawing.Point::get_X()
                  IL_02f1:  box        [mscorlib]System.Int32
                  IL_02f6:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get__Default(object,
                                                                                                          object)
                  IL_02fb:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>::Invoke(!0,
                                                                                                                                                                                                        !1)
                  IL_0300:  stloc.s    V_18
                  IL_0302:  ldarg.0
                  IL_0303:  ldflda     valuetype [mscorlib]System.Threading.CancellationToken iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::token
                  IL_0308:  call       instance bool [mscorlib]System.Threading.CancellationToken::get_IsCancellationRequested()
                  IL_030d:  brfalse.s  IL_0314

                  IL_030f:  leave      IL_04ec

                  IL_0314:  ldloca.s   V_17
                  IL_0316:  call       instance !1 valuetype [mscorlib]System.Collections.Generic.KeyValuePair`2<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify>::get_Value()
                  IL_031b:  ldfld      string [Cimtops.Excel]Cimtops.Excel.CellModify::Comment
                  IL_0320:  brfalse.s  IL_0359

                  IL_0322:  ldloc.s    V_18
                  IL_0324:  callvirt   instance void Microsoft.Office.Interop.Excel.Range::ClearComments()
                  IL_0329:  ldloc.s    V_18
                  IL_032b:  ldsfld     object [mscorlib]System.Type::Missing
                  IL_0330:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::AddComment(object)
                  IL_0335:  pop
                  IL_0336:  ldloc.s    V_18
                  IL_0338:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::get_Comment()
                  IL_033d:  ldloca.s   V_17
                  IL_033f:  call       instance !1 valuetype [mscorlib]System.Collections.Generic.KeyValuePair`2<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify>::get_Value()
                  IL_0344:  ldfld      string [Cimtops.Excel]Cimtops.Excel.CellModify::Comment
                  IL_0349:  ldsfld     object [mscorlib]System.Type::Missing
                  IL_034e:  ldsfld     object [mscorlib]System.Type::Missing
                  IL_0353:  callvirt   instance string Microsoft.Office.Interop.Excel.Comment::Text(object,
                                                                                                    object,
                                                                                                    object)
                  IL_0358:  pop
                  IL_0359:  ldarg.0
                  IL_035a:  ldflda     valuetype [mscorlib]System.Threading.CancellationToken iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::token
                  IL_035f:  call       instance bool [mscorlib]System.Threading.CancellationToken::get_IsCancellationRequested()
                  IL_0364:  brfalse.s  IL_036b

                  IL_0366:  leave      IL_04ec

                  IL_036b:  ldloca.s   V_17
                  IL_036d:  call       instance !1 valuetype [mscorlib]System.Collections.Generic.KeyValuePair`2<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify>::get_Value()
                  IL_0372:  ldflda     valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Color> [Cimtops.Excel]Cimtops.Excel.CellModify::Color
                  IL_0377:  call       instance bool valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Color>::get_HasValue()
                  IL_037c:  brfalse.s  IL_03be

                  IL_037e:  ldloca.s   V_17
                  IL_0380:  call       instance !1 valuetype [mscorlib]System.Collections.Generic.KeyValuePair`2<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify>::get_Value()
                  IL_0385:  ldflda     valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Color> [Cimtops.Excel]Cimtops.Excel.CellModify::Color
                  IL_038a:  call       instance !0 valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Color>::get_Value()
                  IL_038f:  stloc.s    V_20
                  IL_0391:  ldloc.s    V_18
                  IL_0393:  callvirt   instance class Microsoft.Office.Interop.Excel.Interior Microsoft.Office.Interop.Excel.Range::get_Interior()
                  IL_0398:  ldloca.s   V_20
                  IL_039a:  call       instance uint8 [System.Drawing]System.Drawing.Color::get_R()
                  IL_039f:  ldloca.s   V_20
                  IL_03a1:  call       instance uint8 [System.Drawing]System.Drawing.Color::get_G()
                  IL_03a6:  ldc.i4.8
                  IL_03a7:  shl
                  IL_03a8:  or
                  IL_03a9:  ldloca.s   V_20
                  IL_03ab:  call       instance uint8 [System.Drawing]System.Drawing.Color::get_B()
                  IL_03b0:  ldc.i4.s   16
                  IL_03b2:  shl
                  IL_03b3:  or
                  IL_03b4:  box        [mscorlib]System.Int32
                  IL_03b9:  callvirt   instance void Microsoft.Office.Interop.Excel.Interior::set_Color(object)
                  IL_03be:  ldloca.s   V_16
                  IL_03c0:  call       instance bool valuetype [mscorlib]System.Collections.Generic.Dictionary`2/Enumerator<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify>::MoveNext()
                  IL_03c5:  brtrue     IL_0282

                  IL_03ca:  leave.s    IL_03de

                }  // end .try
                finally
                {
                  IL_03cc:  ldloc.0
                  IL_03cd:  ldc.i4.0
                  IL_03ce:  bge.s      IL_03dd

                  IL_03d0:  ldloca.s   V_16
                  IL_03d2:  constrained. valuetype [mscorlib]System.Collections.Generic.Dictionary`2/Enumerator<valuetype [System.Drawing]System.Drawing.Point,class [Cimtops.Excel]Cimtops.Excel.CellModify>
                  IL_03d8:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
                  IL_03dd:  endfinally
                }  // end handler
                IL_03de:  ldloc.s    V_13
                IL_03e0:  callvirt   instance bool [mscorlib]System.Collections.IEnumerator::MoveNext()
                IL_03e5:  brtrue     IL_0246

                IL_03ea:  leave.s    IL_0405

              }  // end .try
              finally
              {
                IL_03ec:  ldloc.0
                IL_03ed:  ldc.i4.0
                IL_03ee:  bge.s      IL_0404

                IL_03f0:  ldloc.s    V_13
                IL_03f2:  isinst     [mscorlib]System.IDisposable
                IL_03f7:  stloc.s    V_21
                IL_03f9:  ldloc.s    V_21
                IL_03fb:  brfalse.s  IL_0404

                IL_03fd:  ldloc.s    V_21
                IL_03ff:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
                IL_0404:  endfinally
              }  // end handler
              IL_0405:  leave.s    IL_0413

            }  // end .try
            catch [mscorlib]System.Exception 
            {
              IL_0407:  stloc.s    V_22
              IL_0409:  ldloc.3
              IL_040a:  ldloc.s    V_22
              IL_040c:  stfld      class [mscorlib]System.Exception iReporterExcelAddIn2019.ThisAddIn/Judged::ex
              IL_0411:  leave.s    IL_0413

            }  // end handler
            IL_0413:  ldarg.0
            IL_0414:  ldflda     valuetype [mscorlib]System.Threading.CancellationToken iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::token
            IL_0419:  call       instance bool [mscorlib]System.Threading.CancellationToken::get_IsCancellationRequested()
            IL_041e:  brfalse.s  IL_0425

            IL_0420:  leave      IL_04ec

            IL_0425:  ldarg.0
            IL_0426:  ldfld      class [System.Windows.Forms]System.Windows.Forms.Form iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::wait
            IL_042b:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.Form::Close()
            IL_0430:  ldloc.1
            IL_0431:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ResetInput()
            IL_0436:  ldarg.0
            IL_0437:  ldflda     valuetype [mscorlib]System.Threading.CancellationToken iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::token
            IL_043c:  call       instance bool [mscorlib]System.Threading.CancellationToken::get_IsCancellationRequested()
            IL_0441:  brfalse.s  IL_0448

            IL_0443:  leave      IL_04ec

            IL_0448:  ldloc.3
            IL_0449:  ldfld      class [mscorlib]System.Exception iReporterExcelAddIn2019.ThisAddIn/Judged::ex
            IL_044e:  brfalse.s  IL_045d

            IL_0450:  call       string iReporterExcelAddIn2019.Properties.Resources::get_EXCEL_ERROR()
            IL_0455:  call       valuetype [System.Windows.Forms]System.Windows.Forms.DialogResult [System.Windows.Forms]System.Windows.Forms.MessageBox::Show(string)
            IL_045a:  pop
            IL_045b:  br.s       IL_04b3

            IL_045d:  ldloc.3
            IL_045e:  ldfld      class [mscorlib]System.Collections.Generic.List`1<string> iReporterExcelAddIn2019.ThisAddIn/Judged::msgs
            IL_0463:  callvirt   instance valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<!0> class [mscorlib]System.Collections.Generic.List`1<string>::GetEnumerator()
            IL_0468:  stloc.s    V_23
            .try
            {
              IL_046a:  br.s       IL_0479

              IL_046c:  ldloca.s   V_23
              IL_046e:  call       instance !0 valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<string>::get_Current()
              IL_0473:  call       valuetype [System.Windows.Forms]System.Windows.Forms.DialogResult [System.Windows.Forms]System.Windows.Forms.MessageBox::Show(string)
              IL_0478:  pop
              IL_0479:  ldloca.s   V_23
              IL_047b:  call       instance bool valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<string>::MoveNext()
              IL_0480:  brtrue.s   IL_046c

              IL_0482:  leave.s    IL_0496

            }  // end .try
            finally
            {
              IL_0484:  ldloc.0
              IL_0485:  ldc.i4.0
              IL_0486:  bge.s      IL_0495

              IL_0488:  ldloca.s   V_23
              IL_048a:  constrained. valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<string>
              IL_0490:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
              IL_0495:  endfinally
            }  // end handler
            IL_0496:  ldarg.0
            IL_0497:  ldflda     valuetype [mscorlib]System.Threading.CancellationToken iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::token
            IL_049c:  call       instance bool [mscorlib]System.Threading.CancellationToken::get_IsCancellationRequested()
            IL_04a1:  brfalse.s  IL_04a5

            IL_04a3:  leave.s    IL_04ec

            IL_04a5:  ldarg.0
            IL_04a6:  ldfld      bool iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::openDlg
            IL_04ab:  brfalse.s  IL_04b3

            IL_04ad:  ldloc.1
            IL_04ae:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ManualCluster()
            IL_04b3:  leave.s    IL_04c3

          }  // end .try
          finally
          {
            IL_04b5:  ldloc.0
            IL_04b6:  ldc.i4.0
            IL_04b7:  bge.s      IL_04c2

            IL_04b9:  ldloc.2
            IL_04ba:  brfalse.s  IL_04c2

            IL_04bc:  ldloc.2
            IL_04bd:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
            IL_04c2:  endfinally
          }  // end handler
          IL_04c3:  leave.s    IL_04d1

        }  // end .try
        finally
        {
          IL_04c5:  ldloc.0
          IL_04c6:  ldc.i4.0
          IL_04c7:  bge.s      IL_04d0

          IL_04c9:  ldloc.1
          IL_04ca:  ldc.i4.0
          IL_04cb:  call       instance void iReporterExcelAddIn2019.ThisAddIn::set_IsJudging(bool)
          IL_04d0:  endfinally
        }  // end handler
        IL_04d1:  leave.s    IL_04ec

      }  // end .try
      catch [mscorlib]System.Exception 
      {
        IL_04d3:  stloc.s    V_24
        IL_04d5:  ldarg.0
        IL_04d6:  ldc.i4.s   -2
        IL_04d8:  stfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::'<>1__state'
        IL_04dd:  ldarg.0
        IL_04de:  ldflda     valuetype [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::'<>t__builder'
        IL_04e3:  ldloc.s    V_24
        IL_04e5:  call       instance void [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder::SetException(class [mscorlib]System.Exception)
        IL_04ea:  leave.s    IL_04ff

      }  // end handler
      IL_04ec:  ldarg.0
      IL_04ed:  ldc.i4.s   -2
      IL_04ef:  stfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::'<>1__state'
      IL_04f4:  ldarg.0
      IL_04f5:  ldflda     valuetype [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::'<>t__builder'
      IL_04fa:  call       instance void [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder::SetResult()
      IL_04ff:  ret
    } // end of method '<AutoJudgeAsync>d__25'::MoveNext

    .method private hidebysig newslot virtual final 
            instance void  SetStateMachine(class [mscorlib]System.Runtime.CompilerServices.IAsyncStateMachine stateMachine) cil managed
    {
      .custom instance void [mscorlib]System.Diagnostics.DebuggerHiddenAttribute::.ctor() = ( 01 00 00 00 ) 
      .override [mscorlib]System.Runtime.CompilerServices.IAsyncStateMachine::SetStateMachine
      // Code size       13 (0xd)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  ldflda     valuetype [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::'<>t__builder'
      IL_0006:  ldarg.1
      IL_0007:  call       instance void [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder::SetStateMachine(class [mscorlib]System.Runtime.CompilerServices.IAsyncStateMachine)
      IL_000c:  ret
    } // end of method '<AutoJudgeAsync>d__25'::SetStateMachine

  } // end of class '<AutoJudgeAsync>d__25'

  .class auto ansi sealed nested private beforefieldinit '<>c__DisplayClass27_0'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public class iReporterExcelAddIn2019.ThisAddIn '<>4__this'
    .field public class [System.Windows.Forms]System.Windows.Forms.IWin32Window window
    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
      IL_0006:  ret
    } // end of method '<>c__DisplayClass27_0'::.ctor

    .method assembly hidebysig instance void 
            '<ShowForm>b__0'(object _) cil managed
    {
      // Code size       12 (0xc)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  ldfld      class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass27_0'::'<>4__this'
      IL_0006:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ResetInput()
      IL_000b:  ret
    } // end of method '<>c__DisplayClass27_0'::'<ShowForm>b__0'

    .method assembly hidebysig instance void 
            '<ShowForm>b__1'(object o,
                             class [System.Windows.Forms]System.Windows.Forms.FormClosedEventArgs e) cil managed
    {
      // Code size       18 (0x12)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  ldfld      class [System.Windows.Forms]System.Windows.Forms.IWin32Window iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass27_0'::window
      IL_0006:  callvirt   instance native int [System.Windows.Forms]System.Windows.Forms.IWin32Window::get_Handle()
      IL_000b:  call       bool iReporterExcelAddIn2019.ThisAddIn::SetForegroundWindow(native int)
      IL_0010:  pop
      IL_0011:  ret
    } // end of method '<>c__DisplayClass27_0'::'<ShowForm>b__1'

  } // end of class '<>c__DisplayClass27_0'

  .class abstract auto ansi sealed nested private beforefieldinit '<>o__34'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class iReporterExcelAddIn2019.ThisAddIn,object,object>> '<>p__0'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype [System.Drawing]System.Drawing.Color>> '<>p__1'
  } // end of class '<>o__34'

  .class abstract auto ansi sealed nested private beforefieldinit '<>o__36'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> '<>p__0'
  } // end of class '<>o__36'

  .class auto ansi sealed nested private beforefieldinit '<>c__DisplayClass36_0'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public int32 row
    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
      IL_0006:  ret
    } // end of method '<>c__DisplayClass36_0'::.ctor

  } // end of class '<>c__DisplayClass36_0'

  .class auto ansi sealed nested private beforefieldinit '<>c__DisplayClass36_1'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public int32 col
    .field public class iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_0' 'CS$<>8__locals1'
    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
      IL_0006:  ret
    } // end of method '<>c__DisplayClass36_1'::.ctor

    .method assembly hidebysig instance bool 
            '<WriteComment>b__2'(valuetype [System.Drawing]System.Drawing.Rectangle r) cil managed
    {
      // Code size       25 (0x19)
      .maxstack  8
      IL_0000:  ldarga.s   r
      IL_0002:  ldarg.0
      IL_0003:  ldfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1'::col
      IL_0008:  ldarg.0
      IL_0009:  ldfld      class iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_0' iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1'::'CS$<>8__locals1'
      IL_000e:  ldfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_0'::row
      IL_0013:  call       instance bool [System.Drawing]System.Drawing.Rectangle::Contains(int32,
                                                                                            int32)
      IL_0018:  ret
    } // end of method '<>c__DisplayClass36_1'::'<WriteComment>b__2'

  } // end of class '<>c__DisplayClass36_1'

  .class auto ansi sealed nested private beforefieldinit '<>c__DisplayClass37_0'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table table
    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
      IL_0006:  ret
    } // end of method '<>c__DisplayClass37_0'::.ctor

    .method assembly hidebysig instance bool 
            '<WriteTable>b__0'(class Microsoft.Office.Interop.Excel.Worksheet s) cil managed
    {
      // Code size       23 (0x17)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
      IL_0006:  ldarg.0
      IL_0007:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass37_0'::table
      IL_000c:  callvirt   instance string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::get_SheetName()
      IL_0011:  call       bool [mscorlib]System.String::op_Equality(string,
                                                                     string)
      IL_0016:  ret
    } // end of method '<>c__DisplayClass37_0'::'<WriteTable>b__0'

  } // end of class '<>c__DisplayClass37_0'

  .class abstract auto ansi sealed nested private beforefieldinit '<>o__42'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> '<>p__0'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> '<>p__1'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32,int32,object>> '<>p__2'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> '<>p__3'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object,object>> '<>p__4'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> '<>p__5'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>> '<>p__6'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> '<>p__7'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> '<>p__8'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> '<>p__9'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> '<>p__10'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>> '<>p__11'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> '<>p__12'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [System.Core]System.Collections.Generic.HashSet`1<string>,object,object>> '<>p__13'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> '<>p__14'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> '<>p__15'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Action`3<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Collections.Generic.List`1<string>,object>> '<>p__16'
  } // end of class '<>o__42'

  .class auto ansi sealed nested private beforefieldinit '<>c__DisplayClass42_0'
         extends [mscorlib]System.ValueType
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public class [mscorlib]System.Collections.Generic.List`1<string> list
  } // end of class '<>c__DisplayClass42_0'

  .class auto ansi sealed nested private beforefieldinit '<>c__DisplayClass42_1'
         extends [mscorlib]System.ValueType
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public class [System.Core]System.Collections.Generic.HashSet`1<string> set
  } // end of class '<>c__DisplayClass42_1'

  .class auto ansi sealed nested private beforefieldinit '<>c__DisplayClass45_0'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem item
    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
      IL_0006:  ret
    } // end of method '<>c__DisplayClass45_0'::.ctor

    .method assembly hidebysig instance bool 
            '<ShowTableForm>b__1'(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table t) cil managed
    {
      // Code size       23 (0x17)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::get_No()
      IL_0006:  ldarg.0
      IL_0007:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass45_0'::item
      IL_000c:  callvirt   instance string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem::get_TableNo()
      IL_0011:  call       bool [mscorlib]System.String::op_Equality(string,
                                                                     string)
      IL_0016:  ret
    } // end of method '<>c__DisplayClass45_0'::'<ShowTableForm>b__1'

  } // end of class '<>c__DisplayClass45_0'

  .class abstract auto ansi sealed nested private beforefieldinit '<>o__46'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> '<>p__0'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> '<>p__1'
  } // end of class '<>o__46'

  .class abstract auto ansi sealed nested private beforefieldinit '<>o__47'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> '<>p__0'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> '<>p__1'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> '<>p__2'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> '<>p__3'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> '<>p__4'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> '<>p__5'
  } // end of class '<>o__47'

  .field private static literal uint32 GA_ROOT = uint32(0x00000002)
  .field private class Microsoft.Office.Interop.Excel.Application excelApplication
  .field private class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton> MenuTexts
  .field private class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 _input2
  .field private class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm tableForm
  .field private object _inputLocker
  .field private bool _selectionChangedSet
  .field private class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm _color
  .field private object judgeLock
  .field private bool '<IsJudging>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field assembly class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.CustomTaskPaneCollection CustomTaskPanes
  .field assembly class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.SmartTagCollection VstoSmartTags
  .field private object missing
  .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                      string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                  75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                  4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                  6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                  2E 30 2E 30 00 00 )                               // .0.0..
  .field assembly class Microsoft.Office.Interop.Excel.Application Application
  .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                      string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                  75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                  4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                  6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                  2E 30 2E 30 00 00 )                               // .0.0..
  .method private hidebysig static pinvokeimpl("user32.dll" winapi) 
          int32  GetWindowRect(native int hWnd,
                               [out] valuetype RECT& lpRect) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("user32.dll" winapi) 
          bool  marshal( bool)  SetForegroundWindow(native int hWnd) cil managed preservesig
  {
  }
  .method private hidebysig instance class [System.Windows.Forms]System.Windows.Forms.IWin32Window 
          CreateWinrowHandle() cil managed
  {
    // Code size       28 (0x1c)
    .maxstack  8
    IL_0000:  newobj     instance void iReporterExcelAddIn2019.ThisAddIn/Win32Window::.ctor()
    IL_0005:  dup
    IL_0006:  ldarg.0
    IL_0007:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_000c:  callvirt   instance int32 Microsoft.Office.Interop.Excel._Application::get_Hwnd()
    IL_0011:  newobj     instance void [mscorlib]System.IntPtr::.ctor(int32)
    IL_0016:  stfld      native int iReporterExcelAddIn2019.ThisAddIn/Win32Window::_handle
    IL_001b:  ret
  } // end of method ThisAddIn::CreateWinrowHandle

  .method private hidebysig static pinvokeimpl("user32.dll" winapi) 
          native int  GetAncestor(native int hwnd,
                                  uint32 gaFlags) cil managed preservesig
  {
  }
  .method private hidebysig instance void 
          SetUpMenu() cil managed
  {
    // Code size       586 (0x24a)
    .maxstack  7
    .locals init (int32 V_0,
             class [mscorlib]System.Collections.IEnumerator V_1,
             class Microsoft.Office.Core.CommandBar V_2,
             class Microsoft.Office.Core.CommandBarButton V_3,
             int32 V_4,
             class [mscorlib]System.IDisposable V_5)
    IL_0000:  ldc.i4.s   10
    IL_0002:  stloc.0
    IL_0003:  ldarg.0
    IL_0004:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0009:  callvirt   instance class Microsoft.Office.Core.CommandBars Microsoft.Office.Interop.Excel._Application::get_CommandBars()
    IL_000e:  callvirt   instance class [mscorlib]System.Collections.IEnumerator Microsoft.Office.Core._CommandBars::GetEnumerator()
    IL_0013:  stloc.1
    .try
    {
      IL_0014:  br         IL_021b

      IL_0019:  ldloc.1
      IL_001a:  callvirt   instance object [mscorlib]System.Collections.IEnumerator::get_Current()
      IL_001f:  castclass  Microsoft.Office.Core.CommandBar
      IL_0024:  stloc.2
      IL_0025:  ldloc.2
      IL_0026:  callvirt   instance string Microsoft.Office.Core.CommandBar::get_Name()
      IL_002b:  ldstr      "Cell"
      IL_0030:  callvirt   instance bool [mscorlib]System.String::Equals(string)
      IL_0035:  brfalse    IL_021b

      IL_003a:  ldloc.2
      IL_003b:  callvirt   instance class Microsoft.Office.Core.CommandBarControls Microsoft.Office.Core.CommandBar::get_Controls()
      IL_0040:  dup
      IL_0041:  ldc.i4.1
      IL_0042:  box        Microsoft.Office.Core.MsoControlType
      IL_0047:  ldarg.0
      IL_0048:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::missing
      IL_004d:  ldarg.0
      IL_004e:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::missing
      IL_0053:  ldarg.0
      IL_0054:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::missing
      IL_0059:  ldc.i4.1
      IL_005a:  box        [mscorlib]System.Boolean
      IL_005f:  callvirt   instance class Microsoft.Office.Core.CommandBarControl Microsoft.Office.Core.CommandBarControls::Add(object,
                                                                                                                                object,
                                                                                                                                object,
                                                                                                                                object,
                                                                                                                                object)
      IL_0064:  castclass  Microsoft.Office.Core.CommandBarButton
      IL_0069:  stloc.3
      IL_006a:  ldloc.3
      IL_006b:  ldc.i4.2
      IL_006c:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Style(valuetype Microsoft.Office.Core.MsoButtonStyle)
      IL_0071:  ldloc.3
      IL_0072:  ldloc.0
      IL_0073:  dup
      IL_0074:  ldc.i4.1
      IL_0075:  add
      IL_0076:  stloc.0
      IL_0077:  stloc.s    V_4
      IL_0079:  ldloca.s   V_4
      IL_007b:  call       instance string [mscorlib]System.Int32::ToString()
      IL_0080:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Tag(string)
      IL_0085:  ldtoken    Microsoft.Office.Core._CommandBarButtonEvents_Event
      IL_008a:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_008f:  ldstr      "Click"
      IL_0094:  newobj     instance void [System.Core]System.Runtime.InteropServices.ComAwareEventInfo::.ctor(class [mscorlib]System.Type,
                                                                                                              string)
      IL_0099:  ldloc.3
      IL_009a:  ldarg.0
      IL_009b:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::'<SetUpMenu>b__13_0'(class Microsoft.Office.Core.CommandBarButton,
                                                                                                 bool&)
      IL_00a1:  newobj     instance void Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler::.ctor(object,
                                                                                                                native uint)
      IL_00a6:  callvirt   instance void [mscorlib]System.Reflection.EventInfo::AddEventHandler(object,
                                                                                                class [mscorlib]System.Delegate)
      IL_00ab:  ldarg.0
      IL_00ac:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton> iReporterExcelAddIn2019.ThisAddIn::MenuTexts
      IL_00b1:  ldloc.3
      IL_00b2:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton>::Add(!0)
      IL_00b7:  dup
      IL_00b8:  ldc.i4.1
      IL_00b9:  box        Microsoft.Office.Core.MsoControlType
      IL_00be:  ldarg.0
      IL_00bf:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::missing
      IL_00c4:  ldarg.0
      IL_00c5:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::missing
      IL_00ca:  ldarg.0
      IL_00cb:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::missing
      IL_00d0:  ldc.i4.1
      IL_00d1:  box        [mscorlib]System.Boolean
      IL_00d6:  callvirt   instance class Microsoft.Office.Core.CommandBarControl Microsoft.Office.Core.CommandBarControls::Add(object,
                                                                                                                                object,
                                                                                                                                object,
                                                                                                                                object,
                                                                                                                                object)
      IL_00db:  castclass  Microsoft.Office.Core.CommandBarButton
      IL_00e0:  stloc.3
      IL_00e1:  ldloc.3
      IL_00e2:  ldc.i4.2
      IL_00e3:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Style(valuetype Microsoft.Office.Core.MsoButtonStyle)
      IL_00e8:  ldloc.3
      IL_00e9:  ldloc.0
      IL_00ea:  dup
      IL_00eb:  ldc.i4.1
      IL_00ec:  add
      IL_00ed:  stloc.0
      IL_00ee:  stloc.s    V_4
      IL_00f0:  ldloca.s   V_4
      IL_00f2:  call       instance string [mscorlib]System.Int32::ToString()
      IL_00f7:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Tag(string)
      IL_00fc:  ldtoken    Microsoft.Office.Core._CommandBarButtonEvents_Event
      IL_0101:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0106:  ldstr      "Click"
      IL_010b:  newobj     instance void [System.Core]System.Runtime.InteropServices.ComAwareEventInfo::.ctor(class [mscorlib]System.Type,
                                                                                                              string)
      IL_0110:  ldloc.3
      IL_0111:  ldarg.0
      IL_0112:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::'<SetUpMenu>b__13_1'(class Microsoft.Office.Core.CommandBarButton,
                                                                                                 bool&)
      IL_0118:  newobj     instance void Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler::.ctor(object,
                                                                                                                native uint)
      IL_011d:  callvirt   instance void [mscorlib]System.Reflection.EventInfo::AddEventHandler(object,
                                                                                                class [mscorlib]System.Delegate)
      IL_0122:  ldarg.0
      IL_0123:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton> iReporterExcelAddIn2019.ThisAddIn::MenuTexts
      IL_0128:  ldloc.3
      IL_0129:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton>::Add(!0)
      IL_012e:  dup
      IL_012f:  ldc.i4.1
      IL_0130:  box        Microsoft.Office.Core.MsoControlType
      IL_0135:  ldarg.0
      IL_0136:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::missing
      IL_013b:  ldarg.0
      IL_013c:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::missing
      IL_0141:  ldarg.0
      IL_0142:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::missing
      IL_0147:  ldc.i4.1
      IL_0148:  box        [mscorlib]System.Boolean
      IL_014d:  callvirt   instance class Microsoft.Office.Core.CommandBarControl Microsoft.Office.Core.CommandBarControls::Add(object,
                                                                                                                                object,
                                                                                                                                object,
                                                                                                                                object,
                                                                                                                                object)
      IL_0152:  castclass  Microsoft.Office.Core.CommandBarButton
      IL_0157:  stloc.3
      IL_0158:  ldloc.3
      IL_0159:  ldc.i4.2
      IL_015a:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Style(valuetype Microsoft.Office.Core.MsoButtonStyle)
      IL_015f:  ldloc.3
      IL_0160:  ldloc.0
      IL_0161:  dup
      IL_0162:  ldc.i4.1
      IL_0163:  add
      IL_0164:  stloc.0
      IL_0165:  stloc.s    V_4
      IL_0167:  ldloca.s   V_4
      IL_0169:  call       instance string [mscorlib]System.Int32::ToString()
      IL_016e:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Tag(string)
      IL_0173:  ldtoken    Microsoft.Office.Core._CommandBarButtonEvents_Event
      IL_0178:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_017d:  ldstr      "Click"
      IL_0182:  newobj     instance void [System.Core]System.Runtime.InteropServices.ComAwareEventInfo::.ctor(class [mscorlib]System.Type,
                                                                                                              string)
      IL_0187:  ldloc.3
      IL_0188:  ldarg.0
      IL_0189:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::'<SetUpMenu>b__13_2'(class Microsoft.Office.Core.CommandBarButton,
                                                                                                 bool&)
      IL_018f:  newobj     instance void Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler::.ctor(object,
                                                                                                                native uint)
      IL_0194:  callvirt   instance void [mscorlib]System.Reflection.EventInfo::AddEventHandler(object,
                                                                                                class [mscorlib]System.Delegate)
      IL_0199:  ldarg.0
      IL_019a:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton> iReporterExcelAddIn2019.ThisAddIn::MenuTexts
      IL_019f:  ldloc.3
      IL_01a0:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton>::Add(!0)
      IL_01a5:  ldc.i4.1
      IL_01a6:  box        Microsoft.Office.Core.MsoControlType
      IL_01ab:  ldarg.0
      IL_01ac:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::missing
      IL_01b1:  ldarg.0
      IL_01b2:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::missing
      IL_01b7:  ldarg.0
      IL_01b8:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::missing
      IL_01bd:  ldc.i4.1
      IL_01be:  box        [mscorlib]System.Boolean
      IL_01c3:  callvirt   instance class Microsoft.Office.Core.CommandBarControl Microsoft.Office.Core.CommandBarControls::Add(object,
                                                                                                                                object,
                                                                                                                                object,
                                                                                                                                object,
                                                                                                                                object)
      IL_01c8:  castclass  Microsoft.Office.Core.CommandBarButton
      IL_01cd:  stloc.3
      IL_01ce:  ldloc.3
      IL_01cf:  ldc.i4.2
      IL_01d0:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Style(valuetype Microsoft.Office.Core.MsoButtonStyle)
      IL_01d5:  ldloc.3
      IL_01d6:  ldloc.0
      IL_01d7:  dup
      IL_01d8:  ldc.i4.1
      IL_01d9:  add
      IL_01da:  stloc.0
      IL_01db:  stloc.s    V_4
      IL_01dd:  ldloca.s   V_4
      IL_01df:  call       instance string [mscorlib]System.Int32::ToString()
      IL_01e4:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Tag(string)
      IL_01e9:  ldtoken    Microsoft.Office.Core._CommandBarButtonEvents_Event
      IL_01ee:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_01f3:  ldstr      "Click"
      IL_01f8:  newobj     instance void [System.Core]System.Runtime.InteropServices.ComAwareEventInfo::.ctor(class [mscorlib]System.Type,
                                                                                                              string)
      IL_01fd:  ldloc.3
      IL_01fe:  ldarg.0
      IL_01ff:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::'<SetUpMenu>b__13_3'(class Microsoft.Office.Core.CommandBarButton,
                                                                                                 bool&)
      IL_0205:  newobj     instance void Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler::.ctor(object,
                                                                                                                native uint)
      IL_020a:  callvirt   instance void [mscorlib]System.Reflection.EventInfo::AddEventHandler(object,
                                                                                                class [mscorlib]System.Delegate)
      IL_020f:  ldarg.0
      IL_0210:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton> iReporterExcelAddIn2019.ThisAddIn::MenuTexts
      IL_0215:  ldloc.3
      IL_0216:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton>::Add(!0)
      IL_021b:  ldloc.1
      IL_021c:  callvirt   instance bool [mscorlib]System.Collections.IEnumerator::MoveNext()
      IL_0221:  brtrue     IL_0019

      IL_0226:  leave.s    IL_023c

    }  // end .try
    finally
    {
      IL_0228:  ldloc.1
      IL_0229:  isinst     [mscorlib]System.IDisposable
      IL_022e:  stloc.s    V_5
      IL_0230:  ldloc.s    V_5
      IL_0232:  brfalse.s  IL_023b

      IL_0234:  ldloc.s    V_5
      IL_0236:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
      IL_023b:  endfinally
    }  // end handler
    IL_023c:  ldc.i4.0
    IL_023d:  call       class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.UserConfig [iReporterExcelAddInCommon]iReporterExcelAddInCommon.UserConfig::GetInstance(valuetype [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ConfigKind)
    IL_0242:  pop
    IL_0243:  ldarg.0
    IL_0244:  call       instance void iReporterExcelAddIn2019.ThisAddIn::SetupMenuCaptions()
    IL_0249:  ret
  } // end of method ThisAddIn::SetUpMenu

  .method public hidebysig instance void 
          SetupMenuCaptions() cil managed
  {
    // Code size       479 (0x1df)
    .maxstack  4
    .locals init (valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<class Microsoft.Office.Core.CommandBarButton> V_0,
             class Microsoft.Office.Core.CommandBarButton V_1,
             int32 V_2)
    IL_0000:  call       class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::get_Ribbons()
    IL_0005:  callvirt   instance class iReporterExcelAddIn2019.Ribbon1 iReporterExcelAddIn2019.ThisRibbonCollection::get_Ribbon1()
    IL_000a:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::autoJudgeButton
    IL_000f:  call       string iReporterExcelAddIn2019.Properties.Resources::get_MENU_EXTRACTION()
    IL_0014:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_Label(string)
    IL_0019:  call       class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::get_Ribbons()
    IL_001e:  callvirt   instance class iReporterExcelAddIn2019.Ribbon1 iReporterExcelAddIn2019.ThisRibbonCollection::get_Ribbon1()
    IL_0023:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::clusterSettingButton
    IL_0028:  call       string iReporterExcelAddIn2019.Properties.Resources::get_MENU_KIND()
    IL_002d:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_Label(string)
    IL_0032:  call       class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::get_Ribbons()
    IL_0037:  callvirt   instance class iReporterExcelAddIn2019.Ribbon1 iReporterExcelAddIn2019.ThisRibbonCollection::get_Ribbon1()
    IL_003c:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::colorSettingButton
    IL_0041:  call       string iReporterExcelAddIn2019.Properties.Resources::get_MENU_SETTING()
    IL_0046:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_Label(string)
    IL_004b:  call       class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::get_Ribbons()
    IL_0050:  callvirt   instance class iReporterExcelAddIn2019.Ribbon1 iReporterExcelAddIn2019.ThisRibbonCollection::get_Ribbon1()
    IL_0055:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::tableSettingButton
    IL_005a:  call       string iReporterExcelAddIn2019.Properties.Resources::get_MENU_TABLE()
    IL_005f:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_Label(string)
    IL_0064:  call       class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::get_Ribbons()
    IL_0069:  callvirt   instance class iReporterExcelAddIn2019.Ribbon1 iReporterExcelAddIn2019.ThisRibbonCollection::get_Ribbon1()
    IL_006e:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::checkSheetNameButton
    IL_0073:  call       string iReporterExcelAddIn2019.Properties.Resources::get_CHECK_SHEET_NAME()
    IL_0078:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::set_Label(string)
    IL_007d:  ldarg.0
    IL_007e:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton> iReporterExcelAddIn2019.ThisAddIn::MenuTexts
    IL_0083:  call       class [mscorlib]System.Collections.Generic.List`1<!!0> [System.Core]System.Linq.Enumerable::ToList<class Microsoft.Office.Core.CommandBarButton>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
    IL_0088:  callvirt   instance valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<!0> class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton>::GetEnumerator()
    IL_008d:  stloc.0
    .try
    {
      IL_008e:  br         IL_01c2

      IL_0093:  ldloca.s   V_0
      IL_0095:  call       instance !0 valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<class Microsoft.Office.Core.CommandBarButton>::get_Current()
      IL_009a:  stloc.1
      .try
      {
        IL_009b:  ldloc.1
        IL_009c:  callvirt   instance string Microsoft.Office.Core._CommandBarButton::get_Tag()
        IL_00a1:  ldloca.s   V_2
        IL_00a3:  call       bool [mscorlib]System.Int32::TryParse(string,
                                                                   int32&)
        IL_00a8:  brtrue.s   IL_00af

        IL_00aa:  leave      IL_01c2

        IL_00af:  ldloc.2
        IL_00b0:  ldc.i4.s   10
        IL_00b2:  sub
        IL_00b3:  switch     ( 
                              IL_00e5,
                              IL_0113,
                              IL_013e,
                              IL_0169,
                              IL_0194,
                              IL_00e5,
                              IL_0113,
                              IL_013e,
                              IL_0169,
                              IL_0194)
        IL_00e0:  br         IL_01bd

        IL_00e5:  ldloc.1
        IL_00e6:  call       class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::get_Ribbons()
        IL_00eb:  callvirt   instance class iReporterExcelAddIn2019.Ribbon1 iReporterExcelAddIn2019.ThisRibbonCollection::get_Ribbon1()
        IL_00f0:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::autoJudgeButton
        IL_00f5:  callvirt   instance string [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::get_Label()
        IL_00fa:  ldstr      "\n"
        IL_00ff:  ldstr      ""
        IL_0104:  callvirt   instance string [mscorlib]System.String::Replace(string,
                                                                              string)
        IL_0109:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Caption(string)
        IL_010e:  br         IL_01bd

        IL_0113:  ldloc.1
        IL_0114:  call       class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::get_Ribbons()
        IL_0119:  callvirt   instance class iReporterExcelAddIn2019.Ribbon1 iReporterExcelAddIn2019.ThisRibbonCollection::get_Ribbon1()
        IL_011e:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::clusterSettingButton
        IL_0123:  callvirt   instance string [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::get_Label()
        IL_0128:  ldstr      "\n"
        IL_012d:  ldstr      ""
        IL_0132:  callvirt   instance string [mscorlib]System.String::Replace(string,
                                                                              string)
        IL_0137:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Caption(string)
        IL_013c:  br.s       IL_01bd

        IL_013e:  ldloc.1
        IL_013f:  call       class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::get_Ribbons()
        IL_0144:  callvirt   instance class iReporterExcelAddIn2019.Ribbon1 iReporterExcelAddIn2019.ThisRibbonCollection::get_Ribbon1()
        IL_0149:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::colorSettingButton
        IL_014e:  callvirt   instance string [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::get_Label()
        IL_0153:  ldstr      "\n"
        IL_0158:  ldstr      ""
        IL_015d:  callvirt   instance string [mscorlib]System.String::Replace(string,
                                                                              string)
        IL_0162:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Caption(string)
        IL_0167:  br.s       IL_01bd

        IL_0169:  ldloc.1
        IL_016a:  call       class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::get_Ribbons()
        IL_016f:  callvirt   instance class iReporterExcelAddIn2019.Ribbon1 iReporterExcelAddIn2019.ThisRibbonCollection::get_Ribbon1()
        IL_0174:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::tableSettingButton
        IL_0179:  callvirt   instance string [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::get_Label()
        IL_017e:  ldstr      "\n"
        IL_0183:  ldstr      ""
        IL_0188:  callvirt   instance string [mscorlib]System.String::Replace(string,
                                                                              string)
        IL_018d:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Caption(string)
        IL_0192:  br.s       IL_01bd

        IL_0194:  ldloc.1
        IL_0195:  call       class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::get_Ribbons()
        IL_019a:  callvirt   instance class iReporterExcelAddIn2019.Ribbon1 iReporterExcelAddIn2019.ThisRibbonCollection::get_Ribbon1()
        IL_019f:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton iReporterExcelAddIn2019.Ribbon1::checkSheetNameButton
        IL_01a4:  callvirt   instance string [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonButton::get_Label()
        IL_01a9:  ldstr      "\n"
        IL_01ae:  ldstr      ""
        IL_01b3:  callvirt   instance string [mscorlib]System.String::Replace(string,
                                                                              string)
        IL_01b8:  callvirt   instance void Microsoft.Office.Core._CommandBarButton::set_Caption(string)
        IL_01bd:  leave.s    IL_01c2

      }  // end .try
      catch [mscorlib]System.Runtime.InteropServices.COMException 
      {
        IL_01bf:  pop
        IL_01c0:  leave.s    IL_01c2

      }  // end handler
      IL_01c2:  ldloca.s   V_0
      IL_01c4:  call       instance bool valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<class Microsoft.Office.Core.CommandBarButton>::MoveNext()
      IL_01c9:  brtrue     IL_0093

      IL_01ce:  leave.s    IL_01de

    }  // end .try
    finally
    {
      IL_01d0:  ldloca.s   V_0
      IL_01d2:  constrained. valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<class Microsoft.Office.Core.CommandBarButton>
      IL_01d8:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
      IL_01dd:  endfinally
    }  // end handler
    IL_01de:  ret
  } // end of method ThisAddIn::SetupMenuCaptions

  .method private hidebysig instance void 
          ThisAddIn_Startup(object sender,
                            class [mscorlib]System.EventArgs e) cil managed
  {
    // Code size       23 (0x17)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.Globals::get_ThisAddIn()
    IL_0006:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_000b:  stfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::excelApplication
    IL_0010:  ldarg.0
    IL_0011:  call       instance void iReporterExcelAddIn2019.ThisAddIn::SetUpMenu()
    IL_0016:  ret
  } // end of method ThisAddIn::ThisAddIn_Startup

  .method private hidebysig instance void 
          ThisAddIn_Shutdown(object sender,
                             class [mscorlib]System.EventArgs e) cil managed
  {
    // Code size       12 (0xc)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton> iReporterExcelAddIn2019.ThisAddIn::MenuTexts
    IL_0006:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton>::Clear()
    IL_000b:  ret
  } // end of method ThisAddIn::ThisAddIn_Shutdown

  .method public hidebysig instance void 
          ManualCluster() cil managed
  {
    // Code size       363 (0x16b)
    .maxstack  4
    .locals init (object V_0,
             bool V_1,
             class [System.Windows.Forms]System.Windows.Forms.IWin32Window V_2,
             valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point> V_3,
             valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point> V_4)
    IL_0000:  ldarg.0
    IL_0001:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::_inputLocker
    IL_0006:  stloc.0
    IL_0007:  ldc.i4.0
    IL_0008:  stloc.1
    .try
    {
      IL_0009:  ldloc.0
      IL_000a:  ldloca.s   V_1
      IL_000c:  call       void [mscorlib]System.Threading.Monitor::Enter(object,
                                                                          bool&)
      IL_0011:  ldarg.0
      IL_0012:  call       instance class [System.Windows.Forms]System.Windows.Forms.IWin32Window iReporterExcelAddIn2019.ThisAddIn::CreateWinrowHandle()
      IL_0017:  stloc.2
      IL_0018:  ldarg.0
      IL_0019:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_001e:  dup
      IL_001f:  brtrue.s   IL_002e

      IL_0021:  pop
      IL_0022:  ldloca.s   V_4
      IL_0024:  initobj    valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point>
      IL_002a:  ldloc.s    V_4
      IL_002c:  br.s       IL_0038

      IL_002e:  call       instance valuetype [System.Drawing]System.Drawing.Point [System.Windows.Forms]System.Windows.Forms.Form::get_DesktopLocation()
      IL_0033:  newobj     instance void valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point>::.ctor(!0)
      IL_0038:  stloc.3
      IL_0039:  ldarg.0
      IL_003a:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
      IL_003f:  callvirt   instance object Microsoft.Office.Interop.Excel._Application::get_ActiveSheet()
      IL_0044:  isinst     Microsoft.Office.Interop.Excel.Worksheet
      IL_0049:  ldarg.0
      IL_004a:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_004f:  dup
      IL_0050:  brtrue.s   IL_0055

      IL_0052:  pop
      IL_0053:  br.s       IL_005a

      IL_0055:  call       instance void [System.Windows.Forms]System.Windows.Forms.Form::Close()
      IL_005a:  dup
      IL_005b:  brtrue.s   IL_0060

      IL_005d:  pop
      IL_005e:  br.s       IL_0065

      IL_0060:  callvirt   instance void Microsoft.Office.Interop.Excel._Worksheet::Activate()
      IL_0065:  ldarg.0
      IL_0066:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::.ctor()
      IL_006b:  stfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_0070:  ldarg.0
      IL_0071:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_0076:  ldarg.0
      IL_0077:  call       instance class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ClusterList iReporterExcelAddIn2019.ThisAddIn::InitClusterList()
      IL_007c:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::set_AllClusters(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ClusterList)
      IL_0081:  ldarg.0
      IL_0082:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_0087:  ldarg.0
      IL_0088:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::WriteComment(class [mscorlib]System.Tuple`2<int32,string>[],
                                                                                         bool)
      IL_008e:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/WriteCommentEventHandler::.ctor(object,
                                                                                                                                         native int)
      IL_0093:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::add_WriteCommentEvent(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/WriteCommentEventHandler)
      IL_0098:  ldarg.0
      IL_0099:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_009e:  ldarg.0
      IL_009f:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::ClearComments()
      IL_00a5:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/ClearCommentEventHandler::.ctor(object,
                                                                                                                                         native int)
      IL_00aa:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::add_ClearCommentEvent(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/ClearCommentEventHandler)
      IL_00af:  ldarg.0
      IL_00b0:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_00b5:  ldarg.0
      IL_00b6:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::CancelAll()
      IL_00bc:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/ClearCommentEventHandler::.ctor(object,
                                                                                                                                         native int)
      IL_00c1:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::add_CancelAllEvent(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/ClearCommentEventHandler)
      IL_00c6:  ldarg.0
      IL_00c7:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_00cc:  ldarg.0
      IL_00cd:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::ManualCluster()
      IL_00d3:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/ClearCommentEventHandler::.ctor(object,
                                                                                                                                         native int)
      IL_00d8:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::add_LanguageChangeEvent(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/ClearCommentEventHandler)
      IL_00dd:  ldarg.0
      IL_00de:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_00e3:  ldarg.0
      IL_00e4:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::SetupMenuCaptions()
      IL_00ea:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/ClearCommentEventHandler::.ctor(object,
                                                                                                                                         native int)
      IL_00ef:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::add_LanguageChangeEvent(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/ClearCommentEventHandler)
      IL_00f4:  ldarg.0
      IL_00f5:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_00fa:  ldarg.0
      IL_00fb:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::SelectCell(string)
      IL_0101:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/SelectCellEventHandler::.ctor(object,
                                                                                                                                       native int)
      IL_0106:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::add_SelectCellEvent(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/SelectCellEventHandler)
      IL_010b:  ldarg.0
      IL_010c:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_0111:  ldarg.0
      IL_0112:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::'<ManualCluster>b__17_0'()
      IL_0118:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/ClearCommentEventHandler::.ctor(object,
                                                                                                                                         native int)
      IL_011d:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::add_ExtractEvent(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/ClearCommentEventHandler)
      IL_0122:  ldarg.0
      IL_0123:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_0128:  ldarg.0
      IL_0129:  ldftn      instance string[] iReporterExcelAddIn2019.ThisAddIn::GetCellText(bool,
                                                                                            bool,
                                                                                            bool,
                                                                                            bool)
      IL_012f:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/GetCellTextEventHandler::.ctor(object,
                                                                                                                                        native int)
      IL_0134:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::add_GetCellTextEvent(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/GetCellTextEventHandler)
      IL_0139:  ldarg.0
      IL_013a:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_013f:  ldarg.0
      IL_0140:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::SetSheetName(string)
      IL_0146:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/SelectCellEventHandler::.ctor(object,
                                                                                                                                       native int)
      IL_014b:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::add_SetSheetNameEvent(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2/SelectCellEventHandler)
      IL_0150:  ldarg.0
      IL_0151:  ldarg.0
      IL_0152:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
      IL_0157:  ldloc.3
      IL_0158:  ldloc.2
      IL_0159:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ShowForm(class [System.Windows.Forms]System.Windows.Forms.Form,
                                                                                     valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point>,
                                                                                     class [System.Windows.Forms]System.Windows.Forms.IWin32Window)
      IL_015e:  leave.s    IL_016a

    }  // end .try
    finally
    {
      IL_0160:  ldloc.1
      IL_0161:  brfalse.s  IL_0169

      IL_0163:  ldloc.0
      IL_0164:  call       void [mscorlib]System.Threading.Monitor::Exit(object)
      IL_0169:  endfinally
    }  // end handler
    IL_016a:  ret
  } // end of method ThisAddIn::ManualCluster

  .method private hidebysig specialname instance bool 
          get_IsJudging() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      bool iReporterExcelAddIn2019.ThisAddIn::'<IsJudging>k__BackingField'
    IL_0006:  ret
  } // end of method ThisAddIn::get_IsJudging

  .method private hidebysig specialname instance void 
          set_IsJudging(bool 'value') cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       8 (0x8)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  stfld      bool iReporterExcelAddIn2019.ThisAddIn::'<IsJudging>k__BackingField'
    IL_0007:  ret
  } // end of method ThisAddIn::set_IsJudging

  .method public hidebysig instance void 
          AutoJudge(bool openDlg) cil managed
  {
    // Code size       88 (0x58)
    .maxstack  4
    .locals init (class [mscorlib]System.Threading.CancellationTokenSource V_0,
             class iReporterExcelAddIn2019.WaitDialog V_1,
             object V_2,
             bool V_3)
    IL_0000:  ldarg.0
    IL_0001:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::judgeLock
    IL_0006:  stloc.2
    IL_0007:  ldc.i4.0
    IL_0008:  stloc.3
    .try
    {
      IL_0009:  ldloc.2
      IL_000a:  ldloca.s   V_3
      IL_000c:  call       void [mscorlib]System.Threading.Monitor::Enter(object,
                                                                          bool&)
      IL_0011:  ldarg.0
      IL_0012:  call       instance bool iReporterExcelAddIn2019.ThisAddIn::get_IsJudging()
      IL_0017:  brfalse.s  IL_001b

      IL_0019:  leave.s    IL_0057

      IL_001b:  ldarg.0
      IL_001c:  ldc.i4.1
      IL_001d:  call       instance void iReporterExcelAddIn2019.ThisAddIn::set_IsJudging(bool)
      IL_0022:  leave.s    IL_002e

    }  // end .try
    finally
    {
      IL_0024:  ldloc.3
      IL_0025:  brfalse.s  IL_002d

      IL_0027:  ldloc.2
      IL_0028:  call       void [mscorlib]System.Threading.Monitor::Exit(object)
      IL_002d:  endfinally
    }  // end handler
    IL_002e:  nop
    .try
    {
      IL_002f:  newobj     instance void [mscorlib]System.Threading.CancellationTokenSource::.ctor()
      IL_0034:  stloc.0
      IL_0035:  ldloc.0
      IL_0036:  newobj     instance void iReporterExcelAddIn2019.WaitDialog::.ctor(class [mscorlib]System.Threading.CancellationTokenSource)
      IL_003b:  stloc.1
      IL_003c:  leave.s    IL_0048

    }  // end .try
    catch [mscorlib]System.Object 
    {
      IL_003e:  pop
      IL_003f:  ldarg.0
      IL_0040:  ldc.i4.0
      IL_0041:  call       instance void iReporterExcelAddIn2019.ThisAddIn::set_IsJudging(bool)
      IL_0046:  rethrow
    }  // end handler
    IL_0048:  ldarg.0
    IL_0049:  ldloc.0
    IL_004a:  callvirt   instance valuetype [mscorlib]System.Threading.CancellationToken [mscorlib]System.Threading.CancellationTokenSource::get_Token()
    IL_004f:  ldloc.1
    IL_0050:  ldarg.1
    IL_0051:  call       instance class [mscorlib]System.Threading.Tasks.Task iReporterExcelAddIn2019.ThisAddIn::AutoJudgeAsync(valuetype [mscorlib]System.Threading.CancellationToken,
                                                                                                                                class [System.Windows.Forms]System.Windows.Forms.Form,
                                                                                                                                bool)
    IL_0056:  pop
    IL_0057:  ret
  } // end of method ThisAddIn::AutoJudge

  .method private hidebysig instance class [mscorlib]System.Threading.Tasks.Task 
          AutoJudgeAsync(valuetype [mscorlib]System.Threading.CancellationToken token,
                         class [System.Windows.Forms]System.Windows.Forms.Form wait,
                         bool openDlg) cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.AsyncStateMachineAttribute::.ctor(class [mscorlib]System.Type) = ( 01 00 37 69 52 65 70 6F 72 74 65 72 45 78 63 65   // ..7iReporterExce
                                                                                                                                       6C 41 64 64 49 6E 32 30 31 39 2E 54 68 69 73 41   // lAddIn2019.ThisA
                                                                                                                                       64 64 49 6E 2B 3C 41 75 74 6F 4A 75 64 67 65 41   // ddIn+<AutoJudgeA
                                                                                                                                       73 79 6E 63 3E 64 5F 5F 32 35 00 00 )             // sync>d__25..
    // Code size       81 (0x51)
    .maxstack  2
    .locals init (valuetype iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25' V_0,
             valuetype [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder V_1)
    IL_0000:  ldloca.s   V_0
    IL_0002:  ldarg.0
    IL_0003:  stfld      class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::'<>4__this'
    IL_0008:  ldloca.s   V_0
    IL_000a:  ldarg.1
    IL_000b:  stfld      valuetype [mscorlib]System.Threading.CancellationToken iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::token
    IL_0010:  ldloca.s   V_0
    IL_0012:  ldarg.2
    IL_0013:  stfld      class [System.Windows.Forms]System.Windows.Forms.Form iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::wait
    IL_0018:  ldloca.s   V_0
    IL_001a:  ldarg.3
    IL_001b:  stfld      bool iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::openDlg
    IL_0020:  ldloca.s   V_0
    IL_0022:  call       valuetype [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder::Create()
    IL_0027:  stfld      valuetype [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::'<>t__builder'
    IL_002c:  ldloca.s   V_0
    IL_002e:  ldc.i4.m1
    IL_002f:  stfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::'<>1__state'
    IL_0034:  ldloc.0
    IL_0035:  ldfld      valuetype [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::'<>t__builder'
    IL_003a:  stloc.1
    IL_003b:  ldloca.s   V_1
    IL_003d:  ldloca.s   V_0
    IL_003f:  call       instance void [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder::Start<valuetype iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'>(!!0&)
    IL_0044:  ldloca.s   V_0
    IL_0046:  ldflda     valuetype [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder iReporterExcelAddIn2019.ThisAddIn/'<AutoJudgeAsync>d__25'::'<>t__builder'
    IL_004b:  call       instance class [mscorlib]System.Threading.Tasks.Task [mscorlib]System.Runtime.CompilerServices.AsyncTaskMethodBuilder::get_Task()
    IL_0050:  ret
  } // end of method ThisAddIn::AutoJudgeAsync

  .method public hidebysig instance void 
          SetClusterColors() cil managed
  {
    // Code size       145 (0x91)
    .maxstack  4
    .locals init (object V_0,
             bool V_1,
             valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point> V_2)
    IL_0000:  ldarg.0
    IL_0001:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::_inputLocker
    IL_0006:  stloc.0
    IL_0007:  ldc.i4.0
    IL_0008:  stloc.1
    .try
    {
      IL_0009:  ldloc.0
      IL_000a:  ldloca.s   V_1
      IL_000c:  call       void [mscorlib]System.Threading.Monitor::Enter(object,
                                                                          bool&)
      IL_0011:  ldarg.0
      IL_0012:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm iReporterExcelAddIn2019.ThisAddIn::_color
      IL_0017:  dup
      IL_0018:  brtrue.s   IL_001e

      IL_001a:  pop
      IL_001b:  ldc.i4.0
      IL_001c:  br.s       IL_0026

      IL_001e:  call       instance bool [System.Windows.Forms]System.Windows.Forms.Control::get_IsDisposed()
      IL_0023:  ldc.i4.0
      IL_0024:  ceq
      IL_0026:  brfalse.s  IL_0035

      IL_0028:  ldarg.0
      IL_0029:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm iReporterExcelAddIn2019.ThisAddIn::_color
      IL_002e:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.Control::Show()
      IL_0033:  leave.s    IL_0090

      IL_0035:  ldarg.0
      IL_0036:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm::.ctor()
      IL_003b:  stfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm iReporterExcelAddIn2019.ThisAddIn::_color
      IL_0040:  ldarg.0
      IL_0041:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm iReporterExcelAddIn2019.ThisAddIn::_color
      IL_0046:  ldarg.0
      IL_0047:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::SetSelectedColor(valuetype [System.Drawing]System.Drawing.Color)
      IL_004d:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm/SetColorEventHandler::.ctor(object,
                                                                                                                                    native int)
      IL_0052:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm::add_SetColorEvent(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm/SetColorEventHandler)
      IL_0057:  ldarg.0
      IL_0058:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm iReporterExcelAddIn2019.ThisAddIn::_color
      IL_005d:  ldarg.0
      IL_005e:  ldftn      instance valuetype [System.Drawing]System.Drawing.Color iReporterExcelAddIn2019.ThisAddIn::GetSelectedColor()
      IL_0064:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm/GetColorEventHandler::.ctor(object,
                                                                                                                                    native int)
      IL_0069:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm::add_GetColorEvent(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm/GetColorEventHandler)
      IL_006e:  ldarg.0
      IL_006f:  ldarg.0
      IL_0070:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ColorForm iReporterExcelAddIn2019.ThisAddIn::_color
      IL_0075:  ldloca.s   V_2
      IL_0077:  initobj    valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point>
      IL_007d:  ldloc.2
      IL_007e:  ldnull
      IL_007f:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ShowForm(class [System.Windows.Forms]System.Windows.Forms.Form,
                                                                                     valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point>,
                                                                                     class [System.Windows.Forms]System.Windows.Forms.IWin32Window)
      IL_0084:  leave.s    IL_0090

    }  // end .try
    finally
    {
      IL_0086:  ldloc.1
      IL_0087:  brfalse.s  IL_008f

      IL_0089:  ldloc.0
      IL_008a:  call       void [mscorlib]System.Threading.Monitor::Exit(object)
      IL_008f:  endfinally
    }  // end handler
    IL_0090:  ret
  } // end of method ThisAddIn::SetClusterColors

  .method private hidebysig instance void 
          ShowForm(class [System.Windows.Forms]System.Windows.Forms.Form form,
                   valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point> location,
                   class [System.Windows.Forms]System.Windows.Forms.IWin32Window window) cil managed
  {
    // Code size       298 (0x12a)
    .maxstack  4
    .locals init (class iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass27_0' V_0,
             valuetype RECT V_1)
    IL_0000:  newobj     instance void iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass27_0'::.ctor()
    IL_0005:  stloc.0
    IL_0006:  ldloc.0
    IL_0007:  ldarg.0
    IL_0008:  stfld      class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass27_0'::'<>4__this'
    IL_000d:  ldloc.0
    IL_000e:  ldarg.3
    IL_000f:  stfld      class [System.Windows.Forms]System.Windows.Forms.IWin32Window iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass27_0'::window
    IL_0014:  ldloc.0
    IL_0015:  ldloc.0
    IL_0016:  ldfld      class [System.Windows.Forms]System.Windows.Forms.IWin32Window iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass27_0'::window
    IL_001b:  dup
    IL_001c:  brtrue.s   IL_0025

    IL_001e:  pop
    IL_001f:  ldarg.0
    IL_0020:  call       instance class [System.Windows.Forms]System.Windows.Forms.IWin32Window iReporterExcelAddIn2019.ThisAddIn::CreateWinrowHandle()
    IL_0025:  stfld      class [System.Windows.Forms]System.Windows.Forms.IWin32Window iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass27_0'::window
    IL_002a:  ldarg.0
    IL_002b:  ldfld      bool iReporterExcelAddIn2019.ThisAddIn::_selectionChangedSet
    IL_0030:  brtrue     IL_00bd

    IL_0035:  ldarg.0
    IL_0036:  ldc.i4.1
    IL_0037:  stfld      bool iReporterExcelAddIn2019.ThisAddIn::_selectionChangedSet
    IL_003c:  ldtoken    Microsoft.Office.Interop.Excel.AppEvents_Event
    IL_0041:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_0046:  ldstr      "SheetSelectionChange"
    IL_004b:  newobj     instance void [System.Core]System.Runtime.InteropServices.ComAwareEventInfo::.ctor(class [mscorlib]System.Type,
                                                                                                            string)
    IL_0050:  ldarg.0
    IL_0051:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0056:  ldarg.0
    IL_0057:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::OnSelectionChanged(object,
                                                                                             class Microsoft.Office.Interop.Excel.Range)
    IL_005d:  newobj     instance void Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler::.ctor(object,
                                                                                                                        native uint)
    IL_0062:  callvirt   instance void [mscorlib]System.Reflection.EventInfo::AddEventHandler(object,
                                                                                              class [mscorlib]System.Delegate)
    IL_0067:  ldtoken    Microsoft.Office.Interop.Excel.AppEvents_Event
    IL_006c:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_0071:  ldstr      "SheetChange"
    IL_0076:  newobj     instance void [System.Core]System.Runtime.InteropServices.ComAwareEventInfo::.ctor(class [mscorlib]System.Type,
                                                                                                            string)
    IL_007b:  ldarg.0
    IL_007c:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0081:  ldarg.0
    IL_0082:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::OnSelectionChanged(object,
                                                                                             class Microsoft.Office.Interop.Excel.Range)
    IL_0088:  newobj     instance void Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler::.ctor(object,
                                                                                                               native uint)
    IL_008d:  callvirt   instance void [mscorlib]System.Reflection.EventInfo::AddEventHandler(object,
                                                                                              class [mscorlib]System.Delegate)
    IL_0092:  ldtoken    Microsoft.Office.Interop.Excel.AppEvents_Event
    IL_0097:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_009c:  ldstr      "SheetActivate"
    IL_00a1:  newobj     instance void [System.Core]System.Runtime.InteropServices.ComAwareEventInfo::.ctor(class [mscorlib]System.Type,
                                                                                                            string)
    IL_00a6:  ldarg.0
    IL_00a7:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_00ac:  ldloc.0
    IL_00ad:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass27_0'::'<ShowForm>b__0'(object)
    IL_00b3:  newobj     instance void Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler::.ctor(object,
                                                                                                                 native uint)
    IL_00b8:  callvirt   instance void [mscorlib]System.Reflection.EventInfo::AddEventHandler(object,
                                                                                              class [mscorlib]System.Delegate)
    IL_00bd:  ldarga.s   location
    IL_00bf:  call       instance bool valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point>::get_HasValue()
    IL_00c4:  brtrue.s   IL_00fe

    IL_00c6:  ldloc.0
    IL_00c7:  ldfld      class [System.Windows.Forms]System.Windows.Forms.IWin32Window iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass27_0'::window
    IL_00cc:  callvirt   instance native int [System.Windows.Forms]System.Windows.Forms.IWin32Window::get_Handle()
    IL_00d1:  ldloca.s   V_1
    IL_00d3:  call       int32 iReporterExcelAddIn2019.ThisAddIn::GetWindowRect(native int,
                                                                                valuetype RECT&)
    IL_00d8:  pop
    IL_00d9:  ldarga.s   location
    IL_00db:  ldloc.1
    IL_00dc:  ldfld      int32 RECT::Right
    IL_00e1:  ldarg.1
    IL_00e2:  callvirt   instance int32 [System.Windows.Forms]System.Windows.Forms.Control::get_Width()
    IL_00e7:  sub
    IL_00e8:  ldloc.1
    IL_00e9:  ldfld      int32 RECT::Top
    IL_00ee:  ldc.i4.0
    IL_00ef:  call       int32 [mscorlib]System.Math::Max(int32,
                                                          int32)
    IL_00f4:  newobj     instance void [System.Drawing]System.Drawing.Point::.ctor(int32,
                                                                                   int32)
    IL_00f9:  call       instance void valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point>::.ctor(!0)
    IL_00fe:  ldarg.1
    IL_00ff:  ldarga.s   location
    IL_0101:  call       instance !0 valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point>::get_Value()
    IL_0106:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.Form::set_DesktopLocation(valuetype [System.Drawing]System.Drawing.Point)
    IL_010b:  ldarg.1
    IL_010c:  ldloc.0
    IL_010d:  ldfld      class [System.Windows.Forms]System.Windows.Forms.IWin32Window iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass27_0'::window
    IL_0112:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.Form::Show(class [System.Windows.Forms]System.Windows.Forms.IWin32Window)
    IL_0117:  ldarg.1
    IL_0118:  ldloc.0
    IL_0119:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass27_0'::'<ShowForm>b__1'(object,
                                                                                                                   class [System.Windows.Forms]System.Windows.Forms.FormClosedEventArgs)
    IL_011f:  newobj     instance void [System.Windows.Forms]System.Windows.Forms.FormClosedEventHandler::.ctor(object,
                                                                                                                native int)
    IL_0124:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.Form::add_FormClosed(class [System.Windows.Forms]System.Windows.Forms.FormClosedEventHandler)
    IL_0129:  ret
  } // end of method ThisAddIn::ShowForm

  .method private hidebysig instance void 
          ResetInput() cil managed
  {
    // Code size       54 (0x36)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
    IL_0006:  dup
    IL_0007:  brtrue.s   IL_000d

    IL_0009:  pop
    IL_000a:  ldc.i4.0
    IL_000b:  br.s       IL_0012

    IL_000d:  call       instance bool [System.Windows.Forms]System.Windows.Forms.Control::get_Visible()
    IL_0012:  brfalse.s  IL_0035

    IL_0014:  ldarg.0
    IL_0015:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
    IL_001a:  callvirt   instance class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ClusterList [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::get_AllClusters()
    IL_001f:  ldarg.0
    IL_0020:  call       instance class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ClusterList iReporterExcelAddIn2019.ThisAddIn::InitClusterList()
    IL_0025:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ClusterList::Copy(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ClusterList)
    IL_002a:  ldarg.0
    IL_002b:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2 iReporterExcelAddIn2019.ThisAddIn::_input2
    IL_0030:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.InputForm2::Setup()
    IL_0035:  ret
  } // end of method ThisAddIn::ResetInput

  .method private hidebysig instance valuetype [System.Drawing]System.Drawing.Color 
          ToColor(object color) cil managed
  {
    // Code size       166 (0xa6)
    .maxstack  4
    .locals init (int32 V_0,
             float64 V_1,
             object V_2,
             int64 V_3,
             int32 V_4,
             int16 V_5,
             uint8 V_6)
    IL_0000:  ldarg.1
    IL_0001:  dup
    IL_0002:  stloc.2
    IL_0003:  isinst     [mscorlib]System.Double
    IL_0008:  brfalse.s  IL_0016

    IL_000a:  ldloc.2
    IL_000b:  unbox.any  [mscorlib]System.Double
    IL_0010:  stloc.1
    IL_0011:  ldloc.1
    IL_0012:  conv.i4
    IL_0013:  stloc.0
    IL_0014:  br.s       IL_0086

    IL_0016:  ldarg.1
    IL_0017:  dup
    IL_0018:  stloc.2
    IL_0019:  isinst     [mscorlib]System.Int64
    IL_001e:  brfalse.s  IL_002c

    IL_0020:  ldloc.2
    IL_0021:  unbox.any  [mscorlib]System.Int64
    IL_0026:  stloc.3
    IL_0027:  ldloc.3
    IL_0028:  conv.i4
    IL_0029:  stloc.0
    IL_002a:  br.s       IL_0086

    IL_002c:  ldarg.1
    IL_002d:  dup
    IL_002e:  stloc.2
    IL_002f:  isinst     [mscorlib]System.Int32
    IL_0034:  brfalse.s  IL_0043

    IL_0036:  ldloc.2
    IL_0037:  unbox.any  [mscorlib]System.Int32
    IL_003c:  stloc.s    V_4
    IL_003e:  ldloc.s    V_4
    IL_0040:  stloc.0
    IL_0041:  br.s       IL_0086

    IL_0043:  ldarg.1
    IL_0044:  dup
    IL_0045:  stloc.2
    IL_0046:  isinst     [mscorlib]System.Int16
    IL_004b:  brfalse.s  IL_005a

    IL_004d:  ldloc.2
    IL_004e:  unbox.any  [mscorlib]System.Int16
    IL_0053:  stloc.s    V_5
    IL_0055:  ldloc.s    V_5
    IL_0057:  stloc.0
    IL_0058:  br.s       IL_0086

    IL_005a:  ldarg.1
    IL_005b:  dup
    IL_005c:  stloc.2
    IL_005d:  isinst     [mscorlib]System.Byte
    IL_0062:  brfalse.s  IL_0071

    IL_0064:  ldloc.2
    IL_0065:  unbox.any  [mscorlib]System.Byte
    IL_006a:  stloc.s    V_6
    IL_006c:  ldloc.s    V_6
    IL_006e:  stloc.0
    IL_006f:  br.s       IL_0086

    IL_0071:  ldc.i4     0xff
    IL_0076:  ldc.i4     0xff
    IL_007b:  ldc.i4     0xff
    IL_0080:  call       valuetype [System.Drawing]System.Drawing.Color [System.Drawing]System.Drawing.Color::FromArgb(int32,
                                                                                                                       int32,
                                                                                                                       int32)
    IL_0085:  ret

    IL_0086:  ldloc.0
    IL_0087:  ldc.i4     0xff
    IL_008c:  and
    IL_008d:  ldloc.0
    IL_008e:  ldc.i4.8
    IL_008f:  shr
    IL_0090:  ldc.i4     0xff
    IL_0095:  and
    IL_0096:  ldloc.0
    IL_0097:  ldc.i4.s   16
    IL_0099:  shr
    IL_009a:  ldc.i4     0xff
    IL_009f:  and
    IL_00a0:  call       valuetype [System.Drawing]System.Drawing.Color [System.Drawing]System.Drawing.Color::FromArgb(int32,
                                                                                                                       int32,
                                                                                                                       int32)
    IL_00a5:  ret
  } // end of method ThisAddIn::ToColor

  .method private hidebysig instance void 
          OnSelectionChanged(object Sh,
                             class Microsoft.Office.Interop.Excel.Range Target) cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ResetInput()
    IL_0006:  ret
  } // end of method ThisAddIn::OnSelectionChanged

  .method private hidebysig instance bool 
          GetSelected([out] class Microsoft.Office.Interop.Excel.Range& sel) cil managed
  {
    // Code size       25 (0x19)
    .maxstack  3
    .locals init (class Microsoft.Office.Interop.Excel.Range V_0)
    IL_0000:  ldarg.1
    IL_0001:  ldarg.0
    IL_0002:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0007:  callvirt   instance object Microsoft.Office.Interop.Excel._Application::get_Selection()
    IL_000c:  isinst     Microsoft.Office.Interop.Excel.Range
    IL_0011:  dup
    IL_0012:  stloc.0
    IL_0013:  stind.ref
    IL_0014:  ldloc.0
    IL_0015:  ldnull
    IL_0016:  cgt.un
    IL_0018:  ret
  } // end of method ThisAddIn::GetSelected

  .method private hidebysig instance bool 
          GetSelectedOneCell([out] class Microsoft.Office.Interop.Excel.Range& cell) cil managed
  {
    // Code size       32 (0x20)
    .maxstack  2
    .locals init (class Microsoft.Office.Interop.Excel.Range V_0)
    IL_0000:  ldarg.1
    IL_0001:  ldnull
    IL_0002:  stind.ref
    IL_0003:  ldarg.0
    IL_0004:  ldloca.s   V_0
    IL_0006:  call       instance bool iReporterExcelAddIn2019.ThisAddIn::GetSelected(class Microsoft.Office.Interop.Excel.Range&)
    IL_000b:  brfalse.s  IL_001a

    IL_000d:  ldarg.1
    IL_000e:  ldloc.0
    IL_000f:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0> [System.Core]System.Linq.Enumerable::OfType<class Microsoft.Office.Interop.Excel.Range>(class [mscorlib]System.Collections.IEnumerable)
    IL_0014:  call       !!0 [System.Core]System.Linq.Enumerable::FirstOrDefault<class Microsoft.Office.Interop.Excel.Range>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
    IL_0019:  stind.ref
    IL_001a:  ldarg.1
    IL_001b:  ldind.ref
    IL_001c:  ldnull
    IL_001d:  cgt.un
    IL_001f:  ret
  } // end of method ThisAddIn::GetSelectedOneCell

  .method private hidebysig instance void 
          SetSelectedColor(valuetype [System.Drawing]System.Drawing.Color color) cil managed
  {
    // Code size       55 (0x37)
    .maxstack  4
    .locals init (class Microsoft.Office.Interop.Excel.Range V_0)
    IL_0000:  ldarg.0
    IL_0001:  ldloca.s   V_0
    IL_0003:  call       instance bool iReporterExcelAddIn2019.ThisAddIn::GetSelected(class Microsoft.Office.Interop.Excel.Range&)
    IL_0008:  brfalse.s  IL_0036

    IL_000a:  ldloc.0
    IL_000b:  callvirt   instance class Microsoft.Office.Interop.Excel.Interior Microsoft.Office.Interop.Excel.Range::get_Interior()
    IL_0010:  ldarga.s   color
    IL_0012:  call       instance uint8 [System.Drawing]System.Drawing.Color::get_R()
    IL_0017:  ldarga.s   color
    IL_0019:  call       instance uint8 [System.Drawing]System.Drawing.Color::get_G()
    IL_001e:  ldc.i4.8
    IL_001f:  shl
    IL_0020:  or
    IL_0021:  ldarga.s   color
    IL_0023:  call       instance uint8 [System.Drawing]System.Drawing.Color::get_B()
    IL_0028:  ldc.i4.s   16
    IL_002a:  shl
    IL_002b:  or
    IL_002c:  box        [mscorlib]System.Int32
    IL_0031:  callvirt   instance void Microsoft.Office.Interop.Excel.Interior::set_Color(object)
    IL_0036:  ret
  } // end of method ThisAddIn::SetSelectedColor

  .method private hidebysig instance valuetype [System.Drawing]System.Drawing.Color 
          GetSelectedColor() cil managed
  {
    // Code size       189 (0xbd)
    .maxstack  11
    .locals init (class Microsoft.Office.Interop.Excel.Range V_0,
             valuetype [System.Drawing]System.Drawing.Color V_1)
    .try
    {
      IL_0000:  ldarg.0
      IL_0001:  ldloca.s   V_0
      IL_0003:  call       instance bool iReporterExcelAddIn2019.ThisAddIn::GetSelectedOneCell(class Microsoft.Office.Interop.Excel.Range&)
      IL_0008:  brfalse    IL_00b0

      IL_000d:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype [System.Drawing]System.Drawing.Color>> iReporterExcelAddIn2019.ThisAddIn/'<>o__34'::'<>p__1'
      IL_0012:  brtrue.s   IL_0038

      IL_0014:  ldc.i4.0
      IL_0015:  ldtoken    [System.Drawing]System.Drawing.Color
      IL_001a:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_001f:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_0024:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0029:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::Convert(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                              class [mscorlib]System.Type,
                                                                                                                                                              class [mscorlib]System.Type)
      IL_002e:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype [System.Drawing]System.Drawing.Color>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_0033:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype [System.Drawing]System.Drawing.Color>> iReporterExcelAddIn2019.ThisAddIn/'<>o__34'::'<>p__1'
      IL_0038:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype [System.Drawing]System.Drawing.Color>> iReporterExcelAddIn2019.ThisAddIn/'<>o__34'::'<>p__1'
      IL_003d:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype [System.Drawing]System.Drawing.Color>>::Target
      IL_0042:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype [System.Drawing]System.Drawing.Color>> iReporterExcelAddIn2019.ThisAddIn/'<>o__34'::'<>p__1'
      IL_0047:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class iReporterExcelAddIn2019.ThisAddIn,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__34'::'<>p__0'
      IL_004c:  brtrue.s   IL_0088

      IL_004e:  ldc.i4.2
      IL_004f:  ldstr      "ToColor"
      IL_0054:  ldnull
      IL_0055:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_005a:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_005f:  ldc.i4.2
      IL_0060:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_0065:  dup
      IL_0066:  ldc.i4.0
      IL_0067:  ldc.i4.1
      IL_0068:  ldnull
      IL_0069:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_006e:  stelem.ref
      IL_006f:  dup
      IL_0070:  ldc.i4.1
      IL_0071:  ldc.i4.0
      IL_0072:  ldnull
      IL_0073:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0078:  stelem.ref
      IL_0079:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   string,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_007e:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class iReporterExcelAddIn2019.ThisAddIn,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_0083:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class iReporterExcelAddIn2019.ThisAddIn,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__34'::'<>p__0'
      IL_0088:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class iReporterExcelAddIn2019.ThisAddIn,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__34'::'<>p__0'
      IL_008d:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class iReporterExcelAddIn2019.ThisAddIn,object,object>>::Target
      IL_0092:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class iReporterExcelAddIn2019.ThisAddIn,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__34'::'<>p__0'
      IL_0097:  ldarg.0
      IL_0098:  ldloc.0
      IL_0099:  callvirt   instance class Microsoft.Office.Interop.Excel.Interior Microsoft.Office.Interop.Excel.Range::get_Interior()
      IL_009e:  callvirt   instance object Microsoft.Office.Interop.Excel.Interior::get_Color()
      IL_00a3:  callvirt   instance !3 class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class iReporterExcelAddIn2019.ThisAddIn,object,object>::Invoke(!0,
                                                                                                                                                                                                !1,
                                                                                                                                                                                                !2)
      IL_00a8:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype [System.Drawing]System.Drawing.Color>::Invoke(!0,
                                                                                                                                                                                                !1)
      IL_00ad:  stloc.1
      IL_00ae:  leave.s    IL_00bb

      IL_00b0:  leave.s    IL_00b5

    }  // end .try
    catch [mscorlib]System.Object 
    {
      IL_00b2:  pop
      IL_00b3:  leave.s    IL_00b5

    }  // end handler
    IL_00b5:  ldsfld     valuetype [System.Drawing]System.Drawing.Color [System.Drawing]System.Drawing.Color::Empty
    IL_00ba:  ret

    IL_00bb:  ldloc.1
    IL_00bc:  ret
  } // end of method ThisAddIn::GetSelectedColor

  .method private hidebysig instance class [mscorlib]System.Collections.Generic.IEnumerable`1<string> 
          SplitLine(string baseText) cil managed
  {
    // Code size       42 (0x2a)
    .maxstack  8
    IL_0000:  ldarg.1
    IL_0001:  dup
    IL_0002:  brtrue.s   IL_000a

    IL_0004:  pop
    IL_0005:  ldstr      ""
    IL_000a:  ldstr      "\r"
    IL_000f:  ldstr      ""
    IL_0014:  callvirt   instance string [mscorlib]System.String::Replace(string,
                                                                          string)
    IL_0019:  ldc.i4.1
    IL_001a:  newarr     [mscorlib]System.Char
    IL_001f:  dup
    IL_0020:  ldc.i4.0
    IL_0021:  ldc.i4.s   10
    IL_0023:  stelem.i2
    IL_0024:  callvirt   instance string[] [mscorlib]System.String::Split(char[])
    IL_0029:  ret
  } // end of method ThisAddIn::SplitLine

  .method private hidebysig instance void 
          WriteComment(class [mscorlib]System.Tuple`2<int32,string>[] keyAndTexts,
                       bool focusExcel) cil managed
  {
    // Code size       1037 (0x40d)
    .maxstack  8
    .locals init (bool V_0,
             int32 V_1,
             class Microsoft.Office.Interop.Excel.Range V_2,
             class [mscorlib]System.Collections.Generic.IEnumerator`1<string> V_3,
             class [mscorlib]System.Collections.Generic.IEnumerable`1<string> V_4,
             string V_5,
             class [mscorlib]System.Tuple`2<int32,int32> V_6,
             class [mscorlib]System.Tuple`2<int32,int32> V_7,
             int32 V_8,
             class [mscorlib]System.Collections.Generic.List`1<valuetype [System.Drawing]System.Drawing.Rectangle> V_9,
             class iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_0' V_10,
             class iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1' V_11,
             int32 V_12,
             class Microsoft.Office.Interop.Excel.Range V_13,
             class [mscorlib]System.Collections.Generic.List`1<string> V_14,
             string V_15,
             class Microsoft.Office.Interop.Excel.Range V_16,
             class [mscorlib]System.Tuple`2<int32,string>[] V_17,
             int32 V_18,
             class [mscorlib]System.Tuple`2<int32,string> V_19,
             valuetype [System.Drawing]System.Drawing.Rectangle V_20)
    IL_0000:  ldarg.1
    IL_0001:  ldsfld     class [mscorlib]System.Func`2<class [mscorlib]System.Tuple`2<int32,string>,int32> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__36_0'
    IL_0006:  dup
    IL_0007:  brtrue.s   IL_0020

    IL_0009:  pop
    IL_000a:  ldsfld     class iReporterExcelAddIn2019.ThisAddIn/'<>c' iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9'
    IL_000f:  ldftn      instance int32 iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<WriteComment>b__36_0'(class [mscorlib]System.Tuple`2<int32,string>)
    IL_0015:  newobj     instance void class [mscorlib]System.Func`2<class [mscorlib]System.Tuple`2<int32,string>,int32>::.ctor(object,
                                                                                                                                native int)
    IL_001a:  dup
    IL_001b:  stsfld     class [mscorlib]System.Func`2<class [mscorlib]System.Tuple`2<int32,string>,int32> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__36_0'
    IL_0020:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!1> [System.Core]System.Linq.Enumerable::Select<class [mscorlib]System.Tuple`2<int32,string>,int32>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                                                                       class [mscorlib]System.Func`2<!!0,!!1>)
    IL_0025:  ldsfld     class [mscorlib]System.Func`2<int32,bool> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__36_1'
    IL_002a:  dup
    IL_002b:  brtrue.s   IL_0044

    IL_002d:  pop
    IL_002e:  ldsfld     class iReporterExcelAddIn2019.ThisAddIn/'<>c' iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9'
    IL_0033:  ldftn      instance bool iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<WriteComment>b__36_1'(int32)
    IL_0039:  newobj     instance void class [mscorlib]System.Func`2<int32,bool>::.ctor(object,
                                                                                        native int)
    IL_003e:  dup
    IL_003f:  stsfld     class [mscorlib]System.Func`2<int32,bool> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__36_1'
    IL_0044:  call       bool [System.Core]System.Linq.Enumerable::Any<int32>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                              class [mscorlib]System.Func`2<!!0,bool>)
    IL_0049:  stloc.0
    IL_004a:  ldc.i4.0
    IL_004b:  stloc.1
    IL_004c:  ldarg.0
    IL_004d:  ldloca.s   V_2
    IL_004f:  call       instance bool iReporterExcelAddIn2019.ThisAddIn::GetSelected(class Microsoft.Office.Interop.Excel.Range&)
    IL_0054:  brfalse    IL_03eb

    IL_0059:  ldloc.2
    IL_005a:  ldc.i4.0
    IL_005b:  box        [mscorlib]System.Boolean
    IL_0060:  ldc.i4.0
    IL_0061:  box        [mscorlib]System.Boolean
    IL_0066:  ldc.i4.1
    IL_0067:  ldsfld     object [mscorlib]System.Type::Missing
    IL_006c:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0071:  callvirt   instance string Microsoft.Office.Interop.Excel.Range::get_Address(object,
                                                                                           object,
                                                                                           valuetype Microsoft.Office.Interop.Excel.XlReferenceStyle,
                                                                                           object,
                                                                                           object)
    IL_0076:  dup
    IL_0077:  brtrue.s   IL_007d

    IL_0079:  pop
    IL_007a:  ldnull
    IL_007b:  br.s       IL_008d

    IL_007d:  ldc.i4.1
    IL_007e:  newarr     [mscorlib]System.Char
    IL_0083:  dup
    IL_0084:  ldc.i4.0
    IL_0085:  ldc.i4.s   44
    IL_0087:  stelem.i2
    IL_0088:  call       instance string[] [mscorlib]System.String::Split(char[])
    IL_008d:  stloc.s    V_4
    IL_008f:  ldloc.s    V_4
    IL_0091:  dup
    IL_0092:  brtrue.s   IL_009a

    IL_0094:  pop
    IL_0095:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0> [System.Core]System.Linq.Enumerable::Empty<string>()
    IL_009a:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerator`1<!0> class [mscorlib]System.Collections.Generic.IEnumerable`1<string>::GetEnumerator()
    IL_009f:  stloc.3
    .try
    {
      IL_00a0:  br         IL_03d4

      IL_00a5:  ldloc.3
      IL_00a6:  callvirt   instance !0 class [mscorlib]System.Collections.Generic.IEnumerator`1<string>::get_Current()
      IL_00ab:  stloc.s    V_5
      IL_00ad:  ldloc.s    V_5
      IL_00af:  ldc.i4.s   58
      IL_00b1:  callvirt   instance int32 [mscorlib]System.String::IndexOf(char)
      IL_00b6:  stloc.s    V_8
      IL_00b8:  ldloc.s    V_8
      IL_00ba:  ldc.i4.0
      IL_00bb:  bge.s      IL_00cb

      IL_00bd:  ldloc.s    V_5
      IL_00bf:  call       class [mscorlib]System.Tuple`2<int32,int32> [Cimtops.Excel]Cimtops.Excel.XLRangeUtil::ToRowCol(string)
      IL_00c4:  dup
      IL_00c5:  stloc.s    V_7
      IL_00c7:  stloc.s    V_6
      IL_00c9:  br.s       IL_00ee

      IL_00cb:  ldloc.s    V_5
      IL_00cd:  ldc.i4.0
      IL_00ce:  ldloc.s    V_8
      IL_00d0:  callvirt   instance string [mscorlib]System.String::Substring(int32,
                                                                              int32)
      IL_00d5:  call       class [mscorlib]System.Tuple`2<int32,int32> [Cimtops.Excel]Cimtops.Excel.XLRangeUtil::ToRowCol(string)
      IL_00da:  stloc.s    V_6
      IL_00dc:  ldloc.s    V_5
      IL_00de:  ldloc.s    V_8
      IL_00e0:  ldc.i4.1
      IL_00e1:  add
      IL_00e2:  callvirt   instance string [mscorlib]System.String::Substring(int32)
      IL_00e7:  call       class [mscorlib]System.Tuple`2<int32,int32> [Cimtops.Excel]Cimtops.Excel.XLRangeUtil::ToRowCol(string)
      IL_00ec:  stloc.s    V_7
      IL_00ee:  ldloc.s    V_6
      IL_00f0:  callvirt   instance !0 class [mscorlib]System.Tuple`2<int32,int32>::get_Item1()
      IL_00f5:  ldc.i4.0
      IL_00f6:  ble        IL_03d4

      IL_00fb:  ldloc.s    V_6
      IL_00fd:  callvirt   instance !1 class [mscorlib]System.Tuple`2<int32,int32>::get_Item2()
      IL_0102:  ldc.i4.0
      IL_0103:  ble        IL_03d4

      IL_0108:  newobj     instance void class [mscorlib]System.Collections.Generic.List`1<valuetype [System.Drawing]System.Drawing.Rectangle>::.ctor()
      IL_010d:  stloc.s    V_9
      IL_010f:  newobj     instance void iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_0'::.ctor()
      IL_0114:  stloc.s    V_10
      IL_0116:  ldloc.s    V_10
      IL_0118:  ldloc.s    V_6
      IL_011a:  callvirt   instance !0 class [mscorlib]System.Tuple`2<int32,int32>::get_Item1()
      IL_011f:  stfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_0'::row
      IL_0124:  br         IL_03c1

      IL_0129:  ldc.i4     0x3e8
      IL_012e:  ldloc.1
      IL_012f:  blt        IL_03d4

      IL_0134:  newobj     instance void iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1'::.ctor()
      IL_0139:  stloc.s    V_11
      IL_013b:  ldloc.s    V_11
      IL_013d:  ldloc.s    V_10
      IL_013f:  stfld      class iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_0' iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1'::'CS$<>8__locals1'
      IL_0144:  ldloc.s    V_11
      IL_0146:  ldloc.s    V_6
      IL_0148:  callvirt   instance !1 class [mscorlib]System.Tuple`2<int32,int32>::get_Item2()
      IL_014d:  stfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1'::col
      IL_0152:  br         IL_039a

      IL_0157:  ldloc.s    V_9
      IL_0159:  ldloc.s    V_11
      IL_015b:  ldftn      instance bool iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1'::'<WriteComment>b__2'(valuetype [System.Drawing]System.Drawing.Rectangle)
      IL_0161:  newobj     instance void class [mscorlib]System.Predicate`1<valuetype [System.Drawing]System.Drawing.Rectangle>::.ctor(object,
                                                                                                                                       native int)
      IL_0166:  callvirt   instance int32 class [mscorlib]System.Collections.Generic.List`1<valuetype [System.Drawing]System.Drawing.Rectangle>::FindIndex(class [mscorlib]System.Predicate`1<!0>)
      IL_016b:  stloc.s    V_12
      IL_016d:  ldloc.s    V_12
      IL_016f:  ldc.i4.0
      IL_0170:  bge        IL_036b

      IL_0175:  ldarg.0
      IL_0176:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
      IL_017b:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Application::get_Cells()
      IL_0180:  ldloc.s    V_11
      IL_0182:  ldfld      class iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_0' iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1'::'CS$<>8__locals1'
      IL_0187:  ldfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_0'::row
      IL_018c:  box        [mscorlib]System.Int32
      IL_0191:  ldloc.s    V_11
      IL_0193:  ldfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1'::col
      IL_0198:  box        [mscorlib]System.Int32
      IL_019d:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get__Default(object,
                                                                                              object)
      IL_01a2:  isinst     Microsoft.Office.Interop.Excel.Range
      IL_01a7:  dup
      IL_01a8:  stloc.s    V_13
      IL_01aa:  brfalse    IL_0386

      IL_01af:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__36'::'<>p__0'
      IL_01b4:  brtrue.s   IL_01e2

      IL_01b6:  ldc.i4.0
      IL_01b7:  ldc.i4.s   83
      IL_01b9:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_01be:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_01c3:  ldc.i4.1
      IL_01c4:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_01c9:  dup
      IL_01ca:  ldc.i4.0
      IL_01cb:  ldc.i4.0
      IL_01cc:  ldnull
      IL_01cd:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_01d2:  stelem.ref
      IL_01d3:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::UnaryOperation(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                     valuetype [System.Core]System.Linq.Expressions.ExpressionType,
                                                                                                                                                                     class [mscorlib]System.Type,
                                                                                                                                                                     class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_01d8:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_01dd:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__36'::'<>p__0'
      IL_01e2:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__36'::'<>p__0'
      IL_01e7:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>>::Target
      IL_01ec:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__36'::'<>p__0'
      IL_01f1:  ldloc.s    V_13
      IL_01f3:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get_MergeCells()
      IL_01f8:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>::Invoke(!0,
                                                                                                                                                      !1)
      IL_01fd:  brfalse.s  IL_023a

      IL_01ff:  ldloc.s    V_13
      IL_0201:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_MergeArea()
      IL_0206:  stloc.s    V_16
      IL_0208:  ldloc.s    V_9
      IL_020a:  ldloc.s    V_16
      IL_020c:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Column()
      IL_0211:  ldloc.s    V_16
      IL_0213:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Row()
      IL_0218:  ldloc.s    V_16
      IL_021a:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_Columns()
      IL_021f:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Count()
      IL_0224:  ldloc.s    V_16
      IL_0226:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_Rows()
      IL_022b:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Count()
      IL_0230:  newobj     instance void [System.Drawing]System.Drawing.Rectangle::.ctor(int32,
                                                                                         int32,
                                                                                         int32,
                                                                                         int32)
      IL_0235:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<valuetype [System.Drawing]System.Drawing.Rectangle>::Add(!0)
      IL_023a:  ldc.i4     0x3e8
      IL_023f:  ldloc.1
      IL_0240:  ldc.i4.1
      IL_0241:  add
      IL_0242:  dup
      IL_0243:  stloc.1
      IL_0244:  bge.s      IL_0265

      IL_0246:  call       string iReporterExcelAddIn2019.Properties.Resources::get_TOO_MANY_SELECT()
      IL_024b:  ldc.i4     0x3e8
      IL_0250:  box        [mscorlib]System.Int32
      IL_0255:  call       string [mscorlib]System.String::Format(string,
                                                                  object)
      IL_025a:  call       valuetype [System.Windows.Forms]System.Windows.Forms.DialogResult [System.Windows.Forms]System.Windows.Forms.MessageBox::Show(string)
      IL_025f:  pop
      IL_0260:  br         IL_03ad

      IL_0265:  ldarg.0
      IL_0266:  ldloc.s    V_13
      IL_0268:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::get_Comment()
      IL_026d:  dup
      IL_026e:  brtrue.s   IL_0274

      IL_0270:  pop
      IL_0271:  ldnull
      IL_0272:  br.s       IL_0288

      IL_0274:  ldsfld     object [mscorlib]System.Type::Missing
      IL_0279:  ldsfld     object [mscorlib]System.Type::Missing
      IL_027e:  ldsfld     object [mscorlib]System.Type::Missing
      IL_0283:  callvirt   instance string Microsoft.Office.Interop.Excel.Comment::Text(object,
                                                                                        object,
                                                                                        object)
      IL_0288:  call       instance class [mscorlib]System.Collections.Generic.IEnumerable`1<string> iReporterExcelAddIn2019.ThisAddIn::SplitLine(string)
      IL_028d:  call       class [mscorlib]System.Collections.Generic.List`1<!!0> [System.Core]System.Linq.Enumerable::ToList<string>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
      IL_0292:  stloc.s    V_14
      IL_0294:  br.s       IL_02bd

      IL_0296:  ldloc.s    V_14
      IL_0298:  ldloc.s    V_14
      IL_029a:  callvirt   instance int32 class [mscorlib]System.Collections.Generic.List`1<string>::get_Count()
      IL_029f:  ldc.i4.3
      IL_02a0:  beq.s      IL_02b3

      IL_02a2:  ldloc.s    V_14
      IL_02a4:  callvirt   instance int32 class [mscorlib]System.Collections.Generic.List`1<string>::get_Count()
      IL_02a9:  ldc.i4.4
      IL_02aa:  beq.s      IL_02b3

      IL_02ac:  ldstr      ""
      IL_02b1:  br.s       IL_02b8

      IL_02b3:  ldstr      "0"
      IL_02b8:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<string>::Add(!0)
      IL_02bd:  ldloc.s    V_14
      IL_02bf:  callvirt   instance int32 class [mscorlib]System.Collections.Generic.List`1<string>::get_Count()
      IL_02c4:  ldc.i4.s   16
      IL_02c6:  blt.s      IL_0296

      IL_02c8:  ldarg.1
      IL_02c9:  stloc.s    V_17
      IL_02cb:  ldc.i4.0
      IL_02cc:  stloc.s    V_18
      IL_02ce:  br.s       IL_0307

      IL_02d0:  ldloc.s    V_17
      IL_02d2:  ldloc.s    V_18
      IL_02d4:  ldelem.ref
      IL_02d5:  stloc.s    V_19
      IL_02d7:  ldloc.s    V_19
      IL_02d9:  callvirt   instance !0 class [mscorlib]System.Tuple`2<int32,string>::get_Item1()
      IL_02de:  ldc.i4.0
      IL_02df:  blt.s      IL_0301

      IL_02e1:  ldc.i4.s   16
      IL_02e3:  ldloc.s    V_19
      IL_02e5:  callvirt   instance !0 class [mscorlib]System.Tuple`2<int32,string>::get_Item1()
      IL_02ea:  ble.s      IL_0301

      IL_02ec:  ldloc.s    V_14
      IL_02ee:  ldloc.s    V_19
      IL_02f0:  callvirt   instance !0 class [mscorlib]System.Tuple`2<int32,string>::get_Item1()
      IL_02f5:  ldloc.s    V_19
      IL_02f7:  callvirt   instance !1 class [mscorlib]System.Tuple`2<int32,string>::get_Item2()
      IL_02fc:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<string>::set_Item(int32,
                                                                                                             !0)
      IL_0301:  ldloc.s    V_18
      IL_0303:  ldc.i4.1
      IL_0304:  add
      IL_0305:  stloc.s    V_18
      IL_0307:  ldloc.s    V_18
      IL_0309:  ldloc.s    V_17
      IL_030b:  ldlen
      IL_030c:  conv.i4
      IL_030d:  blt.s      IL_02d0

      IL_030f:  call       string [mscorlib]System.Environment::get_NewLine()
      IL_0314:  ldloc.s    V_14
      IL_0316:  call       string [mscorlib]System.String::Join(string,
                                                                class [mscorlib]System.Collections.Generic.IEnumerable`1<string>)
      IL_031b:  stloc.s    V_15
      IL_031d:  ldloc.s    V_13
      IL_031f:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::get_Comment()
      IL_0324:  brtrue.s   IL_0350

      IL_0326:  ldc.i4.1
      IL_0327:  stloc.0
      IL_0328:  ldloc.s    V_13
      IL_032a:  ldsfld     object [mscorlib]System.Type::Missing
      IL_032f:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::AddComment(object)
      IL_0334:  pop
      IL_0335:  ldloc.s    V_13
      IL_0337:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::get_Comment()
      IL_033c:  ldloc.s    V_15
      IL_033e:  ldsfld     object [mscorlib]System.Type::Missing
      IL_0343:  ldsfld     object [mscorlib]System.Type::Missing
      IL_0348:  callvirt   instance string Microsoft.Office.Interop.Excel.Comment::Text(object,
                                                                                        object,
                                                                                        object)
      IL_034d:  pop
      IL_034e:  br.s       IL_0386

      IL_0350:  ldloc.s    V_13
      IL_0352:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::get_Comment()
      IL_0357:  ldloc.s    V_15
      IL_0359:  ldsfld     object [mscorlib]System.Type::Missing
      IL_035e:  ldsfld     object [mscorlib]System.Type::Missing
      IL_0363:  callvirt   instance string Microsoft.Office.Interop.Excel.Comment::Text(object,
                                                                                        object,
                                                                                        object)
      IL_0368:  pop
      IL_0369:  br.s       IL_0386

      IL_036b:  ldloc.s    V_11
      IL_036d:  ldloc.s    V_9
      IL_036f:  ldloc.s    V_12
      IL_0371:  callvirt   instance !0 class [mscorlib]System.Collections.Generic.List`1<valuetype [System.Drawing]System.Drawing.Rectangle>::get_Item(int32)
      IL_0376:  stloc.s    V_20
      IL_0378:  ldloca.s   V_20
      IL_037a:  call       instance int32 [System.Drawing]System.Drawing.Rectangle::get_Right()
      IL_037f:  ldc.i4.1
      IL_0380:  sub
      IL_0381:  stfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1'::col
      IL_0386:  ldloc.s    V_11
      IL_0388:  ldfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1'::col
      IL_038d:  ldc.i4.1
      IL_038e:  add
      IL_038f:  stloc.s    V_18
      IL_0391:  ldloc.s    V_11
      IL_0393:  ldloc.s    V_18
      IL_0395:  stfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1'::col
      IL_039a:  ldloc.s    V_11
      IL_039c:  ldfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_1'::col
      IL_03a1:  ldloc.s    V_7
      IL_03a3:  callvirt   instance !1 class [mscorlib]System.Tuple`2<int32,int32>::get_Item2()
      IL_03a8:  ble        IL_0157

      IL_03ad:  ldloc.s    V_10
      IL_03af:  ldfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_0'::row
      IL_03b4:  ldc.i4.1
      IL_03b5:  add
      IL_03b6:  stloc.s    V_18
      IL_03b8:  ldloc.s    V_10
      IL_03ba:  ldloc.s    V_18
      IL_03bc:  stfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_0'::row
      IL_03c1:  ldloc.s    V_10
      IL_03c3:  ldfld      int32 iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass36_0'::row
      IL_03c8:  ldloc.s    V_7
      IL_03ca:  callvirt   instance !0 class [mscorlib]System.Tuple`2<int32,int32>::get_Item1()
      IL_03cf:  ble        IL_0129

      IL_03d4:  ldloc.3
      IL_03d5:  callvirt   instance bool [mscorlib]System.Collections.IEnumerator::MoveNext()
      IL_03da:  brtrue     IL_00a5

      IL_03df:  leave.s    IL_03eb

    }  // end .try
    finally
    {
      IL_03e1:  ldloc.3
      IL_03e2:  brfalse.s  IL_03ea

      IL_03e4:  ldloc.3
      IL_03e5:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
      IL_03ea:  endfinally
    }  // end handler
    IL_03eb:  ldloc.0
    IL_03ec:  brfalse.s  IL_03f8

    IL_03ee:  ldc.i4.0
    IL_03ef:  ldloc.1
    IL_03f0:  bge.s      IL_03f8

    IL_03f2:  ldarg.0
    IL_03f3:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ResetInput()
    IL_03f8:  ldarg.2
    IL_03f9:  brfalse.s  IL_040c

    IL_03fb:  ldarg.0
    IL_03fc:  call       instance class [System.Windows.Forms]System.Windows.Forms.IWin32Window iReporterExcelAddIn2019.ThisAddIn::CreateWinrowHandle()
    IL_0401:  callvirt   instance native int [System.Windows.Forms]System.Windows.Forms.IWin32Window::get_Handle()
    IL_0406:  call       bool iReporterExcelAddIn2019.ThisAddIn::SetForegroundWindow(native int)
    IL_040b:  pop
    IL_040c:  ret
  } // end of method ThisAddIn::WriteComment

  .method private hidebysig instance void 
          WriteTable(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table table) cil managed
  {
    // Code size       362 (0x16a)
    .maxstack  4
    .locals init (class iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass37_0' V_0,
             class Microsoft.Office.Interop.Excel.Worksheet V_1,
             class [mscorlib]System.Collections.Generic.IEnumerator`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableRow> V_2,
             class [mscorlib]System.Collections.Generic.IEnumerator`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem> V_3,
             class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem V_4,
             class Microsoft.Office.Interop.Excel.Range V_5,
             string V_6)
    IL_0000:  newobj     instance void iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass37_0'::.ctor()
    IL_0005:  stloc.0
    IL_0006:  ldloc.0
    IL_0007:  ldarg.1
    IL_0008:  stfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass37_0'::table
    .try
    {
      IL_000d:  ldarg.0
      IL_000e:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
      IL_0013:  callvirt   instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
      IL_0018:  dup
      IL_0019:  brtrue.s   IL_001f

      IL_001b:  pop
      IL_001c:  ldnull
      IL_001d:  br.s       IL_0041

      IL_001f:  callvirt   instance class Microsoft.Office.Interop.Excel.Sheets Microsoft.Office.Interop.Excel._Workbook::get_Sheets()
      IL_0024:  dup
      IL_0025:  brtrue.s   IL_002b

      IL_0027:  pop
      IL_0028:  ldnull
      IL_0029:  br.s       IL_0041

      IL_002b:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0> [System.Core]System.Linq.Enumerable::OfType<class Microsoft.Office.Interop.Excel.Worksheet>(class [mscorlib]System.Collections.IEnumerable)
      IL_0030:  ldloc.0
      IL_0031:  ldftn      instance bool iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass37_0'::'<WriteTable>b__0'(class Microsoft.Office.Interop.Excel.Worksheet)
      IL_0037:  newobj     instance void class [mscorlib]System.Func`2<class Microsoft.Office.Interop.Excel.Worksheet,bool>::.ctor(object,
                                                                                                                                   native int)
      IL_003c:  call       !!0 [System.Core]System.Linq.Enumerable::FirstOrDefault<class Microsoft.Office.Interop.Excel.Worksheet>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                   class [mscorlib]System.Func`2<!!0,bool>)
      IL_0041:  stloc.1
      IL_0042:  ldloc.1
      IL_0043:  brtrue.s   IL_0065

      IL_0045:  call       string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Properties.Resources::get_NOT_FOUND_SHEET()
      IL_004a:  ldloc.0
      IL_004b:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass37_0'::table
      IL_0050:  callvirt   instance string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::get_SheetName()
      IL_0055:  call       string [mscorlib]System.String::Format(string,
                                                                  object)
      IL_005a:  call       valuetype [System.Windows.Forms]System.Windows.Forms.DialogResult [System.Windows.Forms]System.Windows.Forms.MessageBox::Show(string)
      IL_005f:  pop
      IL_0060:  leave      IL_0169

      IL_0065:  ldloc.0
      IL_0066:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass37_0'::table
      IL_006b:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableRow> [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::get_Rows()
      IL_0070:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerator`1<!0> class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableRow>::GetEnumerator()
      IL_0075:  stloc.2
      .try
      {
        IL_0076:  br         IL_013b

        IL_007b:  ldloc.2
        IL_007c:  callvirt   instance !0 class [mscorlib]System.Collections.Generic.IEnumerator`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableRow>::get_Current()
        IL_0081:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem> [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableLine::get_Items()
        IL_0086:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerator`1<!0> class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem>::GetEnumerator()
        IL_008b:  stloc.3
        .try
        {
          IL_008c:  br         IL_0124

          IL_0091:  ldloc.3
          IL_0092:  callvirt   instance !0 class [mscorlib]System.Collections.Generic.IEnumerator`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem>::get_Current()
          IL_0097:  stloc.s    V_4
          IL_0099:  ldloc.1
          IL_009a:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Cells()
          IL_009f:  ldloc.s    V_4
          IL_00a1:  callvirt   instance int32 [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem::get_Y()
          IL_00a6:  box        [mscorlib]System.Int32
          IL_00ab:  ldloc.s    V_4
          IL_00ad:  callvirt   instance int32 [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem::get_X()
          IL_00b2:  box        [mscorlib]System.Int32
          IL_00b7:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get__Default(object,
                                                                                                  object)
          IL_00bc:  isinst     Microsoft.Office.Interop.Excel.Range
          IL_00c1:  stloc.s    V_5
          IL_00c3:  ldloc.s    V_5
          IL_00c5:  brfalse.s  IL_0124

          IL_00c7:  call       string [mscorlib]System.Environment::get_NewLine()
          IL_00cc:  ldloc.s    V_4
          IL_00ce:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerable`1<string> [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem::get_Comments()
          IL_00d3:  call       string [mscorlib]System.String::Join(string,
                                                                    class [mscorlib]System.Collections.Generic.IEnumerable`1<string>)
          IL_00d8:  stloc.s    V_6
          IL_00da:  ldloc.s    V_5
          IL_00dc:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::get_Comment()
          IL_00e1:  brtrue.s   IL_010b

          IL_00e3:  ldloc.s    V_5
          IL_00e5:  ldsfld     object [mscorlib]System.Type::Missing
          IL_00ea:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::AddComment(object)
          IL_00ef:  pop
          IL_00f0:  ldloc.s    V_5
          IL_00f2:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::get_Comment()
          IL_00f7:  ldloc.s    V_6
          IL_00f9:  ldsfld     object [mscorlib]System.Type::Missing
          IL_00fe:  ldsfld     object [mscorlib]System.Type::Missing
          IL_0103:  callvirt   instance string Microsoft.Office.Interop.Excel.Comment::Text(object,
                                                                                            object,
                                                                                            object)
          IL_0108:  pop
          IL_0109:  br.s       IL_0124

          IL_010b:  ldloc.s    V_5
          IL_010d:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::get_Comment()
          IL_0112:  ldloc.s    V_6
          IL_0114:  ldsfld     object [mscorlib]System.Type::Missing
          IL_0119:  ldsfld     object [mscorlib]System.Type::Missing
          IL_011e:  callvirt   instance string Microsoft.Office.Interop.Excel.Comment::Text(object,
                                                                                            object,
                                                                                            object)
          IL_0123:  pop
          IL_0124:  ldloc.3
          IL_0125:  callvirt   instance bool [mscorlib]System.Collections.IEnumerator::MoveNext()
          IL_012a:  brtrue     IL_0091

          IL_012f:  leave.s    IL_013b

        }  // end .try
        finally
        {
          IL_0131:  ldloc.3
          IL_0132:  brfalse.s  IL_013a

          IL_0134:  ldloc.3
          IL_0135:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
          IL_013a:  endfinally
        }  // end handler
        IL_013b:  ldloc.2
        IL_013c:  callvirt   instance bool [mscorlib]System.Collections.IEnumerator::MoveNext()
        IL_0141:  brtrue     IL_007b

        IL_0146:  leave.s    IL_0152

      }  // end .try
      finally
      {
        IL_0148:  ldloc.2
        IL_0149:  brfalse.s  IL_0151

        IL_014b:  ldloc.2
        IL_014c:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
        IL_0151:  endfinally
      }  // end handler
      IL_0152:  leave.s    IL_0169

    }  // end .try
    catch [mscorlib]System.Exception 
    {
      IL_0154:  callvirt   instance string [mscorlib]System.Object::ToString()
      IL_0159:  ldstr      "error"
      IL_015e:  ldc.i4.0
      IL_015f:  ldc.i4.s   16
      IL_0161:  call       valuetype [System.Windows.Forms]System.Windows.Forms.DialogResult [System.Windows.Forms]System.Windows.Forms.MessageBox::Show(string,
                                                                                                                                                         string,
                                                                                                                                                         valuetype [System.Windows.Forms]System.Windows.Forms.MessageBoxButtons,
                                                                                                                                                         valuetype [System.Windows.Forms]System.Windows.Forms.MessageBoxIcon)
      IL_0166:  pop
      IL_0167:  leave.s    IL_0169

    }  // end handler
    IL_0169:  ret
  } // end of method ThisAddIn::WriteTable

  .method private hidebysig instance void 
          ClearComments() cil managed
  {
    // Code size       23 (0x17)
    .maxstack  2
    .locals init (class Microsoft.Office.Interop.Excel.Range V_0)
    IL_0000:  ldarg.0
    IL_0001:  ldloca.s   V_0
    IL_0003:  call       instance bool iReporterExcelAddIn2019.ThisAddIn::GetSelected(class Microsoft.Office.Interop.Excel.Range&)
    IL_0008:  brfalse.s  IL_0016

    IL_000a:  ldloc.0
    IL_000b:  callvirt   instance void Microsoft.Office.Interop.Excel.Range::ClearComments()
    IL_0010:  ldarg.0
    IL_0011:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ResetInput()
    IL_0016:  ret
  } // end of method ThisAddIn::ClearComments

  .method private hidebysig instance void 
          CancelAll() cil managed
  {
    // Code size       61 (0x3d)
    .maxstack  2
    .locals init (class Microsoft.Office.Interop.Excel.Worksheet V_0)
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0006:  callvirt   instance object Microsoft.Office.Interop.Excel._Application::get_ActiveSheet()
    IL_000b:  isinst     Microsoft.Office.Interop.Excel.Worksheet
    IL_0010:  dup
    IL_0011:  stloc.0
    IL_0012:  brfalse.s  IL_003c

    IL_0014:  ldloc.0
    IL_0015:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Cells()
    IL_001a:  dup
    IL_001b:  brtrue.s   IL_0020

    IL_001d:  pop
    IL_001e:  br.s       IL_0025

    IL_0020:  callvirt   instance void Microsoft.Office.Interop.Excel.Range::ClearComments()
    IL_0025:  ldarg.0
    IL_0026:  call       instance string iReporterExcelAddIn2019.ThisAddIn::GetFilePath()
    IL_002b:  ldloc.0
    IL_002c:  callvirt   instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
    IL_0031:  call       void [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder::ClearTitle(string,
                                                                                       string)
    IL_0036:  ldarg.0
    IL_0037:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ResetInput()
    IL_003c:  ret
  } // end of method ThisAddIn::CancelAll

  .method private hidebysig instance void 
          SelectCell(string rangeName) cil managed
  {
    // Code size       57 (0x39)
    .maxstack  3
    .locals init (class Microsoft.Office.Interop.Excel.Worksheet V_0,
             class Microsoft.Office.Interop.Excel.Workbook V_1,
             class Microsoft.Office.Interop.Excel.Range V_2)
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0006:  callvirt   instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
    IL_000b:  stloc.1
    IL_000c:  ldloc.1
    IL_000d:  brtrue.s   IL_0012

    IL_000f:  ldnull
    IL_0010:  br.s       IL_0018

    IL_0012:  ldloc.1
    IL_0013:  callvirt   instance object Microsoft.Office.Interop.Excel._Workbook::get_ActiveSheet()
    IL_0018:  isinst     Microsoft.Office.Interop.Excel.Worksheet
    IL_001d:  dup
    IL_001e:  stloc.0
    IL_001f:  brfalse.s  IL_0038

    IL_0021:  ldloc.0
    IL_0022:  ldarg.1
    IL_0023:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0028:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Range(object,
                                                                                                                                  object)
    IL_002d:  stloc.2
    IL_002e:  ldloc.2
    IL_002f:  brfalse.s  IL_0038

    IL_0031:  ldloc.2
    IL_0032:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::Select()
    IL_0037:  pop
    IL_0038:  ret
  } // end of method ThisAddIn::SelectCell

  .method private hidebysig instance void 
          SetSheetName(string sheetName) cil managed
  {
    // Code size       41 (0x29)
    .maxstack  2
    .locals init (class Microsoft.Office.Interop.Excel.Worksheet V_0,
             class Microsoft.Office.Interop.Excel.Workbook V_1)
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0006:  callvirt   instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
    IL_000b:  stloc.1
    IL_000c:  ldloc.1
    IL_000d:  brtrue.s   IL_0012

    IL_000f:  ldnull
    IL_0010:  br.s       IL_0018

    IL_0012:  ldloc.1
    IL_0013:  callvirt   instance object Microsoft.Office.Interop.Excel._Workbook::get_ActiveSheet()
    IL_0018:  isinst     Microsoft.Office.Interop.Excel.Worksheet
    IL_001d:  dup
    IL_001e:  stloc.0
    IL_001f:  brfalse.s  IL_0028

    IL_0021:  ldloc.0
    IL_0022:  ldarg.1
    IL_0023:  callvirt   instance void Microsoft.Office.Interop.Excel._Worksheet::set_Name(string)
    IL_0028:  ret
  } // end of method ThisAddIn::SetSheetName

  .method private hidebysig instance string[] 
          GetCellText(bool left,
                      bool right,
                      bool up,
                      bool down) cil managed
  {
    // Code size       333 (0x14d)
    .maxstack  5
    .locals init (valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0' V_0,
             class Microsoft.Office.Interop.Excel.Range V_1,
             class Microsoft.Office.Interop.Excel.Worksheet V_2,
             class Microsoft.Office.Interop.Excel.Window V_3,
             valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1' V_4,
             int32 V_5,
             int32 V_6,
             int32 V_7,
             int32 V_8)
    IL_0000:  ldloca.s   V_0
    IL_0002:  newobj     instance void class [mscorlib]System.Collections.Generic.List`1<string>::.ctor()
    IL_0007:  stfld      class [mscorlib]System.Collections.Generic.List`1<string> iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'::list
    IL_000c:  ldarg.0
    IL_000d:  ldloca.s   V_1
    IL_000f:  call       instance bool iReporterExcelAddIn2019.ThisAddIn::GetSelectedOneCell(class Microsoft.Office.Interop.Excel.Range&)
    IL_0014:  brfalse    IL_0141

    IL_0019:  ldarg.0
    IL_001a:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_001f:  callvirt   instance class Microsoft.Office.Interop.Excel.Window Microsoft.Office.Interop.Excel._Application::get_ActiveWindow()
    IL_0024:  stloc.3
    IL_0025:  ldloc.3
    IL_0026:  brtrue.s   IL_002b

    IL_0028:  ldnull
    IL_0029:  br.s       IL_0031

    IL_002b:  ldloc.3
    IL_002c:  callvirt   instance object Microsoft.Office.Interop.Excel.Window::get_ActiveSheet()
    IL_0031:  isinst     Microsoft.Office.Interop.Excel.Worksheet
    IL_0036:  dup
    IL_0037:  stloc.2
    IL_0038:  brfalse    IL_0141

    IL_003d:  ldloca.s   V_4
    IL_003f:  newobj     instance void class [System.Core]System.Collections.Generic.HashSet`1<string>::.ctor()
    IL_0044:  stfld      class [System.Core]System.Collections.Generic.HashSet`1<string> iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1'::set
    IL_0049:  ldloc.1
    IL_004a:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Row()
    IL_004f:  stloc.s    V_5
    IL_0051:  ldloc.1
    IL_0052:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Column()
    IL_0057:  stloc.s    V_6
    IL_0059:  ldloc.2
    IL_005a:  ldloc.s    V_5
    IL_005c:  ldloc.s    V_6
    IL_005e:  ldloca.s   V_0
    IL_0060:  ldloca.s   V_4
    IL_0062:  call       void iReporterExcelAddIn2019.ThisAddIn::'<GetCellText>g__Add|42_0'(class Microsoft.Office.Interop.Excel.Worksheet,
                                                                                            int32,
                                                                                            int32,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'&,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1'&)
    IL_0067:  ldc.i4.0
    IL_0068:  stloc.s    V_7
    IL_006a:  br         IL_0139

    IL_006f:  ldc.i4.0
    IL_0070:  stloc.s    V_8
    IL_0072:  br         IL_012a

    IL_0077:  ldarg.1
    IL_0078:  brfalse.s  IL_00a2

    IL_007a:  ldloc.2
    IL_007b:  ldloc.s    V_5
    IL_007d:  ldloc.s    V_8
    IL_007f:  sub
    IL_0080:  ldloc.s    V_6
    IL_0082:  ldloc.s    V_7
    IL_0084:  sub
    IL_0085:  ldloca.s   V_0
    IL_0087:  ldloca.s   V_4
    IL_0089:  call       void iReporterExcelAddIn2019.ThisAddIn::'<GetCellText>g__Add|42_0'(class Microsoft.Office.Interop.Excel.Worksheet,
                                                                                            int32,
                                                                                            int32,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'&,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1'&)
    IL_008e:  ldloc.2
    IL_008f:  ldloc.s    V_5
    IL_0091:  ldloc.s    V_8
    IL_0093:  add
    IL_0094:  ldloc.s    V_6
    IL_0096:  ldloc.s    V_7
    IL_0098:  sub
    IL_0099:  ldloca.s   V_0
    IL_009b:  ldloca.s   V_4
    IL_009d:  call       void iReporterExcelAddIn2019.ThisAddIn::'<GetCellText>g__Add|42_0'(class Microsoft.Office.Interop.Excel.Worksheet,
                                                                                            int32,
                                                                                            int32,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'&,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1'&)
    IL_00a2:  ldarg.2
    IL_00a3:  brfalse.s  IL_00cd

    IL_00a5:  ldloc.2
    IL_00a6:  ldloc.s    V_5
    IL_00a8:  ldloc.s    V_8
    IL_00aa:  sub
    IL_00ab:  ldloc.s    V_6
    IL_00ad:  ldloc.s    V_7
    IL_00af:  add
    IL_00b0:  ldloca.s   V_0
    IL_00b2:  ldloca.s   V_4
    IL_00b4:  call       void iReporterExcelAddIn2019.ThisAddIn::'<GetCellText>g__Add|42_0'(class Microsoft.Office.Interop.Excel.Worksheet,
                                                                                            int32,
                                                                                            int32,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'&,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1'&)
    IL_00b9:  ldloc.2
    IL_00ba:  ldloc.s    V_5
    IL_00bc:  ldloc.s    V_8
    IL_00be:  add
    IL_00bf:  ldloc.s    V_6
    IL_00c1:  ldloc.s    V_7
    IL_00c3:  add
    IL_00c4:  ldloca.s   V_0
    IL_00c6:  ldloca.s   V_4
    IL_00c8:  call       void iReporterExcelAddIn2019.ThisAddIn::'<GetCellText>g__Add|42_0'(class Microsoft.Office.Interop.Excel.Worksheet,
                                                                                            int32,
                                                                                            int32,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'&,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1'&)
    IL_00cd:  ldarg.3
    IL_00ce:  brfalse.s  IL_00f8

    IL_00d0:  ldloc.2
    IL_00d1:  ldloc.s    V_5
    IL_00d3:  ldloc.s    V_7
    IL_00d5:  sub
    IL_00d6:  ldloc.s    V_6
    IL_00d8:  ldloc.s    V_8
    IL_00da:  sub
    IL_00db:  ldloca.s   V_0
    IL_00dd:  ldloca.s   V_4
    IL_00df:  call       void iReporterExcelAddIn2019.ThisAddIn::'<GetCellText>g__Add|42_0'(class Microsoft.Office.Interop.Excel.Worksheet,
                                                                                            int32,
                                                                                            int32,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'&,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1'&)
    IL_00e4:  ldloc.2
    IL_00e5:  ldloc.s    V_5
    IL_00e7:  ldloc.s    V_7
    IL_00e9:  sub
    IL_00ea:  ldloc.s    V_6
    IL_00ec:  ldloc.s    V_8
    IL_00ee:  add
    IL_00ef:  ldloca.s   V_0
    IL_00f1:  ldloca.s   V_4
    IL_00f3:  call       void iReporterExcelAddIn2019.ThisAddIn::'<GetCellText>g__Add|42_0'(class Microsoft.Office.Interop.Excel.Worksheet,
                                                                                            int32,
                                                                                            int32,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'&,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1'&)
    IL_00f8:  ldarg.s    down
    IL_00fa:  brfalse.s  IL_0124

    IL_00fc:  ldloc.2
    IL_00fd:  ldloc.s    V_5
    IL_00ff:  ldloc.s    V_7
    IL_0101:  add
    IL_0102:  ldloc.s    V_6
    IL_0104:  ldloc.s    V_8
    IL_0106:  sub
    IL_0107:  ldloca.s   V_0
    IL_0109:  ldloca.s   V_4
    IL_010b:  call       void iReporterExcelAddIn2019.ThisAddIn::'<GetCellText>g__Add|42_0'(class Microsoft.Office.Interop.Excel.Worksheet,
                                                                                            int32,
                                                                                            int32,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'&,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1'&)
    IL_0110:  ldloc.2
    IL_0111:  ldloc.s    V_5
    IL_0113:  ldloc.s    V_7
    IL_0115:  add
    IL_0116:  ldloc.s    V_6
    IL_0118:  ldloc.s    V_8
    IL_011a:  add
    IL_011b:  ldloca.s   V_0
    IL_011d:  ldloca.s   V_4
    IL_011f:  call       void iReporterExcelAddIn2019.ThisAddIn::'<GetCellText>g__Add|42_0'(class Microsoft.Office.Interop.Excel.Worksheet,
                                                                                            int32,
                                                                                            int32,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'&,
                                                                                            valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1'&)
    IL_0124:  ldloc.s    V_8
    IL_0126:  ldc.i4.1
    IL_0127:  add
    IL_0128:  stloc.s    V_8
    IL_012a:  ldloc.s    V_8
    IL_012c:  ldloc.s    V_7
    IL_012e:  ble        IL_0077

    IL_0133:  ldloc.s    V_7
    IL_0135:  ldc.i4.1
    IL_0136:  add
    IL_0137:  stloc.s    V_7
    IL_0139:  ldloc.s    V_7
    IL_013b:  ldc.i4.5
    IL_013c:  blt        IL_006f

    IL_0141:  ldloc.0
    IL_0142:  ldfld      class [mscorlib]System.Collections.Generic.List`1<string> iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'::list
    IL_0147:  callvirt   instance !0[] class [mscorlib]System.Collections.Generic.List`1<string>::ToArray()
    IL_014c:  ret
  } // end of method ThisAddIn::GetCellText

  .method private hidebysig instance string 
          GetFilePath() cil managed
  {
    // Code size       70 (0x46)
    .maxstack  3
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0006:  callvirt   instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
    IL_000b:  dup
    IL_000c:  brtrue.s   IL_0012

    IL_000e:  pop
    IL_000f:  ldnull
    IL_0010:  br.s       IL_0017

    IL_0012:  callvirt   instance string Microsoft.Office.Interop.Excel._Workbook::get_Path()
    IL_0017:  dup
    IL_0018:  brtrue.s   IL_0020

    IL_001a:  pop
    IL_001b:  ldstr      ""
    IL_0020:  ldarg.0
    IL_0021:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0026:  callvirt   instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
    IL_002b:  dup
    IL_002c:  brtrue.s   IL_0032

    IL_002e:  pop
    IL_002f:  ldnull
    IL_0030:  br.s       IL_0037

    IL_0032:  callvirt   instance string Microsoft.Office.Interop.Excel._Workbook::get_Name()
    IL_0037:  dup
    IL_0038:  brtrue.s   IL_0040

    IL_003a:  pop
    IL_003b:  ldstr      ""
    IL_0040:  call       string [mscorlib]System.IO.Path::Combine(string,
                                                                  string)
    IL_0045:  ret
  } // end of method ThisAddIn::GetFilePath

  .method private hidebysig instance class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ClusterList 
          InitClusterList() cil managed
  {
    // Code size       534 (0x216)
    .maxstack  31
    .locals init (string V_0,
             class Microsoft.Office.Interop.Excel.Worksheet V_1,
             class Microsoft.Office.Interop.Excel.Workbook V_2,
             class [Cimtops.R2Cluster]Cimtops.R2Cluster.TitleInfo V_3,
             class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ClusterList V_4,
             class Microsoft.Office.Interop.Excel.Comments V_5,
             valuetype [mscorlib]System.Nullable`1<int32> V_6,
             class [mscorlib]System.Exception V_7,
             valuetype [mscorlib]System.Nullable`1<int32> V_8,
             class Microsoft.Office.Interop.Excel.Range V_9,
             int32 V_10,
             class Microsoft.Office.Interop.Excel.Comment V_11,
             class Microsoft.Office.Interop.Excel.Range V_12,
             bool V_13)
    IL_0000:  ldarg.0
    IL_0001:  call       instance string iReporterExcelAddIn2019.ThisAddIn::GetFilePath()
    IL_0006:  stloc.0
    IL_0007:  ldarg.0
    IL_0008:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_000d:  callvirt   instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
    IL_0012:  stloc.2
    IL_0013:  ldloc.2
    IL_0014:  brtrue.s   IL_0019

    IL_0016:  ldnull
    IL_0017:  br.s       IL_001f

    IL_0019:  ldloc.2
    IL_001a:  callvirt   instance object Microsoft.Office.Interop.Excel._Workbook::get_ActiveSheet()
    IL_001f:  isinst     Microsoft.Office.Interop.Excel.Worksheet
    IL_0024:  dup
    IL_0025:  stloc.1
    IL_0026:  brfalse    IL_0209

    IL_002b:  ldloc.0
    IL_002c:  ldloc.1
    IL_002d:  callvirt   instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
    IL_0032:  call       class [Cimtops.R2Cluster]Cimtops.R2Cluster.TitleInfo [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder::FindTitle(string,
                                                                                                                                      string)
    IL_0037:  dup
    IL_0038:  brtrue.s   IL_0053

    IL_003a:  pop
    IL_003b:  ldarg.0
    IL_003c:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0041:  callvirt   instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
    IL_0046:  ldloc.1
    IL_0047:  ldloca.s   V_7
    IL_0049:  call       class [Cimtops.Excel]Cimtops.Excel.Sheet [Cimtops.Excel]Cimtops.Excel.BookFactory::Create(class Microsoft.Office.Interop.Excel.Workbook,
                                                                                                                   class Microsoft.Office.Interop.Excel.Worksheet,
                                                                                                                   class [mscorlib]System.Exception&)
    IL_004e:  call       class [Cimtops.R2Cluster]Cimtops.R2Cluster.TitleInfo [Cimtops.R2Cluster]Cimtops.R2Cluster.Decoder::GetSheetTitleS(class [Cimtops.Excel]Cimtops.Excel.Sheet)
    IL_0053:  stloc.3
    IL_0054:  ldloc.0
    IL_0055:  ldloc.1
    IL_0056:  callvirt   instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
    IL_005b:  ldloc.3
    IL_005c:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ClusterList::.ctor(string,
                                                                                                               string,
                                                                                                               class [Cimtops.R2Cluster]Cimtops.R2Cluster.TitleInfo)
    IL_0061:  stloc.s    V_4
    IL_0063:  ldloc.1
    IL_0064:  callvirt   instance class Microsoft.Office.Interop.Excel.Comments Microsoft.Office.Interop.Excel._Worksheet::get_Comments()
    IL_0069:  stloc.s    V_5
    IL_006b:  ldloc.s    V_5
    IL_006d:  brtrue.s   IL_007b

    IL_006f:  ldloca.s   V_8
    IL_0071:  initobj    valuetype [mscorlib]System.Nullable`1<int32>
    IL_0077:  ldloc.s    V_8
    IL_0079:  br.s       IL_0087

    IL_007b:  ldloc.s    V_5
    IL_007d:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Comments::get_Count()
    IL_0082:  newobj     instance void valuetype [mscorlib]System.Nullable`1<int32>::.ctor(!0)
    IL_0087:  stloc.s    V_6
    IL_0089:  ldc.i4.0
    IL_008a:  ldloc.s    V_6
    IL_008c:  stloc.s    V_8
    IL_008e:  ldloca.s   V_8
    IL_0090:  call       instance !0 valuetype [mscorlib]System.Nullable`1<int32>::GetValueOrDefault()
    IL_0095:  clt
    IL_0097:  ldloca.s   V_8
    IL_0099:  call       instance bool valuetype [mscorlib]System.Nullable`1<int32>::get_HasValue()
    IL_009e:  and
    IL_009f:  brfalse    IL_01ff

    IL_00a4:  ldarg.0
    IL_00a5:  ldloca.s   V_9
    IL_00a7:  call       instance bool iReporterExcelAddIn2019.ThisAddIn::GetSelected(class Microsoft.Office.Interop.Excel.Range&)
    IL_00ac:  pop
    IL_00ad:  ldc.i4.0
    IL_00ae:  stloc.s    V_10
    IL_00b0:  br         IL_01e3

    IL_00b5:  ldloc.s    V_5
    IL_00b7:  ldloc.s    V_10
    IL_00b9:  ldc.i4.1
    IL_00ba:  add
    IL_00bb:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Comments::get__Default(int32)
    IL_00c0:  stloc.s    V_11
    IL_00c2:  ldloc.s    V_11
    IL_00c4:  callvirt   instance object Microsoft.Office.Interop.Excel.Comment::get_Parent()
    IL_00c9:  isinst     Microsoft.Office.Interop.Excel.Range
    IL_00ce:  dup
    IL_00cf:  stloc.s    V_12
    IL_00d1:  brfalse    IL_01dd

    IL_00d6:  ldloc.s    V_9
    IL_00d8:  brfalse    IL_017d

    IL_00dd:  ldarg.0
    IL_00de:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_00e3:  ldloc.s    V_9
    IL_00e5:  ldloc.s    V_12
    IL_00e7:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00ec:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00f1:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00f6:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00fb:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0100:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0105:  ldsfld     object [mscorlib]System.Type::Missing
    IL_010a:  ldsfld     object [mscorlib]System.Type::Missing
    IL_010f:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0114:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0119:  ldsfld     object [mscorlib]System.Type::Missing
    IL_011e:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0123:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0128:  ldsfld     object [mscorlib]System.Type::Missing
    IL_012d:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0132:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0137:  ldsfld     object [mscorlib]System.Type::Missing
    IL_013c:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0141:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0146:  ldsfld     object [mscorlib]System.Type::Missing
    IL_014b:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0150:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0155:  ldsfld     object [mscorlib]System.Type::Missing
    IL_015a:  ldsfld     object [mscorlib]System.Type::Missing
    IL_015f:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0164:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0169:  ldsfld     object [mscorlib]System.Type::Missing
    IL_016e:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0173:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Application::Intersect(class Microsoft.Office.Interop.Excel.Range,
                                                                                                                                    class Microsoft.Office.Interop.Excel.Range,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object)
    IL_0178:  ldnull
    IL_0179:  cgt.un
    IL_017b:  br.s       IL_017e

    IL_017d:  ldc.i4.0
    IL_017e:  stloc.s    V_13
    IL_0180:  ldloc.s    V_12
    IL_0182:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_MergeArea()
    IL_0187:  dup
    IL_0188:  brtrue.s   IL_018d

    IL_018a:  pop
    IL_018b:  ldloc.s    V_12
    IL_018d:  stloc.s    V_12
    IL_018f:  ldloc.s    V_4
    IL_0191:  ldloc.s    V_12
    IL_0193:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Row()
    IL_0198:  ldloc.s    V_12
    IL_019a:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Column()
    IL_019f:  ldloc.s    V_12
    IL_01a1:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_Rows()
    IL_01a6:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Count()
    IL_01ab:  ldloc.s    V_12
    IL_01ad:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_Columns()
    IL_01b2:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Count()
    IL_01b7:  ldloc.s    V_11
    IL_01b9:  ldsfld     object [mscorlib]System.Type::Missing
    IL_01be:  ldsfld     object [mscorlib]System.Type::Missing
    IL_01c3:  ldsfld     object [mscorlib]System.Type::Missing
    IL_01c8:  callvirt   instance string Microsoft.Office.Interop.Excel.Comment::Text(object,
                                                                                      object,
                                                                                      object)
    IL_01cd:  dup
    IL_01ce:  brtrue.s   IL_01d6

    IL_01d0:  pop
    IL_01d1:  ldstr      ""
    IL_01d6:  ldloc.s    V_13
    IL_01d8:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ClusterList::Add(int32,
                                                                                                             int32,
                                                                                                             int32,
                                                                                                             int32,
                                                                                                             string,
                                                                                                             bool)
    IL_01dd:  ldloc.s    V_10
    IL_01df:  ldc.i4.1
    IL_01e0:  add
    IL_01e1:  stloc.s    V_10
    IL_01e3:  ldloc.s    V_10
    IL_01e5:  ldloc.s    V_6
    IL_01e7:  stloc.s    V_8
    IL_01e9:  ldloca.s   V_8
    IL_01eb:  call       instance !0 valuetype [mscorlib]System.Nullable`1<int32>::GetValueOrDefault()
    IL_01f0:  clt
    IL_01f2:  ldloca.s   V_8
    IL_01f4:  call       instance bool valuetype [mscorlib]System.Nullable`1<int32>::get_HasValue()
    IL_01f9:  and
    IL_01fa:  brtrue     IL_00b5

    IL_01ff:  ldloc.s    V_4
    IL_0201:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ClusterList::Sort()
    IL_0206:  ldloc.s    V_4
    IL_0208:  ret

    IL_0209:  ldloc.0
    IL_020a:  ldstr      ""
    IL_020f:  ldnull
    IL_0210:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ClusterList::.ctor(string,
                                                                                                               string,
                                                                                                               class [Cimtops.R2Cluster]Cimtops.R2Cluster.TitleInfo)
    IL_0215:  ret
  } // end of method ThisAddIn::InitClusterList

  .method public hidebysig instance void 
          ShowTableForm() cil managed
  {
    // Code size       669 (0x29d)
    .maxstack  4
    .locals init (class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table> V_0,
             class [mscorlib]System.Collections.IEnumerator V_1,
             class Microsoft.Office.Interop.Excel.Worksheet V_2,
             class Microsoft.Office.Interop.Excel.Comments V_3,
             valuetype [mscorlib]System.Nullable`1<int32> V_4,
             valuetype [mscorlib]System.Nullable`1<int32> V_5,
             int32 V_6,
             class iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass45_0' V_7,
             class Microsoft.Office.Interop.Excel.Comment V_8,
             class Microsoft.Office.Interop.Excel.Range V_9,
             int32 V_10,
             class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table V_11,
             class [mscorlib]System.IDisposable V_12,
             valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table> V_13,
             object V_14,
             bool V_15,
             valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point> V_16)
    .try
    {
      IL_0000:  ldarg.0
      IL_0001:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
      IL_0006:  callvirt   instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
      IL_000b:  dup
      IL_000c:  brtrue.s   IL_0012

      IL_000e:  pop
      IL_000f:  ldnull
      IL_0010:  br.s       IL_0017

      IL_0012:  callvirt   instance class Microsoft.Office.Interop.Excel.Sheets Microsoft.Office.Interop.Excel._Workbook::get_Sheets()
      IL_0017:  brtrue.s   IL_001e

      IL_0019:  leave      IL_029c

      IL_001e:  newobj     instance void class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table>::.ctor()
      IL_0023:  stloc.0
      IL_0024:  ldarg.0
      IL_0025:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
      IL_002a:  callvirt   instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
      IL_002f:  callvirt   instance class Microsoft.Office.Interop.Excel.Sheets Microsoft.Office.Interop.Excel._Workbook::get_Sheets()
      IL_0034:  callvirt   instance class [mscorlib]System.Collections.IEnumerator Microsoft.Office.Interop.Excel.Sheets::GetEnumerator()
      IL_0039:  stloc.1
      .try
      {
        IL_003a:  br         IL_0156

        IL_003f:  ldloc.1
        IL_0040:  callvirt   instance object [mscorlib]System.Collections.IEnumerator::get_Current()
        IL_0045:  isinst     Microsoft.Office.Interop.Excel.Worksheet
        IL_004a:  stloc.2
        IL_004b:  ldloc.2
        IL_004c:  brtrue.s   IL_0051

        IL_004e:  ldnull
        IL_004f:  br.s       IL_0057

        IL_0051:  ldloc.2
        IL_0052:  callvirt   instance class Microsoft.Office.Interop.Excel.Comments Microsoft.Office.Interop.Excel._Worksheet::get_Comments()
        IL_0057:  stloc.3
        IL_0058:  ldloc.3
        IL_0059:  brtrue.s   IL_0067

        IL_005b:  ldloca.s   V_5
        IL_005d:  initobj    valuetype [mscorlib]System.Nullable`1<int32>
        IL_0063:  ldloc.s    V_5
        IL_0065:  br.s       IL_0072

        IL_0067:  ldloc.3
        IL_0068:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Comments::get_Count()
        IL_006d:  newobj     instance void valuetype [mscorlib]System.Nullable`1<int32>::.ctor(!0)
        IL_0072:  stloc.s    V_4
        IL_0074:  ldloca.s   V_4
        IL_0076:  call       instance !0 valuetype [mscorlib]System.Nullable`1<int32>::GetValueOrDefault()
        IL_007b:  brfalse    IL_0156

        IL_0080:  ldc.i4.0
        IL_0081:  stloc.s    V_6
        IL_0083:  br         IL_013a

        IL_0088:  newobj     instance void iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass45_0'::.ctor()
        IL_008d:  stloc.s    V_7
        IL_008f:  ldloc.3
        IL_0090:  ldloc.s    V_6
        IL_0092:  ldc.i4.1
        IL_0093:  add
        IL_0094:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Comments::get__Default(int32)
        IL_0099:  stloc.s    V_8
        IL_009b:  ldloc.s    V_8
        IL_009d:  callvirt   instance object Microsoft.Office.Interop.Excel.Comment::get_Parent()
        IL_00a2:  isinst     Microsoft.Office.Interop.Excel.Range
        IL_00a7:  stloc.s    V_9
        IL_00a9:  ldloc.s    V_9
        IL_00ab:  brfalse    IL_0134

        IL_00b0:  ldloc.s    V_7
        IL_00b2:  ldarg.0
        IL_00b3:  ldloc.s    V_9
        IL_00b5:  ldloc.s    V_8
        IL_00b7:  call       instance class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem iReporterExcelAddIn2019.ThisAddIn::CreateTableItem(class Microsoft.Office.Interop.Excel.Range,
                                                                                                                                                                     class Microsoft.Office.Interop.Excel.Comment)
        IL_00bc:  stfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass45_0'::item
        IL_00c1:  ldloc.s    V_7
        IL_00c3:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass45_0'::item
        IL_00c8:  callvirt   instance string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem::get_TableNo()
        IL_00cd:  ldloca.s   V_10
        IL_00cf:  call       bool [mscorlib]System.Int32::TryParse(string,
                                                                   int32&)
        IL_00d4:  pop
        IL_00d5:  ldloc.s    V_10
        IL_00d7:  ldc.i4.1
        IL_00d8:  blt.s      IL_0134

        IL_00da:  ldloc.0
        IL_00db:  ldloc.s    V_7
        IL_00dd:  ldftn      instance bool iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass45_0'::'<ShowTableForm>b__1'(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table)
        IL_00e3:  newobj     instance void class [mscorlib]System.Predicate`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table>::.ctor(object,
                                                                                                                                                              native int)
        IL_00e8:  callvirt   instance !0 class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table>::Find(class [mscorlib]System.Predicate`1<!0>)
        IL_00ed:  stloc.s    V_11
        IL_00ef:  ldloc.s    V_11
        IL_00f1:  brtrue.s   IL_0112

        IL_00f3:  ldloc.2
        IL_00f4:  callvirt   instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
        IL_00f9:  ldloc.s    V_7
        IL_00fb:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass45_0'::item
        IL_0100:  ldc.i4.1
        IL_0101:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::.ctor(string,
                                                                                                                    class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem,
                                                                                                                    bool)
        IL_0106:  stloc.s    V_11
        IL_0108:  ldloc.0
        IL_0109:  ldloc.s    V_11
        IL_010b:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table>::Add(!0)
        IL_0110:  br.s       IL_0134

        IL_0112:  ldloc.s    V_11
        IL_0114:  callvirt   instance string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::get_SheetName()
        IL_0119:  ldloc.2
        IL_011a:  callvirt   instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
        IL_011f:  call       bool [mscorlib]System.String::op_Inequality(string,
                                                                         string)
        IL_0124:  brtrue.s   IL_0134

        IL_0126:  ldloc.s    V_11
        IL_0128:  ldloc.s    V_7
        IL_012a:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass45_0'::item
        IL_012f:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::Add(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem)
        IL_0134:  ldloc.s    V_6
        IL_0136:  ldc.i4.1
        IL_0137:  add
        IL_0138:  stloc.s    V_6
        IL_013a:  ldloc.s    V_6
        IL_013c:  ldloc.s    V_4
        IL_013e:  stloc.s    V_5
        IL_0140:  ldloca.s   V_5
        IL_0142:  call       instance !0 valuetype [mscorlib]System.Nullable`1<int32>::GetValueOrDefault()
        IL_0147:  clt
        IL_0149:  ldloca.s   V_5
        IL_014b:  call       instance bool valuetype [mscorlib]System.Nullable`1<int32>::get_HasValue()
        IL_0150:  and
        IL_0151:  brtrue     IL_0088

        IL_0156:  ldloc.1
        IL_0157:  callvirt   instance bool [mscorlib]System.Collections.IEnumerator::MoveNext()
        IL_015c:  brtrue     IL_003f

        IL_0161:  leave.s    IL_0177

      }  // end .try
      finally
      {
        IL_0163:  ldloc.1
        IL_0164:  isinst     [mscorlib]System.IDisposable
        IL_0169:  stloc.s    V_12
        IL_016b:  ldloc.s    V_12
        IL_016d:  brfalse.s  IL_0176

        IL_016f:  ldloc.s    V_12
        IL_0171:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
        IL_0176:  endfinally
      }  // end handler
      IL_0177:  ldloc.0
      IL_0178:  callvirt   instance valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<!0> class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table>::GetEnumerator()
      IL_017d:  stloc.s    V_13
      .try
      {
        IL_017f:  br.s       IL_018d

        IL_0181:  ldloca.s   V_13
        IL_0183:  call       instance !0 valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table>::get_Current()
        IL_0188:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::Adjust()
        IL_018d:  ldloca.s   V_13
        IL_018f:  call       instance bool valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table>::MoveNext()
        IL_0194:  brtrue.s   IL_0181

        IL_0196:  leave.s    IL_01a6

      }  // end .try
      finally
      {
        IL_0198:  ldloca.s   V_13
        IL_019a:  constrained. valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table>
        IL_01a0:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
        IL_01a5:  endfinally
      }  // end handler
      IL_01a6:  ldloc.0
      IL_01a7:  ldsfld     class [mscorlib]System.Comparison`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__45_0'
      IL_01ac:  dup
      IL_01ad:  brtrue.s   IL_01c6

      IL_01af:  pop
      IL_01b0:  ldsfld     class iReporterExcelAddIn2019.ThisAddIn/'<>c' iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9'
      IL_01b5:  ldftn      instance int32 iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<ShowTableForm>b__45_0'(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table,
                                                                                                            class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table)
      IL_01bb:  newobj     instance void class [mscorlib]System.Comparison`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table>::.ctor(object,
                                                                                                                                                             native int)
      IL_01c0:  dup
      IL_01c1:  stsfld     class [mscorlib]System.Comparison`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__45_0'
      IL_01c6:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table>::Sort(class [mscorlib]System.Comparison`1<!0>)
      IL_01cb:  ldarg.0
      IL_01cc:  ldfld      object iReporterExcelAddIn2019.ThisAddIn::_inputLocker
      IL_01d1:  stloc.s    V_14
      IL_01d3:  ldc.i4.0
      IL_01d4:  stloc.s    V_15
      .try
      {
        IL_01d6:  ldloc.s    V_14
        IL_01d8:  ldloca.s   V_15
        IL_01da:  call       void [mscorlib]System.Threading.Monitor::Enter(object,
                                                                            bool&)
        IL_01df:  ldarg.0
        IL_01e0:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm iReporterExcelAddIn2019.ThisAddIn::tableForm
        IL_01e5:  dup
        IL_01e6:  brtrue.s   IL_01ec

        IL_01e8:  pop
        IL_01e9:  ldc.i4.0
        IL_01ea:  br.s       IL_01f4

        IL_01ec:  call       instance bool [System.Windows.Forms]System.Windows.Forms.Control::get_IsDisposed()
        IL_01f1:  ldc.i4.0
        IL_01f2:  ceq
        IL_01f4:  brfalse.s  IL_0212

        IL_01f6:  ldarg.0
        IL_01f7:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm iReporterExcelAddIn2019.ThisAddIn::tableForm
        IL_01fc:  ldloc.0
        IL_01fd:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm::SetupTableList(class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table>)
        IL_0202:  ldarg.0
        IL_0203:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm iReporterExcelAddIn2019.ThisAddIn::tableForm
        IL_0208:  callvirt   instance void [System.Windows.Forms]System.Windows.Forms.Control::Show()
        IL_020d:  leave      IL_029c

        IL_0212:  ldarg.0
        IL_0213:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm::.ctor()
        IL_0218:  stfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm iReporterExcelAddIn2019.ThisAddIn::tableForm
        IL_021d:  ldarg.0
        IL_021e:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm iReporterExcelAddIn2019.ThisAddIn::tableForm
        IL_0223:  ldloc.0
        IL_0224:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm::SetupTableList(class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table>)
        IL_0229:  ldarg.0
        IL_022a:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm iReporterExcelAddIn2019.ThisAddIn::tableForm
        IL_022f:  ldarg.0
        IL_0230:  ldftn      instance class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table iReporterExcelAddIn2019.ThisAddIn::ReadTables()
        IL_0236:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm/ReadTableEventHandler::.ctor(object,
                                                                                                                                       native int)
        IL_023b:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm::add_ReadTable(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm/ReadTableEventHandler)
        IL_0240:  ldarg.0
        IL_0241:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm iReporterExcelAddIn2019.ThisAddIn::tableForm
        IL_0246:  ldarg.0
        IL_0247:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::WriteTable(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table)
        IL_024d:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm/WriteTableEventHandler::.ctor(object,
                                                                                                                                        native int)
        IL_0252:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm::add_WriteTable(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm/WriteTableEventHandler)
        IL_0257:  ldarg.0
        IL_0258:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm iReporterExcelAddIn2019.ThisAddIn::tableForm
        IL_025d:  ldarg.0
        IL_025e:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::'<ShowTableForm>b__45_2'(object,
                                                                                                       class [mscorlib]System.EventArgs)
        IL_0264:  newobj     instance void [mscorlib]System.EventHandler::.ctor(object,
                                                                                native int)
        IL_0269:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm::add_Reload(class [mscorlib]System.EventHandler)
        IL_026e:  ldarg.0
        IL_026f:  ldarg.0
        IL_0270:  ldfld      class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.TableForm iReporterExcelAddIn2019.ThisAddIn::tableForm
        IL_0275:  ldloca.s   V_16
        IL_0277:  initobj    valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point>
        IL_027d:  ldloc.s    V_16
        IL_027f:  ldnull
        IL_0280:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ShowForm(class [System.Windows.Forms]System.Windows.Forms.Form,
                                                                                       valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Point>,
                                                                                       class [System.Windows.Forms]System.Windows.Forms.IWin32Window)
        IL_0285:  leave.s    IL_0293

      }  // end .try
      finally
      {
        IL_0287:  ldloc.s    V_15
        IL_0289:  brfalse.s  IL_0292

        IL_028b:  ldloc.s    V_14
        IL_028d:  call       void [mscorlib]System.Threading.Monitor::Exit(object)
        IL_0292:  endfinally
      }  // end handler
      IL_0293:  leave.s    IL_029c

    }  // end .try
    catch [mscorlib]System.Exception 
    {
      IL_0295:  call       void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Util::Catched(class [mscorlib]System.Exception)
      IL_029a:  leave.s    IL_029c

    }  // end handler
    IL_029c:  ret
  } // end of method ThisAddIn::ShowTableForm

  .method private hidebysig instance class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem 
          CreateTableItem(class Microsoft.Office.Interop.Excel.Range cell,
                          class Microsoft.Office.Interop.Excel.Comment cm) cil managed
  {
    // Code size       267 (0x10b)
    .maxstack  11
    .locals init (int32 V_0,
             int32 V_1,
             int32 V_2,
             int32 V_3,
             string V_4,
             object V_5)
    IL_0000:  ldarg.1
    IL_0001:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_MergeArea()
    IL_0006:  dup
    IL_0007:  brtrue.s   IL_000b

    IL_0009:  pop
    IL_000a:  ldarg.1
    IL_000b:  starg.s    cell
    IL_000d:  ldarg.1
    IL_000e:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Column()
    IL_0013:  stloc.0
    IL_0014:  ldarg.1
    IL_0015:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Row()
    IL_001a:  stloc.1
    IL_001b:  ldloc.0
    IL_001c:  ldarg.1
    IL_001d:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_Columns()
    IL_0022:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Count()
    IL_0027:  add
    IL_0028:  ldc.i4.1
    IL_0029:  sub
    IL_002a:  stloc.2
    IL_002b:  ldloc.1
    IL_002c:  ldarg.1
    IL_002d:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_Rows()
    IL_0032:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Count()
    IL_0037:  add
    IL_0038:  ldc.i4.1
    IL_0039:  sub
    IL_003a:  stloc.3
    IL_003b:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> iReporterExcelAddIn2019.ThisAddIn/'<>o__46'::'<>p__1'
    IL_0040:  brtrue.s   IL_0066

    IL_0042:  ldc.i4.0
    IL_0043:  ldtoken    [mscorlib]System.String
    IL_0048:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_004d:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
    IL_0052:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_0057:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::Convert(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                            class [mscorlib]System.Type,
                                                                                                                                                            class [mscorlib]System.Type)
    IL_005c:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
    IL_0061:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> iReporterExcelAddIn2019.ThisAddIn/'<>o__46'::'<>p__1'
    IL_0066:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> iReporterExcelAddIn2019.ThisAddIn/'<>o__46'::'<>p__1'
    IL_006b:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>>::Target
    IL_0070:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> iReporterExcelAddIn2019.ThisAddIn/'<>o__46'::'<>p__1'
    IL_0075:  ldarg.1
    IL_0076:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get_NumberFormatLocal()
    IL_007b:  stloc.s    V_5
    IL_007d:  ldloc.s    V_5
    IL_007f:  brtrue.s   IL_0084

    IL_0081:  ldnull
    IL_0082:  br.s       IL_00d1

    IL_0084:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__46'::'<>p__0'
    IL_0089:  brtrue.s   IL_00bb

    IL_008b:  ldc.i4.0
    IL_008c:  ldstr      "ToString"
    IL_0091:  ldnull
    IL_0092:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
    IL_0097:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_009c:  ldc.i4.1
    IL_009d:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
    IL_00a2:  dup
    IL_00a3:  ldc.i4.0
    IL_00a4:  ldc.i4.0
    IL_00a5:  ldnull
    IL_00a6:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                               string)
    IL_00ab:  stelem.ref
    IL_00ac:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                 string,
                                                                                                                                                                 class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                 class [mscorlib]System.Type,
                                                                                                                                                                 class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
    IL_00b1:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
    IL_00b6:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__46'::'<>p__0'
    IL_00bb:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__46'::'<>p__0'
    IL_00c0:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Target
    IL_00c5:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__46'::'<>p__0'
    IL_00ca:  ldloc.s    V_5
    IL_00cc:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>::Invoke(!0,
                                                                                                                                                      !1)
    IL_00d1:  dup
    IL_00d2:  brtrue.s   IL_00da

    IL_00d4:  pop
    IL_00d5:  ldstr      ""
    IL_00da:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>::Invoke(!0,
                                                                                                                                                      !1)
    IL_00df:  stloc.s    V_4
    IL_00e1:  ldloc.0
    IL_00e2:  ldloc.1
    IL_00e3:  ldloc.2
    IL_00e4:  ldloc.3
    IL_00e5:  ldarg.2
    IL_00e6:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00eb:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00f0:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00f5:  callvirt   instance string Microsoft.Office.Interop.Excel.Comment::Text(object,
                                                                                      object,
                                                                                      object)
    IL_00fa:  dup
    IL_00fb:  brtrue.s   IL_0103

    IL_00fd:  pop
    IL_00fe:  ldstr      ""
    IL_0103:  ldloc.s    V_4
    IL_0105:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem::.ctor(int32,
                                                                                                                    int32,
                                                                                                                    int32,
                                                                                                                    int32,
                                                                                                                    string,
                                                                                                                    string)
    IL_010a:  ret
  } // end of method ThisAddIn::CreateTableItem

  .method private hidebysig instance class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table 
          ReadTables() cil managed
  {
    // Code size       1222 (0x4c6)
    .maxstack  31
    .locals init (class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table V_0,
             class Microsoft.Office.Interop.Excel.Worksheet V_1,
             class Microsoft.Office.Interop.Excel.Comments V_2,
             int32 V_3,
             class Microsoft.Office.Interop.Excel.Range V_4,
             class Microsoft.Office.Interop.Excel.Workbook V_5,
             int32 V_6,
             class Microsoft.Office.Interop.Excel.Comment V_7,
             class Microsoft.Office.Interop.Excel.Range V_8,
             class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem V_9,
             class [mscorlib]System.Collections.Generic.IEnumerator`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableCol> V_10,
             class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableCol V_11,
             class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem V_12,
             class Microsoft.Office.Interop.Excel.Range V_13,
             int32 V_14,
             class Microsoft.Office.Interop.Excel.Range V_15,
             object V_16,
             class [mscorlib]System.Collections.Generic.IEnumerator`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableRow> V_17,
             class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableRow V_18,
             class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem V_19,
             class Microsoft.Office.Interop.Excel.Range V_20,
             class Microsoft.Office.Interop.Excel.Range V_21)
    IL_0000:  ldnull
    IL_0001:  stloc.0
    IL_0002:  ldarg.0
    IL_0003:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0008:  callvirt   instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
    IL_000d:  stloc.s    V_5
    IL_000f:  ldloc.s    V_5
    IL_0011:  brtrue.s   IL_0016

    IL_0013:  ldnull
    IL_0014:  br.s       IL_001d

    IL_0016:  ldloc.s    V_5
    IL_0018:  callvirt   instance object Microsoft.Office.Interop.Excel._Workbook::get_ActiveSheet()
    IL_001d:  isinst     Microsoft.Office.Interop.Excel.Worksheet
    IL_0022:  stloc.1
    IL_0023:  ldloc.1
    IL_0024:  brtrue.s   IL_0029

    IL_0026:  ldnull
    IL_0027:  br.s       IL_002f

    IL_0029:  ldloc.1
    IL_002a:  callvirt   instance class Microsoft.Office.Interop.Excel.Comments Microsoft.Office.Interop.Excel._Worksheet::get_Comments()
    IL_002f:  stloc.2
    IL_0030:  ldloc.2
    IL_0031:  brtrue.s   IL_0036

    IL_0033:  ldc.i4.0
    IL_0034:  br.s       IL_003c

    IL_0036:  ldloc.2
    IL_0037:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Comments::get_Count()
    IL_003c:  stloc.3
    IL_003d:  ldloc.3
    IL_003e:  brfalse.s  IL_004a

    IL_0040:  ldarg.0
    IL_0041:  ldloca.s   V_4
    IL_0043:  call       instance bool iReporterExcelAddIn2019.ThisAddIn::GetSelected(class Microsoft.Office.Interop.Excel.Range&)
    IL_0048:  brtrue.s   IL_004c

    IL_004a:  ldloc.0
    IL_004b:  ret

    IL_004c:  ldc.i4.0
    IL_004d:  stloc.s    V_6
    IL_004f:  br         IL_0140

    IL_0054:  ldloc.2
    IL_0055:  ldloc.s    V_6
    IL_0057:  ldc.i4.1
    IL_0058:  add
    IL_0059:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Comments::get__Default(int32)
    IL_005e:  stloc.s    V_7
    IL_0060:  ldloc.s    V_7
    IL_0062:  callvirt   instance object Microsoft.Office.Interop.Excel.Comment::get_Parent()
    IL_0067:  isinst     Microsoft.Office.Interop.Excel.Range
    IL_006c:  stloc.s    V_8
    IL_006e:  ldloc.s    V_8
    IL_0070:  brfalse    IL_013a

    IL_0075:  ldarg.0
    IL_0076:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_007b:  ldloc.s    V_4
    IL_007d:  ldloc.s    V_8
    IL_007f:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0084:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0089:  ldsfld     object [mscorlib]System.Type::Missing
    IL_008e:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0093:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0098:  ldsfld     object [mscorlib]System.Type::Missing
    IL_009d:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00a2:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00a7:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00ac:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00b1:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00b6:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00bb:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00c0:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00c5:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00ca:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00cf:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00d4:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00d9:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00de:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00e3:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00e8:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00ed:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00f2:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00f7:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00fc:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0101:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0106:  ldsfld     object [mscorlib]System.Type::Missing
    IL_010b:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Application::Intersect(class Microsoft.Office.Interop.Excel.Range,
                                                                                                                                    class Microsoft.Office.Interop.Excel.Range,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object,
                                                                                                                                    object)
    IL_0110:  brfalse.s  IL_013a

    IL_0112:  ldarg.0
    IL_0113:  ldloc.s    V_8
    IL_0115:  ldloc.s    V_7
    IL_0117:  call       instance class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem iReporterExcelAddIn2019.ThisAddIn::CreateTableItem(class Microsoft.Office.Interop.Excel.Range,
                                                                                                                                                                 class Microsoft.Office.Interop.Excel.Comment)
    IL_011c:  stloc.s    V_9
    IL_011e:  ldloc.0
    IL_011f:  brtrue.s   IL_0132

    IL_0121:  ldloc.1
    IL_0122:  callvirt   instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
    IL_0127:  ldloc.s    V_9
    IL_0129:  ldc.i4.0
    IL_012a:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::.ctor(string,
                                                                                                                class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem,
                                                                                                                bool)
    IL_012f:  stloc.0
    IL_0130:  br.s       IL_013a

    IL_0132:  ldloc.0
    IL_0133:  ldloc.s    V_9
    IL_0135:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::Add(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem)
    IL_013a:  ldloc.s    V_6
    IL_013c:  ldc.i4.1
    IL_013d:  add
    IL_013e:  stloc.s    V_6
    IL_0140:  ldloc.s    V_6
    IL_0142:  ldloc.3
    IL_0143:  blt        IL_0054

    IL_0148:  ldloc.0
    IL_0149:  brfalse    IL_04c4

    IL_014e:  ldloc.0
    IL_014f:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::Adjust()
    IL_0154:  ldloc.0
    IL_0155:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableCol> [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::get_Cols()
    IL_015a:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerator`1<!0> class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableCol>::GetEnumerator()
    IL_015f:  stloc.s    V_10
    .try
    {
      IL_0161:  br         IL_030c

      IL_0166:  ldloc.s    V_10
      IL_0168:  callvirt   instance !0 class [mscorlib]System.Collections.Generic.IEnumerator`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableCol>::get_Current()
      IL_016d:  stloc.s    V_11
      IL_016f:  ldloc.s    V_11
      IL_0171:  callvirt   instance string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableCol::get_Key()
      IL_0176:  call       bool [mscorlib]System.String::IsNullOrEmpty(string)
      IL_017b:  brfalse.s  IL_01a3

      IL_017d:  ldloc.s    V_11
      IL_017f:  ldstr      "F_"
      IL_0184:  ldloc.s    V_11
      IL_0186:  callvirt   instance int32 [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableLine::get_Index()
      IL_018b:  stloc.s    V_14
      IL_018d:  ldloca.s   V_14
      IL_018f:  ldstr      "000"
      IL_0194:  call       instance string [mscorlib]System.Int32::ToString(string)
      IL_0199:  call       string [mscorlib]System.String::Concat(string,
                                                                  string)
      IL_019e:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableCol::set_Key(string)
      IL_01a3:  ldloc.s    V_11
      IL_01a5:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem> [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableLine::get_Items()
      IL_01aa:  call       !!0 [System.Core]System.Linq.Enumerable::First<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
      IL_01af:  stloc.s    V_12
      IL_01b1:  ldc.i4.1
      IL_01b2:  ldloc.s    V_12
      IL_01b4:  callvirt   instance int32 [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem::get_Y()
      IL_01b9:  bge        IL_030c

      IL_01be:  ldloc.1
      IL_01bf:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Cells()
      IL_01c4:  ldloc.s    V_12
      IL_01c6:  callvirt   instance int32 [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem::get_Y()
      IL_01cb:  ldc.i4.1
      IL_01cc:  sub
      IL_01cd:  box        [mscorlib]System.Int32
      IL_01d2:  ldloc.s    V_12
      IL_01d4:  callvirt   instance int32 [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem::get_X()
      IL_01d9:  box        [mscorlib]System.Int32
      IL_01de:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get__Default(object,
                                                                                              object)
      IL_01e3:  isinst     Microsoft.Office.Interop.Excel.Range
      IL_01e8:  dup
      IL_01e9:  stloc.s    V_13
      IL_01eb:  brfalse    IL_030c

      IL_01f0:  ldloc.s    V_13
      IL_01f2:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_MergeArea()
      IL_01f7:  dup
      IL_01f8:  stloc.s    V_15
      IL_01fa:  brfalse.s  IL_0260

      IL_01fc:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__0'
      IL_0201:  brtrue.s   IL_0227

      IL_0203:  ldc.i4.0
      IL_0204:  ldtoken    Microsoft.Office.Interop.Excel.Range
      IL_0209:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_020e:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_0213:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0218:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::Convert(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                              class [mscorlib]System.Type,
                                                                                                                                                              class [mscorlib]System.Type)
      IL_021d:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_0222:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__0'
      IL_0227:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__0'
      IL_022c:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>>::Target
      IL_0231:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__0'
      IL_0236:  ldloc.1
      IL_0237:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Cells()
      IL_023c:  ldloc.s    V_15
      IL_023e:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Row()
      IL_0243:  box        [mscorlib]System.Int32
      IL_0248:  ldloc.s    V_15
      IL_024a:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Column()
      IL_024f:  box        [mscorlib]System.Int32
      IL_0254:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get__Default(object,
                                                                                              object)
      IL_0259:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>::Invoke(!0,
                                                                                                                                                                                            !1)
      IL_025e:  stloc.s    V_13
      IL_0260:  ldloc.s    V_11
      IL_0262:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__2'
      IL_0267:  brtrue.s   IL_028d

      IL_0269:  ldc.i4.0
      IL_026a:  ldtoken    [mscorlib]System.String
      IL_026f:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0274:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_0279:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_027e:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::Convert(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                              class [mscorlib]System.Type,
                                                                                                                                                              class [mscorlib]System.Type)
      IL_0283:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_0288:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__2'
      IL_028d:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__2'
      IL_0292:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>>::Target
      IL_0297:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__2'
      IL_029c:  ldloc.s    V_13
      IL_029e:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get_Text()
      IL_02a3:  stloc.s    V_16
      IL_02a5:  ldloc.s    V_16
      IL_02a7:  brtrue.s   IL_02ac

      IL_02a9:  ldnull
      IL_02aa:  br.s       IL_02f9

      IL_02ac:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__1'
      IL_02b1:  brtrue.s   IL_02e3

      IL_02b3:  ldc.i4.0
      IL_02b4:  ldstr      "ToString"
      IL_02b9:  ldnull
      IL_02ba:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_02bf:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_02c4:  ldc.i4.1
      IL_02c5:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_02ca:  dup
      IL_02cb:  ldc.i4.0
      IL_02cc:  ldc.i4.0
      IL_02cd:  ldnull
      IL_02ce:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_02d3:  stelem.ref
      IL_02d4:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   string,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_02d9:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_02de:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__1'
      IL_02e3:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__1'
      IL_02e8:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Target
      IL_02ed:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__1'
      IL_02f2:  ldloc.s    V_16
      IL_02f4:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>::Invoke(!0,
                                                                                                                                                        !1)
      IL_02f9:  dup
      IL_02fa:  brtrue.s   IL_0302

      IL_02fc:  pop
      IL_02fd:  ldstr      ""
      IL_0302:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>::Invoke(!0,
                                                                                                                                                        !1)
      IL_0307:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableCol::set_Name(string)
      IL_030c:  ldloc.s    V_10
      IL_030e:  callvirt   instance bool [mscorlib]System.Collections.IEnumerator::MoveNext()
      IL_0313:  brtrue     IL_0166

      IL_0318:  leave.s    IL_0326

    }  // end .try
    finally
    {
      IL_031a:  ldloc.s    V_10
      IL_031c:  brfalse.s  IL_0325

      IL_031e:  ldloc.s    V_10
      IL_0320:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
      IL_0325:  endfinally
    }  // end handler
    IL_0326:  ldloc.0
    IL_0327:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableRow> [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.Table::get_Rows()
    IL_032c:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerator`1<!0> class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableRow>::GetEnumerator()
    IL_0331:  stloc.s    V_17
    .try
    {
      IL_0333:  br         IL_04aa

      IL_0338:  ldloc.s    V_17
      IL_033a:  callvirt   instance !0 class [mscorlib]System.Collections.Generic.IEnumerator`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableRow>::get_Current()
      IL_033f:  stloc.s    V_18
      IL_0341:  ldloc.s    V_18
      IL_0343:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem> [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableLine::get_Items()
      IL_0348:  call       !!0 [System.Core]System.Linq.Enumerable::First<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
      IL_034d:  stloc.s    V_19
      IL_034f:  ldc.i4.1
      IL_0350:  ldloc.s    V_19
      IL_0352:  callvirt   instance int32 [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem::get_X()
      IL_0357:  bge        IL_04aa

      IL_035c:  ldloc.1
      IL_035d:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Cells()
      IL_0362:  ldloc.s    V_19
      IL_0364:  callvirt   instance int32 [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem::get_Y()
      IL_0369:  box        [mscorlib]System.Int32
      IL_036e:  ldloc.s    V_19
      IL_0370:  callvirt   instance int32 [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableItem::get_X()
      IL_0375:  ldc.i4.1
      IL_0376:  sub
      IL_0377:  box        [mscorlib]System.Int32
      IL_037c:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get__Default(object,
                                                                                              object)
      IL_0381:  isinst     Microsoft.Office.Interop.Excel.Range
      IL_0386:  dup
      IL_0387:  stloc.s    V_20
      IL_0389:  brfalse    IL_04aa

      IL_038e:  ldloc.s    V_20
      IL_0390:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_MergeArea()
      IL_0395:  dup
      IL_0396:  stloc.s    V_21
      IL_0398:  brfalse.s  IL_03fe

      IL_039a:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__3'
      IL_039f:  brtrue.s   IL_03c5

      IL_03a1:  ldc.i4.0
      IL_03a2:  ldtoken    Microsoft.Office.Interop.Excel.Range
      IL_03a7:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_03ac:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_03b1:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_03b6:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::Convert(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                              class [mscorlib]System.Type,
                                                                                                                                                              class [mscorlib]System.Type)
      IL_03bb:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_03c0:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__3'
      IL_03c5:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__3'
      IL_03ca:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>>::Target
      IL_03cf:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__3'
      IL_03d4:  ldloc.1
      IL_03d5:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Cells()
      IL_03da:  ldloc.s    V_21
      IL_03dc:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Row()
      IL_03e1:  box        [mscorlib]System.Int32
      IL_03e6:  ldloc.s    V_21
      IL_03e8:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Column()
      IL_03ed:  box        [mscorlib]System.Int32
      IL_03f2:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get__Default(object,
                                                                                              object)
      IL_03f7:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>::Invoke(!0,
                                                                                                                                                                                            !1)
      IL_03fc:  stloc.s    V_20
      IL_03fe:  ldloc.s    V_18
      IL_0400:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__5'
      IL_0405:  brtrue.s   IL_042b

      IL_0407:  ldc.i4.0
      IL_0408:  ldtoken    [mscorlib]System.String
      IL_040d:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0412:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_0417:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_041c:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::Convert(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                              class [mscorlib]System.Type,
                                                                                                                                                              class [mscorlib]System.Type)
      IL_0421:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_0426:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__5'
      IL_042b:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__5'
      IL_0430:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>>::Target
      IL_0435:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__5'
      IL_043a:  ldloc.s    V_20
      IL_043c:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get_Text()
      IL_0441:  stloc.s    V_16
      IL_0443:  ldloc.s    V_16
      IL_0445:  brtrue.s   IL_044a

      IL_0447:  ldnull
      IL_0448:  br.s       IL_0497

      IL_044a:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__4'
      IL_044f:  brtrue.s   IL_0481

      IL_0451:  ldc.i4.0
      IL_0452:  ldstr      "ToString"
      IL_0457:  ldnull
      IL_0458:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_045d:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0462:  ldc.i4.1
      IL_0463:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_0468:  dup
      IL_0469:  ldc.i4.0
      IL_046a:  ldc.i4.0
      IL_046b:  ldnull
      IL_046c:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0471:  stelem.ref
      IL_0472:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   string,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_0477:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_047c:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__4'
      IL_0481:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__4'
      IL_0486:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Target
      IL_048b:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__47'::'<>p__4'
      IL_0490:  ldloc.s    V_16
      IL_0492:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>::Invoke(!0,
                                                                                                                                                        !1)
      IL_0497:  dup
      IL_0498:  brtrue.s   IL_04a0

      IL_049a:  pop
      IL_049b:  ldstr      ""
      IL_04a0:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>::Invoke(!0,
                                                                                                                                                        !1)
      IL_04a5:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Tables.TableRow::set_Name(string)
      IL_04aa:  ldloc.s    V_17
      IL_04ac:  callvirt   instance bool [mscorlib]System.Collections.IEnumerator::MoveNext()
      IL_04b1:  brtrue     IL_0338

      IL_04b6:  leave.s    IL_04c4

    }  // end .try
    finally
    {
      IL_04b8:  ldloc.s    V_17
      IL_04ba:  brfalse.s  IL_04c3

      IL_04bc:  ldloc.s    V_17
      IL_04be:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
      IL_04c3:  endfinally
    }  // end handler
    IL_04c4:  ldloc.0
    IL_04c5:  ret
  } // end of method ThisAddIn::ReadTables

  .method public hidebysig instance void 
          CheckAllSheetName() cil managed
  {
    // Code size       445 (0x1bd)
    .maxstack  13
    .locals init (class Microsoft.Office.Interop.Excel.Workbook V_0,
             class [mscorlib]System.Collections.Generic.List`1<string> V_1,
             class [mscorlib]System.Collections.Generic.List`1<int32> V_2,
             class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Models.SheetNameInvalidCharacter> V_3,
             string V_4,
             class [mscorlib]System.Collections.Generic.List`1<string> V_5,
             string V_6,
             string V_7,
             class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel> V_8,
             class [mscorlib]System.Collections.Generic.List`1<class [mscorlib]System.Collections.Generic.List`1<string>> V_9,
             string V_10,
             class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Services.NotificationService V_11,
             class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.MessageGridViewModel V_12,
             class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Views.MessageGrid V_13,
             class [PresentationFramework]System.Windows.Interop.WindowInteropHelper V_14,
             native int V_15)
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0006:  callvirt   instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
    IL_000b:  stloc.0
    IL_000c:  ldloc.0
    IL_000d:  brtrue.s   IL_0010

    IL_000f:  ret

    IL_0010:  ldarg.0
    IL_0011:  ldloc.0
    IL_0012:  call       instance class [mscorlib]System.Collections.Generic.IEnumerable`1<string> iReporterExcelAddIn2019.ThisAddIn::GetSheetNames(class Microsoft.Office.Interop.Excel.Workbook)
    IL_0017:  call       class [mscorlib]System.Collections.Generic.List`1<!!0> [System.Core]System.Linq.Enumerable::ToList<string>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
    IL_001c:  stloc.1
    IL_001d:  ldloc.1
    IL_001e:  ldloca.s   V_2
    IL_0020:  call       bool [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Common.WorksheetChecker::ValidateSheetNamesNotBlank(class [mscorlib]System.Collections.Generic.IEnumerable`1<string>,
                                                                                                                                       class [mscorlib]System.Collections.Generic.List`1<int32>&)
    IL_0025:  brtrue.s   IL_0050

    IL_0027:  ldstr      ", "
    IL_002c:  ldloc.2
    IL_002d:  call       string [mscorlib]System.String::Join<int32>(string,
                                                                     class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
    IL_0032:  stloc.s    V_6
    IL_0034:  call       string iReporterExcelAddIn2019.Properties.Resources::get_ExcelMessage65()
    IL_0039:  call       string iReporterExcelAddIn2019.Properties.Resources::get_ExcelMessage64()
    IL_003e:  stloc.s    V_7
    IL_0040:  ldloc.s    V_6
    IL_0042:  ldloc.s    V_7
    IL_0044:  call       string [mscorlib]System.String::Format(string,
                                                                object,
                                                                object)
    IL_0049:  call       valuetype [System.Windows.Forms]System.Windows.Forms.DialogResult [System.Windows.Forms]System.Windows.Forms.MessageBox::Show(string)
    IL_004e:  pop
    IL_004f:  ret

    IL_0050:  call       string [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Helpers.CommonObjects::get_Language()
    IL_0055:  call       class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Models.SheetNameInvalidCharacter> [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Utilities.InvalidCharacterLoader::Load(string)
    IL_005a:  stloc.3
    IL_005b:  call       string iReporterExcelAddIn2019.Properties.Resources::get_ExcelProcessorMessage5()
    IL_0060:  stloc.s    V_4
    IL_0062:  ldloc.1
    IL_0063:  ldloc.3
    IL_0064:  ldloc.s    V_4
    IL_0066:  call       class [mscorlib]System.Collections.Generic.List`1<string> [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Common.WorksheetChecker::ValidateSheetNames(class [mscorlib]System.Collections.Generic.IEnumerable`1<string>,
                                                                                                                                                                                    class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Models.SheetNameInvalidCharacter>,
                                                                                                                                                                                    string)
    IL_006b:  stloc.s    V_5
    IL_006d:  ldloc.s    V_5
    IL_006f:  callvirt   instance int32 class [mscorlib]System.Collections.Generic.List`1<string>::get_Count()
    IL_0074:  ldc.i4.0
    IL_0075:  ble        IL_01a9

    IL_007a:  newobj     instance void class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel>::.ctor()
    IL_007f:  dup
    IL_0080:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel::.ctor()
    IL_0085:  dup
    IL_0086:  call       string iReporterExcelAddIn2019.Properties.Resources::get_SheetNameNoticeGridColumnSymbolHeader()
    IL_008b:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel::set_Header(string)
    IL_0090:  dup
    IL_0091:  ldstr      "Symbol"
    IL_0096:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel::set_PropertyName(string)
    IL_009b:  dup
    IL_009c:  ldc.i4.s   60
    IL_009e:  conv.r8
    IL_009f:  newobj     instance void valuetype [mscorlib]System.Nullable`1<float64>::.ctor(!0)
    IL_00a4:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel::set_Width(valuetype [mscorlib]System.Nullable`1<float64>)
    IL_00a9:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel>::Add(!0)
    IL_00ae:  dup
    IL_00af:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel::.ctor()
    IL_00b4:  dup
    IL_00b5:  call       string iReporterExcelAddIn2019.Properties.Resources::get_SheetNameNoticeGridColumnNameHeader()
    IL_00ba:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel::set_Header(string)
    IL_00bf:  dup
    IL_00c0:  ldstr      "SymbolName"
    IL_00c5:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel::set_PropertyName(string)
    IL_00ca:  dup
    IL_00cb:  ldc.i4     0x1a4
    IL_00d0:  conv.r8
    IL_00d1:  newobj     instance void valuetype [mscorlib]System.Nullable`1<float64>::.ctor(!0)
    IL_00d6:  callvirt   instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel::set_Width(valuetype [mscorlib]System.Nullable`1<float64>)
    IL_00db:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel>::Add(!0)
    IL_00e0:  stloc.s    V_8
    IL_00e2:  ldloc.3
    IL_00e3:  ldsfld     class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Models.SheetNameInvalidCharacter,class [mscorlib]System.Collections.Generic.List`1<string>> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__48_0'
    IL_00e8:  dup
    IL_00e9:  brtrue.s   IL_0102

    IL_00eb:  pop
    IL_00ec:  ldsfld     class iReporterExcelAddIn2019.ThisAddIn/'<>c' iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9'
    IL_00f1:  ldftn      instance class [mscorlib]System.Collections.Generic.List`1<string> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<CheckAllSheetName>b__48_0'(class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Models.SheetNameInvalidCharacter)
    IL_00f7:  newobj     instance void class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Models.SheetNameInvalidCharacter,class [mscorlib]System.Collections.Generic.List`1<string>>::.ctor(object,
                                                                                                                                                                                                                                          native int)
    IL_00fc:  dup
    IL_00fd:  stsfld     class [mscorlib]System.Func`2<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Models.SheetNameInvalidCharacter,class [mscorlib]System.Collections.Generic.List`1<string>> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__48_0'
    IL_0102:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!1> [System.Core]System.Linq.Enumerable::Select<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Domain.Models.SheetNameInvalidCharacter,class [mscorlib]System.Collections.Generic.List`1<string>>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                                                                                                                                                                                 class [mscorlib]System.Func`2<!!0,!!1>)
    IL_0107:  call       class [mscorlib]System.Collections.Generic.List`1<!!0> [System.Core]System.Linq.Enumerable::ToList<class [mscorlib]System.Collections.Generic.List`1<string>>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
    IL_010c:  stloc.s    V_9
    IL_010e:  call       string [mscorlib]System.Environment::get_NewLine()
    IL_0113:  ldloc.s    V_5
    IL_0115:  call       string [mscorlib]System.String::Join(string,
                                                              class [mscorlib]System.Collections.Generic.IEnumerable`1<string>)
    IL_011a:  stloc.s    V_10
    IL_011c:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Services.NotificationService::.ctor()
    IL_0121:  stloc.s    V_11
    IL_0123:  call       string iReporterExcelAddIn2019.Properties.Resources::get_SheetNameNoticeMessageDesc()
    IL_0128:  call       string iReporterExcelAddIn2019.Properties.Resources::get_SheetNameNoticeGridDesc()
    IL_012d:  ldloc.s    V_10
    IL_012f:  ldloc.s    V_8
    IL_0131:  ldloc.s    V_9
    IL_0133:  ldc.i4.2
    IL_0134:  ldloc.s    V_11
    IL_0136:  ldnull
    IL_0137:  ldc.i4.1
    IL_0138:  ldc.i4.0
    IL_0139:  call       string iReporterExcelAddIn2019.Properties.Resources::get_SheetNameNoticeWindow()
    IL_013e:  ldnull
    IL_013f:  ldnull
    IL_0140:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.MessageGridViewModel::.ctor(string,
                                                                                                                                   string,
                                                                                                                                   string,
                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.ColumnDefinitionModel>,
                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Collections.Generic.IEnumerable`1<string>>,
                                                                                                                                   valuetype [iReporterExcelAddInCommon]iReporterExcelAddInCommon.ViewModels.DisplayMode,
                                                                                                                                   class [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Services.INotificationService,
                                                                                                                                   class [mscorlib]System.Action,
                                                                                                                                   bool,
                                                                                                                                   bool,
                                                                                                                                   string,
                                                                                                                                   string,
                                                                                                                                   string)
    IL_0145:  stloc.s    V_12
    IL_0147:  newobj     instance void [iReporterExcelAddInCommon]iReporterExcelAddInCommon.Views.MessageGrid::.ctor()
    IL_014c:  dup
    IL_014d:  ldloc.s    V_12
    IL_014f:  callvirt   instance void [PresentationFramework]System.Windows.FrameworkElement::set_DataContext(object)
    IL_0154:  stloc.s    V_13
    IL_0156:  ldloc.s    V_13
    IL_0158:  newobj     instance void [PresentationFramework]System.Windows.Interop.WindowInteropHelper::.ctor(class [PresentationFramework]System.Windows.Window)
    IL_015d:  stloc.s    V_14
    IL_015f:  ldarg.0
    IL_0160:  call       instance native int iReporterExcelAddIn2019.ThisAddIn::GetExcelMainWindowHandle()
    IL_0165:  stloc.s    V_15
    IL_0167:  ldloc.s    V_15
    IL_0169:  ldsfld     native int [mscorlib]System.IntPtr::Zero
    IL_016e:  call       bool [mscorlib]System.IntPtr::op_Inequality(native int,
                                                                     native int)
    IL_0173:  brfalse.s  IL_0198

    IL_0175:  ldloc.s    V_14
    IL_0177:  ldloc.s    V_15
    IL_0179:  callvirt   instance void [PresentationFramework]System.Windows.Interop.WindowInteropHelper::set_Owner(native int)
    IL_017e:  ldloc.s    V_13
    IL_0180:  ldc.r8     620.
    IL_0189:  callvirt   instance void [PresentationFramework]System.Windows.FrameworkElement::set_Height(float64)
    IL_018e:  ldloc.s    V_13
    IL_0190:  ldc.i4.2
    IL_0191:  callvirt   instance void [PresentationFramework]System.Windows.Window::set_WindowStartupLocation(valuetype [PresentationFramework]System.Windows.WindowStartupLocation)
    IL_0196:  br.s       IL_01a0

    IL_0198:  ldloc.s    V_13
    IL_019a:  ldc.i4.1
    IL_019b:  callvirt   instance void [PresentationFramework]System.Windows.Window::set_WindowStartupLocation(valuetype [PresentationFramework]System.Windows.WindowStartupLocation)
    IL_01a0:  ldloc.s    V_13
    IL_01a2:  callvirt   instance valuetype [mscorlib]System.Nullable`1<bool> [PresentationFramework]System.Windows.Window::ShowDialog()
    IL_01a7:  pop
    IL_01a8:  ret

    IL_01a9:  call       string iReporterExcelAddIn2019.Properties.Resources::get_CheckSheetNameSuccess()
    IL_01ae:  call       string iReporterExcelAddIn2019.Properties.Resources::get_CheckSheetNameComplete()
    IL_01b3:  ldc.i4.0
    IL_01b4:  ldc.i4.s   64
    IL_01b6:  call       valuetype [System.Windows.Forms]System.Windows.Forms.DialogResult [System.Windows.Forms]System.Windows.Forms.MessageBox::Show(string,
                                                                                                                                                       string,
                                                                                                                                                       valuetype [System.Windows.Forms]System.Windows.Forms.MessageBoxButtons,
                                                                                                                                                       valuetype [System.Windows.Forms]System.Windows.Forms.MessageBoxIcon)
    IL_01bb:  pop
    IL_01bc:  ret
  } // end of method ThisAddIn::CheckAllSheetName

  .method public hidebysig instance native int 
          GetExcelMainWindowHandle() cil managed
  {
    // Code size       23 (0x17)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0006:  callvirt   instance int32 Microsoft.Office.Interop.Excel._Application::get_Hwnd()
    IL_000b:  newobj     instance void [mscorlib]System.IntPtr::.ctor(int32)
    IL_0010:  ldc.i4.2
    IL_0011:  call       native int iReporterExcelAddIn2019.ThisAddIn::GetAncestor(native int,
                                                                                   uint32)
    IL_0016:  ret
  } // end of method ThisAddIn::GetExcelMainWindowHandle

  .method public hidebysig instance class [mscorlib]System.Collections.Generic.List`1<int32> 
          GetReportIndex(class Microsoft.Office.Interop.Excel.Workbook workbook) cil managed
  {
    // Code size       115 (0x73)
    .maxstack  3
    .locals init (class [mscorlib]System.Collections.Generic.List`1<int32> V_0,
             int32 V_1,
             class [mscorlib]System.Collections.IEnumerator V_2,
             class Microsoft.Office.Interop.Excel.Worksheet V_3,
             bool V_4,
             class [mscorlib]System.IDisposable V_5)
    IL_0000:  newobj     instance void class [mscorlib]System.Collections.Generic.List`1<int32>::.ctor()
    IL_0005:  stloc.0
    IL_0006:  ldc.i4.1
    IL_0007:  stloc.1
    IL_0008:  ldarg.1
    IL_0009:  callvirt   instance class Microsoft.Office.Interop.Excel.Sheets Microsoft.Office.Interop.Excel._Workbook::get_Sheets()
    IL_000e:  callvirt   instance class [mscorlib]System.Collections.IEnumerator Microsoft.Office.Interop.Excel.Sheets::GetEnumerator()
    IL_0013:  stloc.2
    .try
    {
      IL_0014:  br.s       IL_0053

      IL_0016:  ldloc.2
      IL_0017:  callvirt   instance object [mscorlib]System.Collections.IEnumerator::get_Current()
      IL_001c:  castclass  Microsoft.Office.Interop.Excel.Worksheet
      IL_0021:  stloc.3
      .try
      {
        IL_0022:  ldloc.3
        IL_0023:  callvirt   instance valuetype Microsoft.Office.Interop.Excel.XlSheetVisibility Microsoft.Office.Interop.Excel._Worksheet::get_Visible()
        IL_0028:  ldc.i4.m1
        IL_0029:  ceq
        IL_002b:  ldloc.3
        IL_002c:  callvirt   instance class Microsoft.Office.Interop.Excel.PageSetup Microsoft.Office.Interop.Excel._Worksheet::get_PageSetup()
        IL_0031:  callvirt   instance string Microsoft.Office.Interop.Excel.PageSetup::get_PrintArea()
        IL_0036:  call       bool [mscorlib]System.String::IsNullOrWhiteSpace(string)
        IL_003b:  ldc.i4.0
        IL_003c:  ceq
        IL_003e:  stloc.s    V_4
        IL_0040:  ldloc.s    V_4
        IL_0042:  and
        IL_0043:  brfalse.s  IL_004c

        IL_0045:  ldloc.0
        IL_0046:  ldloc.1
        IL_0047:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<int32>::Add(!0)
        IL_004c:  leave.s    IL_0053

      }  // end .try
      finally
      {
        IL_004e:  ldloc.1
        IL_004f:  ldc.i4.1
        IL_0050:  add
        IL_0051:  stloc.1
        IL_0052:  endfinally
      }  // end handler
      IL_0053:  ldloc.2
      IL_0054:  callvirt   instance bool [mscorlib]System.Collections.IEnumerator::MoveNext()
      IL_0059:  brtrue.s   IL_0016

      IL_005b:  leave.s    IL_0071

    }  // end .try
    finally
    {
      IL_005d:  ldloc.2
      IL_005e:  isinst     [mscorlib]System.IDisposable
      IL_0063:  stloc.s    V_5
      IL_0065:  ldloc.s    V_5
      IL_0067:  brfalse.s  IL_0070

      IL_0069:  ldloc.s    V_5
      IL_006b:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
      IL_0070:  endfinally
    }  // end handler
    IL_0071:  ldloc.0
    IL_0072:  ret
  } // end of method ThisAddIn::GetReportIndex

  .method public hidebysig instance class [mscorlib]System.Collections.Generic.IEnumerable`1<int32> 
          GetSheetIndexes(class Microsoft.Office.Interop.Excel.Workbook workbook) cil managed
  {
    // Code size       48 (0x30)
    .maxstack  8
    IL_0000:  ldarg.1
    IL_0001:  callvirt   instance class Microsoft.Office.Interop.Excel.Sheets Microsoft.Office.Interop.Excel._Workbook::get_Sheets()
    IL_0006:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0> [System.Core]System.Linq.Enumerable::Cast<class Microsoft.Office.Interop.Excel.Worksheet>(class [mscorlib]System.Collections.IEnumerable)
    IL_000b:  ldsfld     class [mscorlib]System.Func`2<class Microsoft.Office.Interop.Excel.Worksheet,int32> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__51_0'
    IL_0010:  dup
    IL_0011:  brtrue.s   IL_002a

    IL_0013:  pop
    IL_0014:  ldsfld     class iReporterExcelAddIn2019.ThisAddIn/'<>c' iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9'
    IL_0019:  ldftn      instance int32 iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<GetSheetIndexes>b__51_0'(class Microsoft.Office.Interop.Excel.Worksheet)
    IL_001f:  newobj     instance void class [mscorlib]System.Func`2<class Microsoft.Office.Interop.Excel.Worksheet,int32>::.ctor(object,
                                                                                                                                  native int)
    IL_0024:  dup
    IL_0025:  stsfld     class [mscorlib]System.Func`2<class Microsoft.Office.Interop.Excel.Worksheet,int32> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__51_0'
    IL_002a:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!1> [System.Core]System.Linq.Enumerable::Select<class Microsoft.Office.Interop.Excel.Worksheet,int32>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                                                                         class [mscorlib]System.Func`2<!!0,!!1>)
    IL_002f:  ret
  } // end of method ThisAddIn::GetSheetIndexes

  .method public hidebysig instance class [mscorlib]System.Collections.Generic.IEnumerable`1<string> 
          GetSheetNames(class Microsoft.Office.Interop.Excel.Workbook workbook) cil managed
  {
    // Code size       48 (0x30)
    .maxstack  8
    IL_0000:  ldarg.1
    IL_0001:  callvirt   instance class Microsoft.Office.Interop.Excel.Sheets Microsoft.Office.Interop.Excel._Workbook::get_Sheets()
    IL_0006:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0> [System.Core]System.Linq.Enumerable::Cast<class Microsoft.Office.Interop.Excel.Worksheet>(class [mscorlib]System.Collections.IEnumerable)
    IL_000b:  ldsfld     class [mscorlib]System.Func`2<class Microsoft.Office.Interop.Excel.Worksheet,string> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__52_0'
    IL_0010:  dup
    IL_0011:  brtrue.s   IL_002a

    IL_0013:  pop
    IL_0014:  ldsfld     class iReporterExcelAddIn2019.ThisAddIn/'<>c' iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9'
    IL_0019:  ldftn      instance string iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<GetSheetNames>b__52_0'(class Microsoft.Office.Interop.Excel.Worksheet)
    IL_001f:  newobj     instance void class [mscorlib]System.Func`2<class Microsoft.Office.Interop.Excel.Worksheet,string>::.ctor(object,
                                                                                                                                   native int)
    IL_0024:  dup
    IL_0025:  stsfld     class [mscorlib]System.Func`2<class Microsoft.Office.Interop.Excel.Worksheet,string> iReporterExcelAddIn2019.ThisAddIn/'<>c'::'<>9__52_0'
    IL_002a:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!1> [System.Core]System.Linq.Enumerable::Select<class Microsoft.Office.Interop.Excel.Worksheet,string>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                                                                          class [mscorlib]System.Func`2<!!0,!!1>)
    IL_002f:  ret
  } // end of method ThisAddIn::GetSheetNames

  .method private hidebysig instance void 
          InternalStartup() cil managed
  {
    // Code size       37 (0x25)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.0
    IL_0002:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::ThisAddIn_Startup(object,
                                                                                            class [mscorlib]System.EventArgs)
    IL_0008:  newobj     instance void [mscorlib]System.EventHandler::.ctor(object,
                                                                            native int)
    IL_000d:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::add_Startup(class [mscorlib]System.EventHandler)
    IL_0012:  ldarg.0
    IL_0013:  ldarg.0
    IL_0014:  ldftn      instance void iReporterExcelAddIn2019.ThisAddIn::ThisAddIn_Shutdown(object,
                                                                                             class [mscorlib]System.EventArgs)
    IL_001a:  newobj     instance void [mscorlib]System.EventHandler::.ctor(object,
                                                                            native int)
    IL_001f:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::add_Shutdown(class [mscorlib]System.EventHandler)
    IL_0024:  ret
  } // end of method ThisAddIn::InternalStartup

  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory factory,
                               class [mscorlib]System.IServiceProvider serviceProvider) cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 01 00 00 00 00 00 ) 
    // Code size       69 (0x45)
    .maxstack  5
    IL_0000:  ldarg.0
    IL_0001:  newobj     instance void class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton>::.ctor()
    IL_0006:  stfld      class [mscorlib]System.Collections.Generic.List`1<class Microsoft.Office.Core.CommandBarButton> iReporterExcelAddIn2019.ThisAddIn::MenuTexts
    IL_000b:  ldarg.0
    IL_000c:  newobj     instance void [mscorlib]System.Object::.ctor()
    IL_0011:  stfld      object iReporterExcelAddIn2019.ThisAddIn::_inputLocker
    IL_0016:  ldarg.0
    IL_0017:  newobj     instance void [mscorlib]System.Object::.ctor()
    IL_001c:  stfld      object iReporterExcelAddIn2019.ThisAddIn::judgeLock
    IL_0021:  ldarg.0
    IL_0022:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0027:  stfld      object iReporterExcelAddIn2019.ThisAddIn::missing
    IL_002c:  ldarg.0
    IL_002d:  ldarg.1
    IL_002e:  ldarg.2
    IL_002f:  ldstr      "AddIn"
    IL_0034:  ldstr      "ThisAddIn"
    IL_0039:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::.ctor(class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Factory,
                                                                                                                               class [mscorlib]System.IServiceProvider,
                                                                                                                               string,
                                                                                                                               string)
    IL_003e:  ldarg.1
    IL_003f:  call       void iReporterExcelAddIn2019.Globals::set_Factory(class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory)
    IL_0044:  ret
  } // end of method ThisAddIn::.ctor

  .method family hidebysig virtual instance void 
          Initialize() cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                        string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                    75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                    4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                    6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                    2E 30 2E 30 00 00 )                               // .0.0..
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 01 00 00 00 00 00 ) 
    // Code size       69 (0x45)
    .maxstack  4
    IL_0000:  ldarg.0
    IL_0001:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::Initialize()
    IL_0006:  ldarg.0
    IL_0007:  ldarg.0
    IL_0008:  ldtoken    Microsoft.Office.Interop.Excel.Application
    IL_000d:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_0012:  ldstr      "Application"
    IL_0017:  call       instance !!0 ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::GetHostItem<class Microsoft.Office.Interop.Excel.Application>(class [mscorlib]System.Type,
                                                                                                                                                                                      string)
    IL_001c:  stfld      class Microsoft.Office.Interop.Excel.Application iReporterExcelAddIn2019.ThisAddIn::Application
    IL_0021:  ldarg.0
    IL_0022:  call       void iReporterExcelAddIn2019.Globals::set_ThisAddIn(class iReporterExcelAddIn2019.ThisAddIn)
    IL_0027:  call       void [System.Windows.Forms]System.Windows.Forms.Application::EnableVisualStyles()
    IL_002c:  ldarg.0
    IL_002d:  call       instance void iReporterExcelAddIn2019.ThisAddIn::InitializeCachedData()
    IL_0032:  ldarg.0
    IL_0033:  call       instance void iReporterExcelAddIn2019.ThisAddIn::InitializeControls()
    IL_0038:  ldarg.0
    IL_0039:  call       instance void iReporterExcelAddIn2019.ThisAddIn::InitializeComponents()
    IL_003e:  ldarg.0
    IL_003f:  call       instance void iReporterExcelAddIn2019.ThisAddIn::InitializeData()
    IL_0044:  ret
  } // end of method ThisAddIn::Initialize

  .method family hidebysig virtual instance void 
          FinishInitialization() cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                        string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                    75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                    4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                    6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                    2E 30 2E 30 00 00 )                               // .0.0..
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 01 00 00 00 00 00 ) 
    // Code size       13 (0xd)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void iReporterExcelAddIn2019.ThisAddIn::InternalStartup()
    IL_0006:  ldarg.0
    IL_0007:  callvirt   instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::OnStartup()
    IL_000c:  ret
  } // end of method ThisAddIn::FinishInitialization

  .method family hidebysig virtual instance void 
          InitializeDataBindings() cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                        string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                    75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                    4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                    6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                    2E 30 2E 30 00 00 )                               // .0.0..
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 01 00 00 00 00 00 ) 
    // Code size       19 (0x13)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void iReporterExcelAddIn2019.ThisAddIn::BeginInitialization()
    IL_0006:  ldarg.0
    IL_0007:  call       instance void iReporterExcelAddIn2019.ThisAddIn::BindToData()
    IL_000c:  ldarg.0
    IL_000d:  call       instance void iReporterExcelAddIn2019.ThisAddIn::EndInitialization()
    IL_0012:  ret
  } // end of method ThisAddIn::InitializeDataBindings

  .method private hidebysig instance void 
          InitializeCachedData() cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                        string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                    75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                    4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                    6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                    2E 30 2E 30 00 00 )                               // .0.0..
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 01 00 00 00 00 00 ) 
    // Code size       35 (0x23)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance class [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::get_DataHost()
    IL_0006:  brtrue.s   IL_0009

    IL_0008:  ret

    IL_0009:  ldarg.0
    IL_000a:  call       instance class [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::get_DataHost()
    IL_000f:  callvirt   instance bool [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider::get_IsCacheInitialized()
    IL_0014:  brfalse.s  IL_0022

    IL_0016:  ldarg.0
    IL_0017:  call       instance class [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::get_DataHost()
    IL_001c:  ldarg.0
    IL_001d:  callvirt   instance void [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider::FillCachedData(object)
    IL_0022:  ret
  } // end of method ThisAddIn::InitializeCachedData

  .method private hidebysig instance void 
          InitializeData() cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                        string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                    75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                    4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                    6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                    2E 30 2E 30 00 00 )                               // .0.0..
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 01 00 00 00 00 00 ) 
    // Code size       1 (0x1)
    .maxstack  8
    IL_0000:  ret
  } // end of method ThisAddIn::InitializeData

  .method private hidebysig instance void 
          BindToData() cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                        string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                    75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                    4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                    6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                    2E 30 2E 30 00 00 )                               // .0.0..
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 01 00 00 00 00 00 ) 
    // Code size       1 (0x1)
    .maxstack  8
    IL_0000:  ret
  } // end of method ThisAddIn::BindToData

  .method private hidebysig instance void 
          StartCaching(string MemberName) cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 02 00 00 00 00 00 ) 
    // Code size       14 (0xe)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance class [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::get_DataHost()
    IL_0006:  ldarg.0
    IL_0007:  ldarg.1
    IL_0008:  callvirt   instance void [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider::StartCaching(object,
                                                                                                                        string)
    IL_000d:  ret
  } // end of method ThisAddIn::StartCaching

  .method private hidebysig instance void 
          StopCaching(string MemberName) cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 02 00 00 00 00 00 ) 
    // Code size       14 (0xe)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance class [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::get_DataHost()
    IL_0006:  ldarg.0
    IL_0007:  ldarg.1
    IL_0008:  callvirt   instance void [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider::StopCaching(object,
                                                                                                                       string)
    IL_000d:  ret
  } // end of method ThisAddIn::StopCaching

  .method private hidebysig instance bool 
          IsCached(string MemberName) cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 02 00 00 00 00 00 ) 
    // Code size       14 (0xe)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance class [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::get_DataHost()
    IL_0006:  ldarg.0
    IL_0007:  ldarg.1
    IL_0008:  callvirt   instance bool [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider::IsCached(object,
                                                                                                                    string)
    IL_000d:  ret
  } // end of method ThisAddIn::IsCached

  .method private hidebysig instance void 
          BeginInitialization() cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                        string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                    75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                    4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                    6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                    2E 30 2E 30 00 00 )                               // .0.0..
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 01 00 00 00 00 00 ) 
    // Code size       29 (0x1d)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  callvirt   instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::BeginInit()
    IL_0006:  ldarg.0
    IL_0007:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.CustomTaskPaneCollection iReporterExcelAddIn2019.ThisAddIn::CustomTaskPanes
    IL_000c:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.CustomTaskPaneCollection::BeginInit()
    IL_0011:  ldarg.0
    IL_0012:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.SmartTagCollection iReporterExcelAddIn2019.ThisAddIn::VstoSmartTags
    IL_0017:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.SmartTagCollection::BeginInit()
    IL_001c:  ret
  } // end of method ThisAddIn::BeginInitialization

  .method private hidebysig instance void 
          EndInitialization() cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                        string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                    75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                    4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                    6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                    2E 30 2E 30 00 00 )                               // .0.0..
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 01 00 00 00 00 00 ) 
    // Code size       29 (0x1d)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.SmartTagCollection iReporterExcelAddIn2019.ThisAddIn::VstoSmartTags
    IL_0006:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.SmartTagCollection::EndInit()
    IL_000b:  ldarg.0
    IL_000c:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.CustomTaskPaneCollection iReporterExcelAddIn2019.ThisAddIn::CustomTaskPanes
    IL_0011:  callvirt   instance void [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.CustomTaskPaneCollection::EndInit()
    IL_0016:  ldarg.0
    IL_0017:  callvirt   instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::EndInit()
    IL_001c:  ret
  } // end of method ThisAddIn::EndInitialization

  .method private hidebysig instance void 
          InitializeControls() cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                        string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                    75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                    4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                    6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                    2E 30 2E 30 00 00 )                               // .0.0..
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 01 00 00 00 00 00 ) 
    // Code size       59 (0x3b)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory iReporterExcelAddIn2019.Globals::get_Factory()
    IL_0006:  ldnull
    IL_0007:  ldnull
    IL_0008:  ldstr      "CustomTaskPanes"
    IL_000d:  ldstr      "CustomTaskPanes"
    IL_0012:  ldarg.0
    IL_0013:  callvirt   instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.CustomTaskPaneCollection [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Factory::CreateCustomTaskPaneCollection(class [mscorlib]System.IServiceProvider,
                                                                                                                                                                                                                     class [Microsoft.Office.Tools]Microsoft.Office.Tools.IHostItemProvider,
                                                                                                                                                                                                                     string,
                                                                                                                                                                                                                     string,
                                                                                                                                                                                                                     object)
    IL_0018:  stfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.CustomTaskPaneCollection iReporterExcelAddIn2019.ThisAddIn::CustomTaskPanes
    IL_001d:  ldarg.0
    IL_001e:  call       class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory iReporterExcelAddIn2019.Globals::get_Factory()
    IL_0023:  ldnull
    IL_0024:  ldnull
    IL_0025:  ldstr      "VstoSmartTags"
    IL_002a:  ldstr      "VstoSmartTags"
    IL_002f:  ldarg.0
    IL_0030:  callvirt   instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.SmartTagCollection [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Factory::CreateSmartTagCollection(class [mscorlib]System.IServiceProvider,
                                                                                                                                                                                                         class [Microsoft.Office.Tools]Microsoft.Office.Tools.IHostItemProvider,
                                                                                                                                                                                                         string,
                                                                                                                                                                                                         string,
                                                                                                                                                                                                         object)
    IL_0035:  stfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.SmartTagCollection iReporterExcelAddIn2019.ThisAddIn::VstoSmartTags
    IL_003a:  ret
  } // end of method ThisAddIn::InitializeControls

  .method private hidebysig instance void 
          InitializeComponents() cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                        string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                    75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                    4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                    6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                    2E 30 2E 30 00 00 )                               // .0.0..
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 01 00 00 00 00 00 ) 
    // Code size       1 (0x1)
    .maxstack  8
    IL_0000:  ret
  } // end of method ThisAddIn::InitializeComponents

  .method private hidebysig instance bool 
          NeedsFill(string MemberName) cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 02 00 00 00 00 00 ) 
    // Code size       14 (0xe)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance class [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::get_DataHost()
    IL_0006:  ldarg.0
    IL_0007:  ldarg.1
    IL_0008:  callvirt   instance bool [Microsoft.Office.Tools]Microsoft.Office.Tools.ICachedDataProvider::NeedsFill(object,
                                                                                                                     string)
    IL_000d:  ret
  } // end of method ThisAddIn::NeedsFill

  .method family hidebysig virtual instance void 
          OnShutdown() cil managed
  {
    .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
    .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                        string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                    75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                    4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                    6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                    2E 30 2E 30 00 00 )                               // .0.0..
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 01 00 00 00 00 00 ) 
    // Code size       29 (0x1d)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.SmartTagCollection iReporterExcelAddIn2019.ThisAddIn::VstoSmartTags
    IL_0006:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
    IL_000b:  ldarg.0
    IL_000c:  ldfld      class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.CustomTaskPaneCollection iReporterExcelAddIn2019.ThisAddIn::CustomTaskPanes
    IL_0011:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
    IL_0016:  ldarg.0
    IL_0017:  call       instance void ['Microsoft.Office.Tools.Common.v4.0.Utilities']Microsoft.Office.Tools.AddInBase::OnShutdown()
    IL_001c:  ret
  } // end of method ThisAddIn::OnShutdown

  .method private hidebysig instance void 
          '<SetUpMenu>b__13_0'(class Microsoft.Office.Core.CommandBarButton Ctrl,
                               bool& CancelDefault) cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       8 (0x8)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldc.i4.1
    IL_0002:  call       instance void iReporterExcelAddIn2019.ThisAddIn::AutoJudge(bool)
    IL_0007:  ret
  } // end of method ThisAddIn::'<SetUpMenu>b__13_0'

  .method private hidebysig instance void 
          '<SetUpMenu>b__13_1'(class Microsoft.Office.Core.CommandBarButton Ctrl,
                               bool& CancelDefault) cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ManualCluster()
    IL_0006:  ret
  } // end of method ThisAddIn::'<SetUpMenu>b__13_1'

  .method private hidebysig instance void 
          '<SetUpMenu>b__13_2'(class Microsoft.Office.Core.CommandBarButton Ctrl,
                               bool& CancelDefault) cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void iReporterExcelAddIn2019.ThisAddIn::SetClusterColors()
    IL_0006:  ret
  } // end of method ThisAddIn::'<SetUpMenu>b__13_2'

  .method private hidebysig instance void 
          '<SetUpMenu>b__13_3'(class Microsoft.Office.Core.CommandBarButton Ctrl,
                               bool& CancelDefault) cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ShowTableForm()
    IL_0006:  ret
  } // end of method ThisAddIn::'<SetUpMenu>b__13_3'

  .method private hidebysig instance void 
          '<ManualCluster>b__17_0'() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       8 (0x8)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldc.i4.0
    IL_0002:  call       instance void iReporterExcelAddIn2019.ThisAddIn::AutoJudge(bool)
    IL_0007:  ret
  } // end of method ThisAddIn::'<ManualCluster>b__17_0'

  .method assembly hidebysig static void 
          '<GetCellText>g__Add|42_0'(class Microsoft.Office.Interop.Excel.Worksheet s,
                                     int32 row,
                                     int32 col,
                                     valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'& A_3,
                                     valuetype iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1'& A_4) cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       1543 (0x607)
    .maxstack  17
    .locals init (object V_0,
             object V_1)
    IL_0000:  ldarg.1
    IL_0001:  ldc.i4.1
    IL_0002:  blt.s      IL_0008

    IL_0004:  ldarg.2
    IL_0005:  ldc.i4.1
    IL_0006:  bge.s      IL_0009

    IL_0008:  ret

    IL_0009:  nop
    .try
    {
      IL_000a:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__3'
      IL_000f:  brtrue.s   IL_0040

      IL_0011:  ldc.i4.0
      IL_0012:  ldstr      "Value2"
      IL_0017:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_001c:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0021:  ldc.i4.1
      IL_0022:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_0027:  dup
      IL_0028:  ldc.i4.0
      IL_0029:  ldc.i4.0
      IL_002a:  ldnull
      IL_002b:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0030:  stelem.ref
      IL_0031:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::GetMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                string,
                                                                                                                                                                class [mscorlib]System.Type,
                                                                                                                                                                class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_0036:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_003b:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__3'
      IL_0040:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__3'
      IL_0045:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Target
      IL_004a:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__3'
      IL_004f:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32,int32,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__2'
      IL_0054:  brtrue.s   IL_0094

      IL_0056:  ldc.i4.0
      IL_0057:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_005c:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0061:  ldc.i4.3
      IL_0062:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_0067:  dup
      IL_0068:  ldc.i4.0
      IL_0069:  ldc.i4.0
      IL_006a:  ldnull
      IL_006b:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0070:  stelem.ref
      IL_0071:  dup
      IL_0072:  ldc.i4.1
      IL_0073:  ldc.i4.3
      IL_0074:  ldnull
      IL_0075:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_007a:  stelem.ref
      IL_007b:  dup
      IL_007c:  ldc.i4.2
      IL_007d:  ldc.i4.3
      IL_007e:  ldnull
      IL_007f:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0084:  stelem.ref
      IL_0085:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::GetIndex(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                               class [mscorlib]System.Type,
                                                                                                                                                               class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_008a:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32,int32,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_008f:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32,int32,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__2'
      IL_0094:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32,int32,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__2'
      IL_0099:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32,int32,object>>::Target
      IL_009e:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32,int32,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__2'
      IL_00a3:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__1'
      IL_00a8:  brtrue.s   IL_00da

      IL_00aa:  ldc.i4.s   64
      IL_00ac:  ldstr      "Cells"
      IL_00b1:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_00b6:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_00bb:  ldc.i4.1
      IL_00bc:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_00c1:  dup
      IL_00c2:  ldc.i4.0
      IL_00c3:  ldc.i4.0
      IL_00c4:  ldnull
      IL_00c5:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_00ca:  stelem.ref
      IL_00cb:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::GetMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                string,
                                                                                                                                                                class [mscorlib]System.Type,
                                                                                                                                                                class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_00d0:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_00d5:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__1'
      IL_00da:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__1'
      IL_00df:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Target
      IL_00e4:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__1'
      IL_00e9:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__0'
      IL_00ee:  brtrue.s   IL_011f

      IL_00f0:  ldc.i4.0
      IL_00f1:  ldstr      "MergeArea"
      IL_00f6:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_00fb:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0100:  ldc.i4.1
      IL_0101:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_0106:  dup
      IL_0107:  ldc.i4.0
      IL_0108:  ldc.i4.0
      IL_0109:  ldnull
      IL_010a:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_010f:  stelem.ref
      IL_0110:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::GetMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                string,
                                                                                                                                                                class [mscorlib]System.Type,
                                                                                                                                                                class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_0115:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_011a:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__0'
      IL_011f:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__0'
      IL_0124:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Target
      IL_0129:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__0'
      IL_012e:  ldarg.0
      IL_012f:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Cells()
      IL_0134:  ldarg.1
      IL_0135:  box        [mscorlib]System.Int32
      IL_013a:  ldarg.2
      IL_013b:  box        [mscorlib]System.Int32
      IL_0140:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get__Default(object,
                                                                                              object)
      IL_0145:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>::Invoke(!0,
                                                                                                                                                        !1)
      IL_014a:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>::Invoke(!0,
                                                                                                                                                        !1)
      IL_014f:  ldc.i4.1
      IL_0150:  ldc.i4.1
      IL_0151:  callvirt   instance !4 class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32,int32,object>::Invoke(!0,
                                                                                                                                                                    !1,
                                                                                                                                                                    !2,
                                                                                                                                                                    !3)
      IL_0156:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>::Invoke(!0,
                                                                                                                                                        !1)
      IL_015b:  stloc.0
      IL_015c:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__5'
      IL_0161:  brtrue.s   IL_018f

      IL_0163:  ldc.i4.0
      IL_0164:  ldc.i4.s   83
      IL_0166:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_016b:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0170:  ldc.i4.1
      IL_0171:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_0176:  dup
      IL_0177:  ldc.i4.0
      IL_0178:  ldc.i4.0
      IL_0179:  ldnull
      IL_017a:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_017f:  stelem.ref
      IL_0180:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::UnaryOperation(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                     valuetype [System.Core]System.Linq.Expressions.ExpressionType,
                                                                                                                                                                     class [mscorlib]System.Type,
                                                                                                                                                                     class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_0185:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_018a:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__5'
      IL_018f:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__5'
      IL_0194:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>>::Target
      IL_0199:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__5'
      IL_019e:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__4'
      IL_01a3:  brtrue.s   IL_01db

      IL_01a5:  ldc.i4.0
      IL_01a6:  ldc.i4.s   13
      IL_01a8:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_01ad:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_01b2:  ldc.i4.2
      IL_01b3:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_01b8:  dup
      IL_01b9:  ldc.i4.0
      IL_01ba:  ldc.i4.0
      IL_01bb:  ldnull
      IL_01bc:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_01c1:  stelem.ref
      IL_01c2:  dup
      IL_01c3:  ldc.i4.1
      IL_01c4:  ldc.i4.2
      IL_01c5:  ldnull
      IL_01c6:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_01cb:  stelem.ref
      IL_01cc:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::BinaryOperation(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                      valuetype [System.Core]System.Linq.Expressions.ExpressionType,
                                                                                                                                                                      class [mscorlib]System.Type,
                                                                                                                                                                      class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_01d1:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_01d6:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__4'
      IL_01db:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__4'
      IL_01e0:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object,object>>::Target
      IL_01e5:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__4'
      IL_01ea:  ldloc.0
      IL_01eb:  ldnull
      IL_01ec:  callvirt   instance !3 class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object,object>::Invoke(!0,
                                                                                                                                                               !1,
                                                                                                                                                               !2)
      IL_01f1:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>::Invoke(!0,
                                                                                                                                                      !1)
      IL_01f6:  brfalse.s  IL_01fd

      IL_01f8:  leave      IL_0606

      IL_01fd:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__10'
      IL_0202:  brtrue.s   IL_0248

      IL_0204:  ldc.i4.0
      IL_0205:  ldstr      "Replace"
      IL_020a:  ldnull
      IL_020b:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_0210:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0215:  ldc.i4.3
      IL_0216:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_021b:  dup
      IL_021c:  ldc.i4.0
      IL_021d:  ldc.i4.0
      IL_021e:  ldnull
      IL_021f:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0224:  stelem.ref
      IL_0225:  dup
      IL_0226:  ldc.i4.1
      IL_0227:  ldc.i4.3
      IL_0228:  ldnull
      IL_0229:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_022e:  stelem.ref
      IL_022f:  dup
      IL_0230:  ldc.i4.2
      IL_0231:  ldc.i4.3
      IL_0232:  ldnull
      IL_0233:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0238:  stelem.ref
      IL_0239:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   string,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_023e:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_0243:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__10'
      IL_0248:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__10'
      IL_024d:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>>::Target
      IL_0252:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__10'
      IL_0257:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__9'
      IL_025c:  brtrue.s   IL_02a2

      IL_025e:  ldc.i4.0
      IL_025f:  ldstr      "Replace"
      IL_0264:  ldnull
      IL_0265:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_026a:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_026f:  ldc.i4.3
      IL_0270:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_0275:  dup
      IL_0276:  ldc.i4.0
      IL_0277:  ldc.i4.0
      IL_0278:  ldnull
      IL_0279:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_027e:  stelem.ref
      IL_027f:  dup
      IL_0280:  ldc.i4.1
      IL_0281:  ldc.i4.3
      IL_0282:  ldnull
      IL_0283:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0288:  stelem.ref
      IL_0289:  dup
      IL_028a:  ldc.i4.2
      IL_028b:  ldc.i4.3
      IL_028c:  ldnull
      IL_028d:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0292:  stelem.ref
      IL_0293:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   string,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_0298:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_029d:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__9'
      IL_02a2:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__9'
      IL_02a7:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>>::Target
      IL_02ac:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__9'
      IL_02b1:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__8'
      IL_02b6:  brtrue.s   IL_02fc

      IL_02b8:  ldc.i4.0
      IL_02b9:  ldstr      "Replace"
      IL_02be:  ldnull
      IL_02bf:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_02c4:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_02c9:  ldc.i4.3
      IL_02ca:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_02cf:  dup
      IL_02d0:  ldc.i4.0
      IL_02d1:  ldc.i4.0
      IL_02d2:  ldnull
      IL_02d3:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_02d8:  stelem.ref
      IL_02d9:  dup
      IL_02da:  ldc.i4.1
      IL_02db:  ldc.i4.3
      IL_02dc:  ldnull
      IL_02dd:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_02e2:  stelem.ref
      IL_02e3:  dup
      IL_02e4:  ldc.i4.2
      IL_02e5:  ldc.i4.3
      IL_02e6:  ldnull
      IL_02e7:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_02ec:  stelem.ref
      IL_02ed:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   string,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_02f2:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_02f7:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__8'
      IL_02fc:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__8'
      IL_0301:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>>::Target
      IL_0306:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__8'
      IL_030b:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__7'
      IL_0310:  brtrue.s   IL_0356

      IL_0312:  ldc.i4.0
      IL_0313:  ldstr      "Replace"
      IL_0318:  ldnull
      IL_0319:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_031e:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0323:  ldc.i4.3
      IL_0324:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_0329:  dup
      IL_032a:  ldc.i4.0
      IL_032b:  ldc.i4.0
      IL_032c:  ldnull
      IL_032d:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0332:  stelem.ref
      IL_0333:  dup
      IL_0334:  ldc.i4.1
      IL_0335:  ldc.i4.3
      IL_0336:  ldnull
      IL_0337:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_033c:  stelem.ref
      IL_033d:  dup
      IL_033e:  ldc.i4.2
      IL_033f:  ldc.i4.3
      IL_0340:  ldnull
      IL_0341:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0346:  stelem.ref
      IL_0347:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   string,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_034c:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_0351:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__7'
      IL_0356:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__7'
      IL_035b:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>>::Target
      IL_0360:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__7'
      IL_0365:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__6'
      IL_036a:  brtrue.s   IL_03a7

      IL_036c:  ldc.i4.0
      IL_036d:  ldstr      "ToString"
      IL_0372:  ldnull
      IL_0373:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_0378:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_037d:  ldc.i4.2
      IL_037e:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_0383:  dup
      IL_0384:  ldc.i4.0
      IL_0385:  ldc.i4.s   33
      IL_0387:  ldnull
      IL_0388:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_038d:  stelem.ref
      IL_038e:  dup
      IL_038f:  ldc.i4.1
      IL_0390:  ldc.i4.0
      IL_0391:  ldnull
      IL_0392:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0397:  stelem.ref
      IL_0398:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   string,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_039d:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_03a2:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__6'
      IL_03a7:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__6'
      IL_03ac:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>>::Target
      IL_03b1:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__6'
      IL_03b6:  ldtoken    [mscorlib]System.Convert
      IL_03bb:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_03c0:  ldloc.0
      IL_03c1:  callvirt   instance !3 class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>::Invoke(!0,
                                                                                                                                                                                    !1,
                                                                                                                                                                                    !2)
      IL_03c6:  ldstr      "\r\n"
      IL_03cb:  ldstr      " "
      IL_03d0:  callvirt   instance !4 class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>::Invoke(!0,
                                                                                                                                                                      !1,
                                                                                                                                                                      !2,
                                                                                                                                                                      !3)
      IL_03d5:  ldstr      "\r"
      IL_03da:  ldstr      " "
      IL_03df:  callvirt   instance !4 class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>::Invoke(!0,
                                                                                                                                                                      !1,
                                                                                                                                                                      !2,
                                                                                                                                                                      !3)
      IL_03e4:  ldstr      "\n"
      IL_03e9:  ldstr      " "
      IL_03ee:  callvirt   instance !4 class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>::Invoke(!0,
                                                                                                                                                                      !1,
                                                                                                                                                                      !2,
                                                                                                                                                                      !3)
      IL_03f3:  ldstr      "\t"
      IL_03f8:  ldstr      " "
      IL_03fd:  callvirt   instance !4 class [mscorlib]System.Func`5<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string,string,object>::Invoke(!0,
                                                                                                                                                                      !1,
                                                                                                                                                                      !2,
                                                                                                                                                                      !3)
      IL_0402:  stloc.1
      IL_0403:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__12'
      IL_0408:  brtrue.s   IL_0436

      IL_040a:  ldc.i4.0
      IL_040b:  ldc.i4.s   83
      IL_040d:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_0412:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0417:  ldc.i4.1
      IL_0418:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_041d:  dup
      IL_041e:  ldc.i4.0
      IL_041f:  ldc.i4.0
      IL_0420:  ldnull
      IL_0421:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0426:  stelem.ref
      IL_0427:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::UnaryOperation(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                     valuetype [System.Core]System.Linq.Expressions.ExpressionType,
                                                                                                                                                                     class [mscorlib]System.Type,
                                                                                                                                                                     class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_042c:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_0431:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__12'
      IL_0436:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__12'
      IL_043b:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>>::Target
      IL_0440:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__12'
      IL_0445:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__11'
      IL_044a:  brtrue.s   IL_0487

      IL_044c:  ldc.i4.0
      IL_044d:  ldstr      "IsNullOrWhiteSpace"
      IL_0452:  ldnull
      IL_0453:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_0458:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_045d:  ldc.i4.2
      IL_045e:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_0463:  dup
      IL_0464:  ldc.i4.0
      IL_0465:  ldc.i4.s   33
      IL_0467:  ldnull
      IL_0468:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_046d:  stelem.ref
      IL_046e:  dup
      IL_046f:  ldc.i4.1
      IL_0470:  ldc.i4.0
      IL_0471:  ldnull
      IL_0472:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0477:  stelem.ref
      IL_0478:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   string,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_047d:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_0482:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__11'
      IL_0487:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__11'
      IL_048c:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>>::Target
      IL_0491:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__11'
      IL_0496:  ldtoken    [mscorlib]System.String
      IL_049b:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_04a0:  ldloc.1
      IL_04a1:  callvirt   instance !3 class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Type,object,object>::Invoke(!0,
                                                                                                                                                                                    !1,
                                                                                                                                                                                    !2)
      IL_04a6:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>::Invoke(!0,
                                                                                                                                                      !1)
      IL_04ab:  brfalse.s  IL_04b2

      IL_04ad:  leave      IL_0606

      IL_04b2:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__15'
      IL_04b7:  brtrue.s   IL_04e5

      IL_04b9:  ldc.i4.0
      IL_04ba:  ldc.i4.s   83
      IL_04bc:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_04c1:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_04c6:  ldc.i4.1
      IL_04c7:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_04cc:  dup
      IL_04cd:  ldc.i4.0
      IL_04ce:  ldc.i4.0
      IL_04cf:  ldnull
      IL_04d0:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_04d5:  stelem.ref
      IL_04d6:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::UnaryOperation(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                     valuetype [System.Core]System.Linq.Expressions.ExpressionType,
                                                                                                                                                                     class [mscorlib]System.Type,
                                                                                                                                                                     class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_04db:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_04e0:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__15'
      IL_04e5:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__15'
      IL_04ea:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>>::Target
      IL_04ef:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__15'
      IL_04f4:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__14'
      IL_04f9:  brtrue.s   IL_0527

      IL_04fb:  ldc.i4.0
      IL_04fc:  ldc.i4.s   34
      IL_04fe:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_0503:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0508:  ldc.i4.1
      IL_0509:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_050e:  dup
      IL_050f:  ldc.i4.0
      IL_0510:  ldc.i4.0
      IL_0511:  ldnull
      IL_0512:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0517:  stelem.ref
      IL_0518:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::UnaryOperation(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                     valuetype [System.Core]System.Linq.Expressions.ExpressionType,
                                                                                                                                                                     class [mscorlib]System.Type,
                                                                                                                                                                     class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_051d:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_0522:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__14'
      IL_0527:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__14'
      IL_052c:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Target
      IL_0531:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__14'
      IL_0536:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [System.Core]System.Collections.Generic.HashSet`1<string>,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__13'
      IL_053b:  brtrue.s   IL_0577

      IL_053d:  ldc.i4.0
      IL_053e:  ldstr      "Add"
      IL_0543:  ldnull
      IL_0544:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_0549:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_054e:  ldc.i4.2
      IL_054f:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_0554:  dup
      IL_0555:  ldc.i4.0
      IL_0556:  ldc.i4.1
      IL_0557:  ldnull
      IL_0558:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_055d:  stelem.ref
      IL_055e:  dup
      IL_055f:  ldc.i4.1
      IL_0560:  ldc.i4.0
      IL_0561:  ldnull
      IL_0562:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_0567:  stelem.ref
      IL_0568:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   string,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_056d:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [System.Core]System.Collections.Generic.HashSet`1<string>,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_0572:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [System.Core]System.Collections.Generic.HashSet`1<string>,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__13'
      IL_0577:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [System.Core]System.Collections.Generic.HashSet`1<string>,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__13'
      IL_057c:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [System.Core]System.Collections.Generic.HashSet`1<string>,object,object>>::Target
      IL_0581:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [System.Core]System.Collections.Generic.HashSet`1<string>,object,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__13'
      IL_0586:  ldarg.s    A_4
      IL_0588:  ldfld      class [System.Core]System.Collections.Generic.HashSet`1<string> iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_1'::set
      IL_058d:  ldloc.1
      IL_058e:  callvirt   instance !3 class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class [System.Core]System.Collections.Generic.HashSet`1<string>,object,object>::Invoke(!0,
                                                                                                                                                                                                                        !1,
                                                                                                                                                                                                                        !2)
      IL_0593:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>::Invoke(!0,
                                                                                                                                                        !1)
      IL_0598:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>::Invoke(!0,
                                                                                                                                                      !1)
      IL_059d:  brfalse.s  IL_05a1

      IL_059f:  leave.s    IL_0606

      IL_05a1:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Action`3<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Collections.Generic.List`1<string>,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__16'
      IL_05a6:  brtrue.s   IL_05e6

      IL_05a8:  ldc.i4     0x100
      IL_05ad:  ldstr      "Add"
      IL_05b2:  ldnull
      IL_05b3:  ldtoken    iReporterExcelAddIn2019.ThisAddIn
      IL_05b8:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_05bd:  ldc.i4.2
      IL_05be:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
      IL_05c3:  dup
      IL_05c4:  ldc.i4.0
      IL_05c5:  ldc.i4.1
      IL_05c6:  ldnull
      IL_05c7:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_05cc:  stelem.ref
      IL_05cd:  dup
      IL_05ce:  ldc.i4.1
      IL_05cf:  ldc.i4.0
      IL_05d0:  ldnull
      IL_05d1:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                                 string)
      IL_05d6:  stelem.ref
      IL_05d7:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   string,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
      IL_05dc:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Action`3<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Collections.Generic.List`1<string>,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_05e1:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Action`3<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Collections.Generic.List`1<string>,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__16'
      IL_05e6:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Action`3<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Collections.Generic.List`1<string>,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__16'
      IL_05eb:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Action`3<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Collections.Generic.List`1<string>,object>>::Target
      IL_05f0:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Action`3<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Collections.Generic.List`1<string>,object>> iReporterExcelAddIn2019.ThisAddIn/'<>o__42'::'<>p__16'
      IL_05f5:  ldarg.3
      IL_05f6:  ldfld      class [mscorlib]System.Collections.Generic.List`1<string> iReporterExcelAddIn2019.ThisAddIn/'<>c__DisplayClass42_0'::list
      IL_05fb:  ldloc.1
      IL_05fc:  callvirt   instance void class [mscorlib]System.Action`3<class [System.Core]System.Runtime.CompilerServices.CallSite,class [mscorlib]System.Collections.Generic.List`1<string>,object>::Invoke(!0,
                                                                                                                                                                                                               !1,
                                                                                                                                                                                                               !2)
      IL_0601:  leave.s    IL_0606

    }  // end .try
    catch [mscorlib]System.Object 
    {
      IL_0603:  pop
      IL_0604:  leave.s    IL_0606

    }  // end handler
    IL_0606:  ret
  } // end of method ThisAddIn::'<GetCellText>g__Add|42_0'

  .method private hidebysig instance void 
          '<ShowTableForm>b__45_2'(object o,
                                   class [mscorlib]System.EventArgs _) cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void iReporterExcelAddIn2019.ThisAddIn::ShowTableForm()
    IL_0006:  ret
  } // end of method ThisAddIn::'<ShowTableForm>b__45_2'

  .property instance bool IsJudging()
  {
    .get instance bool iReporterExcelAddIn2019.ThisAddIn::get_IsJudging()
    .set instance void iReporterExcelAddIn2019.ThisAddIn::set_IsJudging(bool)
  } // end of property ThisAddIn::IsJudging
} // end of class iReporterExcelAddIn2019.ThisAddIn

.class private auto ansi sealed beforefieldinit iReporterExcelAddIn2019.Globals
       extends [mscorlib]System.Object
{
  .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                      string) = ( 01 00 38 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..8Microsoft.Vis
                                                                                                  75 61 6C 53 74 75 64 69 6F 2E 54 6F 6F 6C 73 2E   // ualStudio.Tools.
                                                                                                  4F 66 66 69 63 65 2E 50 72 6F 67 72 61 6D 6D 69   // Office.Programmi
                                                                                                  6E 67 4D 6F 64 65 6C 2E 64 6C 6C 08 31 35 2E 30   // ngModel.dll.15.0
                                                                                                  2E 30 2E 30 00 00 )                               // .0.0..
  .field private static class iReporterExcelAddIn2019.ThisAddIn _ThisAddIn
  .field private static class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory _factory
  .field private static class iReporterExcelAddIn2019.ThisRibbonCollection _ThisRibbonCollection
  .method private hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
    IL_0006:  ret
  } // end of method Globals::.ctor

  .method assembly hidebysig specialname static 
          class iReporterExcelAddIn2019.ThisAddIn 
          get_ThisAddIn() cil managed
  {
    // Code size       6 (0x6)
    .maxstack  8
    IL_0000:  ldsfld     class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.Globals::_ThisAddIn
    IL_0005:  ret
  } // end of method Globals::get_ThisAddIn

  .method assembly hidebysig specialname static 
          void  set_ThisAddIn(class iReporterExcelAddIn2019.ThisAddIn 'value') cil managed
  {
    // Code size       20 (0x14)
    .maxstack  8
    IL_0000:  ldsfld     class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.Globals::_ThisAddIn
    IL_0005:  brtrue.s   IL_000e

    IL_0007:  ldarg.0
    IL_0008:  stsfld     class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.Globals::_ThisAddIn
    IL_000d:  ret

    IL_000e:  newobj     instance void [mscorlib]System.NotSupportedException::.ctor()
    IL_0013:  throw
  } // end of method Globals::set_ThisAddIn

  .method assembly hidebysig specialname static 
          class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory 
          get_Factory() cil managed
  {
    // Code size       6 (0x6)
    .maxstack  8
    IL_0000:  ldsfld     class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory iReporterExcelAddIn2019.Globals::_factory
    IL_0005:  ret
  } // end of method Globals::get_Factory

  .method assembly hidebysig specialname static 
          void  set_Factory(class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory 'value') cil managed
  {
    // Code size       20 (0x14)
    .maxstack  8
    IL_0000:  ldsfld     class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory iReporterExcelAddIn2019.Globals::_factory
    IL_0005:  brtrue.s   IL_000e

    IL_0007:  ldarg.0
    IL_0008:  stsfld     class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory iReporterExcelAddIn2019.Globals::_factory
    IL_000d:  ret

    IL_000e:  newobj     instance void [mscorlib]System.NotSupportedException::.ctor()
    IL_0013:  throw
  } // end of method Globals::set_Factory

  .method assembly hidebysig specialname static 
          class iReporterExcelAddIn2019.ThisRibbonCollection 
          get_Ribbons() cil managed
  {
    // Code size       33 (0x21)
    .maxstack  8
    IL_0000:  ldsfld     class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::_ThisRibbonCollection
    IL_0005:  brtrue.s   IL_001b

    IL_0007:  ldsfld     class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory iReporterExcelAddIn2019.Globals::_factory
    IL_000c:  callvirt   instance class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Factory::GetRibbonFactory()
    IL_0011:  newobj     instance void iReporterExcelAddIn2019.ThisRibbonCollection::.ctor(class [Microsoft.Office.Tools.Common]Microsoft.Office.Tools.Ribbon.RibbonFactory)
    IL_0016:  stsfld     class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::_ThisRibbonCollection
    IL_001b:  ldsfld     class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::_ThisRibbonCollection
    IL_0020:  ret
  } // end of method Globals::get_Ribbons

  .property class iReporterExcelAddIn2019.ThisAddIn
          ThisAddIn()
  {
    .get class iReporterExcelAddIn2019.ThisAddIn iReporterExcelAddIn2019.Globals::get_ThisAddIn()
    .set void iReporterExcelAddIn2019.Globals::set_ThisAddIn(class iReporterExcelAddIn2019.ThisAddIn)
  } // end of property Globals::ThisAddIn
  .property class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory
          Factory()
  {
    .get class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory iReporterExcelAddIn2019.Globals::get_Factory()
    .set void iReporterExcelAddIn2019.Globals::set_Factory(class [Microsoft.Office.Tools.Excel]Microsoft.Office.Tools.Excel.ApplicationFactory)
  } // end of property Globals::Factory
  .property class iReporterExcelAddIn2019.ThisRibbonCollection
          Ribbons()
  {
    .get class iReporterExcelAddIn2019.ThisRibbonCollection iReporterExcelAddIn2019.Globals::get_Ribbons()
  } // end of property Globals::Ribbons
} // end of class iReporterExcelAddIn2019.Globals

.class private auto ansi beforefieldinit iReporterExcelAddIn2019.Properties.Resources
       extends [mscorlib]System.Object
{
  .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                      string) = ( 01 00 33 53 79 73 74 65 6D 2E 52 65 73 6F 75 72   // ..3System.Resour
                                                                                                  63 65 73 2E 54 6F 6F 6C 73 2E 53 74 72 6F 6E 67   // ces.Tools.Strong
                                                                                                  6C 79 54 79 70 65 64 52 65 73 6F 75 72 63 65 42   // lyTypedResourceB
                                                                                                  75 69 6C 64 65 72 08 31 35 2E 30 2E 30 2E 30 00   // uilder.15.0.0.0.
                                                                                                  00 ) 
  .custom instance void [mscorlib]System.Diagnostics.DebuggerNonUserCodeAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private static class [mscorlib]System.Resources.ResourceManager resourceMan
  .field private static class [mscorlib]System.Globalization.CultureInfo resourceCulture
  .method assembly hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
    IL_0006:  ret
  } // end of method Resources::.ctor

  .method assembly hidebysig specialname static 
          class [mscorlib]System.Resources.ResourceManager 
          get_ResourceManager() cil managed
  {
    // Code size       43 (0x2b)
    .maxstack  8
    IL_0000:  ldsfld     class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::resourceMan
    IL_0005:  brtrue.s   IL_0025

    IL_0007:  ldstr      "iReporterExcelAddIn2019.Properties.Resources"
    IL_000c:  ldtoken    iReporterExcelAddIn2019.Properties.Resources
    IL_0011:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_0016:  callvirt   instance class [mscorlib]System.Reflection.Assembly [mscorlib]System.Type::get_Assembly()
    IL_001b:  newobj     instance void [mscorlib]System.Resources.ResourceManager::.ctor(string,
                                                                                         class [mscorlib]System.Reflection.Assembly)
    IL_0020:  stsfld     class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::resourceMan
    IL_0025:  ldsfld     class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::resourceMan
    IL_002a:  ret
  } // end of method Resources::get_ResourceManager

  .method assembly hidebysig specialname static 
          class [mscorlib]System.Globalization.CultureInfo 
          get_Culture() cil managed
  {
    // Code size       6 (0x6)
    .maxstack  8
    IL_0000:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_0005:  ret
  } // end of method Resources::get_Culture

  .method assembly hidebysig specialname static 
          void  set_Culture(class [mscorlib]System.Globalization.CultureInfo 'value') cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  stsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_0006:  ret
  } // end of method Resources::set_Culture

  .method assembly hidebysig specialname static 
          string  get_API_ERROR() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "API_ERROR"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_API_ERROR

  .method assembly hidebysig specialname static 
          string  get_CHECK_SHEET_NAME() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "CHECK_SHEET_NAME"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_CHECK_SHEET_NAME

  .method assembly hidebysig specialname static 
          string  get_CheckSheetNameComplete() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "CheckSheetNameComplete"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_CheckSheetNameComplete

  .method assembly hidebysig specialname static 
          string  get_CheckSheetNameSuccess() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "CheckSheetNameSuccess"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_CheckSheetNameSuccess

  .method assembly hidebysig specialname static 
          string  get_EMPTY_SHEET() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "EMPTY_SHEET"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_EMPTY_SHEET

  .method assembly hidebysig specialname static 
          string  get_EXCEL_ERROR() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "EXCEL_ERROR"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_EXCEL_ERROR

  .method assembly hidebysig specialname static 
          string  get_ExcelMessage64() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "ExcelMessage64"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_ExcelMessage64

  .method assembly hidebysig specialname static 
          string  get_ExcelMessage65() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "ExcelMessage65"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_ExcelMessage65

  .method assembly hidebysig specialname static 
          string  get_ExcelProcessorMessage5() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "ExcelProcessorMessage5"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_ExcelProcessorMessage5

  .method assembly hidebysig specialname static 
          string  get_MENU_EXTRACTION() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "MENU_EXTRACTION"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_MENU_EXTRACTION

  .method assembly hidebysig specialname static 
          string  get_MENU_KIND() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "MENU_KIND"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_MENU_KIND

  .method assembly hidebysig specialname static 
          string  get_MENU_SETTING() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "MENU_SETTING"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_MENU_SETTING

  .method assembly hidebysig specialname static 
          string  get_MENU_TABLE() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "MENU_TABLE"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_MENU_TABLE

  .method assembly hidebysig specialname static 
          string  get_MULTI_PRINT() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "MULTI_PRINT"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_MULTI_PRINT

  .method assembly hidebysig specialname static 
          string  get_SHEET_NAME() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "SHEET_NAME"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_SHEET_NAME

  .method assembly hidebysig specialname static 
          string  get_SheetNameNoticeGridColumnNameHeader() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "SheetNameNoticeGridColumnNameHeader"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_SheetNameNoticeGridColumnNameHeader

  .method assembly hidebysig specialname static 
          string  get_SheetNameNoticeGridColumnSymbolHeader() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "SheetNameNoticeGridColumnSymbolHeader"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_SheetNameNoticeGridColumnSymbolHeader

  .method assembly hidebysig specialname static 
          string  get_SheetNameNoticeGridDesc() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "SheetNameNoticeGridDesc"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_SheetNameNoticeGridDesc

  .method assembly hidebysig specialname static 
          string  get_SheetNameNoticeMessageDesc() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "SheetNameNoticeMessageDesc"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_SheetNameNoticeMessageDesc

  .method assembly hidebysig specialname static 
          string  get_SheetNameNoticeWindow() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "SheetNameNoticeWindow"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_SheetNameNoticeWindow

  .method assembly hidebysig specialname static 
          string  get_TOO_MANY_SELECT() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  call       class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
    IL_0005:  ldstr      "TOO_MANY_SELECT"
    IL_000a:  ldsfld     class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::resourceCulture
    IL_000f:  callvirt   instance string [mscorlib]System.Resources.ResourceManager::GetString(string,
                                                                                               class [mscorlib]System.Globalization.CultureInfo)
    IL_0014:  ret
  } // end of method Resources::get_TOO_MANY_SELECT

  .property class [mscorlib]System.Resources.ResourceManager
          ResourceManager()
  {
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 02 00 00 00 00 00 ) 
    .get class [mscorlib]System.Resources.ResourceManager iReporterExcelAddIn2019.Properties.Resources::get_ResourceManager()
  } // end of property Resources::ResourceManager
  .property class [mscorlib]System.Globalization.CultureInfo
          Culture()
  {
    .custom instance void [System]System.ComponentModel.EditorBrowsableAttribute::.ctor(valuetype [System]System.ComponentModel.EditorBrowsableState) = ( 01 00 02 00 00 00 00 00 ) 
    .get class [mscorlib]System.Globalization.CultureInfo iReporterExcelAddIn2019.Properties.Resources::get_Culture()
    .set void iReporterExcelAddIn2019.Properties.Resources::set_Culture(class [mscorlib]System.Globalization.CultureInfo)
  } // end of property Resources::Culture
  .property string API_ERROR()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_API_ERROR()
  } // end of property Resources::API_ERROR
  .property string CHECK_SHEET_NAME()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_CHECK_SHEET_NAME()
  } // end of property Resources::CHECK_SHEET_NAME
  .property string CheckSheetNameComplete()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_CheckSheetNameComplete()
  } // end of property Resources::CheckSheetNameComplete
  .property string CheckSheetNameSuccess()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_CheckSheetNameSuccess()
  } // end of property Resources::CheckSheetNameSuccess
  .property string EMPTY_SHEET()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_EMPTY_SHEET()
  } // end of property Resources::EMPTY_SHEET
  .property string EXCEL_ERROR()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_EXCEL_ERROR()
  } // end of property Resources::EXCEL_ERROR
  .property string ExcelMessage64()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_ExcelMessage64()
  } // end of property Resources::ExcelMessage64
  .property string ExcelMessage65()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_ExcelMessage65()
  } // end of property Resources::ExcelMessage65
  .property string ExcelProcessorMessage5()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_ExcelProcessorMessage5()
  } // end of property Resources::ExcelProcessorMessage5
  .property string MENU_EXTRACTION()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_MENU_EXTRACTION()
  } // end of property Resources::MENU_EXTRACTION
  .property string MENU_KIND()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_MENU_KIND()
  } // end of property Resources::MENU_KIND
  .property string MENU_SETTING()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_MENU_SETTING()
  } // end of property Resources::MENU_SETTING
  .property string MENU_TABLE()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_MENU_TABLE()
  } // end of property Resources::MENU_TABLE
  .property string MULTI_PRINT()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_MULTI_PRINT()
  } // end of property Resources::MULTI_PRINT
  .property string SHEET_NAME()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_SHEET_NAME()
  } // end of property Resources::SHEET_NAME
  .property string SheetNameNoticeGridColumnNameHeader()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_SheetNameNoticeGridColumnNameHeader()
  } // end of property Resources::SheetNameNoticeGridColumnNameHeader
  .property string SheetNameNoticeGridColumnSymbolHeader()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_SheetNameNoticeGridColumnSymbolHeader()
  } // end of property Resources::SheetNameNoticeGridColumnSymbolHeader
  .property string SheetNameNoticeGridDesc()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_SheetNameNoticeGridDesc()
  } // end of property Resources::SheetNameNoticeGridDesc
  .property string SheetNameNoticeMessageDesc()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_SheetNameNoticeMessageDesc()
  } // end of property Resources::SheetNameNoticeMessageDesc
  .property string SheetNameNoticeWindow()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_SheetNameNoticeWindow()
  } // end of property Resources::SheetNameNoticeWindow
  .property string TOO_MANY_SELECT()
  {
    .get string iReporterExcelAddIn2019.Properties.Resources::get_TOO_MANY_SELECT()
  } // end of property Resources::TOO_MANY_SELECT
} // end of class iReporterExcelAddIn2019.Properties.Resources

.class private auto ansi sealed beforefieldinit iReporterExcelAddIn2019.Properties.Settings
       extends [System]System.Configuration.ApplicationSettingsBase
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [System]System.CodeDom.Compiler.GeneratedCodeAttribute::.ctor(string,
                                                                                      string) = ( 01 00 4B 4D 69 63 72 6F 73 6F 66 74 2E 56 69 73   // ..KMicrosoft.Vis
                                                                                                  75 61 6C 53 74 75 64 69 6F 2E 45 64 69 74 6F 72   // ualStudio.Editor
                                                                                                  73 2E 53 65 74 74 69 6E 67 73 44 65 73 69 67 6E   // s.SettingsDesign
                                                                                                  65 72 2E 53 65 74 74 69 6E 67 73 53 69 6E 67 6C   // er.SettingsSingl
                                                                                                  65 46 69 6C 65 47 65 6E 65 72 61 74 6F 72 08 31   // eFileGenerator.1
                                                                                                  32 2E 30 2E 30 2E 30 00 00 )                      // 2.0.0.0..
  .field private static class iReporterExcelAddIn2019.Properties.Settings defaultInstance
  .method public hidebysig specialname static 
          class iReporterExcelAddIn2019.Properties.Settings 
          get_Default() cil managed
  {
    // Code size       6 (0x6)
    .maxstack  8
    IL_0000:  ldsfld     class iReporterExcelAddIn2019.Properties.Settings iReporterExcelAddIn2019.Properties.Settings::defaultInstance
    IL_0005:  ret
  } // end of method Settings::get_Default

  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [System]System.Configuration.ApplicationSettingsBase::.ctor()
    IL_0006:  ret
  } // end of method Settings::.ctor

  .method private hidebysig specialname rtspecialname static 
          void  .cctor() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  newobj     instance void iReporterExcelAddIn2019.Properties.Settings::.ctor()
    IL_0005:  call       class [System]System.Configuration.SettingsBase [System]System.Configuration.SettingsBase::Synchronized(class [System]System.Configuration.SettingsBase)
    IL_000a:  castclass  iReporterExcelAddIn2019.Properties.Settings
    IL_000f:  stsfld     class iReporterExcelAddIn2019.Properties.Settings iReporterExcelAddIn2019.Properties.Settings::defaultInstance
    IL_0014:  ret
  } // end of method Settings::.cctor

  .property class iReporterExcelAddIn2019.Properties.Settings
          Default()
  {
    .get class iReporterExcelAddIn2019.Properties.Settings iReporterExcelAddIn2019.Properties.Settings::get_Default()
  } // end of property Settings::Default
} // end of class iReporterExcelAddIn2019.Properties.Settings

.class interface public abstract auto ansi import Microsoft.Office.Core.CommandBar
       implements Microsoft.Office.Core._IMsoOleAccDispObj
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 43 30 33 30 34 2D 30 30 30 30   // ..$000C0304-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_26() runtime managed
  {
  } // end of method CommandBar::_VtblGap1_26

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Core.CommandBarControls 
          marshal( interface ) 
          get_Controls() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 03 00 04 60 00 00 )                         // .....`..
  } // end of method CommandBar::get_Controls

  .method public specialname rtspecialname 
          instance void  _VtblGap2_10() runtime managed
  {
  } // end of method CommandBar::_VtblGap2_10

  .method public hidebysig newslot specialname abstract virtual 
          instance string 
          marshal( bstr) 
          get_Name() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 0E 00 04 60 00 00 )                         // .....`..
  } // end of method CommandBar::get_Name

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set_Name([in] string  marshal( bstr) pbstrName) runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 0E 00 04 60 00 00 )                         // .....`..
  } // end of method CommandBar::set_Name

  .property class Microsoft.Office.Core.CommandBarControls
          Controls()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 03 00 04 60 00 00 )                         // .....`..
    .get instance class Microsoft.Office.Core.CommandBarControls Microsoft.Office.Core.CommandBar::get_Controls()
  } // end of property CommandBar::Controls
  .property string Name()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 0E 00 04 60 00 00 )                         // .....`..
    .get instance string Microsoft.Office.Core.CommandBar::get_Name()
    .set instance void Microsoft.Office.Core.CommandBar::set_Name(string)
  } // end of property CommandBar::Name
} // end of class Microsoft.Office.Core.CommandBar

.class interface public abstract auto ansi import Microsoft.Office.Core.CommandBarButton
       implements Microsoft.Office.Core._CommandBarButton,
                  Microsoft.Office.Core._CommandBarButtonEvents_Event
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 43 30 33 30 45 2D 30 30 30 30   // ..$000C030E-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.CoClassAttribute::.ctor(class [mscorlib]System.Type) = ( 01 00 5A 53 79 73 74 65 6D 2E 4F 62 6A 65 63 74   // ..ZSystem.Object
                                                                                                                          2C 20 6D 73 63 6F 72 6C 69 62 2C 20 56 65 72 73   // , mscorlib, Vers
                                                                                                                          69 6F 6E 3D 34 2E 30 2E 30 2E 30 2C 20 43 75 6C   // ion=4.0.0.0, Cul
                                                                                                                          74 75 72 65 3D 6E 65 75 74 72 61 6C 2C 20 50 75   // ture=neutral, Pu
                                                                                                                          62 6C 69 63 4B 65 79 54 6F 6B 65 6E 3D 62 37 37   // blicKeyToken=b77
                                                                                                                          61 35 63 35 36 31 39 33 34 65 30 38 39 00 00 )    // a5c561934e089..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Core.CommandBarButton

.class interface public abstract auto ansi import Microsoft.Office.Core.CommandBarControl
       implements Microsoft.Office.Core._IMsoOleAccDispObj
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 43 30 33 30 38 2D 30 30 30 30   // ..$000C0308-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Core.CommandBarControl

.class interface public abstract auto ansi import Microsoft.Office.Core.CommandBarControls
       implements Microsoft.Office.Core._IMsoDispObj,
                  [mscorlib]System.Collections.IEnumerable
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 43 30 33 30 36 2D 30 30 30 30   // ..$000C0306-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Reflection.DefaultMemberAttribute::.ctor(string) = ( 01 00 04 49 74 65 6D 00 00 )                      // ...Item..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_2() runtime managed
  {
  } // end of method CommandBarControls::_VtblGap1_2

  .method public hidebysig newslot abstract virtual 
          instance class Microsoft.Office.Core.CommandBarControl 
          marshal( interface ) 
          Add([in][opt] object  marshal( struct) Type,
              [in][opt] object  marshal( struct) Id,
              [in][opt] object  marshal( struct) Parameter,
              [in][opt] object  marshal( struct) Before,
              [in][opt] object  marshal( struct) Temporary) runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 00 00 03 60 00 00 )                         // .....`..
  } // end of method CommandBarControls::Add

} // end of class Microsoft.Office.Core.CommandBarControls

.class interface public abstract auto ansi import Microsoft.Office.Core.CommandBars
       implements Microsoft.Office.Core._CommandBars,
                  Microsoft.Office.Core._CommandBarsEvents_Event
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.CoClassAttribute::.ctor(class [mscorlib]System.Type) = ( 01 00 5A 53 79 73 74 65 6D 2E 4F 62 6A 65 63 74   // ..ZSystem.Object
                                                                                                                          2C 20 6D 73 63 6F 72 6C 69 62 2C 20 56 65 72 73   // , mscorlib, Vers
                                                                                                                          69 6F 6E 3D 34 2E 30 2E 30 2E 30 2C 20 43 75 6C   // ion=4.0.0.0, Cul
                                                                                                                          74 75 72 65 3D 6E 65 75 74 72 61 6C 2C 20 50 75   // ture=neutral, Pu
                                                                                                                          62 6C 69 63 4B 65 79 54 6F 6B 65 6E 3D 62 37 37   // blicKeyToken=b77
                                                                                                                          61 35 63 35 36 31 39 33 34 65 30 38 39 00 00 )    // a5c561934e089..
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 43 30 33 30 32 2D 30 30 30 30   // ..$000C0302-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Core.CommandBars

.class interface public abstract auto ansi import Microsoft.Office.Core.IAccessible
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 36 31 38 37 33 36 45 30 2D 33 43 33 44   // ..$618736E0-3C3D
                                                                                                  2D 31 31 43 46 2D 38 31 30 43 2D 30 30 41 41 30   // -11CF-810C-00AA0
                                                                                                  30 33 38 39 42 37 31 00 00 )                      // 0389B71..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Core.IAccessible

.class public auto ansi sealed Microsoft.Office.Core.MsoButtonStyle
       extends [mscorlib]System.Enum
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 32 64 66 38 64 30 34 63 2D 35 62 66 61   // ..$2df8d04c-5bfa
                                                                                                            2D 31 30 31 62 2D 62 64 65 35 2D 30 30 61 61 30   // -101b-bde5-00aa0
                                                                                                            30 34 34 64 65 35 32 24 4D 69 63 72 6F 73 6F 66   // 044de52$Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 43 6F 72 65 2E 4D 73   // t.Office.Core.Ms
                                                                                                            6F 42 75 74 74 6F 6E 53 74 79 6C 65 00 00 )       // oButtonStyle..
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype Microsoft.Office.Core.MsoButtonStyle msoButtonAutomatic = int32(0x00000000)
  .field public static literal valuetype Microsoft.Office.Core.MsoButtonStyle msoButtonIcon = int32(0x00000001)
  .field public static literal valuetype Microsoft.Office.Core.MsoButtonStyle msoButtonCaption = int32(0x00000002)
  .field public static literal valuetype Microsoft.Office.Core.MsoButtonStyle msoButtonIconAndCaption = int32(0x00000003)
  .field public static literal valuetype Microsoft.Office.Core.MsoButtonStyle msoButtonIconAndWrapCaption = int32(0x00000007)
  .field public static literal valuetype Microsoft.Office.Core.MsoButtonStyle msoButtonIconAndCaptionBelow = int32(0x0000000B)
  .field public static literal valuetype Microsoft.Office.Core.MsoButtonStyle msoButtonWrapCaption = int32(0x0000000E)
  .field public static literal valuetype Microsoft.Office.Core.MsoButtonStyle msoButtonIconAndWrapCaptionBelow = int32(0x0000000F)
} // end of class Microsoft.Office.Core.MsoButtonStyle

.class public auto ansi sealed Microsoft.Office.Core.MsoControlType
       extends [mscorlib]System.Enum
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 32 64 66 38 64 30 34 63 2D 35 62 66 61   // ..$2df8d04c-5bfa
                                                                                                            2D 31 30 31 62 2D 62 64 65 35 2D 30 30 61 61 30   // -101b-bde5-00aa0
                                                                                                            30 34 34 64 65 35 32 24 4D 69 63 72 6F 73 6F 66   // 044de52$Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 43 6F 72 65 2E 4D 73   // t.Office.Core.Ms
                                                                                                            6F 43 6F 6E 74 72 6F 6C 54 79 70 65 00 00 )       // oControlType..
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlCustom = int32(0x00000000)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlButton = int32(0x00000001)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlEdit = int32(0x00000002)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlDropdown = int32(0x00000003)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlComboBox = int32(0x00000004)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlButtonDropdown = int32(0x00000005)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlSplitDropdown = int32(0x00000006)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlOCXDropdown = int32(0x00000007)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlGenericDropdown = int32(0x00000008)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlGraphicDropdown = int32(0x00000009)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlPopup = int32(0x0000000A)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlGraphicPopup = int32(0x0000000B)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlButtonPopup = int32(0x0000000C)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlSplitButtonPopup = int32(0x0000000D)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlSplitButtonMRUPopup = int32(0x0000000E)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlLabel = int32(0x0000000F)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlExpandingGrid = int32(0x00000010)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlSplitExpandingGrid = int32(0x00000011)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlGrid = int32(0x00000012)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlGauge = int32(0x00000013)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlGraphicCombo = int32(0x00000014)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlPane = int32(0x00000015)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlActiveX = int32(0x00000016)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlSpinner = int32(0x00000017)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlLabelEx = int32(0x00000018)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlWorkPane = int32(0x00000019)
  .field public static literal valuetype Microsoft.Office.Core.MsoControlType msoControlAutoCompleteCombo = int32(0x0000001A)
} // end of class Microsoft.Office.Core.MsoControlType

.class public auto ansi sealed Microsoft.Office.Core.RibbonControlSize
       extends [mscorlib]System.Enum
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 32 64 66 38 64 30 34 63 2D 35 62 66 61   // ..$2df8d04c-5bfa
                                                                                                            2D 31 30 31 62 2D 62 64 65 35 2D 30 30 61 61 30   // -101b-bde5-00aa0
                                                                                                            30 34 34 64 65 35 32 27 4D 69 63 72 6F 73 6F 66   // 044de52'Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 43 6F 72 65 2E 52 69   // t.Office.Core.Ri
                                                                                                            62 62 6F 6E 43 6F 6E 74 72 6F 6C 53 69 7A 65 00   // bbonControlSize.
                                                                                                            00 ) 
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype Microsoft.Office.Core.RibbonControlSize RibbonControlSizeRegular = int32(0x00000000)
  .field public static literal valuetype Microsoft.Office.Core.RibbonControlSize RibbonControlSizeLarge = int32(0x00000001)
} // end of class Microsoft.Office.Core.RibbonControlSize

.class interface public abstract auto ansi import Microsoft.Office.Core._CommandBarButton
       implements Microsoft.Office.Core.CommandBarControl
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 43 30 33 30 45 2D 30 30 30 30   // ..$000C030E-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_26() runtime managed
  {
  } // end of method _CommandBarButton::_VtblGap1_26

  .method public hidebysig newslot specialname abstract virtual 
          instance string 
          marshal( bstr) 
          get_Caption() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 03 00 04 60 00 00 )                         // .....`..
  } // end of method _CommandBarButton::get_Caption

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set_Caption([in] string  marshal( bstr) pbstrCaption) runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 03 00 04 60 00 00 )                         // .....`..
  } // end of method _CommandBarButton::set_Caption

  .method public specialname rtspecialname 
          instance void  _VtblGap2_30() runtime managed
  {
  } // end of method _CommandBarButton::_VtblGap2_30

  .method public hidebysig newslot specialname abstract virtual 
          instance string 
          marshal( bstr) 
          get_Tag() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 23 00 04 60 00 00 )                         // ..#..`..
  } // end of method _CommandBarButton::get_Tag

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set_Tag([in] string  marshal( bstr) pbstrTag) runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 23 00 04 60 00 00 )                         // ..#..`..
  } // end of method _CommandBarButton::set_Tag

  .method public specialname rtspecialname 
          instance void  _VtblGap3_26() runtime managed
  {
  } // end of method _CommandBarButton::_VtblGap3_26

  .method public hidebysig newslot specialname abstract virtual 
          instance valuetype Microsoft.Office.Core.MsoButtonStyle 
          get_Style() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 0A 00 05 60 00 00 )                         // .....`..
  } // end of method _CommandBarButton::get_Style

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set_Style([in] valuetype Microsoft.Office.Core.MsoButtonStyle pstyle) runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 0A 00 05 60 00 00 )                         // .....`..
  } // end of method _CommandBarButton::set_Style

  .property string Caption()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 03 00 04 60 00 00 )                         // .....`..
    .get instance string Microsoft.Office.Core._CommandBarButton::get_Caption()
    .set instance void Microsoft.Office.Core._CommandBarButton::set_Caption(string)
  } // end of property _CommandBarButton::Caption
  .property string Tag()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 23 00 04 60 00 00 )                         // ..#..`..
    .get instance string Microsoft.Office.Core._CommandBarButton::get_Tag()
    .set instance void Microsoft.Office.Core._CommandBarButton::set_Tag(string)
  } // end of property _CommandBarButton::Tag
  .property valuetype Microsoft.Office.Core.MsoButtonStyle
          Style()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 0A 00 05 60 00 00 )                         // .....`..
    .get instance valuetype Microsoft.Office.Core.MsoButtonStyle Microsoft.Office.Core._CommandBarButton::get_Style()
    .set instance void Microsoft.Office.Core._CommandBarButton::set_Style(valuetype Microsoft.Office.Core.MsoButtonStyle)
  } // end of property _CommandBarButton::Style
} // end of class Microsoft.Office.Core._CommandBarButton

.class interface public abstract auto ansi import Microsoft.Office.Core._CommandBarButtonEvents
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 43 30 33 35 31 2D 30 30 30 30   // ..$000C0351-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  Click([in] class Microsoft.Office.Core.CommandBarButton  marshal( interface ) Ctrl,
                               [in][out] bool& CancelDefault) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 01 00 00 00 00 00 ) 
  } // end of method _CommandBarButtonEvents::Click

} // end of class Microsoft.Office.Core._CommandBarButtonEvents

.class public auto ansi sealed Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler
       extends [mscorlib]System.MulticastDelegate
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 32 64 66 38 64 30 34 63 2D 35 62 66 61   // ..$2df8d04c-5bfa
                                                                                                            2D 31 30 31 62 2D 62 64 65 35 2D 30 30 61 61 30   // -101b-bde5-00aa0
                                                                                                            30 34 34 64 65 35 32 3F 4D 69 63 72 6F 73 6F 66   // 044de52?Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 43 6F 72 65 2E 5F 43   // t.Office.Core._C
                                                                                                            6F 6D 6D 61 6E 64 42 61 72 42 75 74 74 6F 6E 45   // ommandBarButtonE
                                                                                                            76 65 6E 74 73 5F 43 6C 69 63 6B 45 76 65 6E 74   // vents_ClickEvent
                                                                                                            48 61 6E 64 6C 65 72 00 00 )                      // Handler..
  .method public specialname rtspecialname 
          instance void  .ctor(object A_1,
                               native uint A_2) runtime managed
  {
  } // end of method _CommandBarButtonEvents_ClickEventHandler::.ctor

  .method public virtual instance void  Invoke([in] class Microsoft.Office.Core.CommandBarButton  marshal( interface ) Ctrl,
                                               [in][out] bool& CancelDefault) runtime managed preservesig
  {
  } // end of method _CommandBarButtonEvents_ClickEventHandler::Invoke

} // end of class Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler

.class interface public abstract auto ansi import Microsoft.Office.Core._CommandBarButtonEvents_Event
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComEventInterfaceAttribute::.ctor(class [mscorlib]System.Type,
                                                                                                   class [mscorlib]System.Type) = ( 01 00 2D 4D 69 63 72 6F 73 6F 66 74 2E 4F 66 66   // ..-Microsoft.Off
                                                                                                                                    69 63 65 2E 43 6F 72 65 2E 5F 43 6F 6D 6D 61 6E   // ice.Core._Comman
                                                                                                                                    64 42 61 72 42 75 74 74 6F 6E 45 76 65 6E 74 73   // dBarButtonEvents
                                                                                                                                    2D 4D 69 63 72 6F 73 6F 66 74 2E 4F 66 66 69 63   // -Microsoft.Offic
                                                                                                                                    65 2E 43 6F 72 65 2E 5F 43 6F 6D 6D 61 6E 64 42   // e.Core._CommandB
                                                                                                                                    61 72 42 75 74 74 6F 6E 45 76 65 6E 74 73 00 00 ) // arButtonEvents..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 32 64 66 38 64 30 34 63 2D 35 62 66 61   // ..$2df8d04c-5bfa
                                                                                                            2D 31 30 31 62 2D 62 64 65 35 2D 30 30 61 61 30   // -101b-bde5-00aa0
                                                                                                            30 34 34 64 65 35 32 33 4D 69 63 72 6F 73 6F 66   // 044de523Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 43 6F 72 65 2E 5F 43   // t.Office.Core._C
                                                                                                            6F 6D 6D 61 6E 64 42 61 72 42 75 74 74 6F 6E 45   // ommandBarButtonE
                                                                                                            76 65 6E 74 73 5F 45 76 65 6E 74 00 00 )          // vents_Event..
  .method public hidebysig newslot abstract virtual 
          instance void  add_Click(class Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler A_1) runtime managed internalcall
  {
  } // end of method _CommandBarButtonEvents_Event::add_Click

  .method public hidebysig newslot abstract virtual 
          instance void  remove_Click(class Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler A_1) runtime managed internalcall
  {
  } // end of method _CommandBarButtonEvents_Event::remove_Click

  .event Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler Click
  {
    .addon instance void Microsoft.Office.Core._CommandBarButtonEvents_Event::add_Click(class Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler)
    .removeon instance void Microsoft.Office.Core._CommandBarButtonEvents_Event::remove_Click(class Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler)
  } // end of event _CommandBarButtonEvents_Event::Click
} // end of class Microsoft.Office.Core._CommandBarButtonEvents_Event

.class interface public abstract auto ansi import Microsoft.Office.Core._CommandBars
       implements Microsoft.Office.Core._IMsoDispObj,
                  [mscorlib]System.Collections.IEnumerable
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 43 30 33 30 32 2D 30 30 30 30   // ..$000C0302-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Reflection.DefaultMemberAttribute::.ctor(string) = ( 01 00 04 49 74 65 6D 00 00 )                      // ...Item..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_16() runtime managed
  {
  } // end of method _CommandBars::_VtblGap1_16

  .method public hidebysig newslot abstract virtual 
          instance class [mscorlib]System.Collections.IEnumerator 
          marshal( custom ("System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a","")) 
          GetEnumerator() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 FC FF FF FF 00 00 ) 
  } // end of method _CommandBars::GetEnumerator

} // end of class Microsoft.Office.Core._CommandBars

.class interface public abstract auto ansi import Microsoft.Office.Core._CommandBarsEvents
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 43 30 33 35 32 2D 30 30 30 30   // ..$000C0352-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Core._CommandBarsEvents

.class interface public abstract auto ansi import Microsoft.Office.Core._CommandBarsEvents_Event
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComEventInterfaceAttribute::.ctor(class [mscorlib]System.Type,
                                                                                                   class [mscorlib]System.Type) = ( 01 00 28 4D 69 63 72 6F 73 6F 66 74 2E 4F 66 66   // ..(Microsoft.Off
                                                                                                                                    69 63 65 2E 43 6F 72 65 2E 5F 43 6F 6D 6D 61 6E   // ice.Core._Comman
                                                                                                                                    64 42 61 72 73 45 76 65 6E 74 73 28 4D 69 63 72   // dBarsEvents(Micr
                                                                                                                                    6F 73 6F 66 74 2E 4F 66 66 69 63 65 2E 43 6F 72   // osoft.Office.Cor
                                                                                                                                    65 2E 5F 43 6F 6D 6D 61 6E 64 42 61 72 73 45 76   // e._CommandBarsEv
                                                                                                                                    65 6E 74 73 00 00 )                               // ents..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 32 64 66 38 64 30 34 63 2D 35 62 66 61   // ..$2df8d04c-5bfa
                                                                                                            2D 31 30 31 62 2D 62 64 65 35 2D 30 30 61 61 30   // -101b-bde5-00aa0
                                                                                                            30 34 34 64 65 35 32 2E 4D 69 63 72 6F 73 6F 66   // 044de52.Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 43 6F 72 65 2E 5F 43   // t.Office.Core._C
                                                                                                            6F 6D 6D 61 6E 64 42 61 72 73 45 76 65 6E 74 73   // ommandBarsEvents
                                                                                                            5F 45 76 65 6E 74 00 00 )                         // _Event..
} // end of class Microsoft.Office.Core._CommandBarsEvents_Event

.class interface public abstract auto ansi import Microsoft.Office.Core._IMsoDispObj
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 43 30 33 30 30 2D 30 30 30 30   // ..$000C0300-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Core._IMsoDispObj

.class interface public abstract auto ansi import Microsoft.Office.Core._IMsoOleAccDispObj
       implements Microsoft.Office.Core.IAccessible
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 43 30 33 30 31 2D 30 30 30 30   // ..$000C0301-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Core._IMsoOleAccDispObj

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.AppEvents
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 34 34 31 33 2D 30 30 30 30   // ..$00024413-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_1() runtime managed
  {
  } // end of method AppEvents::_VtblGap1_1

  .method public hidebysig newslot abstract virtual 
          instance void  SheetSelectionChange([in] object  marshal( idispatch ) Sh,
                                              [in] class Microsoft.Office.Interop.Excel.Range  marshal( interface ) Target) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 16 06 00 00 00 00 ) 
  } // end of method AppEvents::SheetSelectionChange

  .method public specialname rtspecialname 
          instance void  _VtblGap2_2() runtime managed
  {
  } // end of method AppEvents::_VtblGap2_2

  .method public hidebysig newslot abstract virtual 
          instance void  SheetActivate([in] object  marshal( idispatch ) Sh) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 19 06 00 00 00 00 ) 
  } // end of method AppEvents::SheetActivate

  .method public specialname rtspecialname 
          instance void  _VtblGap3_2() runtime managed
  {
  } // end of method AppEvents::_VtblGap3_2

  .method public hidebysig newslot abstract virtual 
          instance void  SheetChange([in] object  marshal( idispatch ) Sh,
                                     [in] class Microsoft.Office.Interop.Excel.Range  marshal( interface ) Target) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 1C 06 00 00 00 00 ) 
  } // end of method AppEvents::SheetChange

} // end of class Microsoft.Office.Interop.Excel.AppEvents

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.AppEvents_Event
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComEventInterfaceAttribute::.ctor(class [mscorlib]System.Type,
                                                                                                   class [mscorlib]System.Type) = ( 01 00 28 4D 69 63 72 6F 73 6F 66 74 2E 4F 66 66   // ..(Microsoft.Off
                                                                                                                                    69 63 65 2E 49 6E 74 65 72 6F 70 2E 45 78 63 65   // ice.Interop.Exce
                                                                                                                                    6C 2E 41 70 70 45 76 65 6E 74 73 28 4D 69 63 72   // l.AppEvents(Micr
                                                                                                                                    6F 73 6F 66 74 2E 4F 66 66 69 63 65 2E 49 6E 74   // osoft.Office.Int
                                                                                                                                    65 72 6F 70 2E 45 78 63 65 6C 2E 41 70 70 45 76   // erop.Excel.AppEv
                                                                                                                                    65 6E 74 73 00 00 )                               // ents..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 30 30 30 32 30 38 31 33 2D 30 30 30 30   // ..$00020813-0000
                                                                                                            2D 30 30 30 30 2D 63 30 30 30 2D 30 30 30 30 30   // -0000-c000-00000
                                                                                                            30 30 30 30 30 34 36 2E 4D 69 63 72 6F 73 6F 66   // 0000046.Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 49 6E 74 65 72 6F 70   // t.Office.Interop
                                                                                                            2E 45 78 63 65 6C 2E 41 70 70 45 76 65 6E 74 73   // .Excel.AppEvents
                                                                                                            5F 45 76 65 6E 74 00 00 )                         // _Event..
  .method public specialname rtspecialname 
          instance void  _VtblGap1_2() runtime managed
  {
  } // end of method AppEvents_Event::_VtblGap1_2

  .method public hidebysig newslot abstract virtual 
          instance void  add_SheetSelectionChange(class Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler A_1) runtime managed internalcall
  {
  } // end of method AppEvents_Event::add_SheetSelectionChange

  .method public hidebysig newslot abstract virtual 
          instance void  remove_SheetSelectionChange(class Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler A_1) runtime managed internalcall
  {
  } // end of method AppEvents_Event::remove_SheetSelectionChange

  .method public specialname rtspecialname 
          instance void  _VtblGap2_4() runtime managed
  {
  } // end of method AppEvents_Event::_VtblGap2_4

  .method public hidebysig newslot abstract virtual 
          instance void  add_SheetActivate(class Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler A_1) runtime managed internalcall
  {
  } // end of method AppEvents_Event::add_SheetActivate

  .method public hidebysig newslot abstract virtual 
          instance void  remove_SheetActivate(class Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler A_1) runtime managed internalcall
  {
  } // end of method AppEvents_Event::remove_SheetActivate

  .method public specialname rtspecialname 
          instance void  _VtblGap3_4() runtime managed
  {
  } // end of method AppEvents_Event::_VtblGap3_4

  .method public hidebysig newslot abstract virtual 
          instance void  add_SheetChange(class Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler A_1) runtime managed internalcall
  {
  } // end of method AppEvents_Event::add_SheetChange

  .method public hidebysig newslot abstract virtual 
          instance void  remove_SheetChange(class Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler A_1) runtime managed internalcall
  {
  } // end of method AppEvents_Event::remove_SheetChange

  .event Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler SheetSelectionChange
  {
    .addon instance void Microsoft.Office.Interop.Excel.AppEvents_Event::add_SheetSelectionChange(class Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler)
    .removeon instance void Microsoft.Office.Interop.Excel.AppEvents_Event::remove_SheetSelectionChange(class Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler)
  } // end of event AppEvents_Event::SheetSelectionChange
  .event Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler SheetActivate
  {
    .addon instance void Microsoft.Office.Interop.Excel.AppEvents_Event::add_SheetActivate(class Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler)
    .removeon instance void Microsoft.Office.Interop.Excel.AppEvents_Event::remove_SheetActivate(class Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler)
  } // end of event AppEvents_Event::SheetActivate
  .event Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler SheetChange
  {
    .addon instance void Microsoft.Office.Interop.Excel.AppEvents_Event::add_SheetChange(class Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler)
    .removeon instance void Microsoft.Office.Interop.Excel.AppEvents_Event::remove_SheetChange(class Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler)
  } // end of event AppEvents_Event::SheetChange
} // end of class Microsoft.Office.Interop.Excel.AppEvents_Event

.class public auto ansi sealed Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler
       extends [mscorlib]System.MulticastDelegate
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 30 30 30 32 30 38 31 33 2D 30 30 30 30   // ..$00020813-0000
                                                                                                            2D 30 30 30 30 2D 63 30 30 30 2D 30 30 30 30 30   // -0000-c000-00000
                                                                                                            30 30 30 30 30 34 36 42 4D 69 63 72 6F 73 6F 66   // 0000046BMicrosof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 49 6E 74 65 72 6F 70   // t.Office.Interop
                                                                                                            2E 45 78 63 65 6C 2E 41 70 70 45 76 65 6E 74 73   // .Excel.AppEvents
                                                                                                            5F 53 68 65 65 74 41 63 74 69 76 61 74 65 45 76   // _SheetActivateEv
                                                                                                            65 6E 74 48 61 6E 64 6C 65 72 00 00 )             // entHandler..
  .method public specialname rtspecialname 
          instance void  .ctor(object A_1,
                               native uint A_2) runtime managed
  {
  } // end of method AppEvents_SheetActivateEventHandler::.ctor

  .method public virtual instance void  Invoke([in] object  marshal( idispatch ) Sh) runtime managed preservesig
  {
  } // end of method AppEvents_SheetActivateEventHandler::Invoke

} // end of class Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler

.class public auto ansi sealed Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler
       extends [mscorlib]System.MulticastDelegate
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 30 30 30 32 30 38 31 33 2D 30 30 30 30   // ..$00020813-0000
                                                                                                            2D 30 30 30 30 2D 63 30 30 30 2D 30 30 30 30 30   // -0000-c000-00000
                                                                                                            30 30 30 30 30 34 36 40 4D 69 63 72 6F 73 6F 66   // 0000046@Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 49 6E 74 65 72 6F 70   // t.Office.Interop
                                                                                                            2E 45 78 63 65 6C 2E 41 70 70 45 76 65 6E 74 73   // .Excel.AppEvents
                                                                                                            5F 53 68 65 65 74 43 68 61 6E 67 65 45 76 65 6E   // _SheetChangeEven
                                                                                                            74 48 61 6E 64 6C 65 72 00 00 )                   // tHandler..
  .method public specialname rtspecialname 
          instance void  .ctor(object A_1,
                               native uint A_2) runtime managed
  {
  } // end of method AppEvents_SheetChangeEventHandler::.ctor

  .method public virtual instance void  Invoke([in] object  marshal( idispatch ) Sh,
                                               [in] class Microsoft.Office.Interop.Excel.Range  marshal( interface ) Target) runtime managed preservesig
  {
  } // end of method AppEvents_SheetChangeEventHandler::Invoke

} // end of class Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler

.class public auto ansi sealed Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler
       extends [mscorlib]System.MulticastDelegate
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 30 30 30 32 30 38 31 33 2D 30 30 30 30   // ..$00020813-0000
                                                                                                            2D 30 30 30 30 2D 63 30 30 30 2D 30 30 30 30 30   // -0000-c000-00000
                                                                                                            30 30 30 30 30 34 36 49 4D 69 63 72 6F 73 6F 66   // 0000046IMicrosof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 49 6E 74 65 72 6F 70   // t.Office.Interop
                                                                                                            2E 45 78 63 65 6C 2E 41 70 70 45 76 65 6E 74 73   // .Excel.AppEvents
                                                                                                            5F 53 68 65 65 74 53 65 6C 65 63 74 69 6F 6E 43   // _SheetSelectionC
                                                                                                            68 61 6E 67 65 45 76 65 6E 74 48 61 6E 64 6C 65   // hangeEventHandle
                                                                                                            72 00 00 )                                        // r..
  .method public specialname rtspecialname 
          instance void  .ctor(object A_1,
                               native uint A_2) runtime managed
  {
  } // end of method AppEvents_SheetSelectionChangeEventHandler::.ctor

  .method public virtual instance void  Invoke([in] object  marshal( idispatch ) Sh,
                                               [in] class Microsoft.Office.Interop.Excel.Range  marshal( interface ) Target) runtime managed preservesig
  {
  } // end of method AppEvents_SheetSelectionChangeEventHandler::Invoke

} // end of class Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Application
       implements Microsoft.Office.Interop.Excel._Application,
                  Microsoft.Office.Interop.Excel.AppEvents_Event
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 44 35 2D 30 30 30 30   // ..$000208D5-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.CoClassAttribute::.ctor(class [mscorlib]System.Type) = ( 01 00 5A 53 79 73 74 65 6D 2E 4F 62 6A 65 63 74   // ..ZSystem.Object
                                                                                                                          2C 20 6D 73 63 6F 72 6C 69 62 2C 20 56 65 72 73   // , mscorlib, Vers
                                                                                                                          69 6F 6E 3D 34 2E 30 2E 30 2E 30 2C 20 43 75 6C   // ion=4.0.0.0, Cul
                                                                                                                          74 75 72 65 3D 6E 65 75 74 72 61 6C 2C 20 50 75   // ture=neutral, Pu
                                                                                                                          62 6C 69 63 4B 65 79 54 6F 6B 65 6E 3D 62 37 37   // blicKeyToken=b77
                                                                                                                          61 35 63 35 36 31 39 33 34 65 30 38 39 00 00 )    // a5c561934e089..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Interop.Excel.Application

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Comment
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 34 34 32 37 2D 30 30 30 30   // ..$00024427-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_2() runtime managed
  {
  } // end of method Comment::_VtblGap1_2

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( idispatch ) 
          get_Parent() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 96 00 00 00 00 00 ) 
  } // end of method Comment::get_Parent

  .method public specialname rtspecialname 
          instance void  _VtblGap2_4() runtime managed
  {
  } // end of method Comment::_VtblGap2_4

  .method public hidebysig newslot abstract virtual 
          instance string 
          marshal( bstr) 
          Text([in][opt] object  marshal( struct) Text,
               [in][opt] object  marshal( struct) Start,
               [in][opt] object  marshal( struct) Overwrite) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 8A 00 00 00 00 00 ) 
  } // end of method Comment::Text

  .property object Parent()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 96 00 00 00 00 00 ) 
    .get instance object Microsoft.Office.Interop.Excel.Comment::get_Parent()
  } // end of property Comment::Parent
} // end of class Microsoft.Office.Interop.Excel.Comment

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Comments
       implements [mscorlib]System.Collections.IEnumerable
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Reflection.DefaultMemberAttribute::.ctor(string) = ( 01 00 08 5F 44 65 66 61 75 6C 74 00 00 )          // ..._Default..
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 34 34 32 36 2D 30 30 30 30   // ..$00024426-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_3() runtime managed
  {
  } // end of method Comments::_VtblGap1_3

  .method public hidebysig newslot specialname abstract virtual 
          instance int32  get_Count() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 76 00 00 00 00 00 )                         // ..v.....
  } // end of method Comments::get_Count

  .method public specialname rtspecialname 
          instance void  _VtblGap2_1() runtime managed
  {
  } // end of method Comments::_VtblGap2_1

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Comment 
          marshal( interface ) 
          get__Default([in] int32 Index) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
  } // end of method Comments::get__Default

  .property int32 Count()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 76 00 00 00 00 00 )                         // ..v.....
    .get instance int32 Microsoft.Office.Interop.Excel.Comments::get_Count()
  } // end of property Comments::Count
  .property class Microsoft.Office.Interop.Excel.Comment
          _Default(int32)
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Comments::get__Default(int32)
  } // end of property Comments::_Default
} // end of class Microsoft.Office.Interop.Excel.Comments

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.DocEvents
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 34 34 31 31 2D 30 30 30 30   // ..$00024411-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Interop.Excel.DocEvents

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.DocEvents_Event
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComEventInterfaceAttribute::.ctor(class [mscorlib]System.Type,
                                                                                                   class [mscorlib]System.Type) = ( 01 00 28 4D 69 63 72 6F 73 6F 66 74 2E 4F 66 66   // ..(Microsoft.Off
                                                                                                                                    69 63 65 2E 49 6E 74 65 72 6F 70 2E 45 78 63 65   // ice.Interop.Exce
                                                                                                                                    6C 2E 44 6F 63 45 76 65 6E 74 73 28 4D 69 63 72   // l.DocEvents(Micr
                                                                                                                                    6F 73 6F 66 74 2E 4F 66 66 69 63 65 2E 49 6E 74   // osoft.Office.Int
                                                                                                                                    65 72 6F 70 2E 45 78 63 65 6C 2E 44 6F 63 45 76   // erop.Excel.DocEv
                                                                                                                                    65 6E 74 73 00 00 )                               // ents..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 30 30 30 32 30 38 31 33 2D 30 30 30 30   // ..$00020813-0000
                                                                                                            2D 30 30 30 30 2D 63 30 30 30 2D 30 30 30 30 30   // -0000-c000-00000
                                                                                                            30 30 30 30 30 34 36 2E 4D 69 63 72 6F 73 6F 66   // 0000046.Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 49 6E 74 65 72 6F 70   // t.Office.Interop
                                                                                                            2E 45 78 63 65 6C 2E 44 6F 63 45 76 65 6E 74 73   // .Excel.DocEvents
                                                                                                            5F 45 76 65 6E 74 00 00 )                         // _Event..
} // end of class Microsoft.Office.Interop.Excel.DocEvents_Event

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Interior
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 37 30 2D 30 30 30 30   // ..$00020870-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_3() runtime managed
  {
  } // end of method Interior::_VtblGap1_3

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( struct) 
          get_Color() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 63 00 00 00 00 00 )                         // ..c.....
  } // end of method Interior::get_Color

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set_Color([in] object  marshal( struct) A_1) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 63 00 00 00 00 00 )                         // ..c.....
  } // end of method Interior::set_Color

  .property object Color()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 63 00 00 00 00 00 )                         // ..c.....
    .get instance object Microsoft.Office.Interop.Excel.Interior::get_Color()
    .set instance void Microsoft.Office.Interop.Excel.Interior::set_Color(object)
  } // end of property Interior::Color
} // end of class Microsoft.Office.Interop.Excel.Interior

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.PageSetup
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 42 34 2D 30 30 30 30   // ..$000208B4-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_41() runtime managed
  {
  } // end of method PageSetup::_VtblGap1_41

  .method public hidebysig newslot specialname abstract virtual 
          instance string 
          marshal( bstr) 
          get_PrintArea() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 FB 03 00 00 00 00 ) 
  } // end of method PageSetup::get_PrintArea

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set_PrintArea([in] string  marshal( bstr) A_1) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 FB 03 00 00 00 00 ) 
  } // end of method PageSetup::set_PrintArea

  .property string PrintArea()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 FB 03 00 00 00 00 ) 
    .get instance string Microsoft.Office.Interop.Excel.PageSetup::get_PrintArea()
    .set instance void Microsoft.Office.Interop.Excel.PageSetup::set_PrintArea(string)
  } // end of property PageSetup::PrintArea
} // end of class Microsoft.Office.Interop.Excel.PageSetup

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Range
       implements [mscorlib]System.Collections.IEnumerable
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Reflection.DefaultMemberAttribute::.ctor(string) = ( 01 00 08 5F 44 65 66 61 75 6C 74 00 00 )          // ..._Default..
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 34 36 2D 30 30 30 30   // ..$00020846-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_6() runtime managed
  {
  } // end of method Range::_VtblGap1_6

  .method public hidebysig newslot specialname abstract virtual 
          instance string 
          marshal( bstr) 
          get_Address([in][opt] object  marshal( struct) RowAbsolute,
                      [in][opt] object  marshal( struct) ColumnAbsolute,
                      [in][opt] valuetype Microsoft.Office.Interop.Excel.XlReferenceStyle ReferenceStyle,
                      [in][opt] object  marshal( struct) External,
                      [in][opt] object  marshal( struct) RelativeTo) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 EC 00 00 00 00 00 ) 
    .param [3] = int32(0x00000001)
  } // end of method Range::get_Address

  .method public specialname rtspecialname 
          instance void  _VtblGap2_22() runtime managed
  {
  } // end of method Range::_VtblGap2_22

  .method public hidebysig newslot specialname abstract virtual 
          instance int32  get_Column() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 F0 00 00 00 00 00 ) 
  } // end of method Range::get_Column

  .method public specialname rtspecialname 
          instance void  _VtblGap3_1() runtime managed
  {
  } // end of method Range::_VtblGap3_1

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Range 
          marshal( interface ) 
          get_Columns() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 F1 00 00 00 00 00 ) 
  } // end of method Range::get_Columns

  .method public specialname rtspecialname 
          instance void  _VtblGap4_6() runtime managed
  {
  } // end of method Range::_VtblGap4_6

  .method public hidebysig newslot specialname abstract virtual 
          instance int32  get_Count() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 76 00 00 00 00 00 )                         // ..v.....
  } // end of method Range::get_Count

  .method public specialname rtspecialname 
          instance void  _VtblGap5_6() runtime managed
  {
  } // end of method Range::_VtblGap5_6

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( struct) 
          get__Default([in][opt] object  marshal( struct) RowIndex,
                       [in][opt] object  marshal( struct) ColumnIndex) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
  } // end of method Range::get__Default

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set__Default([in][opt] object  marshal( struct) RowIndex,
                                      [in][opt] object  marshal( struct) ColumnIndex,
                                      [in][opt] object  marshal( struct) A_3) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
  } // end of method Range::set__Default

  .method public specialname rtspecialname 
          instance void  _VtblGap6_45() runtime managed
  {
  } // end of method Range::_VtblGap6_45

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Interior 
          marshal( interface ) 
          get_Interior() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 81 00 00 00 00 00 ) 
  } // end of method Range::get_Interior

  .method public specialname rtspecialname 
          instance void  _VtblGap7_11() runtime managed
  {
  } // end of method Range::_VtblGap7_11

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Range 
          marshal( interface ) 
          get_MergeArea() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 69 05 00 00 00 00 )                         // ..i.....
  } // end of method Range::get_MergeArea

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( struct) 
          get_MergeCells() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 D0 00 00 00 00 00 ) 
  } // end of method Range::get_MergeCells

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set_MergeCells([in] object  marshal( struct) A_1) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 D0 00 00 00 00 00 ) 
  } // end of method Range::set_MergeCells

  .method public specialname rtspecialname 
          instance void  _VtblGap8_8() runtime managed
  {
  } // end of method Range::_VtblGap8_8

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( struct) 
          get_NumberFormatLocal() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 49 04 00 00 00 00 )                         // ..I.....
  } // end of method Range::get_NumberFormatLocal

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set_NumberFormatLocal([in] object  marshal( struct) A_1) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 49 04 00 00 00 00 )                         // ..I.....
  } // end of method Range::set_NumberFormatLocal

  .method public specialname rtspecialname 
          instance void  _VtblGap9_22() runtime managed
  {
  } // end of method Range::_VtblGap9_22

  .method public hidebysig newslot specialname abstract virtual 
          instance int32  get_Row() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 01 01 00 00 00 00 ) 
  } // end of method Range::get_Row

  .method public specialname rtspecialname 
          instance void  _VtblGap10_3() runtime managed
  {
  } // end of method Range::_VtblGap10_3

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Range 
          marshal( interface ) 
          get_Rows() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 02 01 00 00 00 00 ) 
  } // end of method Range::get_Rows

  .method public specialname rtspecialname 
          instance void  _VtblGap11_1() runtime managed
  {
  } // end of method Range::_VtblGap11_1

  .method public hidebysig newslot abstract virtual 
          instance object 
          marshal( struct) 
          Select() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 EB 00 00 00 00 00 ) 
  } // end of method Range::Select

  .method public specialname rtspecialname 
          instance void  _VtblGap12_18() runtime managed
  {
  } // end of method Range::_VtblGap12_18

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( struct) 
          get_Text() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 8A 00 00 00 00 00 ) 
  } // end of method Range::get_Text

  .method public specialname rtspecialname 
          instance void  _VtblGap13_18() runtime managed
  {
  } // end of method Range::_VtblGap13_18

  .method public hidebysig newslot abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Comment 
          marshal( interface ) 
          AddComment([in][opt] object  marshal( struct) Text) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 6D 05 00 00 00 00 )                         // ..m.....
  } // end of method Range::AddComment

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Comment 
          marshal( interface ) 
          get_Comment() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 8E 03 00 00 00 00 ) 
  } // end of method Range::get_Comment

  .method public hidebysig newslot abstract virtual 
          instance void  ClearComments() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 6E 05 00 00 00 00 )                         // ..n.....
  } // end of method Range::ClearComments

  .property string Address(object,
                           object,
                           valuetype Microsoft.Office.Interop.Excel.XlReferenceStyle,
                           object,
                           object)
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 EC 00 00 00 00 00 ) 
    .get instance string Microsoft.Office.Interop.Excel.Range::get_Address(object,
                                                                           object,
                                                                           valuetype Microsoft.Office.Interop.Excel.XlReferenceStyle,
                                                                           object,
                                                                           object)
  } // end of property Range::Address
  .property int32 Column()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 F0 00 00 00 00 00 ) 
    .get instance int32 Microsoft.Office.Interop.Excel.Range::get_Column()
  } // end of property Range::Column
  .property class Microsoft.Office.Interop.Excel.Range
          Columns()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 F1 00 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_Columns()
  } // end of property Range::Columns
  .property int32 Count()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 76 00 00 00 00 00 )                         // ..v.....
    .get instance int32 Microsoft.Office.Interop.Excel.Range::get_Count()
  } // end of property Range::Count
  .property object _Default(object,
                            object)
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
    .get instance object Microsoft.Office.Interop.Excel.Range::get__Default(object,
                                                                            object)
    .set instance void Microsoft.Office.Interop.Excel.Range::set__Default(object,
                                                                          object,
                                                                          object)
  } // end of property Range::_Default
  .property class Microsoft.Office.Interop.Excel.Interior
          Interior()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 81 00 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Interior Microsoft.Office.Interop.Excel.Range::get_Interior()
  } // end of property Range::Interior
  .property class Microsoft.Office.Interop.Excel.Range
          MergeArea()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 69 05 00 00 00 00 )                         // ..i.....
    .get instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_MergeArea()
  } // end of property Range::MergeArea
  .property object MergeCells()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 D0 00 00 00 00 00 ) 
    .get instance object Microsoft.Office.Interop.Excel.Range::get_MergeCells()
    .set instance void Microsoft.Office.Interop.Excel.Range::set_MergeCells(object)
  } // end of property Range::MergeCells
  .property object NumberFormatLocal()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 49 04 00 00 00 00 )                         // ..I.....
    .get instance object Microsoft.Office.Interop.Excel.Range::get_NumberFormatLocal()
    .set instance void Microsoft.Office.Interop.Excel.Range::set_NumberFormatLocal(object)
  } // end of property Range::NumberFormatLocal
  .property int32 Row()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 01 01 00 00 00 00 ) 
    .get instance int32 Microsoft.Office.Interop.Excel.Range::get_Row()
  } // end of property Range::Row
  .property class Microsoft.Office.Interop.Excel.Range
          Rows()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 02 01 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_Rows()
  } // end of property Range::Rows
  .property object Text()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 8A 00 00 00 00 00 ) 
    .get instance object Microsoft.Office.Interop.Excel.Range::get_Text()
  } // end of property Range::Text
  .property class Microsoft.Office.Interop.Excel.Comment
          Comment()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 8E 03 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::get_Comment()
  } // end of property Range::Comment
} // end of class Microsoft.Office.Interop.Excel.Range

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Sheets
       implements [mscorlib]System.Collections.IEnumerable
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Reflection.DefaultMemberAttribute::.ctor(string) = ( 01 00 08 5F 44 65 66 61 75 6C 74 00 00 )          // ..._Default..
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 44 37 2D 30 30 30 30   // ..$000208D7-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_5() runtime managed
  {
  } // end of method Sheets::_VtblGap1_5

  .method public hidebysig newslot specialname abstract virtual 
          instance int32  get_Count() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 76 00 00 00 00 00 )                         // ..v.....
  } // end of method Sheets::get_Count

  .method public specialname rtspecialname 
          instance void  _VtblGap2_4() runtime managed
  {
  } // end of method Sheets::_VtblGap2_4

  .method public hidebysig newslot abstract virtual 
          instance class [mscorlib]System.Collections.IEnumerator 
          marshal( custom ("System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a","")) 
          GetEnumerator() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 FC FF FF FF 00 00 ) 
  } // end of method Sheets::GetEnumerator

  .method public specialname rtspecialname 
          instance void  _VtblGap3_7() runtime managed
  {
  } // end of method Sheets::_VtblGap3_7

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( idispatch ) 
          get__Default([in] object  marshal( struct) Index) runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
  } // end of method Sheets::get__Default

  .property int32 Count()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 76 00 00 00 00 00 )                         // ..v.....
    .get instance int32 Microsoft.Office.Interop.Excel.Sheets::get_Count()
  } // end of property Sheets::Count
  .property object _Default(object)
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
    .get instance object Microsoft.Office.Interop.Excel.Sheets::get__Default(object)
  } // end of property Sheets::_Default
} // end of class Microsoft.Office.Interop.Excel.Sheets

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Window
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 39 33 2D 30 30 30 30   // ..$00020893-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_9() runtime managed
  {
  } // end of method Window::_VtblGap1_9

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( idispatch ) 
          get_ActiveSheet() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 33 01 00 00 00 00 )                         // ..3.....
  } // end of method Window::get_ActiveSheet

  .method public specialname rtspecialname 
          instance void  _VtblGap2_47() runtime managed
  {
  } // end of method Window::_VtblGap2_47

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Sheets 
          marshal( interface ) 
          get_SelectedSheets() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 90 02 00 00 00 00 ) 
  } // end of method Window::get_SelectedSheets

  .property object ActiveSheet()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 33 01 00 00 00 00 )                         // ..3.....
    .get instance object Microsoft.Office.Interop.Excel.Window::get_ActiveSheet()
  } // end of property Window::ActiveSheet
  .property class Microsoft.Office.Interop.Excel.Sheets
          SelectedSheets()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 90 02 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Sheets Microsoft.Office.Interop.Excel.Window::get_SelectedSheets()
  } // end of property Window::SelectedSheets
} // end of class Microsoft.Office.Interop.Excel.Window

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Windows
       implements [mscorlib]System.Collections.IEnumerable
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Reflection.DefaultMemberAttribute::.ctor(string) = ( 01 00 08 5F 44 65 66 61 75 6C 74 00 00 )          // ..._Default..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 39 32 2D 30 30 30 30   // ..$00020892-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Interop.Excel.Windows

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Workbook
       implements Microsoft.Office.Interop.Excel._Workbook,
                  Microsoft.Office.Interop.Excel.WorkbookEvents_Event
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 44 41 2D 30 30 30 30   // ..$000208DA-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.CoClassAttribute::.ctor(class [mscorlib]System.Type) = ( 01 00 5A 53 79 73 74 65 6D 2E 4F 62 6A 65 63 74   // ..ZSystem.Object
                                                                                                                          2C 20 6D 73 63 6F 72 6C 69 62 2C 20 56 65 72 73   // , mscorlib, Vers
                                                                                                                          69 6F 6E 3D 34 2E 30 2E 30 2E 30 2C 20 43 75 6C   // ion=4.0.0.0, Cul
                                                                                                                          74 75 72 65 3D 6E 65 75 74 72 61 6C 2C 20 50 75   // ture=neutral, Pu
                                                                                                                          62 6C 69 63 4B 65 79 54 6F 6B 65 6E 3D 62 37 37   // blicKeyToken=b77
                                                                                                                          61 35 63 35 36 31 39 33 34 65 30 38 39 00 00 )    // a5c561934e089..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Interop.Excel.Workbook

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.WorkbookEvents
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 34 34 31 32 2D 30 30 30 30   // ..$00024412-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Interop.Excel.WorkbookEvents

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.WorkbookEvents_Event
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComEventInterfaceAttribute::.ctor(class [mscorlib]System.Type,
                                                                                                   class [mscorlib]System.Type) = ( 01 00 2D 4D 69 63 72 6F 73 6F 66 74 2E 4F 66 66   // ..-Microsoft.Off
                                                                                                                                    69 63 65 2E 49 6E 74 65 72 6F 70 2E 45 78 63 65   // ice.Interop.Exce
                                                                                                                                    6C 2E 57 6F 72 6B 62 6F 6F 6B 45 76 65 6E 74 73   // l.WorkbookEvents
                                                                                                                                    2D 4D 69 63 72 6F 73 6F 66 74 2E 4F 66 66 69 63   // -Microsoft.Offic
                                                                                                                                    65 2E 49 6E 74 65 72 6F 70 2E 45 78 63 65 6C 2E   // e.Interop.Excel.
                                                                                                                                    57 6F 72 6B 62 6F 6F 6B 45 76 65 6E 74 73 00 00 ) // WorkbookEvents..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 30 30 30 32 30 38 31 33 2D 30 30 30 30   // ..$00020813-0000
                                                                                                            2D 30 30 30 30 2D 63 30 30 30 2D 30 30 30 30 30   // -0000-c000-00000
                                                                                                            30 30 30 30 30 34 36 33 4D 69 63 72 6F 73 6F 66   // 00000463Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 49 6E 74 65 72 6F 70   // t.Office.Interop
                                                                                                            2E 45 78 63 65 6C 2E 57 6F 72 6B 62 6F 6F 6B 45   // .Excel.WorkbookE
                                                                                                            76 65 6E 74 73 5F 45 76 65 6E 74 00 00 )          // vents_Event..
} // end of class Microsoft.Office.Interop.Excel.WorkbookEvents_Event

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Worksheet
       implements Microsoft.Office.Interop.Excel._Worksheet,
                  Microsoft.Office.Interop.Excel.DocEvents_Event
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.CoClassAttribute::.ctor(class [mscorlib]System.Type) = ( 01 00 5A 53 79 73 74 65 6D 2E 4F 62 6A 65 63 74   // ..ZSystem.Object
                                                                                                                          2C 20 6D 73 63 6F 72 6C 69 62 2C 20 56 65 72 73   // , mscorlib, Vers
                                                                                                                          69 6F 6E 3D 34 2E 30 2E 30 2E 30 2C 20 43 75 6C   // ion=4.0.0.0, Cul
                                                                                                                          74 75 72 65 3D 6E 65 75 74 72 61 6C 2C 20 50 75   // ture=neutral, Pu
                                                                                                                          62 6C 69 63 4B 65 79 54 6F 6B 65 6E 3D 62 37 37   // blicKeyToken=b77
                                                                                                                          61 35 63 35 36 31 39 33 34 65 30 38 39 00 00 )    // a5c561934e089..
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 44 38 2D 30 30 30 30   // ..$000208D8-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
} // end of class Microsoft.Office.Interop.Excel.Worksheet

.class public auto ansi sealed Microsoft.Office.Interop.Excel.XlReferenceStyle
       extends [mscorlib]System.Enum
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 30 30 30 32 30 38 31 33 2D 30 30 30 30   // ..$00020813-0000
                                                                                                            2D 30 30 30 30 2D 63 30 30 30 2D 30 30 30 30 30   // -0000-c000-00000
                                                                                                            30 30 30 30 30 34 36 2F 4D 69 63 72 6F 73 6F 66   // 0000046/Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 49 6E 74 65 72 6F 70   // t.Office.Interop
                                                                                                            2E 45 78 63 65 6C 2E 58 6C 52 65 66 65 72 65 6E   // .Excel.XlReferen
                                                                                                            63 65 53 74 79 6C 65 00 00 )                      // ceStyle..
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlReferenceStyle xlA1 = int32(0x00000001)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlReferenceStyle xlR1C1 = int32(0xFFFFEFCA)
} // end of class Microsoft.Office.Interop.Excel.XlReferenceStyle

.class public auto ansi sealed Microsoft.Office.Interop.Excel.XlSheetVisibility
       extends [mscorlib]System.Enum
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 30 30 30 32 30 38 31 33 2D 30 30 30 30   // ..$00020813-0000
                                                                                                            2D 30 30 30 30 2D 63 30 30 30 2D 30 30 30 30 30   // -0000-c000-00000
                                                                                                            30 30 30 30 30 34 36 30 4D 69 63 72 6F 73 6F 66   // 00000460Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 49 6E 74 65 72 6F 70   // t.Office.Interop
                                                                                                            2E 45 78 63 65 6C 2E 58 6C 53 68 65 65 74 56 69   // .Excel.XlSheetVi
                                                                                                            73 69 62 69 6C 69 74 79 00 00 )                   // sibility..
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlSheetVisibility xlSheetVisible = int32(0xFFFFFFFF)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlSheetVisibility xlSheetHidden = int32(0x00000000)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlSheetVisibility xlSheetVeryHidden = int32(0x00000002)
} // end of class Microsoft.Office.Interop.Excel.XlSheetVisibility

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel._Application
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Reflection.DefaultMemberAttribute::.ctor(string) = ( 01 00 08 5F 44 65 66 61 75 6C 74 00 00 )          // ..._Default..
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 44 35 2D 30 30 30 30   // ..$000208D5-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_9() runtime managed
  {
  } // end of method _Application::_VtblGap1_9

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( idispatch ) 
          get_ActiveSheet() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 33 01 00 00 00 00 )                         // ..3.....
  } // end of method _Application::get_ActiveSheet

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Window 
          marshal( interface ) 
          get_ActiveWindow() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 F7 02 00 00 00 00 ) 
  } // end of method _Application::get_ActiveWindow

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Workbook 
          marshal( interface ) 
          get_ActiveWorkbook() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 34 01 00 00 00 00 )                         // ..4.....
  } // end of method _Application::get_ActiveWorkbook

  .method public specialname rtspecialname 
          instance void  _VtblGap2_3() runtime managed
  {
  } // end of method _Application::_VtblGap2_3

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Range 
          marshal( interface ) 
          get_Cells() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 EE 00 00 00 00 00 ) 
  } // end of method _Application::get_Cells

  .method public specialname rtspecialname 
          instance void  _VtblGap3_2() runtime managed
  {
  } // end of method _Application::_VtblGap3_2

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Core.CommandBars 
          marshal( interface ) 
          get_CommandBars() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 9F 05 00 00 00 00 ) 
  } // end of method _Application::get_CommandBars

  .method public specialname rtspecialname 
          instance void  _VtblGap4_10() runtime managed
  {
  } // end of method _Application::_VtblGap4_10

  .method public hidebysig newslot abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Range 
          marshal( interface ) 
          Intersect([in] class Microsoft.Office.Interop.Excel.Range  marshal( interface ) Arg1,
                    [in] class Microsoft.Office.Interop.Excel.Range  marshal( interface ) Arg2,
                    [in][opt] object  marshal( struct) Arg3,
                    [in][opt] object  marshal( struct) Arg4,
                    [in][opt] object  marshal( struct) Arg5,
                    [in][opt] object  marshal( struct) Arg6,
                    [in][opt] object  marshal( struct) Arg7,
                    [in][opt] object  marshal( struct) Arg8,
                    [in][opt] object  marshal( struct) Arg9,
                    [in][opt] object  marshal( struct) Arg10,
                    [in][opt] object  marshal( struct) Arg11,
                    [in][opt] object  marshal( struct) Arg12,
                    [in][opt] object  marshal( struct) Arg13,
                    [in][opt] object  marshal( struct) Arg14,
                    [in][opt] object  marshal( struct) Arg15,
                    [in][opt] object  marshal( struct) Arg16,
                    [in][opt] object  marshal( struct) Arg17,
                    [in][opt] object  marshal( struct) Arg18,
                    [in][opt] object  marshal( struct) Arg19,
                    [in][opt] object  marshal( struct) Arg20,
                    [in][opt] object  marshal( struct) Arg21,
                    [in][opt] object  marshal( struct) Arg22,
                    [in][opt] object  marshal( struct) Arg23,
                    [in][opt] object  marshal( struct) Arg24,
                    [in][opt] object  marshal( struct) Arg25,
                    [in][opt] object  marshal( struct) Arg26,
                    [in][opt] object  marshal( struct) Arg27,
                    [in][opt] object  marshal( struct) Arg28,
                    [in][opt] object  marshal( struct) Arg29,
                    [in][opt] object  marshal( struct) Arg30) runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 FE 02 00 00 00 00 ) 
    .custom instance void [mscorlib]System.Runtime.InteropServices.LCIDConversionAttribute::.ctor(int32) = ( 01 00 1E 00 00 00 00 00 ) 
  } // end of method _Application::Intersect

  .method public specialname rtspecialname 
          instance void  _VtblGap5_7() runtime managed
  {
  } // end of method _Application::_VtblGap5_7

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( idispatch ) 
          get_Selection() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 93 00 00 00 00 00 ) 
    .custom instance void [mscorlib]System.Runtime.InteropServices.LCIDConversionAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
  } // end of method _Application::get_Selection

  .method public specialname rtspecialname 
          instance void  _VtblGap6_6() runtime managed
  {
  } // end of method _Application::_VtblGap6_6

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Windows 
          marshal( interface ) 
          get_Windows() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 AE 01 00 00 00 00 ) 
  } // end of method _Application::get_Windows

  .method public specialname rtspecialname 
          instance void  _VtblGap7_61() runtime managed
  {
  } // end of method _Application::_VtblGap7_61

  .method public hidebysig newslot specialname abstract virtual 
          instance string 
          marshal( bstr) 
          get__Default() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
  } // end of method _Application::get__Default

  .method public specialname rtspecialname 
          instance void  _VtblGap8_237() runtime managed
  {
  } // end of method _Application::_VtblGap8_237

  .method public hidebysig newslot specialname abstract virtual 
          instance int32  get_Hwnd() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 9E 07 00 00 00 00 ) 
  } // end of method _Application::get_Hwnd

  .property object ActiveSheet()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 33 01 00 00 00 00 )                         // ..3.....
    .get instance object Microsoft.Office.Interop.Excel._Application::get_ActiveSheet()
  } // end of property _Application::ActiveSheet
  .property class Microsoft.Office.Interop.Excel.Window
          ActiveWindow()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 F7 02 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Window Microsoft.Office.Interop.Excel._Application::get_ActiveWindow()
  } // end of property _Application::ActiveWindow
  .property class Microsoft.Office.Interop.Excel.Workbook
          ActiveWorkbook()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 34 01 00 00 00 00 )                         // ..4.....
    .get instance class Microsoft.Office.Interop.Excel.Workbook Microsoft.Office.Interop.Excel._Application::get_ActiveWorkbook()
  } // end of property _Application::ActiveWorkbook
  .property class Microsoft.Office.Interop.Excel.Range
          Cells()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 EE 00 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Application::get_Cells()
  } // end of property _Application::Cells
  .property class Microsoft.Office.Core.CommandBars
          CommandBars()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 9F 05 00 00 00 00 ) 
    .get instance class Microsoft.Office.Core.CommandBars Microsoft.Office.Interop.Excel._Application::get_CommandBars()
  } // end of property _Application::CommandBars
  .property object Selection()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 93 00 00 00 00 00 ) 
    .get instance object Microsoft.Office.Interop.Excel._Application::get_Selection()
  } // end of property _Application::Selection
  .property class Microsoft.Office.Interop.Excel.Windows
          Windows()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 AE 01 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Windows Microsoft.Office.Interop.Excel._Application::get_Windows()
  } // end of property _Application::Windows
  .property string _Default()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
    .get instance string Microsoft.Office.Interop.Excel._Application::get__Default()
  } // end of property _Application::_Default
  .property int32 Hwnd()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 9E 07 00 00 00 00 ) 
    .get instance int32 Microsoft.Office.Interop.Excel._Application::get_Hwnd()
  } // end of property _Application::Hwnd
} // end of class Microsoft.Office.Interop.Excel._Application

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel._Workbook
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 44 41 2D 30 30 30 30   // ..$000208DA-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_7() runtime managed
  {
  } // end of method _Workbook::_VtblGap1_7

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( idispatch ) 
          get_ActiveSheet() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 33 01 00 00 00 00 )                         // ..3.....
  } // end of method _Workbook::get_ActiveSheet

  .method public specialname rtspecialname 
          instance void  _VtblGap2_51() runtime managed
  {
  } // end of method _Workbook::_VtblGap2_51

  .method public hidebysig newslot specialname abstract virtual 
          instance string 
          marshal( bstr) 
          get_Name() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 6E 00 00 00 00 00 )                         // ..n.....
  } // end of method _Workbook::get_Name

  .method public specialname rtspecialname 
          instance void  _VtblGap3_9() runtime managed
  {
  } // end of method _Workbook::_VtblGap3_9

  .method public hidebysig newslot specialname abstract virtual 
          instance string 
          marshal( bstr) 
          get_Path() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 23 01 00 00 00 00 )                         // ..#.....
    .custom instance void [mscorlib]System.Runtime.InteropServices.LCIDConversionAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
  } // end of method _Workbook::get_Path

  .method public specialname rtspecialname 
          instance void  _VtblGap4_35() runtime managed
  {
  } // end of method _Workbook::_VtblGap4_35

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Sheets 
          marshal( interface ) 
          get_Sheets() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 E5 01 00 00 00 00 ) 
  } // end of method _Workbook::get_Sheets

  .property object ActiveSheet()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 33 01 00 00 00 00 )                         // ..3.....
    .get instance object Microsoft.Office.Interop.Excel._Workbook::get_ActiveSheet()
  } // end of property _Workbook::ActiveSheet
  .property string Name()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 6E 00 00 00 00 00 )                         // ..n.....
    .get instance string Microsoft.Office.Interop.Excel._Workbook::get_Name()
  } // end of property _Workbook::Name
  .property string Path()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 23 01 00 00 00 00 )                         // ..#.....
    .get instance string Microsoft.Office.Interop.Excel._Workbook::get_Path()
  } // end of property _Workbook::Path
  .property class Microsoft.Office.Interop.Excel.Sheets
          Sheets()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 E5 01 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Sheets Microsoft.Office.Interop.Excel._Workbook::get_Sheets()
  } // end of property _Workbook::Sheets
} // end of class Microsoft.Office.Interop.Excel._Workbook

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel._Worksheet
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 44 38 2D 30 30 30 30   // ..$000208D8-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_3() runtime managed
  {
  } // end of method _Worksheet::_VtblGap1_3

  .method public hidebysig newslot abstract virtual 
          instance void  Activate() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.LCIDConversionAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 30 01 00 00 00 00 )                         // ..0.....
  } // end of method _Worksheet::Activate

  .method public specialname rtspecialname 
          instance void  _VtblGap2_5() runtime managed
  {
  } // end of method _Worksheet::_VtblGap2_5

  .method public hidebysig newslot specialname abstract virtual 
          instance int32  get_Index() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.LCIDConversionAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 E6 01 00 00 00 00 ) 
  } // end of method _Worksheet::get_Index

  .method public specialname rtspecialname 
          instance void  _VtblGap3_1() runtime managed
  {
  } // end of method _Worksheet::_VtblGap3_1

  .method public hidebysig newslot specialname abstract virtual 
          instance string 
          marshal( bstr) 
          get_Name() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 6E 00 00 00 00 00 )                         // ..n.....
  } // end of method _Worksheet::get_Name

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set_Name([in] string  marshal( bstr) RHS) runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 6E 00 00 00 00 00 )                         // ..n.....
  } // end of method _Worksheet::set_Name

  .method public specialname rtspecialname 
          instance void  _VtblGap4_7() runtime managed
  {
  } // end of method _Worksheet::_VtblGap4_7

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.PageSetup 
          marshal( interface ) 
          get_PageSetup() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 E6 03 00 00 00 00 ) 
  } // end of method _Worksheet::get_PageSetup

  .method public specialname rtspecialname 
          instance void  _VtblGap5_9() runtime managed
  {
  } // end of method _Worksheet::_VtblGap5_9

  .method public hidebysig newslot abstract virtual 
          instance void  Select([in][opt] object  marshal( struct) Replace) runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.LCIDConversionAttribute::.ctor(int32) = ( 01 00 01 00 00 00 00 00 ) 
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 EB 00 00 00 00 00 ) 
  } // end of method _Worksheet::Select

  .method public specialname rtspecialname 
          instance void  _VtblGap6_1() runtime managed
  {
  } // end of method _Worksheet::_VtblGap6_1

  .method public hidebysig newslot specialname abstract virtual 
          instance valuetype Microsoft.Office.Interop.Excel.XlSheetVisibility 
          get_Visible() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.LCIDConversionAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 2E 02 00 00 00 00 )                         // ........
  } // end of method _Worksheet::get_Visible

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set_Visible([in] valuetype Microsoft.Office.Interop.Excel.XlSheetVisibility RHS) runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 2E 02 00 00 00 00 )                         // ........
    .custom instance void [mscorlib]System.Runtime.InteropServices.LCIDConversionAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
  } // end of method _Worksheet::set_Visible

  .method public specialname rtspecialname 
          instance void  _VtblGap7_11() runtime managed
  {
  } // end of method _Worksheet::_VtblGap7_11

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Range 
          marshal( interface ) 
          get_Cells() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 EE 00 00 00 00 00 ) 
  } // end of method _Worksheet::get_Cells

  .method public specialname rtspecialname 
          instance void  _VtblGap8_47() runtime managed
  {
  } // end of method _Worksheet::_VtblGap8_47

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Range 
          marshal( interface ) 
          get_Range([in] object  marshal( struct) Cell1,
                    [in][opt] object  marshal( struct) Cell2) runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 C5 00 00 00 00 00 ) 
  } // end of method _Worksheet::get_Range

  .method public specialname rtspecialname 
          instance void  _VtblGap9_22() runtime managed
  {
  } // end of method _Worksheet::_VtblGap9_22

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Comments 
          marshal( interface ) 
          get_Comments() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 3F 02 00 00 00 00 )                         // ..?.....
  } // end of method _Worksheet::get_Comments

  .property int32 Index()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 E6 01 00 00 00 00 ) 
    .get instance int32 Microsoft.Office.Interop.Excel._Worksheet::get_Index()
  } // end of property _Worksheet::Index
  .property string Name()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 6E 00 00 00 00 00 )                         // ..n.....
    .get instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
    .set instance void Microsoft.Office.Interop.Excel._Worksheet::set_Name(string)
  } // end of property _Worksheet::Name
  .property class Microsoft.Office.Interop.Excel.PageSetup
          PageSetup()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 E6 03 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.PageSetup Microsoft.Office.Interop.Excel._Worksheet::get_PageSetup()
  } // end of property _Worksheet::PageSetup
  .property valuetype Microsoft.Office.Interop.Excel.XlSheetVisibility
          Visible()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 2E 02 00 00 00 00 )                         // ........
    .get instance valuetype Microsoft.Office.Interop.Excel.XlSheetVisibility Microsoft.Office.Interop.Excel._Worksheet::get_Visible()
    .set instance void Microsoft.Office.Interop.Excel._Worksheet::set_Visible(valuetype Microsoft.Office.Interop.Excel.XlSheetVisibility)
  } // end of property _Worksheet::Visible
  .property class Microsoft.Office.Interop.Excel.Range
          Cells()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 EE 00 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Cells()
  } // end of property _Worksheet::Cells
  .property class Microsoft.Office.Interop.Excel.Range
          Range(object,
                object)
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 C5 00 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Range(object,
                                                                                                                  object)
  } // end of property _Worksheet::Range
  .property class Microsoft.Office.Interop.Excel.Comments
          Comments()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 3F 02 00 00 00 00 )                         // ..?.....
    .get instance class Microsoft.Office.Interop.Excel.Comments Microsoft.Office.Interop.Excel._Worksheet::get_Comments()
  } // end of property _Worksheet::Comments
} // end of class Microsoft.Office.Interop.Excel._Worksheet


// =============================================================

// *********** DISASSEMBLY COMPLETE ***********************
// WARNING: Created Win32 resource file C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise\iReporterExcelAddIn.res
