
//  Microsoft (R) .NET Framework IL Disassembler.  Version 4.8.3928.0
//  Copyright (c) Microsoft Corporation.  All rights reserved.



// Metadata version: v4.0.30319
.assembly extern mscorlib
{
  .publickeytoken = (B7 7A 5C 56 19 34 E0 89 )                         // .z\V.4..
  .ver 4:0:0:0
}
.assembly extern System.Drawing
{
  .publickeytoken = (B0 3F 5F 7F 11 D5 0A 3A )                         // .?_....:
  .ver 4:0:0:0
}
.assembly extern System.Core
{
  .publickeytoken = (B7 7A 5C 56 19 34 E0 89 )                         // .z\V.4..
  .ver 4:0:0:0
}
.assembly extern Microsoft.CSharp
{
  .publickeytoken = (B0 3F 5F 7F 11 D5 0A 3A )                         // .?_....:
  .ver 4:0:0:0
}
.assembly Cimtops.Excel
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.ExtensionAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilationRelaxationsAttribute::.ctor(int32) = ( 01 00 08 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.CompilerServices.RuntimeCompatibilityAttribute::.ctor() = ( 01 00 01 00 54 02 16 57 72 61 70 4E 6F 6E 45 78   // ....T..WrapNonEx
                                                                                                             63 65 70 74 69 6F 6E 54 68 72 6F 77 73 01 )       // ceptionThrows.

  // --- The following custom attribute is added automatically, do not uncomment -------
  //  .custom instance void [mscorlib]System.Diagnostics.DebuggableAttribute::.ctor(valuetype [mscorlib]System.Diagnostics.DebuggableAttribute/DebuggingModes) = ( 01 00 02 00 00 00 00 00 ) 

  .custom instance void [mscorlib]System.Reflection.AssemblyTitleAttribute::.ctor(string) = ( 01 00 0D 43 69 6D 74 6F 70 73 2E 45 78 63 65 6C   // ...Cimtops.Excel
                                                                                              00 00 ) 
  .custom instance void [mscorlib]System.Reflection.AssemblyDescriptionAttribute::.ctor(string) = ( 01 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Reflection.AssemblyConfigurationAttribute::.ctor(string) = ( 01 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Reflection.AssemblyCompanyAttribute::.ctor(string) = ( 01 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Reflection.AssemblyProductAttribute::.ctor(string) = ( 01 00 0D 43 69 6D 74 6F 70 73 2E 45 78 63 65 6C   // ...Cimtops.Excel
                                                                                                00 00 ) 
  .custom instance void [mscorlib]System.Reflection.AssemblyCopyrightAttribute::.ctor(string) = ( 01 00 12 43 6F 70 79 72 69 67 68 74 20 C2 A9 20   // ...Copyright .. 
                                                                                                  20 32 30 31 38 00 00 )                            //  2018..
  .custom instance void [mscorlib]System.Reflection.AssemblyTrademarkAttribute::.ctor(string) = ( 01 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComVisibleAttribute::.ctor(bool) = ( 01 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 64 62 37 61 65 36 64 34 2D 33 63 33 38   // ..$db7ae6d4-3c38
                                                                                                  2D 34 62 66 31 2D 62 39 36 65 2D 37 62 64 35 31   // -4bf1-b96e-7bd51
                                                                                                  30 66 36 37 35 36 63 00 00 )                      // 0f6756c..
  .custom instance void [mscorlib]System.Reflection.AssemblyFileVersionAttribute::.ctor(string) = ( 01 00 0B 38 2E 32 2E 32 36 30 36 30 2E 31 00 00 ) // ...8.2.26060.1..
  .custom instance void [mscorlib]System.Runtime.Versioning.TargetFrameworkAttribute::.ctor(string) = ( 01 00 1A 2E 4E 45 54 46 72 61 6D 65 77 6F 72 6B   // ....NETFramework
                                                                                                        2C 56 65 72 73 69 6F 6E 3D 76 34 2E 35 01 00 54   // ,Version=v4.5..T
                                                                                                        0E 14 46 72 61 6D 65 77 6F 72 6B 44 69 73 70 6C   // ..FrameworkDispl
                                                                                                        61 79 4E 61 6D 65 12 2E 4E 45 54 20 46 72 61 6D   // ayName..NET Fram
                                                                                                        65 77 6F 72 6B 20 34 2E 35 )                      // ework 4.5
  .hash algorithm 0x00008004
  .ver 8:2:26060:1
}
.module Cimtops.Excel.dll
// MVID: {BF183737-F970-4003-AD4F-0FA16F458AE0}
.imagebase 0x10000000
.file alignment 0x00000200
.stackreserve 0x00100000
.subsystem 0x0003       // WINDOWS_CUI
.corflags 0x00000001    //  ILONLY
// Image base: 0x07120000


// =============== CLASS MEMBERS DECLARATION ===================

.class public abstract auto ansi sealed beforefieldinit Cimtops.Excel.ActionTypes
       extends [mscorlib]System.Object
{
  .field public static literal string DOCUMENT = "document"
  .field public static literal string SHEET_JUMP = "sheetJump"
  .field public static literal string MENU = "menu"
  .field public static literal string OPEN_URL = "openURL"
  .field public static literal string NO_NEED_TO_FILL_OUT = "noNeedToFillOut"
  .field public static literal string SHEET_COPY = "sheetCopy"
  .field public static literal string BACK_TO_LIBRARY = "backToLibrary"
  .field public static literal string CONMAS_IOT_GATEWAY = "conmasIotGateway"
  .field public static literal string TIMER_START = "timerStart"
  .field public static literal string RUN_COMMAND = "runCommand"
  .field public static literal string OUTPUT_TEXT = "outputText"
  .field public static literal string AUTO_INPUT_CLUSTER = "autoInputCluster"
  .field public static literal string CREATE_QR_CODE = "createQRCode"
  .field public static literal string BIOMETRICS = "biometrics"
  .field public static literal string BATCHCLEAR = "batchClear"
  .method public hidebysig static valuetype Cimtops.Excel.DeviceType[] 
          GetDevices(string actionType) cil managed
  {
    // Code size       74 (0x4a)
    .maxstack  4
    IL_0000:  ldarg.0
    IL_0001:  ldstr      "runCommand"
    IL_0006:  call       bool [mscorlib]System.String::op_Equality(string,
                                                                   string)
    IL_000b:  brtrue.s   IL_0036

    IL_000d:  ldarg.0
    IL_000e:  ldstr      "outputText"
    IL_0013:  call       bool [mscorlib]System.String::op_Equality(string,
                                                                   string)
    IL_0018:  brtrue.s   IL_0036

    IL_001a:  ldarg.0
    IL_001b:  ldstr      "autoInputCluster"
    IL_0020:  call       bool [mscorlib]System.String::op_Equality(string,
                                                                   string)
    IL_0025:  brtrue.s   IL_0036

    IL_0027:  ldarg.0
    IL_0028:  ldstr      "biometrics"
    IL_002d:  call       bool [mscorlib]System.String::op_Equality(string,
                                                                   string)
    IL_0032:  brtrue.s   IL_003d

    IL_0034:  br.s       IL_0048

    IL_0036:  ldc.i4.1
    IL_0037:  newarr     Cimtops.Excel.DeviceType
    IL_003c:  ret

    IL_003d:  ldc.i4.1
    IL_003e:  newarr     Cimtops.Excel.DeviceType
    IL_0043:  dup
    IL_0044:  ldc.i4.0
    IL_0045:  ldc.i4.1
    IL_0046:  stelem.i4
    IL_0047:  ret

    IL_0048:  ldnull
    IL_0049:  ret
  } // end of method ActionTypes::GetDevices

} // end of class Cimtops.Excel.ActionTypes

.class public auto ansi beforefieldinit Cimtops.Excel.Cell
       extends Cimtops.Excel.CellRange
{
  .class auto ansi serializable sealed nested private beforefieldinit '<>c'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static initonly class Cimtops.Excel.Cell/'<>c' '<>9'
    .field public static class [mscorlib]System.Func`2<class Cimtops.Excel.Col,float64> '<>9__19_1'
    .field public static class [mscorlib]System.Func`2<class Cimtops.Excel.Row,float64> '<>9__21_1'
    .method private hidebysig specialname rtspecialname static 
            void  .cctor() cil managed
    {
      // Code size       11 (0xb)
      .maxstack  8
      IL_0000:  newobj     instance void Cimtops.Excel.Cell/'<>c'::.ctor()
      IL_0005:  stsfld     class Cimtops.Excel.Cell/'<>c' Cimtops.Excel.Cell/'<>c'::'<>9'
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

    .method assembly hidebysig instance float64 
            '<get_Width>b__19_1'(class Cimtops.Excel.Col c) cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance float64 Cimtops.Excel.Col::get_Width()
      IL_0006:  ret
    } // end of method '<>c'::'<get_Width>b__19_1'

    .method assembly hidebysig instance float64 
            '<get_Height>b__21_1'(class Cimtops.Excel.Row r) cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance float64 Cimtops.Excel.Row::get_Height()
      IL_0006:  ret
    } // end of method '<>c'::'<get_Height>b__21_1'

  } // end of class '<>c'

  .class abstract auto ansi sealed nested private beforefieldinit '<>o__24'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> '<>p__0'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> '<>p__1'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> '<>p__2'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class Cimtops.Excel.Cell,object,object>> '<>p__3'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32>> '<>p__4'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> '<>p__5'
  } // end of class '<>o__24'

  .class abstract auto ansi sealed nested private beforefieldinit '<>o__25'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype Microsoft.Office.Interop.Excel.XlLineStyle>> '<>p__0'
  } // end of class '<>o__25'

  .field private initonly class Cimtops.Excel.Sheet '<Sheet>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly string '<Text>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly bool '<IsFormula>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly string '<Comment>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly int32 '<BGColor>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly class Cimtops.Excel.BorderQ '<Border>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig specialname instance class Cimtops.Excel.Sheet 
          get_Sheet() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class Cimtops.Excel.Sheet Cimtops.Excel.Cell::'<Sheet>k__BackingField'
    IL_0006:  ret
  } // end of method Cell::get_Sheet

  .method public hidebysig specialname instance string 
          get_Text() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      string Cimtops.Excel.Cell::'<Text>k__BackingField'
    IL_0006:  ret
  } // end of method Cell::get_Text

  .method public hidebysig specialname instance bool 
          get_IsFormula() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      bool Cimtops.Excel.Cell::'<IsFormula>k__BackingField'
    IL_0006:  ret
  } // end of method Cell::get_IsFormula

  .method public hidebysig specialname instance string 
          get_Comment() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      string Cimtops.Excel.Cell::'<Comment>k__BackingField'
    IL_0006:  ret
  } // end of method Cell::get_Comment

  .method public hidebysig specialname instance int32 
          get_BGColor() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      int32 Cimtops.Excel.Cell::'<BGColor>k__BackingField'
    IL_0006:  ret
  } // end of method Cell::get_BGColor

  .method public hidebysig specialname instance class Cimtops.Excel.BorderQ 
          get_Border() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class Cimtops.Excel.BorderQ Cimtops.Excel.Cell::'<Border>k__BackingField'
    IL_0006:  ret
  } // end of method Cell::get_Border

  .method public hidebysig specialname instance float64 
          get_Width() cil managed
  {
    // Code size       65 (0x41)
    .maxstack  3
    IL_0000:  ldarg.0
    IL_0001:  call       instance class Cimtops.Excel.Sheet Cimtops.Excel.Cell::get_Sheet()
    IL_0006:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Col> Cimtops.Excel.Sheet::get_Cols()
    IL_000b:  ldarg.0
    IL_000c:  ldftn      instance bool Cimtops.Excel.Cell::'<get_Width>b__19_0'(class Cimtops.Excel.Col)
    IL_0012:  newobj     instance void class [mscorlib]System.Func`2<class Cimtops.Excel.Col,bool>::.ctor(object,
                                                                                                          native int)
    IL_0017:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0> [System.Core]System.Linq.Enumerable::Where<class Cimtops.Excel.Col>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                                           class [mscorlib]System.Func`2<!!0,bool>)
    IL_001c:  ldsfld     class [mscorlib]System.Func`2<class Cimtops.Excel.Col,float64> Cimtops.Excel.Cell/'<>c'::'<>9__19_1'
    IL_0021:  dup
    IL_0022:  brtrue.s   IL_003b

    IL_0024:  pop
    IL_0025:  ldsfld     class Cimtops.Excel.Cell/'<>c' Cimtops.Excel.Cell/'<>c'::'<>9'
    IL_002a:  ldftn      instance float64 Cimtops.Excel.Cell/'<>c'::'<get_Width>b__19_1'(class Cimtops.Excel.Col)
    IL_0030:  newobj     instance void class [mscorlib]System.Func`2<class Cimtops.Excel.Col,float64>::.ctor(object,
                                                                                                             native int)
    IL_0035:  dup
    IL_0036:  stsfld     class [mscorlib]System.Func`2<class Cimtops.Excel.Col,float64> Cimtops.Excel.Cell/'<>c'::'<>9__19_1'
    IL_003b:  call       float64 [System.Core]System.Linq.Enumerable::Sum<class Cimtops.Excel.Col>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                   class [mscorlib]System.Func`2<!!0,float64>)
    IL_0040:  ret
  } // end of method Cell::get_Width

  .method public hidebysig specialname instance float64 
          get_Height() cil managed
  {
    // Code size       65 (0x41)
    .maxstack  3
    IL_0000:  ldarg.0
    IL_0001:  call       instance class Cimtops.Excel.Sheet Cimtops.Excel.Cell::get_Sheet()
    IL_0006:  callvirt   instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Row> Cimtops.Excel.Sheet::get_Rows()
    IL_000b:  ldarg.0
    IL_000c:  ldftn      instance bool Cimtops.Excel.Cell::'<get_Height>b__21_0'(class Cimtops.Excel.Row)
    IL_0012:  newobj     instance void class [mscorlib]System.Func`2<class Cimtops.Excel.Row,bool>::.ctor(object,
                                                                                                          native int)
    IL_0017:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0> [System.Core]System.Linq.Enumerable::Where<class Cimtops.Excel.Row>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                                           class [mscorlib]System.Func`2<!!0,bool>)
    IL_001c:  ldsfld     class [mscorlib]System.Func`2<class Cimtops.Excel.Row,float64> Cimtops.Excel.Cell/'<>c'::'<>9__21_1'
    IL_0021:  dup
    IL_0022:  brtrue.s   IL_003b

    IL_0024:  pop
    IL_0025:  ldsfld     class Cimtops.Excel.Cell/'<>c' Cimtops.Excel.Cell/'<>c'::'<>9'
    IL_002a:  ldftn      instance float64 Cimtops.Excel.Cell/'<>c'::'<get_Height>b__21_1'(class Cimtops.Excel.Row)
    IL_0030:  newobj     instance void class [mscorlib]System.Func`2<class Cimtops.Excel.Row,float64>::.ctor(object,
                                                                                                             native int)
    IL_0035:  dup
    IL_0036:  stsfld     class [mscorlib]System.Func`2<class Cimtops.Excel.Row,float64> Cimtops.Excel.Cell/'<>c'::'<>9__21_1'
    IL_003b:  call       float64 [System.Core]System.Linq.Enumerable::Sum<class Cimtops.Excel.Row>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                   class [mscorlib]System.Func`2<!!0,float64>)
    IL_0040:  ret
  } // end of method Cell::get_Height

  .method public hidebysig specialname instance bool 
          get_IsMerged() cil managed
  {
    // Code size       31 (0x1f)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance int32 Cimtops.Excel.CellRange::get_Left()
    IL_0006:  ldarg.0
    IL_0007:  call       instance int32 Cimtops.Excel.CellRange::get_Right()
    IL_000c:  blt.s      IL_001d

    IL_000e:  ldarg.0
    IL_000f:  call       instance int32 Cimtops.Excel.CellRange::get_Top()
    IL_0014:  ldarg.0
    IL_0015:  call       instance int32 Cimtops.Excel.CellRange::get_Bottom()
    IL_001a:  clt
    IL_001c:  ret

    IL_001d:  ldc.i4.1
    IL_001e:  ret
  } // end of method Cell::get_IsMerged

  .method assembly hidebysig specialname rtspecialname 
          instance void  .ctor(class Cimtops.Excel.Sheet sheet,
                               class Microsoft.Office.Interop.Excel.Range cell) cil managed
  {
    // Code size       816 (0x330)
    .maxstack  12
    .locals init (bool V_0,
             int32 V_1,
             int32 V_2,
             int32 V_3,
             class Microsoft.Office.Interop.Excel.Borders V_4,
             object V_5,
             valuetype [System.Drawing]System.Drawing.Color V_6)
    IL_0000:  ldarg.0
    IL_0001:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> Cimtops.Excel.Cell/'<>o__24'::'<>p__0'
    IL_0006:  brtrue.s   IL_0034

    IL_0008:  ldc.i4.0
    IL_0009:  ldc.i4.s   83
    IL_000b:  ldtoken    Cimtops.Excel.Cell
    IL_0010:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_0015:  ldc.i4.1
    IL_0016:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
    IL_001b:  dup
    IL_001c:  ldc.i4.0
    IL_001d:  ldc.i4.0
    IL_001e:  ldnull
    IL_001f:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                               string)
    IL_0024:  stelem.ref
    IL_0025:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::UnaryOperation(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   valuetype [System.Core]System.Linq.Expressions.ExpressionType,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
    IL_002a:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
    IL_002f:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> Cimtops.Excel.Cell/'<>o__24'::'<>p__0'
    IL_0034:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> Cimtops.Excel.Cell/'<>o__24'::'<>p__0'
    IL_0039:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>>::Target
    IL_003e:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> Cimtops.Excel.Cell/'<>o__24'::'<>p__0'
    IL_0043:  ldarg.2
    IL_0044:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get_MergeCells()
    IL_0049:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>::Invoke(!0,
                                                                                                                                                    !1)
    IL_004e:  brtrue.s   IL_0053

    IL_0050:  ldarg.2
    IL_0051:  br.s       IL_0059

    IL_0053:  ldarg.2
    IL_0054:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_MergeArea()
    IL_0059:  call       instance void Cimtops.Excel.CellRange::.ctor(class Microsoft.Office.Interop.Excel.Range)
    IL_005e:  ldarg.0
    IL_005f:  ldarg.1
    IL_0060:  stfld      class Cimtops.Excel.Sheet Cimtops.Excel.Cell::'<Sheet>k__BackingField'
    IL_0065:  ldarg.0
    IL_0066:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> Cimtops.Excel.Cell/'<>o__24'::'<>p__2'
    IL_006b:  brtrue.s   IL_0091

    IL_006d:  ldc.i4.0
    IL_006e:  ldtoken    [mscorlib]System.String
    IL_0073:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_0078:  ldtoken    Cimtops.Excel.Cell
    IL_007d:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_0082:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::Convert(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                            class [mscorlib]System.Type,
                                                                                                                                                            class [mscorlib]System.Type)
    IL_0087:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
    IL_008c:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> Cimtops.Excel.Cell/'<>o__24'::'<>p__2'
    IL_0091:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> Cimtops.Excel.Cell/'<>o__24'::'<>p__2'
    IL_0096:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>>::Target
    IL_009b:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>> Cimtops.Excel.Cell/'<>o__24'::'<>p__2'
    IL_00a0:  ldarg.2
    IL_00a1:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get_Value2()
    IL_00a6:  stloc.s    V_5
    IL_00a8:  ldloc.s    V_5
    IL_00aa:  brtrue.s   IL_00af

    IL_00ac:  ldnull
    IL_00ad:  br.s       IL_00fc

    IL_00af:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> Cimtops.Excel.Cell/'<>o__24'::'<>p__1'
    IL_00b4:  brtrue.s   IL_00e6

    IL_00b6:  ldc.i4.0
    IL_00b7:  ldstr      "ToString"
    IL_00bc:  ldnull
    IL_00bd:  ldtoken    Cimtops.Excel.Cell
    IL_00c2:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_00c7:  ldc.i4.1
    IL_00c8:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
    IL_00cd:  dup
    IL_00ce:  ldc.i4.0
    IL_00cf:  ldc.i4.0
    IL_00d0:  ldnull
    IL_00d1:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                               string)
    IL_00d6:  stelem.ref
    IL_00d7:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                 string,
                                                                                                                                                                 class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                 class [mscorlib]System.Type,
                                                                                                                                                                 class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
    IL_00dc:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
    IL_00e1:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> Cimtops.Excel.Cell/'<>o__24'::'<>p__1'
    IL_00e6:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> Cimtops.Excel.Cell/'<>o__24'::'<>p__1'
    IL_00eb:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Target
    IL_00f0:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> Cimtops.Excel.Cell/'<>o__24'::'<>p__1'
    IL_00f5:  ldloc.s    V_5
    IL_00f7:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>::Invoke(!0,
                                                                                                                                                      !1)
    IL_00fc:  dup
    IL_00fd:  brtrue.s   IL_0105

    IL_00ff:  pop
    IL_0100:  ldstr      ""
    IL_0105:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,string>::Invoke(!0,
                                                                                                                                                      !1)
    IL_010a:  stfld      string Cimtops.Excel.Cell::'<Text>k__BackingField'
    IL_010f:  ldarg.2
    IL_0110:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get_HasFormula()
    IL_0115:  dup
    IL_0116:  stloc.s    V_5
    IL_0118:  isinst     [mscorlib]System.Boolean
    IL_011d:  brfalse.s  IL_012e

    IL_011f:  ldloc.s    V_5
    IL_0121:  unbox.any  [mscorlib]System.Boolean
    IL_0126:  stloc.0
    IL_0127:  ldarg.0
    IL_0128:  ldloc.0
    IL_0129:  stfld      bool Cimtops.Excel.Cell::'<IsFormula>k__BackingField'
    IL_012e:  ldarg.0
    IL_012f:  ldarg.2
    IL_0130:  callvirt   instance class Microsoft.Office.Interop.Excel.Comment Microsoft.Office.Interop.Excel.Range::get_Comment()
    IL_0135:  dup
    IL_0136:  brtrue.s   IL_013c

    IL_0138:  pop
    IL_0139:  ldnull
    IL_013a:  br.s       IL_0150

    IL_013c:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0141:  ldsfld     object [mscorlib]System.Type::Missing
    IL_0146:  ldsfld     object [mscorlib]System.Type::Missing
    IL_014b:  callvirt   instance string Microsoft.Office.Interop.Excel.Comment::Text(object,
                                                                                      object,
                                                                                      object)
    IL_0150:  dup
    IL_0151:  brtrue.s   IL_0159

    IL_0153:  pop
    IL_0154:  ldstr      ""
    IL_0159:  stfld      string Cimtops.Excel.Cell::'<Comment>k__BackingField'
    IL_015e:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32>> Cimtops.Excel.Cell/'<>o__24'::'<>p__4'
    IL_0163:  brtrue.s   IL_0189

    IL_0165:  ldc.i4.0
    IL_0166:  ldtoken    [mscorlib]System.Int32
    IL_016b:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_0170:  ldtoken    Cimtops.Excel.Cell
    IL_0175:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_017a:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::Convert(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                            class [mscorlib]System.Type,
                                                                                                                                                            class [mscorlib]System.Type)
    IL_017f:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
    IL_0184:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32>> Cimtops.Excel.Cell/'<>o__24'::'<>p__4'
    IL_0189:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32>> Cimtops.Excel.Cell/'<>o__24'::'<>p__4'
    IL_018e:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32>>::Target
    IL_0193:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32>> Cimtops.Excel.Cell/'<>o__24'::'<>p__4'
    IL_0198:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class Cimtops.Excel.Cell,object,object>> Cimtops.Excel.Cell/'<>o__24'::'<>p__3'
    IL_019d:  brtrue.s   IL_01d9

    IL_019f:  ldc.i4.2
    IL_01a0:  ldstr      "C2I"
    IL_01a5:  ldnull
    IL_01a6:  ldtoken    Cimtops.Excel.Cell
    IL_01ab:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_01b0:  ldc.i4.2
    IL_01b1:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
    IL_01b6:  dup
    IL_01b7:  ldc.i4.0
    IL_01b8:  ldc.i4.1
    IL_01b9:  ldnull
    IL_01ba:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                               string)
    IL_01bf:  stelem.ref
    IL_01c0:  dup
    IL_01c1:  ldc.i4.1
    IL_01c2:  ldc.i4.0
    IL_01c3:  ldnull
    IL_01c4:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                               string)
    IL_01c9:  stelem.ref
    IL_01ca:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::InvokeMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                 string,
                                                                                                                                                                 class [mscorlib]System.Collections.Generic.IEnumerable`1<class [mscorlib]System.Type>,
                                                                                                                                                                 class [mscorlib]System.Type,
                                                                                                                                                                 class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
    IL_01cf:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class Cimtops.Excel.Cell,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
    IL_01d4:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class Cimtops.Excel.Cell,object,object>> Cimtops.Excel.Cell/'<>o__24'::'<>p__3'
    IL_01d9:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class Cimtops.Excel.Cell,object,object>> Cimtops.Excel.Cell/'<>o__24'::'<>p__3'
    IL_01de:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class Cimtops.Excel.Cell,object,object>>::Target
    IL_01e3:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class Cimtops.Excel.Cell,object,object>> Cimtops.Excel.Cell/'<>o__24'::'<>p__3'
    IL_01e8:  ldarg.0
    IL_01e9:  ldarg.2
    IL_01ea:  callvirt   instance class Microsoft.Office.Interop.Excel.Interior Microsoft.Office.Interop.Excel.Range::get_Interior()
    IL_01ef:  callvirt   instance object Microsoft.Office.Interop.Excel.Interior::get_Color()
    IL_01f4:  callvirt   instance !3 class [mscorlib]System.Func`4<class [System.Core]System.Runtime.CompilerServices.CallSite,class Cimtops.Excel.Cell,object,object>::Invoke(!0,
                                                                                                                                                                               !1,
                                                                                                                                                                               !2)
    IL_01f9:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,int32>::Invoke(!0,
                                                                                                                                                     !1)
    IL_01fe:  dup
    IL_01ff:  ldc.i4     0xff
    IL_0204:  and
    IL_0205:  stloc.1
    IL_0206:  dup
    IL_0207:  ldc.i4.8
    IL_0208:  shr
    IL_0209:  ldc.i4     0xff
    IL_020e:  and
    IL_020f:  stloc.2
    IL_0210:  ldc.i4.s   16
    IL_0212:  shr
    IL_0213:  ldc.i4     0xff
    IL_0218:  and
    IL_0219:  stloc.3
    IL_021a:  ldarg.0
    IL_021b:  ldloc.1
    IL_021c:  ldloc.2
    IL_021d:  ldloc.3
    IL_021e:  call       valuetype [System.Drawing]System.Drawing.Color [System.Drawing]System.Drawing.Color::FromArgb(int32,
                                                                                                                       int32,
                                                                                                                       int32)
    IL_0223:  stloc.s    V_6
    IL_0225:  ldloca.s   V_6
    IL_0227:  call       instance int32 [System.Drawing]System.Drawing.Color::ToArgb()
    IL_022c:  stfld      int32 Cimtops.Excel.Cell::'<BGColor>k__BackingField'
    IL_0231:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> Cimtops.Excel.Cell/'<>o__24'::'<>p__5'
    IL_0236:  brtrue.s   IL_0264

    IL_0238:  ldc.i4.0
    IL_0239:  ldc.i4.s   83
    IL_023b:  ldtoken    Cimtops.Excel.Cell
    IL_0240:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_0245:  ldc.i4.1
    IL_0246:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
    IL_024b:  dup
    IL_024c:  ldc.i4.0
    IL_024d:  ldc.i4.0
    IL_024e:  ldnull
    IL_024f:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                               string)
    IL_0254:  stelem.ref
    IL_0255:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::UnaryOperation(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                                   valuetype [System.Core]System.Linq.Expressions.ExpressionType,
                                                                                                                                                                   class [mscorlib]System.Type,
                                                                                                                                                                   class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
    IL_025a:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
    IL_025f:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> Cimtops.Excel.Cell/'<>o__24'::'<>p__5'
    IL_0264:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> Cimtops.Excel.Cell/'<>o__24'::'<>p__5'
    IL_0269:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>>::Target
    IL_026e:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>> Cimtops.Excel.Cell/'<>o__24'::'<>p__5'
    IL_0273:  ldarg.2
    IL_0274:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get_MergeCells()
    IL_0279:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,bool>::Invoke(!0,
                                                                                                                                                    !1)
    IL_027e:  brtrue.s   IL_0283

    IL_0280:  ldarg.2
    IL_0281:  br.s       IL_0289

    IL_0283:  ldarg.2
    IL_0284:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_MergeArea()
    IL_0289:  callvirt   instance class Microsoft.Office.Interop.Excel.Borders Microsoft.Office.Interop.Excel.Range::get_Borders()
    IL_028e:  stloc.s    V_4
    IL_0290:  ldarg.0
    IL_0291:  ldarg.0
    IL_0292:  ldloc.s    V_4
    IL_0294:  ldc.i4.6
    IL_0295:  call       instance valuetype Microsoft.Office.Interop.Excel.XlLineStyle Cimtops.Excel.Cell::ToStyle(class Microsoft.Office.Interop.Excel.Borders,
                                                                                                                   valuetype Microsoft.Office.Interop.Excel.XlBordersIndex)
    IL_029a:  ldc.i4     0xffffefd2
    IL_029f:  bne.un.s   IL_02b6

    IL_02a1:  ldarg.0
    IL_02a2:  ldloc.s    V_4
    IL_02a4:  ldc.i4.5
    IL_02a5:  call       instance valuetype Microsoft.Office.Interop.Excel.XlLineStyle Cimtops.Excel.Cell::ToStyle(class Microsoft.Office.Interop.Excel.Borders,
                                                                                                                   valuetype Microsoft.Office.Interop.Excel.XlBordersIndex)
    IL_02aa:  ldc.i4     0xffffefd2
    IL_02af:  ceq
    IL_02b1:  ldc.i4.0
    IL_02b2:  ceq
    IL_02b4:  br.s       IL_02b7

    IL_02b6:  ldc.i4.1
    IL_02b7:  newobj     instance void Cimtops.Excel.BorderQ::.ctor(bool)
    IL_02bc:  stfld      class Cimtops.Excel.BorderQ Cimtops.Excel.Cell::'<Border>k__BackingField'
    IL_02c1:  ldarg.0
    IL_02c2:  call       instance class Cimtops.Excel.BorderQ Cimtops.Excel.Cell::get_Border()
    IL_02c7:  callvirt   instance class Cimtops.Excel.CellBorder[] Cimtops.Excel.BorderQ::get_Borders()
    IL_02cc:  ldc.i4.0
    IL_02cd:  ldelem.ref
    IL_02ce:  ldarg.0
    IL_02cf:  ldloc.s    V_4
    IL_02d1:  ldc.i4.8
    IL_02d2:  call       instance valuetype Microsoft.Office.Interop.Excel.XlLineStyle Cimtops.Excel.Cell::ToStyle(class Microsoft.Office.Interop.Excel.Borders,
                                                                                                                   valuetype Microsoft.Office.Interop.Excel.XlBordersIndex)
    IL_02d7:  callvirt   instance void Cimtops.Excel.CellBorder::SetStyle(valuetype Microsoft.Office.Interop.Excel.XlLineStyle)
    IL_02dc:  ldarg.0
    IL_02dd:  call       instance class Cimtops.Excel.BorderQ Cimtops.Excel.Cell::get_Border()
    IL_02e2:  callvirt   instance class Cimtops.Excel.CellBorder[] Cimtops.Excel.BorderQ::get_Borders()
    IL_02e7:  ldc.i4.1
    IL_02e8:  ldelem.ref
    IL_02e9:  ldarg.0
    IL_02ea:  ldloc.s    V_4
    IL_02ec:  ldc.i4.7
    IL_02ed:  call       instance valuetype Microsoft.Office.Interop.Excel.XlLineStyle Cimtops.Excel.Cell::ToStyle(class Microsoft.Office.Interop.Excel.Borders,
                                                                                                                   valuetype Microsoft.Office.Interop.Excel.XlBordersIndex)
    IL_02f2:  callvirt   instance void Cimtops.Excel.CellBorder::SetStyle(valuetype Microsoft.Office.Interop.Excel.XlLineStyle)
    IL_02f7:  ldarg.0
    IL_02f8:  call       instance class Cimtops.Excel.BorderQ Cimtops.Excel.Cell::get_Border()
    IL_02fd:  callvirt   instance class Cimtops.Excel.CellBorder[] Cimtops.Excel.BorderQ::get_Borders()
    IL_0302:  ldc.i4.2
    IL_0303:  ldelem.ref
    IL_0304:  ldarg.0
    IL_0305:  ldloc.s    V_4
    IL_0307:  ldc.i4.s   9
    IL_0309:  call       instance valuetype Microsoft.Office.Interop.Excel.XlLineStyle Cimtops.Excel.Cell::ToStyle(class Microsoft.Office.Interop.Excel.Borders,
                                                                                                                   valuetype Microsoft.Office.Interop.Excel.XlBordersIndex)
    IL_030e:  callvirt   instance void Cimtops.Excel.CellBorder::SetStyle(valuetype Microsoft.Office.Interop.Excel.XlLineStyle)
    IL_0313:  ldarg.0
    IL_0314:  call       instance class Cimtops.Excel.BorderQ Cimtops.Excel.Cell::get_Border()
    IL_0319:  callvirt   instance class Cimtops.Excel.CellBorder[] Cimtops.Excel.BorderQ::get_Borders()
    IL_031e:  ldc.i4.3
    IL_031f:  ldelem.ref
    IL_0320:  ldarg.0
    IL_0321:  ldloc.s    V_4
    IL_0323:  ldc.i4.s   10
    IL_0325:  call       instance valuetype Microsoft.Office.Interop.Excel.XlLineStyle Cimtops.Excel.Cell::ToStyle(class Microsoft.Office.Interop.Excel.Borders,
                                                                                                                   valuetype Microsoft.Office.Interop.Excel.XlBordersIndex)
    IL_032a:  callvirt   instance void Cimtops.Excel.CellBorder::SetStyle(valuetype Microsoft.Office.Interop.Excel.XlLineStyle)
    IL_032f:  ret
  } // end of method Cell::.ctor

  .method private hidebysig instance valuetype Microsoft.Office.Interop.Excel.XlLineStyle 
          ToStyle(class Microsoft.Office.Interop.Excel.Borders borders,
                  valuetype Microsoft.Office.Interop.Excel.XlBordersIndex dir) cil managed
  {
    // Code size       77 (0x4d)
    .maxstack  4
    IL_0000:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype Microsoft.Office.Interop.Excel.XlLineStyle>> Cimtops.Excel.Cell/'<>o__25'::'<>p__0'
    IL_0005:  brtrue.s   IL_002c

    IL_0007:  ldc.i4.s   16
    IL_0009:  ldtoken    Microsoft.Office.Interop.Excel.XlLineStyle
    IL_000e:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_0013:  ldtoken    Cimtops.Excel.Cell
    IL_0018:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_001d:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::Convert(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                            class [mscorlib]System.Type,
                                                                                                                                                            class [mscorlib]System.Type)
    IL_0022:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype Microsoft.Office.Interop.Excel.XlLineStyle>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
    IL_0027:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype Microsoft.Office.Interop.Excel.XlLineStyle>> Cimtops.Excel.Cell/'<>o__25'::'<>p__0'
    IL_002c:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype Microsoft.Office.Interop.Excel.XlLineStyle>> Cimtops.Excel.Cell/'<>o__25'::'<>p__0'
    IL_0031:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype Microsoft.Office.Interop.Excel.XlLineStyle>>::Target
    IL_0036:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype Microsoft.Office.Interop.Excel.XlLineStyle>> Cimtops.Excel.Cell/'<>o__25'::'<>p__0'
    IL_003b:  ldarg.1
    IL_003c:  ldarg.2
    IL_003d:  callvirt   instance class Microsoft.Office.Interop.Excel.Border Microsoft.Office.Interop.Excel.Borders::get_Item(valuetype Microsoft.Office.Interop.Excel.XlBordersIndex)
    IL_0042:  callvirt   instance object Microsoft.Office.Interop.Excel.Border::get_LineStyle()
    IL_0047:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,valuetype Microsoft.Office.Interop.Excel.XlLineStyle>::Invoke(!0,
                                                                                                                                                                                                    !1)
    IL_004c:  ret
  } // end of method Cell::ToStyle

  .method private hidebysig instance int32 
          C2I(object color) cil managed
  {
    // Code size       79 (0x4f)
    .maxstack  1
    IL_0000:  ldarg.1
    IL_0001:  isinst     [mscorlib]System.Double
    IL_0006:  brfalse.s  IL_0010

    IL_0008:  ldarg.1
    IL_0009:  unbox.any  [mscorlib]System.Double
    IL_000e:  conv.i4
    IL_000f:  ret

    IL_0010:  ldarg.1
    IL_0011:  isinst     [mscorlib]System.Int64
    IL_0016:  brfalse.s  IL_0020

    IL_0018:  ldarg.1
    IL_0019:  unbox.any  [mscorlib]System.Int64
    IL_001e:  conv.i4
    IL_001f:  ret

    IL_0020:  ldarg.1
    IL_0021:  isinst     [mscorlib]System.Int32
    IL_0026:  brfalse.s  IL_002f

    IL_0028:  ldarg.1
    IL_0029:  unbox.any  [mscorlib]System.Int32
    IL_002e:  ret

    IL_002f:  ldarg.1
    IL_0030:  isinst     [mscorlib]System.Int16
    IL_0035:  brfalse.s  IL_003e

    IL_0037:  ldarg.1
    IL_0038:  unbox.any  [mscorlib]System.Int16
    IL_003d:  ret

    IL_003e:  ldarg.1
    IL_003f:  isinst     [mscorlib]System.Byte
    IL_0044:  brfalse.s  IL_004d

    IL_0046:  ldarg.1
    IL_0047:  unbox.any  [mscorlib]System.Byte
    IL_004c:  ret

    IL_004d:  ldc.i4.0
    IL_004e:  ret
  } // end of method Cell::C2I

  .method private hidebysig instance bool 
          '<get_Width>b__19_0'(class Cimtops.Excel.Col c) cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       19 (0x13)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.0
    IL_0002:  call       instance int32 Cimtops.Excel.CellRange::get_Top()
    IL_0007:  ldarg.1
    IL_0008:  callvirt   instance int32 Cimtops.Excel.Col::get_Index()
    IL_000d:  call       instance bool Cimtops.Excel.CellRange::Contains(int32,
                                                                         int32)
    IL_0012:  ret
  } // end of method Cell::'<get_Width>b__19_0'

  .method private hidebysig instance bool 
          '<get_Height>b__21_0'(class Cimtops.Excel.Row r) cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       19 (0x13)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  callvirt   instance int32 Cimtops.Excel.Row::get_Index()
    IL_0007:  ldarg.0
    IL_0008:  call       instance int32 Cimtops.Excel.CellRange::get_Left()
    IL_000d:  call       instance bool Cimtops.Excel.CellRange::Contains(int32,
                                                                         int32)
    IL_0012:  ret
  } // end of method Cell::'<get_Height>b__21_0'

  .property instance class Cimtops.Excel.Sheet
          Sheet()
  {
    .get instance class Cimtops.Excel.Sheet Cimtops.Excel.Cell::get_Sheet()
  } // end of property Cell::Sheet
  .property instance string Text()
  {
    .get instance string Cimtops.Excel.Cell::get_Text()
  } // end of property Cell::Text
  .property instance bool IsFormula()
  {
    .get instance bool Cimtops.Excel.Cell::get_IsFormula()
  } // end of property Cell::IsFormula
  .property instance string Comment()
  {
    .get instance string Cimtops.Excel.Cell::get_Comment()
  } // end of property Cell::Comment
  .property instance int32 BGColor()
  {
    .get instance int32 Cimtops.Excel.Cell::get_BGColor()
  } // end of property Cell::BGColor
  .property instance class Cimtops.Excel.BorderQ
          Border()
  {
    .get instance class Cimtops.Excel.BorderQ Cimtops.Excel.Cell::get_Border()
  } // end of property Cell::Border
  .property instance float64 Width()
  {
    .get instance float64 Cimtops.Excel.Cell::get_Width()
  } // end of property Cell::Width
  .property instance float64 Height()
  {
    .get instance float64 Cimtops.Excel.Cell::get_Height()
  } // end of property Cell::Height
  .property instance bool IsMerged()
  {
    .get instance bool Cimtops.Excel.Cell::get_IsMerged()
  } // end of property Cell::IsMerged
} // end of class Cimtops.Excel.Cell

.class public auto ansi beforefieldinit Cimtops.Excel.CellBorder
       extends [mscorlib]System.Object
{
  .field private valuetype Microsoft.Office.Interop.Excel.XlLineStyle '<Style>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig specialname instance valuetype Microsoft.Office.Interop.Excel.XlLineStyle 
          get_Style() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      valuetype Microsoft.Office.Interop.Excel.XlLineStyle Cimtops.Excel.CellBorder::'<Style>k__BackingField'
    IL_0006:  ret
  } // end of method CellBorder::get_Style

  .method private hidebysig specialname instance void 
          set_Style(valuetype Microsoft.Office.Interop.Excel.XlLineStyle 'value') cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       8 (0x8)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  stfld      valuetype Microsoft.Office.Interop.Excel.XlLineStyle Cimtops.Excel.CellBorder::'<Style>k__BackingField'
    IL_0007:  ret
  } // end of method CellBorder::set_Style

  .method public hidebysig instance void 
          SetStyle(valuetype Microsoft.Office.Interop.Excel.XlLineStyle style) cil managed
  {
    // Code size       8 (0x8)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  call       instance void Cimtops.Excel.CellBorder::set_Style(valuetype Microsoft.Office.Interop.Excel.XlLineStyle)
    IL_0007:  ret
  } // end of method CellBorder::SetStyle

  .method public hidebysig virtual instance string 
          ToString() cil managed
  {
    // Code size       21 (0x15)
    .maxstack  1
    .locals init (valuetype Microsoft.Office.Interop.Excel.XlLineStyle V_0)
    IL_0000:  ldarg.0
    IL_0001:  call       instance valuetype Microsoft.Office.Interop.Excel.XlLineStyle Cimtops.Excel.CellBorder::get_Style()
    IL_0006:  stloc.0
    IL_0007:  ldloca.s   V_0
    IL_0009:  constrained. Microsoft.Office.Interop.Excel.XlLineStyle
    IL_000f:  callvirt   instance string [mscorlib]System.Object::ToString()
    IL_0014:  ret
  } // end of method CellBorder::ToString

  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       18 (0x12)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldc.i4     0xffffefd2
    IL_0006:  stfld      valuetype Microsoft.Office.Interop.Excel.XlLineStyle Cimtops.Excel.CellBorder::'<Style>k__BackingField'
    IL_000b:  ldarg.0
    IL_000c:  call       instance void [mscorlib]System.Object::.ctor()
    IL_0011:  ret
  } // end of method CellBorder::.ctor

  .property instance valuetype Microsoft.Office.Interop.Excel.XlLineStyle
          Style()
  {
    .get instance valuetype Microsoft.Office.Interop.Excel.XlLineStyle Cimtops.Excel.CellBorder::get_Style()
    .set instance void Cimtops.Excel.CellBorder::set_Style(valuetype Microsoft.Office.Interop.Excel.XlLineStyle)
  } // end of property CellBorder::Style
} // end of class Cimtops.Excel.CellBorder

.class public auto ansi sealed Cimtops.Excel.Direction
       extends [mscorlib]System.Enum
{
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype Cimtops.Excel.Direction TOP = int32(0x00000000)
  .field public static literal valuetype Cimtops.Excel.Direction LEFT = int32(0x00000001)
  .field public static literal valuetype Cimtops.Excel.Direction BOTTOM = int32(0x00000002)
  .field public static literal valuetype Cimtops.Excel.Direction RIGHT = int32(0x00000003)
  .field public static literal valuetype Cimtops.Excel.Direction NONE = int32(0x00000063)
} // end of class Cimtops.Excel.Direction

.class public auto ansi beforefieldinit Cimtops.Excel.BorderQ
       extends [mscorlib]System.Object
{
  .class auto ansi serializable sealed nested private beforefieldinit '<>c'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static initonly class Cimtops.Excel.BorderQ/'<>c' '<>9'
    .field public static class [mscorlib]System.Func`2<int32,class Cimtops.Excel.CellBorder> '<>9__8_0'
    .method private hidebysig specialname rtspecialname static 
            void  .cctor() cil managed
    {
      // Code size       11 (0xb)
      .maxstack  8
      IL_0000:  newobj     instance void Cimtops.Excel.BorderQ/'<>c'::.ctor()
      IL_0005:  stsfld     class Cimtops.Excel.BorderQ/'<>c' Cimtops.Excel.BorderQ/'<>c'::'<>9'
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

    .method assembly hidebysig instance class Cimtops.Excel.CellBorder 
            '<.ctor>b__8_0'(int32 n) cil managed
    {
      // Code size       6 (0x6)
      .maxstack  8
      IL_0000:  newobj     instance void Cimtops.Excel.CellBorder::.ctor()
      IL_0005:  ret
    } // end of method '<>c'::'<.ctor>b__8_0'

  } // end of class '<>c'

  .class auto ansi sealed nested private beforefieldinit '<>c__DisplayClass13_0'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public class Cimtops.Excel.BorderQ '<>4__this'
    .field public bool onlyReal
    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
      IL_0006:  ret
    } // end of method '<>c__DisplayClass13_0'::.ctor

    .method assembly hidebysig instance bool 
            '<HasAll>b__0'(valuetype Cimtops.Excel.Direction n) cil managed
    {
      // Code size       31 (0x1f)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  ldfld      class Cimtops.Excel.BorderQ Cimtops.Excel.BorderQ/'<>c__DisplayClass13_0'::'<>4__this'
      IL_0006:  ldarg.0
      IL_0007:  ldfld      class Cimtops.Excel.BorderQ Cimtops.Excel.BorderQ/'<>c__DisplayClass13_0'::'<>4__this'
      IL_000c:  call       instance class Cimtops.Excel.CellBorder[] Cimtops.Excel.BorderQ::get_Borders()
      IL_0011:  ldarg.1
      IL_0012:  ldelem.ref
      IL_0013:  ldarg.0
      IL_0014:  ldfld      bool Cimtops.Excel.BorderQ/'<>c__DisplayClass13_0'::onlyReal
      IL_0019:  call       instance bool Cimtops.Excel.BorderQ::Exists(class Cimtops.Excel.CellBorder,
                                                                       bool)
      IL_001e:  ret
    } // end of method '<>c__DisplayClass13_0'::'<HasAll>b__0'

  } // end of class '<>c__DisplayClass13_0'

  .field private class Cimtops.Excel.CellBorder[] '<Borders>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private bool '<IsDiagonal>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig specialname instance class Cimtops.Excel.CellBorder[] 
          get_Borders() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class Cimtops.Excel.CellBorder[] Cimtops.Excel.BorderQ::'<Borders>k__BackingField'
    IL_0006:  ret
  } // end of method BorderQ::get_Borders

  .method private hidebysig specialname instance void 
          set_Borders(class Cimtops.Excel.CellBorder[] 'value') cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       8 (0x8)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  stfld      class Cimtops.Excel.CellBorder[] Cimtops.Excel.BorderQ::'<Borders>k__BackingField'
    IL_0007:  ret
  } // end of method BorderQ::set_Borders

  .method public hidebysig specialname instance bool 
          get_IsDiagonal() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      bool Cimtops.Excel.BorderQ::'<IsDiagonal>k__BackingField'
    IL_0006:  ret
  } // end of method BorderQ::get_IsDiagonal

  .method private hidebysig specialname instance void 
          set_IsDiagonal(bool 'value') cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       8 (0x8)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  stfld      bool Cimtops.Excel.BorderQ::'<IsDiagonal>k__BackingField'
    IL_0007:  ret
  } // end of method BorderQ::set_IsDiagonal

  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(bool isDiagonal) cil managed
  {
    // Code size       68 (0x44)
    .maxstack  4
    IL_0000:  ldarg.0
    IL_0001:  ldc.i4.0
    IL_0002:  ldc.i4.4
    IL_0003:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<int32> [System.Core]System.Linq.Enumerable::Range(int32,
                                                                                                                                    int32)
    IL_0008:  ldsfld     class [mscorlib]System.Func`2<int32,class Cimtops.Excel.CellBorder> Cimtops.Excel.BorderQ/'<>c'::'<>9__8_0'
    IL_000d:  dup
    IL_000e:  brtrue.s   IL_0027

    IL_0010:  pop
    IL_0011:  ldsfld     class Cimtops.Excel.BorderQ/'<>c' Cimtops.Excel.BorderQ/'<>c'::'<>9'
    IL_0016:  ldftn      instance class Cimtops.Excel.CellBorder Cimtops.Excel.BorderQ/'<>c'::'<.ctor>b__8_0'(int32)
    IL_001c:  newobj     instance void class [mscorlib]System.Func`2<int32,class Cimtops.Excel.CellBorder>::.ctor(object,
                                                                                                                  native int)
    IL_0021:  dup
    IL_0022:  stsfld     class [mscorlib]System.Func`2<int32,class Cimtops.Excel.CellBorder> Cimtops.Excel.BorderQ/'<>c'::'<>9__8_0'
    IL_0027:  call       class [mscorlib]System.Collections.Generic.IEnumerable`1<!!1> [System.Core]System.Linq.Enumerable::Select<int32,class Cimtops.Excel.CellBorder>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                                                         class [mscorlib]System.Func`2<!!0,!!1>)
    IL_002c:  call       !!0[] [System.Core]System.Linq.Enumerable::ToArray<class Cimtops.Excel.CellBorder>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
    IL_0031:  stfld      class Cimtops.Excel.CellBorder[] Cimtops.Excel.BorderQ::'<Borders>k__BackingField'
    IL_0036:  ldarg.0
    IL_0037:  call       instance void [mscorlib]System.Object::.ctor()
    IL_003c:  ldarg.0
    IL_003d:  ldarg.1
    IL_003e:  call       instance void Cimtops.Excel.BorderQ::set_IsDiagonal(bool)
    IL_0043:  ret
  } // end of method BorderQ::.ctor

  .method public hidebysig specialname instance bool 
          get_IsEncloded() cil managed
  {
    // Code size       24 (0x18)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance class Cimtops.Excel.CellBorder[] Cimtops.Excel.BorderQ::get_Borders()
    IL_0006:  ldarg.0
    IL_0007:  ldftn      instance bool Cimtops.Excel.BorderQ::'<get_IsEncloded>b__10_0'(class Cimtops.Excel.CellBorder)
    IL_000d:  newobj     instance void class [mscorlib]System.Func`2<class Cimtops.Excel.CellBorder,bool>::.ctor(object,
                                                                                                                 native int)
    IL_0012:  call       bool [System.Core]System.Linq.Enumerable::All<class Cimtops.Excel.CellBorder>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                       class [mscorlib]System.Func`2<!!0,bool>)
    IL_0017:  ret
  } // end of method BorderQ::get_IsEncloded

  .method public hidebysig instance bool 
          Exists(valuetype Cimtops.Excel.Direction[] index) cil managed
  {
    .param [1]
    .custom instance void [mscorlib]System.ParamArrayAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       19 (0x13)
    .maxstack  8
    IL_0000:  ldarg.1
    IL_0001:  ldarg.0
    IL_0002:  ldftn      instance bool Cimtops.Excel.BorderQ::'<Exists>b__11_0'(valuetype Cimtops.Excel.Direction)
    IL_0008:  newobj     instance void class [mscorlib]System.Func`2<valuetype Cimtops.Excel.Direction,bool>::.ctor(object,
                                                                                                                    native int)
    IL_000d:  call       bool [System.Core]System.Linq.Enumerable::Any<valuetype Cimtops.Excel.Direction>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                          class [mscorlib]System.Func`2<!!0,bool>)
    IL_0012:  ret
  } // end of method BorderQ::Exists

  .method public hidebysig instance bool 
          HasAll(valuetype Cimtops.Excel.Direction[] index) cil managed
  {
    .param [1]
    .custom instance void [mscorlib]System.ParamArrayAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       9 (0x9)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldc.i4.0
    IL_0002:  ldarg.1
    IL_0003:  call       instance bool Cimtops.Excel.BorderQ::HasAll(bool,
                                                                     valuetype Cimtops.Excel.Direction[])
    IL_0008:  ret
  } // end of method BorderQ::HasAll

  .method public hidebysig instance bool 
          HasAll(bool onlyReal,
                 valuetype Cimtops.Excel.Direction[] index) cil managed
  {
    .param [2]
    .custom instance void [mscorlib]System.ParamArrayAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       39 (0x27)
    .maxstack  3
    .locals init (class Cimtops.Excel.BorderQ/'<>c__DisplayClass13_0' V_0)
    IL_0000:  newobj     instance void Cimtops.Excel.BorderQ/'<>c__DisplayClass13_0'::.ctor()
    IL_0005:  stloc.0
    IL_0006:  ldloc.0
    IL_0007:  ldarg.0
    IL_0008:  stfld      class Cimtops.Excel.BorderQ Cimtops.Excel.BorderQ/'<>c__DisplayClass13_0'::'<>4__this'
    IL_000d:  ldloc.0
    IL_000e:  ldarg.1
    IL_000f:  stfld      bool Cimtops.Excel.BorderQ/'<>c__DisplayClass13_0'::onlyReal
    IL_0014:  ldarg.2
    IL_0015:  ldloc.0
    IL_0016:  ldftn      instance bool Cimtops.Excel.BorderQ/'<>c__DisplayClass13_0'::'<HasAll>b__0'(valuetype Cimtops.Excel.Direction)
    IL_001c:  newobj     instance void class [mscorlib]System.Func`2<valuetype Cimtops.Excel.Direction,bool>::.ctor(object,
                                                                                                                    native int)
    IL_0021:  call       bool [System.Core]System.Linq.Enumerable::All<valuetype Cimtops.Excel.Direction>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                          class [mscorlib]System.Func`2<!!0,bool>)
    IL_0026:  ret
  } // end of method BorderQ::HasAll

  .method private hidebysig instance bool 
          Exists(class Cimtops.Excel.CellBorder border,
                 bool onlyReal) cil managed
  {
    // Code size       69 (0x45)
    .maxstack  2
    .locals init (valuetype Microsoft.Office.Interop.Excel.XlLineStyle V_0)
    IL_0000:  ldarg.1
    IL_0001:  callvirt   instance valuetype Microsoft.Office.Interop.Excel.XlLineStyle Cimtops.Excel.CellBorder::get_Style()
    IL_0006:  stloc.0
    IL_0007:  ldloc.0
    IL_0008:  ldc.i4.1
    IL_0009:  bgt.s      IL_0031

    IL_000b:  ldloc.0
    IL_000c:  ldc.i4     0xffffefe9
    IL_0011:  sub
    IL_0012:  switch     ( 
                          IL_003c,
                          IL_003e,
                          IL_0043,
                          IL_0043,
                          IL_003e)
    IL_002b:  ldloc.0
    IL_002c:  ldc.i4.1
    IL_002d:  beq.s      IL_003c

    IL_002f:  br.s       IL_0043

    IL_0031:  ldloc.0
    IL_0032:  ldc.i4.4
    IL_0033:  sub
    IL_0034:  ldc.i4.1
    IL_0035:  ble.un.s   IL_003e

    IL_0037:  ldloc.0
    IL_0038:  ldc.i4.s   13
    IL_003a:  bne.un.s   IL_0043

    IL_003c:  ldc.i4.1
    IL_003d:  ret

    IL_003e:  ldarg.2
    IL_003f:  ldc.i4.0
    IL_0040:  ceq
    IL_0042:  ret

    IL_0043:  ldc.i4.0
    IL_0044:  ret
  } // end of method BorderQ::Exists

  .method private hidebysig instance bool 
          '<get_IsEncloded>b__10_0'(class Cimtops.Excel.CellBorder b) cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       9 (0x9)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  ldc.i4.0
    IL_0003:  call       instance bool Cimtops.Excel.BorderQ::Exists(class Cimtops.Excel.CellBorder,
                                                                     bool)
    IL_0008:  ret
  } // end of method BorderQ::'<get_IsEncloded>b__10_0'

  .method private hidebysig instance bool 
          '<Exists>b__11_0'(valuetype Cimtops.Excel.Direction n) cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       16 (0x10)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.0
    IL_0002:  call       instance class Cimtops.Excel.CellBorder[] Cimtops.Excel.BorderQ::get_Borders()
    IL_0007:  ldarg.1
    IL_0008:  ldelem.ref
    IL_0009:  ldc.i4.0
    IL_000a:  call       instance bool Cimtops.Excel.BorderQ::Exists(class Cimtops.Excel.CellBorder,
                                                                     bool)
    IL_000f:  ret
  } // end of method BorderQ::'<Exists>b__11_0'

  .property instance class Cimtops.Excel.CellBorder[]
          Borders()
  {
    .get instance class Cimtops.Excel.CellBorder[] Cimtops.Excel.BorderQ::get_Borders()
    .set instance void Cimtops.Excel.BorderQ::set_Borders(class Cimtops.Excel.CellBorder[])
  } // end of property BorderQ::Borders
  .property instance bool IsDiagonal()
  {
    .get instance bool Cimtops.Excel.BorderQ::get_IsDiagonal()
    .set instance void Cimtops.Excel.BorderQ::set_IsDiagonal(bool)
  } // end of property BorderQ::IsDiagonal
  .property instance bool IsEncloded()
  {
    .get instance bool Cimtops.Excel.BorderQ::get_IsEncloded()
  } // end of property BorderQ::IsEncloded
} // end of class Cimtops.Excel.BorderQ

.class public auto ansi beforefieldinit Cimtops.Excel.CellModify
       extends [mscorlib]System.Object
{
  .field public valuetype [mscorlib]System.Nullable`1<valuetype [System.Drawing]System.Drawing.Color> Color
  .field public string Comment
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
    IL_0006:  ret
  } // end of method CellModify::.ctor

} // end of class Cimtops.Excel.CellModify

.class public auto ansi beforefieldinit Cimtops.Excel.CellRange
       extends [mscorlib]System.Object
{
  .field private initonly int32 '<Left>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly int32 '<Top>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly int32 '<Right>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly int32 '<Bottom>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig specialname instance int32 
          get_Left() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      int32 Cimtops.Excel.CellRange::'<Left>k__BackingField'
    IL_0006:  ret
  } // end of method CellRange::get_Left

  .method public hidebysig specialname instance int32 
          get_Top() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      int32 Cimtops.Excel.CellRange::'<Top>k__BackingField'
    IL_0006:  ret
  } // end of method CellRange::get_Top

  .method public hidebysig specialname instance int32 
          get_Right() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      int32 Cimtops.Excel.CellRange::'<Right>k__BackingField'
    IL_0006:  ret
  } // end of method CellRange::get_Right

  .method public hidebysig specialname instance int32 
          get_Bottom() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      int32 Cimtops.Excel.CellRange::'<Bottom>k__BackingField'
    IL_0006:  ret
  } // end of method CellRange::get_Bottom

  .method public hidebysig specialname instance int32 
          get_FirstNum() cil managed
  {
    // Code size       2 (0x2)
    .maxstack  8
    IL_0000:  ldc.i4.1
    IL_0001:  ret
  } // end of method CellRange::get_FirstNum

  .method assembly hidebysig specialname rtspecialname 
          instance void  .ctor(class Microsoft.Office.Interop.Excel.Range range) cil managed
  {
    // Code size       83 (0x53)
    .maxstack  3
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
    IL_0006:  ldarg.0
    IL_0007:  ldarg.1
    IL_0008:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Column()
    IL_000d:  stfld      int32 Cimtops.Excel.CellRange::'<Left>k__BackingField'
    IL_0012:  ldarg.0
    IL_0013:  ldarg.1
    IL_0014:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Row()
    IL_0019:  stfld      int32 Cimtops.Excel.CellRange::'<Top>k__BackingField'
    IL_001e:  ldarg.0
    IL_001f:  ldarg.0
    IL_0020:  call       instance int32 Cimtops.Excel.CellRange::get_Left()
    IL_0025:  ldarg.1
    IL_0026:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_Columns()
    IL_002b:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Count()
    IL_0030:  add
    IL_0031:  ldc.i4.1
    IL_0032:  sub
    IL_0033:  stfld      int32 Cimtops.Excel.CellRange::'<Right>k__BackingField'
    IL_0038:  ldarg.0
    IL_0039:  ldarg.0
    IL_003a:  call       instance int32 Cimtops.Excel.CellRange::get_Top()
    IL_003f:  ldarg.1
    IL_0040:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_Rows()
    IL_0045:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Count()
    IL_004a:  add
    IL_004b:  ldc.i4.1
    IL_004c:  sub
    IL_004d:  stfld      int32 Cimtops.Excel.CellRange::'<Bottom>k__BackingField'
    IL_0052:  ret
  } // end of method CellRange::.ctor

  .method public hidebysig instance string 
          ToText() cil managed
  {
    // Code size       118 (0x76)
    .maxstack  5
    IL_0000:  ldarg.0
    IL_0001:  call       instance int32 Cimtops.Excel.CellRange::get_Left()
    IL_0006:  ldarg.0
    IL_0007:  call       instance int32 Cimtops.Excel.CellRange::get_Right()
    IL_000c:  bne.un.s   IL_003d

    IL_000e:  ldarg.0
    IL_000f:  call       instance int32 Cimtops.Excel.CellRange::get_Top()
    IL_0014:  ldarg.0
    IL_0015:  call       instance int32 Cimtops.Excel.CellRange::get_Bottom()
    IL_001a:  bne.un.s   IL_003d

    IL_001c:  ldarg.0
    IL_001d:  call       instance int32 Cimtops.Excel.CellRange::get_Top()
    IL_0022:  ldarg.0
    IL_0023:  call       instance int32 Cimtops.Excel.CellRange::get_Left()
    IL_0028:  ldarg.0
    IL_0029:  call       instance int32 Cimtops.Excel.CellRange::get_FirstNum()
    IL_002e:  call       string Cimtops.Excel.XLRangeUtil::ToAddressName(int32,
                                                                         int32,
                                                                         int32)
    IL_0033:  dup
    IL_0034:  brtrue.s   IL_003c

    IL_0036:  pop
    IL_0037:  ldstr      ""
    IL_003c:  ret

    IL_003d:  ldarg.0
    IL_003e:  call       instance int32 Cimtops.Excel.CellRange::get_Top()
    IL_0043:  ldarg.0
    IL_0044:  call       instance int32 Cimtops.Excel.CellRange::get_Left()
    IL_0049:  ldarg.0
    IL_004a:  call       instance int32 Cimtops.Excel.CellRange::get_FirstNum()
    IL_004f:  call       string Cimtops.Excel.XLRangeUtil::ToAddressName(int32,
                                                                         int32,
                                                                         int32)
    IL_0054:  ldstr      ":"
    IL_0059:  ldarg.0
    IL_005a:  call       instance int32 Cimtops.Excel.CellRange::get_Bottom()
    IL_005f:  ldarg.0
    IL_0060:  call       instance int32 Cimtops.Excel.CellRange::get_Right()
    IL_0065:  ldarg.0
    IL_0066:  call       instance int32 Cimtops.Excel.CellRange::get_FirstNum()
    IL_006b:  call       string Cimtops.Excel.XLRangeUtil::ToAddressName(int32,
                                                                         int32,
                                                                         int32)
    IL_0070:  call       string [mscorlib]System.String::Concat(string,
                                                                string,
                                                                string)
    IL_0075:  ret
  } // end of method CellRange::ToText

  .method public hidebysig instance bool 
          IsOverlap(class Cimtops.Excel.CellRange other) cil managed
  {
    // Code size       71 (0x47)
    .maxstack  2
    IL_0000:  ldarg.1
    IL_0001:  brtrue.s   IL_0005

    IL_0003:  ldc.i4.0
    IL_0004:  ret

    IL_0005:  ldarg.0
    IL_0006:  call       instance int32 Cimtops.Excel.CellRange::get_Right()
    IL_000b:  ldarg.1
    IL_000c:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Left()
    IL_0011:  bge.s      IL_0015

    IL_0013:  ldc.i4.0
    IL_0014:  ret

    IL_0015:  ldarg.1
    IL_0016:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Right()
    IL_001b:  ldarg.0
    IL_001c:  call       instance int32 Cimtops.Excel.CellRange::get_Left()
    IL_0021:  bge.s      IL_0025

    IL_0023:  ldc.i4.0
    IL_0024:  ret

    IL_0025:  ldarg.0
    IL_0026:  call       instance int32 Cimtops.Excel.CellRange::get_Bottom()
    IL_002b:  ldarg.1
    IL_002c:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Top()
    IL_0031:  bge.s      IL_0035

    IL_0033:  ldc.i4.0
    IL_0034:  ret

    IL_0035:  ldarg.1
    IL_0036:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Bottom()
    IL_003b:  ldarg.0
    IL_003c:  call       instance int32 Cimtops.Excel.CellRange::get_Top()
    IL_0041:  bge.s      IL_0045

    IL_0043:  ldc.i4.0
    IL_0044:  ret

    IL_0045:  ldc.i4.1
    IL_0046:  ret
  } // end of method CellRange::IsOverlap

  .method public hidebysig instance bool 
          Contains(int32 row,
                   int32 col) cil managed
  {
    // Code size       42 (0x2a)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance int32 Cimtops.Excel.CellRange::get_Left()
    IL_0006:  ldarg.2
    IL_0007:  bgt.s      IL_0028

    IL_0009:  ldarg.2
    IL_000a:  ldarg.0
    IL_000b:  call       instance int32 Cimtops.Excel.CellRange::get_Right()
    IL_0010:  bgt.s      IL_0028

    IL_0012:  ldarg.0
    IL_0013:  call       instance int32 Cimtops.Excel.CellRange::get_Top()
    IL_0018:  ldarg.1
    IL_0019:  bgt.s      IL_0028

    IL_001b:  ldarg.1
    IL_001c:  ldarg.0
    IL_001d:  call       instance int32 Cimtops.Excel.CellRange::get_Bottom()
    IL_0022:  cgt
    IL_0024:  ldc.i4.0
    IL_0025:  ceq
    IL_0027:  ret

    IL_0028:  ldc.i4.0
    IL_0029:  ret
  } // end of method CellRange::Contains

  .method public hidebysig virtual instance string 
          ToString() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance string Cimtops.Excel.CellRange::ToText()
    IL_0006:  ret
  } // end of method CellRange::ToString

  .property instance int32 Left()
  {
    .get instance int32 Cimtops.Excel.CellRange::get_Left()
  } // end of property CellRange::Left
  .property instance int32 Top()
  {
    .get instance int32 Cimtops.Excel.CellRange::get_Top()
  } // end of property CellRange::Top
  .property instance int32 Right()
  {
    .get instance int32 Cimtops.Excel.CellRange::get_Right()
  } // end of property CellRange::Right
  .property instance int32 Bottom()
  {
    .get instance int32 Cimtops.Excel.CellRange::get_Bottom()
  } // end of property CellRange::Bottom
  .property instance int32 FirstNum()
  {
    .get instance int32 Cimtops.Excel.CellRange::get_FirstNum()
  } // end of property CellRange::FirstNum
} // end of class Cimtops.Excel.CellRange

.class public auto ansi beforefieldinit Cimtops.Excel.Book
       extends [mscorlib]System.Object
{
  .field private initonly class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Sheet> sheets
  .field private string '<Path>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig specialname instance string 
          get_Path() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      string Cimtops.Excel.Book::'<Path>k__BackingField'
    IL_0006:  ret
  } // end of method Book::get_Path

  .method private hidebysig specialname instance void 
          set_Path(string 'value') cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       8 (0x8)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  stfld      string Cimtops.Excel.Book::'<Path>k__BackingField'
    IL_0007:  ret
  } // end of method Book::set_Path

  .method public hidebysig specialname instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Sheet> 
          get_Worksheets() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Sheet> Cimtops.Excel.Book::sheets
    IL_0006:  ret
  } // end of method Book::get_Worksheets

  .method assembly hidebysig specialname rtspecialname 
          instance void  .ctor(class Microsoft.Office.Interop.Excel.Workbook book,
                               [opt] string sheetName) cil managed
  {
    .param [2] = nullref
    // Code size       135 (0x87)
    .maxstack  3
    .locals init (class [mscorlib]System.Collections.IEnumerator V_0,
             class Microsoft.Office.Interop.Excel.Worksheet V_1,
             class [mscorlib]System.IDisposable V_2)
    IL_0000:  ldarg.0
    IL_0001:  newobj     instance void class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Sheet>::.ctor()
    IL_0006:  stfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Sheet> Cimtops.Excel.Book::sheets
    IL_000b:  ldarg.0
    IL_000c:  call       instance void [mscorlib]System.Object::.ctor()
    IL_0011:  ldarg.0
    IL_0012:  ldarg.1
    IL_0013:  callvirt   instance string Microsoft.Office.Interop.Excel._Workbook::get_Path()
    IL_0018:  ldarg.1
    IL_0019:  callvirt   instance string Microsoft.Office.Interop.Excel._Workbook::get_Name()
    IL_001e:  call       string [mscorlib]System.IO.Path::Combine(string,
                                                                  string)
    IL_0023:  call       instance void Cimtops.Excel.Book::set_Path(string)
    IL_0028:  ldarg.1
    IL_0029:  callvirt   instance class Microsoft.Office.Interop.Excel.Sheets Microsoft.Office.Interop.Excel._Workbook::get_Sheets()
    IL_002e:  callvirt   instance class [mscorlib]System.Collections.IEnumerator Microsoft.Office.Interop.Excel.Sheets::GetEnumerator()
    IL_0033:  stloc.0
    .try
    {
      IL_0034:  br.s       IL_006b

      IL_0036:  ldloc.0
      IL_0037:  callvirt   instance object [mscorlib]System.Collections.IEnumerator::get_Current()
      IL_003c:  isinst     Microsoft.Office.Interop.Excel.Worksheet
      IL_0041:  stloc.1
      IL_0042:  ldloc.1
      IL_0043:  brfalse.s  IL_006b

      IL_0045:  ldarg.2
      IL_0046:  brtrue.s   IL_004b

      IL_0048:  ldc.i4.1
      IL_0049:  br.s       IL_0057

      IL_004b:  ldarg.2
      IL_004c:  ldloc.1
      IL_004d:  callvirt   instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
      IL_0052:  call       instance bool [mscorlib]System.String::Equals(string)
      IL_0057:  brfalse.s  IL_006b

      IL_0059:  ldarg.0
      IL_005a:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Sheet> Cimtops.Excel.Book::sheets
      IL_005f:  ldarg.0
      IL_0060:  ldloc.1
      IL_0061:  newobj     instance void Cimtops.Excel.Sheet::.ctor(class Cimtops.Excel.Book,
                                                                    class Microsoft.Office.Interop.Excel.Worksheet)
      IL_0066:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Sheet>::Add(!0)
      IL_006b:  ldloc.0
      IL_006c:  callvirt   instance bool [mscorlib]System.Collections.IEnumerator::MoveNext()
      IL_0071:  brtrue.s   IL_0036

      IL_0073:  leave.s    IL_0086

    }  // end .try
    finally
    {
      IL_0075:  ldloc.0
      IL_0076:  isinst     [mscorlib]System.IDisposable
      IL_007b:  stloc.2
      IL_007c:  ldloc.2
      IL_007d:  brfalse.s  IL_0085

      IL_007f:  ldloc.2
      IL_0080:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
      IL_0085:  endfinally
    }  // end handler
    IL_0086:  ret
  } // end of method Book::.ctor

  .property instance string Path()
  {
    .get instance string Cimtops.Excel.Book::get_Path()
    .set instance void Cimtops.Excel.Book::set_Path(string)
  } // end of property Book::Path
  .property instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Sheet>
          Worksheets()
  {
    .get instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Sheet> Cimtops.Excel.Book::get_Worksheets()
  } // end of property Book::Worksheets
} // end of class Cimtops.Excel.Book

.class public auto ansi beforefieldinit Cimtops.Excel.BookFactory
       extends [mscorlib]System.Object
{
  .method public hidebysig static class Cimtops.Excel.Book 
          Create(class Microsoft.Office.Interop.Excel.Workbook book,
                 [out] class [mscorlib]System.Exception& ex) cil managed
  {
    // Code size       29 (0x1d)
    .maxstack  2
    .locals init (class Cimtops.Excel.Book V_0,
             class [mscorlib]System.Exception V_1)
    IL_0000:  ldarg.1
    IL_0001:  ldnull
    IL_0002:  stind.ref
    IL_0003:  ldarg.0
    IL_0004:  brtrue.s   IL_0008

    IL_0006:  ldnull
    IL_0007:  ret

    IL_0008:  nop
    .try
    {
      IL_0009:  ldarg.0
      IL_000a:  ldnull
      IL_000b:  newobj     instance void Cimtops.Excel.Book::.ctor(class Microsoft.Office.Interop.Excel.Workbook,
                                                                   string)
      IL_0010:  stloc.0
      IL_0011:  leave.s    IL_001b

    }  // end .try
    catch [mscorlib]System.Exception 
    {
      IL_0013:  stloc.1
      IL_0014:  ldarg.1
      IL_0015:  ldloc.1
      IL_0016:  stind.ref
      IL_0017:  ldnull
      IL_0018:  stloc.0
      IL_0019:  leave.s    IL_001b

    }  // end handler
    IL_001b:  ldloc.0
    IL_001c:  ret
  } // end of method BookFactory::Create

  .method public hidebysig static class Cimtops.Excel.Sheet 
          Create(class Microsoft.Office.Interop.Excel.Workbook book,
                 class Microsoft.Office.Interop.Excel.Worksheet sheet,
                 [out] class [mscorlib]System.Exception& ex) cil managed
  {
    // Code size       47 (0x2f)
    .maxstack  2
    .locals init (class Cimtops.Excel.Sheet V_0,
             class [mscorlib]System.Exception V_1)
    IL_0000:  ldarg.2
    IL_0001:  ldnull
    IL_0002:  stind.ref
    IL_0003:  ldarg.0
    IL_0004:  brfalse.s  IL_0009

    IL_0006:  ldarg.1
    IL_0007:  brtrue.s   IL_000b

    IL_0009:  ldnull
    IL_000a:  ret

    IL_000b:  nop
    .try
    {
      IL_000c:  ldarg.0
      IL_000d:  ldarg.1
      IL_000e:  callvirt   instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
      IL_0013:  newobj     instance void Cimtops.Excel.Book::.ctor(class Microsoft.Office.Interop.Excel.Workbook,
                                                                   string)
      IL_0018:  call       instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Sheet> Cimtops.Excel.Book::get_Worksheets()
      IL_001d:  call       !!0 [System.Core]System.Linq.Enumerable::FirstOrDefault<class Cimtops.Excel.Sheet>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>)
      IL_0022:  stloc.0
      IL_0023:  leave.s    IL_002d

    }  // end .try
    catch [mscorlib]System.Exception 
    {
      IL_0025:  stloc.1
      IL_0026:  ldarg.2
      IL_0027:  ldloc.1
      IL_0028:  stind.ref
      IL_0029:  ldnull
      IL_002a:  stloc.0
      IL_002b:  leave.s    IL_002d

    }  // end handler
    IL_002d:  ldloc.0
    IL_002e:  ret
  } // end of method BookFactory::Create

  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
    IL_0006:  ret
  } // end of method BookFactory::.ctor

} // end of class Cimtops.Excel.BookFactory

.class public auto ansi sealed Cimtops.Excel.DeviceType
       extends [mscorlib]System.Enum
{
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype Cimtops.Excel.DeviceType Windows = int32(0x00000000)
  .field public static literal valuetype Cimtops.Excel.DeviceType IOS = int32(0x00000001)
  .field public static literal valuetype Cimtops.Excel.DeviceType Android = int32(0x00000002)
} // end of class Cimtops.Excel.DeviceType

.class public auto ansi beforefieldinit Cimtops.Excel.Row
       extends [mscorlib]System.Object
{
  .class auto ansi sealed nested private beforefieldinit '<>c__DisplayClass11_0'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public int32 x
    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
      IL_0006:  ret
    } // end of method '<>c__DisplayClass11_0'::.ctor

    .method assembly hidebysig instance bool 
            '<Find>b__0'(class Cimtops.Excel.Cell c) cil managed
    {
      // Code size       19 (0x13)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  ldarg.1
      IL_0002:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Top()
      IL_0007:  ldarg.0
      IL_0008:  ldfld      int32 Cimtops.Excel.Row/'<>c__DisplayClass11_0'::x
      IL_000d:  callvirt   instance bool Cimtops.Excel.CellRange::Contains(int32,
                                                                           int32)
      IL_0012:  ret
    } // end of method '<>c__DisplayClass11_0'::'<Find>b__0'

  } // end of class '<>c__DisplayClass11_0'

  .field private initonly class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Cell> cells
  .field private initonly int32 '<Index>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly float64 '<Height>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig specialname instance int32 
          get_Index() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      int32 Cimtops.Excel.Row::'<Index>k__BackingField'
    IL_0006:  ret
  } // end of method Row::get_Index

  .method public hidebysig specialname instance float64 
          get_Height() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      float64 Cimtops.Excel.Row::'<Height>k__BackingField'
    IL_0006:  ret
  } // end of method Row::get_Height

  .method public hidebysig specialname instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Cell> 
          get_Cells() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Cell> Cimtops.Excel.Row::cells
    IL_0006:  ret
  } // end of method Row::get_Cells

  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(int32 index,
                               float64 height) cil managed
  {
    // Code size       32 (0x20)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  newobj     instance void class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Cell>::.ctor()
    IL_0006:  stfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Cell> Cimtops.Excel.Row::cells
    IL_000b:  ldarg.0
    IL_000c:  call       instance void [mscorlib]System.Object::.ctor()
    IL_0011:  ldarg.0
    IL_0012:  ldarg.1
    IL_0013:  stfld      int32 Cimtops.Excel.Row::'<Index>k__BackingField'
    IL_0018:  ldarg.0
    IL_0019:  ldarg.2
    IL_001a:  stfld      float64 Cimtops.Excel.Row::'<Height>k__BackingField'
    IL_001f:  ret
  } // end of method Row::.ctor

  .method public hidebysig instance void 
          Add(class Cimtops.Excel.Cell cell) cil managed
  {
    // Code size       13 (0xd)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Cell> Cimtops.Excel.Row::cells
    IL_0006:  ldarg.1
    IL_0007:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Cell>::Add(!0)
    IL_000c:  ret
  } // end of method Row::Add

  .method public hidebysig instance class Cimtops.Excel.Cell 
          Find(int32 x) cil managed
  {
    // Code size       37 (0x25)
    .maxstack  3
    .locals init (class Cimtops.Excel.Row/'<>c__DisplayClass11_0' V_0)
    IL_0000:  newobj     instance void Cimtops.Excel.Row/'<>c__DisplayClass11_0'::.ctor()
    IL_0005:  stloc.0
    IL_0006:  ldloc.0
    IL_0007:  ldarg.1
    IL_0008:  stfld      int32 Cimtops.Excel.Row/'<>c__DisplayClass11_0'::x
    IL_000d:  ldarg.0
    IL_000e:  call       instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Cell> Cimtops.Excel.Row::get_Cells()
    IL_0013:  ldloc.0
    IL_0014:  ldftn      instance bool Cimtops.Excel.Row/'<>c__DisplayClass11_0'::'<Find>b__0'(class Cimtops.Excel.Cell)
    IL_001a:  newobj     instance void class [mscorlib]System.Func`2<class Cimtops.Excel.Cell,bool>::.ctor(object,
                                                                                                           native int)
    IL_001f:  call       !!0 [System.Core]System.Linq.Enumerable::FirstOrDefault<class Cimtops.Excel.Cell>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                           class [mscorlib]System.Func`2<!!0,bool>)
    IL_0024:  ret
  } // end of method Row::Find

  .property instance int32 Index()
  {
    .get instance int32 Cimtops.Excel.Row::get_Index()
  } // end of property Row::Index
  .property instance float64 Height()
  {
    .get instance float64 Cimtops.Excel.Row::get_Height()
  } // end of property Row::Height
  .property instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Cell>
          Cells()
  {
    .get instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Cell> Cimtops.Excel.Row::get_Cells()
  } // end of property Row::Cells
} // end of class Cimtops.Excel.Row

.class public auto ansi beforefieldinit Cimtops.Excel.Col
       extends [mscorlib]System.Object
{
  .field private initonly int32 '<Index>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly float64 '<Width>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig specialname instance int32 
          get_Index() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      int32 Cimtops.Excel.Col::'<Index>k__BackingField'
    IL_0006:  ret
  } // end of method Col::get_Index

  .method public hidebysig specialname instance float64 
          get_Width() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      float64 Cimtops.Excel.Col::'<Width>k__BackingField'
    IL_0006:  ret
  } // end of method Col::get_Width

  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(int32 index,
                               float64 width) cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
    IL_0006:  ldarg.0
    IL_0007:  ldarg.1
    IL_0008:  stfld      int32 Cimtops.Excel.Col::'<Index>k__BackingField'
    IL_000d:  ldarg.0
    IL_000e:  ldarg.2
    IL_000f:  stfld      float64 Cimtops.Excel.Col::'<Width>k__BackingField'
    IL_0014:  ret
  } // end of method Col::.ctor

  .property instance int32 Index()
  {
    .get instance int32 Cimtops.Excel.Col::get_Index()
  } // end of property Col::Index
  .property instance float64 Width()
  {
    .get instance float64 Cimtops.Excel.Col::get_Width()
  } // end of property Col::Width
} // end of class Cimtops.Excel.Col

.class public auto ansi beforefieldinit Cimtops.Excel.Sheet
       extends [mscorlib]System.Object
{
  .class auto ansi serializable sealed nested private beforefieldinit '<>c'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static initonly class Cimtops.Excel.Sheet/'<>c' '<>9'
    .field public static class [mscorlib]System.Func`2<class Cimtops.Excel.Col,float64> '<>9__18_0'
    .field public static class [mscorlib]System.Func`2<class Cimtops.Excel.Row,float64> '<>9__20_0'
    .method private hidebysig specialname rtspecialname static 
            void  .cctor() cil managed
    {
      // Code size       11 (0xb)
      .maxstack  8
      IL_0000:  newobj     instance void Cimtops.Excel.Sheet/'<>c'::.ctor()
      IL_0005:  stsfld     class Cimtops.Excel.Sheet/'<>c' Cimtops.Excel.Sheet/'<>c'::'<>9'
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

    .method assembly hidebysig instance float64 
            '<get_Width>b__18_0'(class Cimtops.Excel.Col c) cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance float64 Cimtops.Excel.Col::get_Width()
      IL_0006:  ret
    } // end of method '<>c'::'<get_Width>b__18_0'

    .method assembly hidebysig instance float64 
            '<get_Height>b__20_0'(class Cimtops.Excel.Row c) cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance float64 Cimtops.Excel.Row::get_Height()
      IL_0006:  ret
    } // end of method '<>c'::'<get_Height>b__20_0'

  } // end of class '<>c'

  .class auto ansi sealed nested private beforefieldinit '<>c__DisplayClass25_0'
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
    } // end of method '<>c__DisplayClass25_0'::.ctor

    .method assembly hidebysig instance bool 
            '<GetCell>b__0'(class Cimtops.Excel.Row r) cil managed
    {
      // Code size       15 (0xf)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance int32 Cimtops.Excel.Row::get_Index()
      IL_0006:  ldarg.0
      IL_0007:  ldfld      int32 Cimtops.Excel.Sheet/'<>c__DisplayClass25_0'::row
      IL_000c:  ceq
      IL_000e:  ret
    } // end of method '<>c__DisplayClass25_0'::'<GetCell>b__0'

  } // end of class '<>c__DisplayClass25_0'

  .class abstract auto ansi sealed nested private beforefieldinit '<>o__26'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> '<>p__0'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> '<>p__1'
    .field public static class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> '<>p__2'
  } // end of class '<>o__26'

  .class auto ansi sealed nested private beforefieldinit '<>c__DisplayClass26_0'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public class Cimtops.Excel.Row row
    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
      IL_0006:  ret
    } // end of method '<>c__DisplayClass26_0'::.ctor

  } // end of class '<>c__DisplayClass26_0'

  .class auto ansi sealed nested private beforefieldinit '<>c__DisplayClass26_1'
         extends [mscorlib]System.Object
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public int32 x
    .field public class Cimtops.Excel.Sheet/'<>c__DisplayClass26_0' 'CS$<>8__locals1'
    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [mscorlib]System.Object::.ctor()
      IL_0006:  ret
    } // end of method '<>c__DisplayClass26_1'::.ctor

    .method assembly hidebysig instance bool 
            '<.ctor>b__0'(class Cimtops.Excel.Cell c) cil managed
    {
      // Code size       29 (0x1d)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  ldarg.0
      IL_0002:  ldfld      class Cimtops.Excel.Sheet/'<>c__DisplayClass26_0' Cimtops.Excel.Sheet/'<>c__DisplayClass26_1'::'CS$<>8__locals1'
      IL_0007:  ldfld      class Cimtops.Excel.Row Cimtops.Excel.Sheet/'<>c__DisplayClass26_0'::row
      IL_000c:  callvirt   instance int32 Cimtops.Excel.Row::get_Index()
      IL_0011:  ldarg.0
      IL_0012:  ldfld      int32 Cimtops.Excel.Sheet/'<>c__DisplayClass26_1'::x
      IL_0017:  callvirt   instance bool Cimtops.Excel.CellRange::Contains(int32,
                                                                           int32)
      IL_001c:  ret
    } // end of method '<>c__DisplayClass26_1'::'<.ctor>b__0'

  } // end of class '<>c__DisplayClass26_1'

  .field private initonly class Cimtops.Excel.Book '<Book>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly string '<Name>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly bool '<IsMultiPrint>k__BackingField'
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .field private initonly class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Row> rows
  .field private initonly class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Col> cols
  .method public hidebysig specialname instance class Cimtops.Excel.Book 
          get_Book() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class Cimtops.Excel.Book Cimtops.Excel.Sheet::'<Book>k__BackingField'
    IL_0006:  ret
  } // end of method Sheet::get_Book

  .method public hidebysig specialname instance string 
          get_Name() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      string Cimtops.Excel.Sheet::'<Name>k__BackingField'
    IL_0006:  ret
  } // end of method Sheet::get_Name

  .method public hidebysig specialname instance bool 
          get_IsMultiPrint() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      bool Cimtops.Excel.Sheet::'<IsMultiPrint>k__BackingField'
    IL_0006:  ret
  } // end of method Sheet::get_IsMultiPrint

  .method public hidebysig specialname instance bool 
          get_IsEmpty() cil managed
  {
    // Code size       30 (0x1e)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Row> Cimtops.Excel.Sheet::rows
    IL_0006:  callvirt   instance int32 class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Row>::get_Count()
    IL_000b:  brfalse.s  IL_001c

    IL_000d:  ldarg.0
    IL_000e:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Col> Cimtops.Excel.Sheet::cols
    IL_0013:  callvirt   instance int32 class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Col>::get_Count()
    IL_0018:  ldc.i4.0
    IL_0019:  ceq
    IL_001b:  ret

    IL_001c:  ldc.i4.1
    IL_001d:  ret
  } // end of method Sheet::get_IsEmpty

  .method public hidebysig specialname instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Row> 
          get_Rows() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Row> Cimtops.Excel.Sheet::rows
    IL_0006:  ret
  } // end of method Sheet::get_Rows

  .method public hidebysig specialname instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Col> 
          get_Cols() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Col> Cimtops.Excel.Sheet::cols
    IL_0006:  ret
  } // end of method Sheet::get_Cols

  .method public hidebysig specialname instance float64 
          get_Width() cil managed
  {
    // Code size       43 (0x2b)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Col> Cimtops.Excel.Sheet::cols
    IL_0006:  ldsfld     class [mscorlib]System.Func`2<class Cimtops.Excel.Col,float64> Cimtops.Excel.Sheet/'<>c'::'<>9__18_0'
    IL_000b:  dup
    IL_000c:  brtrue.s   IL_0025

    IL_000e:  pop
    IL_000f:  ldsfld     class Cimtops.Excel.Sheet/'<>c' Cimtops.Excel.Sheet/'<>c'::'<>9'
    IL_0014:  ldftn      instance float64 Cimtops.Excel.Sheet/'<>c'::'<get_Width>b__18_0'(class Cimtops.Excel.Col)
    IL_001a:  newobj     instance void class [mscorlib]System.Func`2<class Cimtops.Excel.Col,float64>::.ctor(object,
                                                                                                             native int)
    IL_001f:  dup
    IL_0020:  stsfld     class [mscorlib]System.Func`2<class Cimtops.Excel.Col,float64> Cimtops.Excel.Sheet/'<>c'::'<>9__18_0'
    IL_0025:  call       float64 [System.Core]System.Linq.Enumerable::Sum<class Cimtops.Excel.Col>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                   class [mscorlib]System.Func`2<!!0,float64>)
    IL_002a:  ret
  } // end of method Sheet::get_Width

  .method public hidebysig specialname instance float64 
          get_Height() cil managed
  {
    // Code size       43 (0x2b)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Row> Cimtops.Excel.Sheet::rows
    IL_0006:  ldsfld     class [mscorlib]System.Func`2<class Cimtops.Excel.Row,float64> Cimtops.Excel.Sheet/'<>c'::'<>9__20_0'
    IL_000b:  dup
    IL_000c:  brtrue.s   IL_0025

    IL_000e:  pop
    IL_000f:  ldsfld     class Cimtops.Excel.Sheet/'<>c' Cimtops.Excel.Sheet/'<>c'::'<>9'
    IL_0014:  ldftn      instance float64 Cimtops.Excel.Sheet/'<>c'::'<get_Height>b__20_0'(class Cimtops.Excel.Row)
    IL_001a:  newobj     instance void class [mscorlib]System.Func`2<class Cimtops.Excel.Row,float64>::.ctor(object,
                                                                                                             native int)
    IL_001f:  dup
    IL_0020:  stsfld     class [mscorlib]System.Func`2<class Cimtops.Excel.Row,float64> Cimtops.Excel.Sheet/'<>c'::'<>9__20_0'
    IL_0025:  call       float64 [System.Core]System.Linq.Enumerable::Sum<class Cimtops.Excel.Row>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                   class [mscorlib]System.Func`2<!!0,float64>)
    IL_002a:  ret
  } // end of method Sheet::get_Height

  .method public hidebysig instance class Cimtops.Excel.Cell 
          UpOf(class Cimtops.Excel.Cell cell) cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Top()
    IL_0007:  ldc.i4.1
    IL_0008:  sub
    IL_0009:  ldarg.1
    IL_000a:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Left()
    IL_000f:  call       instance class Cimtops.Excel.Cell Cimtops.Excel.Sheet::GetCell(int32,
                                                                                        int32)
    IL_0014:  ret
  } // end of method Sheet::UpOf

  .method public hidebysig instance class Cimtops.Excel.Cell 
          DownOf(class Cimtops.Excel.Cell cell) cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Bottom()
    IL_0007:  ldc.i4.1
    IL_0008:  add
    IL_0009:  ldarg.1
    IL_000a:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Left()
    IL_000f:  call       instance class Cimtops.Excel.Cell Cimtops.Excel.Sheet::GetCell(int32,
                                                                                        int32)
    IL_0014:  ret
  } // end of method Sheet::DownOf

  .method public hidebysig instance class Cimtops.Excel.Cell 
          LeftOf(class Cimtops.Excel.Cell cell) cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Top()
    IL_0007:  ldarg.1
    IL_0008:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Left()
    IL_000d:  ldc.i4.1
    IL_000e:  sub
    IL_000f:  call       instance class Cimtops.Excel.Cell Cimtops.Excel.Sheet::GetCell(int32,
                                                                                        int32)
    IL_0014:  ret
  } // end of method Sheet::LeftOf

  .method public hidebysig instance class Cimtops.Excel.Cell 
          RightOf(class Cimtops.Excel.Cell cell) cil managed
  {
    // Code size       21 (0x15)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldarg.1
    IL_0002:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Top()
    IL_0007:  ldarg.1
    IL_0008:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Right()
    IL_000d:  ldc.i4.1
    IL_000e:  add
    IL_000f:  call       instance class Cimtops.Excel.Cell Cimtops.Excel.Sheet::GetCell(int32,
                                                                                        int32)
    IL_0014:  ret
  } // end of method Sheet::RightOf

  .method public hidebysig instance class Cimtops.Excel.Cell 
          GetCell(int32 row,
                  int32 col) cil managed
  {
    // Code size       49 (0x31)
    .maxstack  3
    .locals init (class Cimtops.Excel.Sheet/'<>c__DisplayClass25_0' V_0)
    IL_0000:  newobj     instance void Cimtops.Excel.Sheet/'<>c__DisplayClass25_0'::.ctor()
    IL_0005:  stloc.0
    IL_0006:  ldloc.0
    IL_0007:  ldarg.1
    IL_0008:  stfld      int32 Cimtops.Excel.Sheet/'<>c__DisplayClass25_0'::row
    IL_000d:  ldarg.0
    IL_000e:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Row> Cimtops.Excel.Sheet::rows
    IL_0013:  ldloc.0
    IL_0014:  ldftn      instance bool Cimtops.Excel.Sheet/'<>c__DisplayClass25_0'::'<GetCell>b__0'(class Cimtops.Excel.Row)
    IL_001a:  newobj     instance void class [mscorlib]System.Func`2<class Cimtops.Excel.Row,bool>::.ctor(object,
                                                                                                          native int)
    IL_001f:  call       !!0 [System.Core]System.Linq.Enumerable::FirstOrDefault<class Cimtops.Excel.Row>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                          class [mscorlib]System.Func`2<!!0,bool>)
    IL_0024:  dup
    IL_0025:  brtrue.s   IL_002a

    IL_0027:  pop
    IL_0028:  ldnull
    IL_0029:  ret

    IL_002a:  ldarg.2
    IL_002b:  call       instance class Cimtops.Excel.Cell Cimtops.Excel.Row::Find(int32)
    IL_0030:  ret
  } // end of method Sheet::GetCell

  .method assembly hidebysig specialname rtspecialname 
          instance void  .ctor(class Cimtops.Excel.Book book,
                               class Microsoft.Office.Interop.Excel.Worksheet sheet) cil managed
  {
    // Code size       893 (0x37d)
    .maxstack  8
    .locals init (string V_0,
             valuetype [mscorlib]System.Nullable`1<int32> V_1,
             valuetype [mscorlib]System.Nullable`1<int32> V_2,
             class Microsoft.Office.Interop.Excel.Range V_3,
             int32 V_4,
             int32 V_5,
             int32 V_6,
             class Cimtops.Excel.Row V_7,
             int32 V_8,
             float64 V_9,
             class Cimtops.Excel.Row V_10,
             object V_11,
             valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<class Cimtops.Excel.Row> V_12,
             class Cimtops.Excel.Sheet/'<>c__DisplayClass26_0' V_13,
             class Cimtops.Excel.Sheet/'<>c__DisplayClass26_1' V_14,
             class Cimtops.Excel.Cell V_15,
             class Microsoft.Office.Interop.Excel.Range V_16,
             int32 V_17,
             float64 V_18,
             class Cimtops.Excel.Col V_19)
    IL_0000:  ldarg.0
    IL_0001:  newobj     instance void class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Row>::.ctor()
    IL_0006:  stfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Row> Cimtops.Excel.Sheet::rows
    IL_000b:  ldarg.0
    IL_000c:  newobj     instance void class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Col>::.ctor()
    IL_0011:  stfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Col> Cimtops.Excel.Sheet::cols
    IL_0016:  ldarg.0
    IL_0017:  call       instance void [mscorlib]System.Object::.ctor()
    IL_001c:  ldarg.0
    IL_001d:  ldarg.1
    IL_001e:  stfld      class Cimtops.Excel.Book Cimtops.Excel.Sheet::'<Book>k__BackingField'
    IL_0023:  ldarg.0
    IL_0024:  ldarg.2
    IL_0025:  callvirt   instance string Microsoft.Office.Interop.Excel._Worksheet::get_Name()
    IL_002a:  stfld      string Cimtops.Excel.Sheet::'<Name>k__BackingField'
    IL_002f:  ldarg.2
    IL_0030:  callvirt   instance class Microsoft.Office.Interop.Excel.PageSetup Microsoft.Office.Interop.Excel._Worksheet::get_PageSetup()
    IL_0035:  callvirt   instance string Microsoft.Office.Interop.Excel.PageSetup::get_PrintArea()
    IL_003a:  stloc.0
    IL_003b:  ldloc.0
    IL_003c:  call       bool [mscorlib]System.String::IsNullOrEmpty(string)
    IL_0041:  brtrue     IL_037c

    IL_0046:  ldloc.0
    IL_0047:  call       string Cimtops.Excel.XLRangeUtil::AdjustRange(string)
    IL_004c:  stloc.0
    IL_004d:  ldarg.0
    IL_004e:  ldloc.0
    IL_004f:  ldc.i4.s   44
    IL_0051:  call       bool [System.Core]System.Linq.Enumerable::Contains<char>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                  !!0)
    IL_0056:  brtrue.s   IL_0091

    IL_0058:  ldc.i4.1
    IL_0059:  ldarg.2
    IL_005a:  callvirt   instance class Microsoft.Office.Interop.Excel.PageSetup Microsoft.Office.Interop.Excel._Worksheet::get_PageSetup()
    IL_005f:  callvirt   instance class Microsoft.Office.Interop.Excel.Pages Microsoft.Office.Interop.Excel.PageSetup::get_Pages()
    IL_0064:  dup
    IL_0065:  brtrue.s   IL_0073

    IL_0067:  pop
    IL_0068:  ldloca.s   V_2
    IL_006a:  initobj    valuetype [mscorlib]System.Nullable`1<int32>
    IL_0070:  ldloc.2
    IL_0071:  br.s       IL_007d

    IL_0073:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Pages::get_Count()
    IL_0078:  newobj     instance void valuetype [mscorlib]System.Nullable`1<int32>::.ctor(!0)
    IL_007d:  stloc.1
    IL_007e:  ldloca.s   V_1
    IL_0080:  call       instance !0 valuetype [mscorlib]System.Nullable`1<int32>::GetValueOrDefault()
    IL_0085:  clt
    IL_0087:  ldloca.s   V_1
    IL_0089:  call       instance bool valuetype [mscorlib]System.Nullable`1<int32>::get_HasValue()
    IL_008e:  and
    IL_008f:  br.s       IL_0092

    IL_0091:  ldc.i4.1
    IL_0092:  stfld      bool Cimtops.Excel.Sheet::'<IsMultiPrint>k__BackingField'
    IL_0097:  ldarg.0
    IL_0098:  call       instance bool Cimtops.Excel.Sheet::get_IsMultiPrint()
    IL_009d:  brtrue     IL_037c

    IL_00a2:  ldarg.2
    IL_00a3:  ldloc.0
    IL_00a4:  ldsfld     object [mscorlib]System.Type::Missing
    IL_00a9:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Range(object,
                                                                                                                                  object)
    IL_00ae:  stloc.3
    IL_00af:  ldloc.3
    IL_00b0:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Row()
    IL_00b5:  dup
    IL_00b6:  ldloc.3
    IL_00b7:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_Rows()
    IL_00bc:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Count()
    IL_00c1:  add
    IL_00c2:  ldc.i4.1
    IL_00c3:  sub
    IL_00c4:  stloc.s    V_4
    IL_00c6:  ldloc.3
    IL_00c7:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Column()
    IL_00cc:  stloc.s    V_5
    IL_00ce:  ldloc.s    V_5
    IL_00d0:  ldloc.3
    IL_00d1:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel.Range::get_Columns()
    IL_00d6:  callvirt   instance int32 Microsoft.Office.Interop.Excel.Range::get_Count()
    IL_00db:  add
    IL_00dc:  ldc.i4.1
    IL_00dd:  sub
    IL_00de:  stloc.s    V_6
    IL_00e0:  stloc.s    V_8
    IL_00e2:  br         IL_0186

    IL_00e7:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> Cimtops.Excel.Sheet/'<>o__26'::'<>p__0'
    IL_00ec:  brtrue.s   IL_011d

    IL_00ee:  ldc.i4.0
    IL_00ef:  ldstr      "Height"
    IL_00f4:  ldtoken    Cimtops.Excel.Sheet
    IL_00f9:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_00fe:  ldc.i4.1
    IL_00ff:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
    IL_0104:  dup
    IL_0105:  ldc.i4.0
    IL_0106:  ldc.i4.0
    IL_0107:  ldnull
    IL_0108:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                               string)
    IL_010d:  stelem.ref
    IL_010e:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::GetMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                              string,
                                                                                                                                                              class [mscorlib]System.Type,
                                                                                                                                                              class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
    IL_0113:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
    IL_0118:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> Cimtops.Excel.Sheet/'<>o__26'::'<>p__0'
    IL_011d:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> Cimtops.Excel.Sheet/'<>o__26'::'<>p__0'
    IL_0122:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Target
    IL_0127:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> Cimtops.Excel.Sheet/'<>o__26'::'<>p__0'
    IL_012c:  ldarg.2
    IL_012d:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Rows()
    IL_0132:  ldloc.s    V_8
    IL_0134:  box        [mscorlib]System.Int32
    IL_0139:  ldsfld     object [mscorlib]System.Type::Missing
    IL_013e:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get__Default(object,
                                                                                            object)
    IL_0143:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>::Invoke(!0,
                                                                                                                                                      !1)
    IL_0148:  dup
    IL_0149:  stloc.s    V_11
    IL_014b:  isinst     [mscorlib]System.Double
    IL_0150:  brfalse.s  IL_015d

    IL_0152:  ldloc.s    V_11
    IL_0154:  unbox.any  [mscorlib]System.Double
    IL_0159:  stloc.s    V_9
    IL_015b:  br.s       IL_0168

    IL_015d:  ldc.r8     1.
    IL_0166:  stloc.s    V_9
    IL_0168:  ldloc.s    V_8
    IL_016a:  ldloc.s    V_9
    IL_016c:  newobj     instance void Cimtops.Excel.Row::.ctor(int32,
                                                                float64)
    IL_0171:  stloc.s    V_10
    IL_0173:  ldarg.0
    IL_0174:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Row> Cimtops.Excel.Sheet::rows
    IL_0179:  ldloc.s    V_10
    IL_017b:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Row>::Add(!0)
    IL_0180:  ldloc.s    V_8
    IL_0182:  ldc.i4.1
    IL_0183:  add
    IL_0184:  stloc.s    V_8
    IL_0186:  ldloc.s    V_8
    IL_0188:  ldloc.s    V_4
    IL_018a:  ble        IL_00e7

    IL_018f:  ldnull
    IL_0190:  stloc.s    V_7
    IL_0192:  ldarg.0
    IL_0193:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Row> Cimtops.Excel.Sheet::rows
    IL_0198:  callvirt   instance valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<!0> class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Row>::GetEnumerator()
    IL_019d:  stloc.s    V_12
    .try
    {
      IL_019f:  br         IL_02af

      IL_01a4:  newobj     instance void Cimtops.Excel.Sheet/'<>c__DisplayClass26_0'::.ctor()
      IL_01a9:  stloc.s    V_13
      IL_01ab:  ldloc.s    V_13
      IL_01ad:  ldloca.s   V_12
      IL_01af:  call       instance !0 valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<class Cimtops.Excel.Row>::get_Current()
      IL_01b4:  stfld      class Cimtops.Excel.Row Cimtops.Excel.Sheet/'<>c__DisplayClass26_0'::row
      IL_01b9:  newobj     instance void Cimtops.Excel.Sheet/'<>c__DisplayClass26_1'::.ctor()
      IL_01be:  stloc.s    V_14
      IL_01c0:  ldloc.s    V_14
      IL_01c2:  ldloc.s    V_13
      IL_01c4:  stfld      class Cimtops.Excel.Sheet/'<>c__DisplayClass26_0' Cimtops.Excel.Sheet/'<>c__DisplayClass26_1'::'CS$<>8__locals1'
      IL_01c9:  ldloc.s    V_14
      IL_01cb:  ldloc.s    V_5
      IL_01cd:  stfld      int32 Cimtops.Excel.Sheet/'<>c__DisplayClass26_1'::x
      IL_01d2:  br         IL_0298

      IL_01d7:  ldloc.s    V_7
      IL_01d9:  brtrue.s   IL_01de

      IL_01db:  ldnull
      IL_01dc:  br.s       IL_01f7

      IL_01de:  ldloc.s    V_7
      IL_01e0:  call       instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Cell> Cimtops.Excel.Row::get_Cells()
      IL_01e5:  ldloc.s    V_14
      IL_01e7:  ldftn      instance bool Cimtops.Excel.Sheet/'<>c__DisplayClass26_1'::'<.ctor>b__0'(class Cimtops.Excel.Cell)
      IL_01ed:  newobj     instance void class [mscorlib]System.Func`2<class Cimtops.Excel.Cell,bool>::.ctor(object,
                                                                                                             native int)
      IL_01f2:  call       !!0 [System.Core]System.Linq.Enumerable::FirstOrDefault<class Cimtops.Excel.Cell>(class [mscorlib]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                             class [mscorlib]System.Func`2<!!0,bool>)
      IL_01f7:  stloc.s    V_15
      IL_01f9:  ldloc.s    V_15
      IL_01fb:  brtrue.s   IL_0275

      IL_01fd:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> Cimtops.Excel.Sheet/'<>o__26'::'<>p__1'
      IL_0202:  brtrue.s   IL_0228

      IL_0204:  ldc.i4.0
      IL_0205:  ldtoken    Microsoft.Office.Interop.Excel.Range
      IL_020a:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_020f:  ldtoken    Cimtops.Excel.Sheet
      IL_0214:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
      IL_0219:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::Convert(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                              class [mscorlib]System.Type,
                                                                                                                                                              class [mscorlib]System.Type)
      IL_021e:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
      IL_0223:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> Cimtops.Excel.Sheet/'<>o__26'::'<>p__1'
      IL_0228:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> Cimtops.Excel.Sheet/'<>o__26'::'<>p__1'
      IL_022d:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>>::Target
      IL_0232:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>> Cimtops.Excel.Sheet/'<>o__26'::'<>p__1'
      IL_0237:  ldarg.2
      IL_0238:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Cells()
      IL_023d:  ldloc.s    V_14
      IL_023f:  ldfld      class Cimtops.Excel.Sheet/'<>c__DisplayClass26_0' Cimtops.Excel.Sheet/'<>c__DisplayClass26_1'::'CS$<>8__locals1'
      IL_0244:  ldfld      class Cimtops.Excel.Row Cimtops.Excel.Sheet/'<>c__DisplayClass26_0'::row
      IL_0249:  callvirt   instance int32 Cimtops.Excel.Row::get_Index()
      IL_024e:  box        [mscorlib]System.Int32
      IL_0253:  ldloc.s    V_14
      IL_0255:  ldfld      int32 Cimtops.Excel.Sheet/'<>c__DisplayClass26_1'::x
      IL_025a:  box        [mscorlib]System.Int32
      IL_025f:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get__Default(object,
                                                                                              object)
      IL_0264:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,class Microsoft.Office.Interop.Excel.Range>::Invoke(!0,
                                                                                                                                                                                            !1)
      IL_0269:  stloc.s    V_16
      IL_026b:  ldarg.0
      IL_026c:  ldloc.s    V_16
      IL_026e:  newobj     instance void Cimtops.Excel.Cell::.ctor(class Cimtops.Excel.Sheet,
                                                                   class Microsoft.Office.Interop.Excel.Range)
      IL_0273:  stloc.s    V_15
      IL_0275:  ldloc.s    V_14
      IL_0277:  ldfld      class Cimtops.Excel.Sheet/'<>c__DisplayClass26_0' Cimtops.Excel.Sheet/'<>c__DisplayClass26_1'::'CS$<>8__locals1'
      IL_027c:  ldfld      class Cimtops.Excel.Row Cimtops.Excel.Sheet/'<>c__DisplayClass26_0'::row
      IL_0281:  ldloc.s    V_15
      IL_0283:  callvirt   instance void Cimtops.Excel.Row::Add(class Cimtops.Excel.Cell)
      IL_0288:  ldloc.s    V_14
      IL_028a:  ldloc.s    V_15
      IL_028c:  callvirt   instance int32 Cimtops.Excel.CellRange::get_Right()
      IL_0291:  ldc.i4.1
      IL_0292:  add
      IL_0293:  stfld      int32 Cimtops.Excel.Sheet/'<>c__DisplayClass26_1'::x
      IL_0298:  ldloc.s    V_14
      IL_029a:  ldfld      int32 Cimtops.Excel.Sheet/'<>c__DisplayClass26_1'::x
      IL_029f:  ldloc.s    V_6
      IL_02a1:  ble        IL_01d7

      IL_02a6:  ldloc.s    V_13
      IL_02a8:  ldfld      class Cimtops.Excel.Row Cimtops.Excel.Sheet/'<>c__DisplayClass26_0'::row
      IL_02ad:  stloc.s    V_7
      IL_02af:  ldloca.s   V_12
      IL_02b1:  call       instance bool valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<class Cimtops.Excel.Row>::MoveNext()
      IL_02b6:  brtrue     IL_01a4

      IL_02bb:  leave.s    IL_02cb

    }  // end .try
    finally
    {
      IL_02bd:  ldloca.s   V_12
      IL_02bf:  constrained. valuetype [mscorlib]System.Collections.Generic.List`1/Enumerator<class Cimtops.Excel.Row>
      IL_02c5:  callvirt   instance void [mscorlib]System.IDisposable::Dispose()
      IL_02ca:  endfinally
    }  // end handler
    IL_02cb:  ldloc.s    V_5
    IL_02cd:  stloc.s    V_17
    IL_02cf:  br         IL_0373

    IL_02d4:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> Cimtops.Excel.Sheet/'<>o__26'::'<>p__2'
    IL_02d9:  brtrue.s   IL_030a

    IL_02db:  ldc.i4.0
    IL_02dc:  ldstr      "Width"
    IL_02e1:  ldtoken    Cimtops.Excel.Sheet
    IL_02e6:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_02eb:  ldc.i4.1
    IL_02ec:  newarr     [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo
    IL_02f1:  dup
    IL_02f2:  ldc.i4.0
    IL_02f3:  ldc.i4.0
    IL_02f4:  ldnull
    IL_02f5:  call       class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo::Create(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfoFlags,
                                                                                                                                                                               string)
    IL_02fa:  stelem.ref
    IL_02fb:  call       class [System.Core]System.Runtime.CompilerServices.CallSiteBinder [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.Binder::GetMember(valuetype [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpBinderFlags,
                                                                                                                                                              string,
                                                                                                                                                              class [mscorlib]System.Type,
                                                                                                                                                              class [mscorlib]System.Collections.Generic.IEnumerable`1<class [Microsoft.CSharp]Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo>)
    IL_0300:  call       class [System.Core]System.Runtime.CompilerServices.CallSite`1<!0> class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Create(class [System.Core]System.Runtime.CompilerServices.CallSiteBinder)
    IL_0305:  stsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> Cimtops.Excel.Sheet/'<>o__26'::'<>p__2'
    IL_030a:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> Cimtops.Excel.Sheet/'<>o__26'::'<>p__2'
    IL_030f:  ldfld      !0 class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>>::Target
    IL_0314:  ldsfld     class [System.Core]System.Runtime.CompilerServices.CallSite`1<class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>> Cimtops.Excel.Sheet/'<>o__26'::'<>p__2'
    IL_0319:  ldarg.2
    IL_031a:  callvirt   instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Columns()
    IL_031f:  ldloc.s    V_17
    IL_0321:  box        [mscorlib]System.Int32
    IL_0326:  ldsfld     object [mscorlib]System.Type::Missing
    IL_032b:  callvirt   instance object Microsoft.Office.Interop.Excel.Range::get__Default(object,
                                                                                            object)
    IL_0330:  callvirt   instance !2 class [mscorlib]System.Func`3<class [System.Core]System.Runtime.CompilerServices.CallSite,object,object>::Invoke(!0,
                                                                                                                                                      !1)
    IL_0335:  dup
    IL_0336:  stloc.s    V_11
    IL_0338:  isinst     [mscorlib]System.Double
    IL_033d:  brfalse.s  IL_034a

    IL_033f:  ldloc.s    V_11
    IL_0341:  unbox.any  [mscorlib]System.Double
    IL_0346:  stloc.s    V_18
    IL_0348:  br.s       IL_0355

    IL_034a:  ldc.r8     1.
    IL_0353:  stloc.s    V_18
    IL_0355:  ldloc.s    V_17
    IL_0357:  ldloc.s    V_18
    IL_0359:  newobj     instance void Cimtops.Excel.Col::.ctor(int32,
                                                                float64)
    IL_035e:  stloc.s    V_19
    IL_0360:  ldarg.0
    IL_0361:  ldfld      class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Col> Cimtops.Excel.Sheet::cols
    IL_0366:  ldloc.s    V_19
    IL_0368:  callvirt   instance void class [mscorlib]System.Collections.Generic.List`1<class Cimtops.Excel.Col>::Add(!0)
    IL_036d:  ldloc.s    V_17
    IL_036f:  ldc.i4.1
    IL_0370:  add
    IL_0371:  stloc.s    V_17
    IL_0373:  ldloc.s    V_17
    IL_0375:  ldloc.s    V_6
    IL_0377:  ble        IL_02d4

    IL_037c:  ret
  } // end of method Sheet::.ctor

  .property instance class Cimtops.Excel.Book
          Book()
  {
    .get instance class Cimtops.Excel.Book Cimtops.Excel.Sheet::get_Book()
  } // end of property Sheet::Book
  .property instance string Name()
  {
    .get instance string Cimtops.Excel.Sheet::get_Name()
  } // end of property Sheet::Name
  .property instance bool IsMultiPrint()
  {
    .get instance bool Cimtops.Excel.Sheet::get_IsMultiPrint()
  } // end of property Sheet::IsMultiPrint
  .property instance bool IsEmpty()
  {
    .get instance bool Cimtops.Excel.Sheet::get_IsEmpty()
  } // end of property Sheet::IsEmpty
  .property instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Row>
          Rows()
  {
    .get instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Row> Cimtops.Excel.Sheet::get_Rows()
  } // end of property Sheet::Rows
  .property instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Col>
          Cols()
  {
    .get instance class [mscorlib]System.Collections.Generic.IEnumerable`1<class Cimtops.Excel.Col> Cimtops.Excel.Sheet::get_Cols()
  } // end of property Sheet::Cols
  .property instance float64 Width()
  {
    .get instance float64 Cimtops.Excel.Sheet::get_Width()
  } // end of property Sheet::Width
  .property instance float64 Height()
  {
    .get instance float64 Cimtops.Excel.Sheet::get_Height()
  } // end of property Sheet::Height
} // end of class Cimtops.Excel.Sheet

.class public abstract auto ansi sealed beforefieldinit Cimtops.Excel.XLRangeUtil
       extends [mscorlib]System.Object
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.ExtensionAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public hidebysig static string 
          ToAddressName(int32 rowNum,
                        int32 colNum,
                        int32 firstNum) cil managed
  {
    // Code size       31 (0x1f)
    .maxstack  8
    IL_0000:  ldstr      "{0}{1}"
    IL_0005:  ldarg.1
    IL_0006:  ldc.i4.1
    IL_0007:  add
    IL_0008:  ldarg.2
    IL_0009:  sub
    IL_000a:  call       string Cimtops.Excel.XLRangeUtil::Num2Alpha(int32)
    IL_000f:  ldarg.0
    IL_0010:  ldc.i4.1
    IL_0011:  add
    IL_0012:  ldarg.2
    IL_0013:  sub
    IL_0014:  box        [mscorlib]System.Int32
    IL_0019:  call       string [mscorlib]System.String::Format(string,
                                                                object,
                                                                object)
    IL_001e:  ret
  } // end of method XLRangeUtil::ToAddressName

  .method private hidebysig static string 
          Num2Alpha(int32 number) cil managed
  {
    // Code size       49 (0x31)
    .maxstack  4
    .locals init (int32 V_0)
    IL_0000:  ldc.i4.s   26
    IL_0002:  ldarg.0
    IL_0003:  bge.s      IL_0023

    IL_0005:  ldarg.0
    IL_0006:  ldc.i4.1
    IL_0007:  sub
    IL_0008:  ldc.i4.s   26
    IL_000a:  div
    IL_000b:  stloc.0
    IL_000c:  ldloc.0
    IL_000d:  call       string Cimtops.Excel.XLRangeUtil::Num2Alpha(int32)
    IL_0012:  ldarg.0
    IL_0013:  ldloc.0
    IL_0014:  ldc.i4.s   26
    IL_0016:  mul
    IL_0017:  sub
    IL_0018:  call       string Cimtops.Excel.XLRangeUtil::Num2Alpha(int32)
    IL_001d:  call       string [mscorlib]System.String::Concat(string,
                                                                string)
    IL_0022:  ret

    IL_0023:  ldc.i4.s   65
    IL_0025:  ldarg.0
    IL_0026:  add
    IL_0027:  ldc.i4.1
    IL_0028:  sub
    IL_0029:  conv.u2
    IL_002a:  ldc.i4.1
    IL_002b:  newobj     instance void [mscorlib]System.String::.ctor(char,
                                                                      int32)
    IL_0030:  ret
  } // end of method XLRangeUtil::Num2Alpha

  .method public hidebysig static string 
          AdjustRange(string area) cil managed
  {
    .custom instance void [mscorlib]System.Runtime.CompilerServices.ExtensionAttribute::.ctor() = ( 01 00 00 00 ) 
    // Code size       149 (0x95)
    .maxstack  5
    .locals init (int32 V_0,
             int32 V_1,
             int32 V_2,
             string V_3,
             int32 V_4,
             int32 V_5)
    IL_0000:  ldarg.0
    IL_0001:  ldc.i4.s   58
    IL_0003:  callvirt   instance int32 [mscorlib]System.String::IndexOf(char)
    IL_0008:  stloc.0
    IL_0009:  ldc.i4.0
    IL_000a:  ldloc.0
    IL_000b:  bgt.s      IL_0033

    IL_000d:  ldarg.0
    IL_000e:  ldc.i4.0
    IL_000f:  ldloc.0
    IL_0010:  callvirt   instance string [mscorlib]System.String::Substring(int32,
                                                                            int32)
    IL_0015:  call       string Cimtops.Excel.XLRangeUtil::AdjustRange(string)
    IL_001a:  ldstr      ":"
    IL_001f:  ldarg.0
    IL_0020:  ldloc.0
    IL_0021:  ldc.i4.1
    IL_0022:  add
    IL_0023:  callvirt   instance string [mscorlib]System.String::Substring(int32)
    IL_0028:  call       string Cimtops.Excel.XLRangeUtil::AdjustRange(string)
    IL_002d:  call       string [mscorlib]System.String::Concat(string,
                                                                string,
                                                                string)
    IL_0032:  ret

    IL_0033:  ldarg.0
    IL_0034:  ldc.i4.s   82
    IL_0036:  callvirt   instance int32 [mscorlib]System.String::IndexOf(char)
    IL_003b:  stloc.1
    IL_003c:  ldarg.0
    IL_003d:  ldc.i4.s   67
    IL_003f:  callvirt   instance int32 [mscorlib]System.String::IndexOf(char)
    IL_0044:  stloc.2
    IL_0045:  ldc.i4.0
    IL_0046:  ldloc.1
    IL_0047:  bgt.s      IL_0093

    IL_0049:  ldloc.1
    IL_004a:  ldc.i4.1
    IL_004b:  add
    IL_004c:  ldloc.2
    IL_004d:  bge.s      IL_0093

    IL_004f:  ldarg.0
    IL_0050:  ldloc.1
    IL_0051:  ldc.i4.1
    IL_0052:  add
    IL_0053:  ldloc.2
    IL_0054:  ldloc.1
    IL_0055:  sub
    IL_0056:  ldc.i4.1
    IL_0057:  sub
    IL_0058:  callvirt   instance string [mscorlib]System.String::Substring(int32,
                                                                            int32)
    IL_005d:  ldarg.0
    IL_005e:  ldloc.2
    IL_005f:  ldc.i4.1
    IL_0060:  add
    IL_0061:  callvirt   instance string [mscorlib]System.String::Substring(int32)
    IL_0066:  stloc.3
    IL_0067:  ldloca.s   V_4
    IL_0069:  call       bool [mscorlib]System.Int32::TryParse(string,
                                                               int32&)
    IL_006e:  brfalse.s  IL_0093

    IL_0070:  ldloc.3
    IL_0071:  ldloca.s   V_5
    IL_0073:  call       bool [mscorlib]System.Int32::TryParse(string,
                                                               int32&)
    IL_0078:  brfalse.s  IL_0093

    IL_007a:  ldstr      "${0}${1}"
    IL_007f:  ldloc.s    V_5
    IL_0081:  call       string Cimtops.Excel.XLRangeUtil::Num2Alpha(int32)
    IL_0086:  ldloc.s    V_4
    IL_0088:  box        [mscorlib]System.Int32
    IL_008d:  call       string [mscorlib]System.String::Format(string,
                                                                object,
                                                                object)
    IL_0092:  ret

    IL_0093:  ldarg.0
    IL_0094:  ret
  } // end of method XLRangeUtil::AdjustRange

  .method public hidebysig static class [mscorlib]System.Tuple`2<int32,int32> 
          ToRowCol(string area) cil managed
  {
    // Code size       194 (0xc2)
    .maxstack  4
    .locals init (int32 V_0,
             int32 V_1,
             int32 V_2,
             int32 V_3,
             string V_4,
             int32 V_5,
             int32 V_6,
             int32 V_7,
             char V_8)
    IL_0000:  ldarg.0
    IL_0001:  callvirt   instance string [mscorlib]System.String::ToUpper()
    IL_0006:  starg.s    area
    IL_0008:  ldarg.0
    IL_0009:  ldc.i4.s   82
    IL_000b:  callvirt   instance int32 [mscorlib]System.String::IndexOf(char)
    IL_0010:  stloc.0
    IL_0011:  ldarg.0
    IL_0012:  ldc.i4.s   67
    IL_0014:  callvirt   instance int32 [mscorlib]System.String::IndexOf(char)
    IL_0019:  stloc.1
    IL_001a:  ldc.i4.0
    IL_001b:  ldloc.0
    IL_001c:  bgt.s      IL_005b

    IL_001e:  ldloc.0
    IL_001f:  ldc.i4.1
    IL_0020:  add
    IL_0021:  ldloc.1
    IL_0022:  bge.s      IL_005b

    IL_0024:  ldarg.0
    IL_0025:  ldloc.0
    IL_0026:  ldc.i4.1
    IL_0027:  add
    IL_0028:  ldloc.1
    IL_0029:  ldloc.0
    IL_002a:  sub
    IL_002b:  ldc.i4.1
    IL_002c:  sub
    IL_002d:  callvirt   instance string [mscorlib]System.String::Substring(int32,
                                                                            int32)
    IL_0032:  ldarg.0
    IL_0033:  ldloc.1
    IL_0034:  ldc.i4.1
    IL_0035:  add
    IL_0036:  callvirt   instance string [mscorlib]System.String::Substring(int32)
    IL_003b:  stloc.s    V_4
    IL_003d:  ldloca.s   V_5
    IL_003f:  call       bool [mscorlib]System.Int32::TryParse(string,
                                                               int32&)
    IL_0044:  brfalse.s  IL_005b

    IL_0046:  ldloc.s    V_4
    IL_0048:  ldloca.s   V_6
    IL_004a:  call       bool [mscorlib]System.Int32::TryParse(string,
                                                               int32&)
    IL_004f:  brfalse.s  IL_005b

    IL_0051:  ldloc.s    V_5
    IL_0053:  ldloc.s    V_6
    IL_0055:  call       class [mscorlib]System.Tuple`2<!!0,!!1> [mscorlib]System.Tuple::Create<int32,int32>(!!0,
                                                                                                             !!1)
    IL_005a:  ret

    IL_005b:  ldc.i4.0
    IL_005c:  stloc.2
    IL_005d:  ldc.i4.0
    IL_005e:  stloc.3
    IL_005f:  ldc.i4.0
    IL_0060:  stloc.s    V_7
    IL_0062:  br.s       IL_00b0

    IL_0064:  ldarg.0
    IL_0065:  ldloc.s    V_7
    IL_0067:  callvirt   instance char [mscorlib]System.String::get_Chars(int32)
    IL_006c:  stloc.s    V_8
    IL_006e:  ldc.i4.s   65
    IL_0070:  ldloc.s    V_8
    IL_0072:  bgt.s      IL_008b

    IL_0074:  ldloc.s    V_8
    IL_0076:  ldc.i4.s   90
    IL_0078:  bgt.s      IL_008b

    IL_007a:  ldloc.3
    IL_007b:  ldc.i4.s   26
    IL_007d:  mul
    IL_007e:  stloc.3
    IL_007f:  ldloc.3
    IL_0080:  ldloc.s    V_8
    IL_0082:  ldc.i4.1
    IL_0083:  add
    IL_0084:  ldc.i4.s   65
    IL_0086:  sub
    IL_0087:  add
    IL_0088:  stloc.3
    IL_0089:  br.s       IL_00aa

    IL_008b:  ldc.i4.s   48
    IL_008d:  ldloc.s    V_8
    IL_008f:  bgt.s      IL_00a6

    IL_0091:  ldloc.s    V_8
    IL_0093:  ldc.i4.s   57
    IL_0095:  bgt.s      IL_00a6

    IL_0097:  ldloc.2
    IL_0098:  ldc.i4.s   10
    IL_009a:  mul
    IL_009b:  stloc.2
    IL_009c:  ldloc.2
    IL_009d:  ldloc.s    V_8
    IL_009f:  ldc.i4.s   48
    IL_00a1:  sub
    IL_00a2:  add
    IL_00a3:  stloc.2
    IL_00a4:  br.s       IL_00aa

    IL_00a6:  ldc.i4.0
    IL_00a7:  ldloc.2
    IL_00a8:  blt.s      IL_00ba

    IL_00aa:  ldloc.s    V_7
    IL_00ac:  ldc.i4.1
    IL_00ad:  add
    IL_00ae:  stloc.s    V_7
    IL_00b0:  ldloc.s    V_7
    IL_00b2:  ldarg.0
    IL_00b3:  callvirt   instance int32 [mscorlib]System.String::get_Length()
    IL_00b8:  blt.s      IL_0064

    IL_00ba:  ldloc.2
    IL_00bb:  ldloc.3
    IL_00bc:  call       class [mscorlib]System.Tuple`2<!!0,!!1> [mscorlib]System.Tuple::Create<int32,int32>(!!0,
                                                                                                             !!1)
    IL_00c1:  ret
  } // end of method XLRangeUtil::ToRowCol

} // end of class Cimtops.Excel.XLRangeUtil

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Border
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 35 34 2D 30 30 30 30   // ..$00020854-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_7() runtime managed
  {
  } // end of method Border::_VtblGap1_7

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( struct) 
          get_LineStyle() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 77 00 00 00 00 00 )                         // ..w.....
  } // end of method Border::get_LineStyle

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set_LineStyle([in] object  marshal( struct) A_1) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 77 00 00 00 00 00 )                         // ..w.....
  } // end of method Border::set_LineStyle

  .property object LineStyle()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 77 00 00 00 00 00 )                         // ..w.....
    .get instance object Microsoft.Office.Interop.Excel.Border::get_LineStyle()
    .set instance void Microsoft.Office.Interop.Excel.Border::set_LineStyle(object)
  } // end of property Border::LineStyle
} // end of class Microsoft.Office.Interop.Excel.Border

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Borders
       implements [mscorlib]System.Collections.IEnumerable
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Reflection.DefaultMemberAttribute::.ctor(string) = ( 01 00 08 5F 44 65 66 61 75 6C 74 00 00 )          // ..._Default..
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 35 35 2D 30 30 30 30   // ..$00020855-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_8() runtime managed
  {
  } // end of method Borders::_VtblGap1_8

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Border 
          marshal( interface ) 
          get_Item([in] valuetype Microsoft.Office.Interop.Excel.XlBordersIndex Index) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 AA 00 00 00 00 00 ) 
  } // end of method Borders::get_Item

  .property class Microsoft.Office.Interop.Excel.Border
          Item(valuetype Microsoft.Office.Interop.Excel.XlBordersIndex)
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 AA 00 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Border Microsoft.Office.Interop.Excel.Borders::get_Item(valuetype Microsoft.Office.Interop.Excel.XlBordersIndex)
  } // end of property Borders::Item
} // end of class Microsoft.Office.Interop.Excel.Borders

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Comment
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 34 34 32 37 2D 30 30 30 30   // ..$00024427-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_7() runtime managed
  {
  } // end of method Comment::_VtblGap1_7

  .method public hidebysig newslot abstract virtual 
          instance string 
          marshal( bstr) 
          Text([in][opt] object  marshal( struct) Text,
               [in][opt] object  marshal( struct) Start,
               [in][opt] object  marshal( struct) Overwrite) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 8A 00 00 00 00 00 ) 
  } // end of method Comment::Text

} // end of class Microsoft.Office.Interop.Excel.Comment

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

  .method public specialname rtspecialname 
          instance void  _VtblGap2_40() runtime managed
  {
  } // end of method PageSetup::_VtblGap2_40

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Pages 
          marshal( interface ) 
          get_Pages() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 2C 0A 00 00 00 00 )                         // ..,.....
  } // end of method PageSetup::get_Pages

  .property string PrintArea()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 FB 03 00 00 00 00 ) 
    .get instance string Microsoft.Office.Interop.Excel.PageSetup::get_PrintArea()
    .set instance void Microsoft.Office.Interop.Excel.PageSetup::set_PrintArea(string)
  } // end of property PageSetup::PrintArea
  .property class Microsoft.Office.Interop.Excel.Pages
          Pages()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 2C 0A 00 00 00 00 )                         // ..,.....
    .get instance class Microsoft.Office.Interop.Excel.Pages Microsoft.Office.Interop.Excel.PageSetup::get_Pages()
  } // end of property PageSetup::Pages
} // end of class Microsoft.Office.Interop.Excel.PageSetup

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel.Pages
       implements [mscorlib]System.Collections.IEnumerable
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 34 34 41 33 2D 30 30 30 30   // ..$000244A3-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Reflection.DefaultMemberAttribute::.ctor(string) = ( 01 00 08 5F 44 65 66 61 75 6C 74 00 00 )          // ..._Default..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16) = ( 01 00 02 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_2() runtime managed
  {
  } // end of method Pages::_VtblGap1_2

  .method public hidebysig newslot specialname abstract virtual 
          instance int32  get_Count() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 76 00 00 00 00 00 )                         // ..v.....
  } // end of method Pages::get_Count

  .property int32 Count()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 76 00 00 00 00 00 )                         // ..v.....
    .get instance int32 Microsoft.Office.Interop.Excel.Pages::get_Count()
  } // end of property Pages::Count
} // end of class Microsoft.Office.Interop.Excel.Pages

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
          instance void  _VtblGap1_19() runtime managed
  {
  } // end of method Range::_VtblGap1_19

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Borders 
          marshal( interface ) 
          get_Borders() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 B3 01 00 00 00 00 ) 
  } // end of method Range::get_Borders

  .method public specialname rtspecialname 
          instance void  _VtblGap2_9() runtime managed
  {
  } // end of method Range::_VtblGap2_9

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
          instance void  _VtblGap6_35() runtime managed
  {
  } // end of method Range::_VtblGap6_35

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( struct) 
          get_HasFormula() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 0B 01 00 00 00 00 ) 
  } // end of method Range::get_HasFormula

  .method public specialname rtspecialname 
          instance void  _VtblGap7_9() runtime managed
  {
  } // end of method Range::_VtblGap7_9

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Interior 
          marshal( interface ) 
          get_Interior() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 81 00 00 00 00 00 ) 
  } // end of method Range::get_Interior

  .method public specialname rtspecialname 
          instance void  _VtblGap8_11() runtime managed
  {
  } // end of method Range::_VtblGap8_11

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
          instance void  _VtblGap9_32() runtime managed
  {
  } // end of method Range::_VtblGap9_32

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
          instance void  _VtblGap11_31() runtime managed
  {
  } // end of method Range::_VtblGap11_31

  .method public hidebysig newslot specialname abstract virtual 
          instance object 
          marshal( struct) 
          get_Value2() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 6C 05 00 00 00 00 )                         // ..l.....
  } // end of method Range::get_Value2

  .method public hidebysig newslot specialname abstract virtual 
          instance void  set_Value2([in] object  marshal( struct) A_1) runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 6C 05 00 00 00 00 )                         // ..l.....
  } // end of method Range::set_Value2

  .method public specialname rtspecialname 
          instance void  _VtblGap12_7() runtime managed
  {
  } // end of method Range::_VtblGap12_7

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Comment 
          marshal( interface ) 
          get_Comment() runtime managed preservesig internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 8E 03 00 00 00 00 ) 
  } // end of method Range::get_Comment

  .property class Microsoft.Office.Interop.Excel.Borders
          Borders()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 B3 01 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Borders Microsoft.Office.Interop.Excel.Range::get_Borders()
  } // end of property Range::Borders
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
  .property object HasFormula()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 0B 01 00 00 00 00 ) 
    .get instance object Microsoft.Office.Interop.Excel.Range::get_HasFormula()
  } // end of property Range::HasFormula
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
  .property object Value2()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 6C 05 00 00 00 00 )                         // ..l.....
    .get instance object Microsoft.Office.Interop.Excel.Range::get_Value2()
    .set instance void Microsoft.Office.Interop.Excel.Range::set_Value2(object)
  } // end of property Range::Value2
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
          instance void  _VtblGap1_10() runtime managed
  {
  } // end of method Sheets::_VtblGap1_10

  .method public hidebysig newslot abstract virtual 
          instance class [mscorlib]System.Collections.IEnumerator 
          marshal( custom ("System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a","")) 
          GetEnumerator() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 FC FF FF FF 00 00 ) 
  } // end of method Sheets::GetEnumerator

} // end of class Microsoft.Office.Interop.Excel.Sheets

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

.class public auto ansi sealed Microsoft.Office.Interop.Excel.XlBordersIndex
       extends [mscorlib]System.Enum
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 30 30 30 32 30 38 31 33 2D 30 30 30 30   // ..$00020813-0000
                                                                                                            2D 30 30 30 30 2D 63 30 30 30 2D 30 30 30 30 30   // -0000-c000-00000
                                                                                                            30 30 30 30 30 34 36 2D 4D 69 63 72 6F 73 6F 66   // 0000046-Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 49 6E 74 65 72 6F 70   // t.Office.Interop
                                                                                                            2E 45 78 63 65 6C 2E 58 6C 42 6F 72 64 65 72 73   // .Excel.XlBorders
                                                                                                            49 6E 64 65 78 00 00 )                            // Index..
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlBordersIndex xlInsideHorizontal = int32(0x0000000C)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlBordersIndex xlInsideVertical = int32(0x0000000B)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlBordersIndex xlDiagonalDown = int32(0x00000005)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlBordersIndex xlDiagonalUp = int32(0x00000006)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlBordersIndex xlEdgeBottom = int32(0x00000009)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlBordersIndex xlEdgeLeft = int32(0x00000007)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlBordersIndex xlEdgeRight = int32(0x0000000A)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlBordersIndex xlEdgeTop = int32(0x00000008)
} // end of class Microsoft.Office.Interop.Excel.XlBordersIndex

.class public auto ansi sealed Microsoft.Office.Interop.Excel.XlLineStyle
       extends [mscorlib]System.Enum
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor(string,
                                                                                                string) = ( 01 00 24 30 30 30 32 30 38 31 33 2D 30 30 30 30   // ..$00020813-0000
                                                                                                            2D 30 30 30 30 2D 63 30 30 30 2D 30 30 30 30 30   // -0000-c000-00000
                                                                                                            30 30 30 30 30 34 36 2A 4D 69 63 72 6F 73 6F 66   // 0000046*Microsof
                                                                                                            74 2E 4F 66 66 69 63 65 2E 49 6E 74 65 72 6F 70   // t.Office.Interop
                                                                                                            2E 45 78 63 65 6C 2E 58 6C 4C 69 6E 65 53 74 79   // .Excel.XlLineSty
                                                                                                            6C 65 00 00 )                                     // le..
  .field public specialname rtspecialname int32 value__
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlLineStyle xlContinuous = int32(0x00000001)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlLineStyle xlDash = int32(0xFFFFEFED)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlLineStyle xlDashDot = int32(0x00000004)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlLineStyle xlDashDotDot = int32(0x00000005)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlLineStyle xlDot = int32(0xFFFFEFEA)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlLineStyle xlDouble = int32(0xFFFFEFE9)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlLineStyle xlSlantDashDot = int32(0x0000000D)
  .field public static literal valuetype Microsoft.Office.Interop.Excel.XlLineStyle xlLineStyleNone = int32(0xFFFFEFD2)
} // end of class Microsoft.Office.Interop.Excel.XlLineStyle

.class interface public abstract auto ansi import Microsoft.Office.Interop.Excel._Workbook
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 32 30 38 44 41 2D 30 30 30 30   // ..$000208DA-0000
                                                                                                  2D 30 30 30 30 2D 43 30 30 30 2D 30 30 30 30 30   // -0000-C000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.TypeIdentifierAttribute::.ctor() = ( 01 00 00 00 ) 
  .method public specialname rtspecialname 
          instance void  _VtblGap1_59() runtime managed
  {
  } // end of method _Workbook::_VtblGap1_59

  .method public hidebysig newslot specialname abstract virtual 
          instance string 
          marshal( bstr) 
          get_Name() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 6E 00 00 00 00 00 )                         // ..n.....
  } // end of method _Workbook::get_Name

  .method public specialname rtspecialname 
          instance void  _VtblGap2_9() runtime managed
  {
  } // end of method _Workbook::_VtblGap2_9

  .method public hidebysig newslot specialname abstract virtual 
          instance string 
          marshal( bstr) 
          get_Path() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 23 01 00 00 00 00 )                         // ..#.....
    .custom instance void [mscorlib]System.Runtime.InteropServices.LCIDConversionAttribute::.ctor(int32) = ( 01 00 00 00 00 00 00 00 ) 
  } // end of method _Workbook::get_Path

  .method public specialname rtspecialname 
          instance void  _VtblGap3_35() runtime managed
  {
  } // end of method _Workbook::_VtblGap3_35

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Sheets 
          marshal( interface ) 
          get_Sheets() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 E5 01 00 00 00 00 ) 
  } // end of method _Workbook::get_Sheets

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
          instance void  _VtblGap1_11() runtime managed
  {
  } // end of method _Worksheet::_VtblGap1_11

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
          instance void  _VtblGap2_7() runtime managed
  {
  } // end of method _Worksheet::_VtblGap2_7

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.PageSetup 
          marshal( interface ) 
          get_PageSetup() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 E6 03 00 00 00 00 ) 
  } // end of method _Worksheet::get_PageSetup

  .method public specialname rtspecialname 
          instance void  _VtblGap3_24() runtime managed
  {
  } // end of method _Worksheet::_VtblGap3_24

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Range 
          marshal( interface ) 
          get_Cells() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 EE 00 00 00 00 00 ) 
  } // end of method _Worksheet::get_Cells

  .method public specialname rtspecialname 
          instance void  _VtblGap4_5() runtime managed
  {
  } // end of method _Worksheet::_VtblGap4_5

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Range 
          marshal( interface ) 
          get_Columns() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 F1 00 00 00 00 00 ) 
  } // end of method _Worksheet::get_Columns

  .method public specialname rtspecialname 
          instance void  _VtblGap5_41() runtime managed
  {
  } // end of method _Worksheet::_VtblGap5_41

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Range 
          marshal( interface ) 
          get_Range([in] object  marshal( struct) Cell1,
                    [in][opt] object  marshal( struct) Cell2) runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 C5 00 00 00 00 00 ) 
  } // end of method _Worksheet::get_Range

  .method public specialname rtspecialname 
          instance void  _VtblGap6_1() runtime managed
  {
  } // end of method _Worksheet::_VtblGap6_1

  .method public hidebysig newslot specialname abstract virtual 
          instance class Microsoft.Office.Interop.Excel.Range 
          marshal( interface ) 
          get_Rows() runtime managed internalcall
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 02 01 00 00 00 00 ) 
  } // end of method _Worksheet::get_Rows

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
  .property class Microsoft.Office.Interop.Excel.Range
          Cells()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 EE 00 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Cells()
  } // end of property _Worksheet::Cells
  .property class Microsoft.Office.Interop.Excel.Range
          Columns()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 F1 00 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Columns()
  } // end of property _Worksheet::Columns
  .property class Microsoft.Office.Interop.Excel.Range
          Range(object,
                object)
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 C5 00 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Range(object,
                                                                                                                  object)
  } // end of property _Worksheet::Range
  .property class Microsoft.Office.Interop.Excel.Range
          Rows()
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32) = ( 01 00 02 01 00 00 00 00 ) 
    .get instance class Microsoft.Office.Interop.Excel.Range Microsoft.Office.Interop.Excel._Worksheet::get_Rows()
  } // end of property _Worksheet::Rows
} // end of class Microsoft.Office.Interop.Excel._Worksheet


// =============================================================

// *********** DISASSEMBLY COMPLETE ***********************
// WARNING: Created Win32 resource file C:\Users\MCF-JOHNELLEEMPUERTO\Documents\Johnell\PaperLess Enterprise\CimtopsExcel.res
