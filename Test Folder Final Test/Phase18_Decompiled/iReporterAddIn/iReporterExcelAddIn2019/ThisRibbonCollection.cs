using System.CodeDom.Compiler;
using System.Diagnostics;
using Microsoft.Office.Tools.Ribbon;

namespace iReporterExcelAddIn2019;

[DebuggerNonUserCode]
[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
internal sealed class ThisRibbonCollection : RibbonCollectionBase
{
	internal Ribbon1 Ribbon1 => ((RibbonCollectionBase)this).GetRibbon<Ribbon1>();

	internal ThisRibbonCollection(RibbonFactory factory)
		: base(factory)
	{
	}
}
