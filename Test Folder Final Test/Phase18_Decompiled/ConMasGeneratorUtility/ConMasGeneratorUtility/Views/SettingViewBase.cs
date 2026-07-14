using System.Windows.Forms;

namespace ConMasGeneratorUtility.Views;

public class SettingViewBase : UserControl
{
	public virtual void SaveTemporary()
	{
	}

	public virtual bool IsValid()
	{
		return false;
	}
}
