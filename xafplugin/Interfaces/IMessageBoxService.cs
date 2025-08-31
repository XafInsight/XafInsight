using System.Windows.Forms;

namespace xafplugin.Interfaces
{
    public interface IMessageBoxService
    {
        DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon);

        void ShowInfo(string text);
        void ShowWarning(string text);
        void ShowError(string text);
    }


}
