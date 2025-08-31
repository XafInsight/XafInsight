using System.Windows.Forms;
using xafplugin.Interfaces;

namespace xafplugin.Modules
{
    public class MessageBoxService : IMessageBoxService
    {
        public DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
            => MessageBox.Show(text, caption, buttons, icon);

        public void ShowInfo(string text)
            => Show(text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

        public void ShowWarning(string text)
            => Show(text, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        public void ShowError(string text)
            => Show(text, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
