using NLog;
using System.ComponentModel;
using xafplugin.Helpers;
using xafplugin.Interfaces;
using xafplugin.Modules;

namespace xafplugin.ViewModels
{
    public abstract class ViewModelBase : INotifyPropertyChanged
    {
        protected readonly Logger _logger = LogManager.GetCurrentClassLogger();
        protected readonly ISettingsProvider _settings;
        protected readonly IEnvironmentService _env;
        protected readonly IMessageBoxService _dialog;


        protected ViewModelBase(IEnvironmentService environment, ISettingsProvider settingsProvider, IMessageBoxService dialog)
        {
            _settings = settingsProvider;
            _env = environment;
            _dialog = dialog;
        }

        // Default constructor for XAML designer or simple instantiation
        protected ViewModelBase()
        {
            _settings = new SettingsProvider();
            _env = new EnvironmentService();
            _dialog = new MessageBoxService();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string property)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
        }
    }
}
