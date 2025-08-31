using NLog;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using xafplugin.Helpers;
using xafplugin.Modules;
using xafplugin.ViewModels;

namespace xafplugin.Form
{
    public partial class OpenFileWindow : UserControl
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private OpenFileWindowViewModel _viewModel { get; }

        public event EventHandler<bool> CloseRequested;

        public OpenFileWindow()
        {
            InitializeComponent();

            try
            {
                var entries = HistoryFileManager.ReadAll(Globals.ThisAddIn.Config.ConfigPath, out var lines, out var error);
                DataContext = _viewModel = new OpenFileWindowViewModel(lines);

                // SUBSCRIBE: Without this, RequestClose from the VM is never observed.
                _viewModel.RequestClose += ViewModel_RequestClose;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Fout bij initialisatie OpenFileWindow");
                MessageBox.Show("Kon het venster niet initialiseren.", "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ViewModel_RequestClose(object sender, bool ok)
        {
            _logger.Info("OpenFileWindow ViewModel verzocht sluiten (ok={0})", ok);

            // Try to close hosting Window if the control is inside one.
            var hostWindow = Window.GetWindow(this);
            if (hostWindow != null)
            {
                try
                {
                    hostWindow.DialogResult = ok; // Only works if shown with ShowDialog()
                }
                catch
                {
                    // Ignore if not a dialog.
                }
                hostWindow.Close();
            }
            else
            {
                // Bubble up so a parent (e.g. custom task pane host) can remove the control.
                CloseRequested?.Invoke(this, ok);
            }
        }

        // Double-click confirms (equivalent to pressing OK)
        private void ListEntries_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (_viewModel == null)
                return;

            var listBox = sender as ListBox;
            if (listBox == null)
                return;

            // Only confirm if the double-click occurred on an actual item
            var container = listBox.ContainerFromElement((DependencyObject)e.OriginalSource) as ListBoxItem;
            if (container == null)
                return;

            if (_viewModel.SelectedEntry != null && _viewModel.ConfirmCommand.CanExecute(null))
            {
                _viewModel.ConfirmCommand.Execute(null);
            }
        }
    }

    public class SelectedEntryToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var s = value as string;
            return !string.IsNullOrWhiteSpace(s);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }
}