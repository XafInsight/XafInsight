using NLog;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using xafplugin.Helpers;
using xafplugin.Interfaces;
using xafplugin.Modules;
using xafplugin.ViewModels;

namespace xafplugin.Form
{
    public partial class TableControl : UserControl
    {
        private readonly TableControlViewModel _viewModel;
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IMessageBoxService _dialog = new MessageBoxService();

        public TableControl()
        {
            InitializeComponent();

            var settings = new SettingsProvider();
            var environmentService = new EnvironmentService();

            if (string.IsNullOrEmpty(environmentService.DatabasePath))
            {
                logger.Warn("No database path found. Load the audit file first.");
                _dialog.ShowInfo("No database is available. Load the audit file first.");
                return;
            }

            DataContext = _viewModel = new TableControlViewModel();
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA1806:Do not ignore method results", Justification = "WindowInteropHelper used for side-effect (Owner).")]
        private void BtnToevoegen_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new TableRelationPadControl();
            new WindowInteropHelper(wnd)
            {
                Owner = new IntPtr(Globals.ThisAddIn.Application.Hwnd)
            };
            if (wnd.ShowDialog() == true)
            {
                var newTable = wnd.Result;
                if (newTable != null)
                {
                    _viewModel.AddTable(newTable);
                }
            }
        }

        private void BtnVerwijderTabel_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.Tag is string tableName)
            {
                if (DataContext is TableControlViewModel vm && vm.ExportTables.Contains(tableName))
                {
                    var result = _dialog.Show(
                        $"Are you sure you want to remove the table '{tableName}'?",
                        "Confirm",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        vm.RemoveTable(tableName);
                    }
                }
            }
        }
    }
}