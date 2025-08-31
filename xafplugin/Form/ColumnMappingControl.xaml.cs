using NLog;
using System;
using System.Windows;
using System.Windows.Controls;
using xafplugin.Interfaces;
using xafplugin.Modules;
using xafplugin.ViewModels;

namespace xafplugin.Form
{
    public partial class ColumnMappingControl : UserControl
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly ColumnMappingViewModel _viewModel;
        private readonly IMessageBoxService _dialog = new MessageBoxService();

        public ColumnMappingControl()
        {
            InitializeComponent();

            var environmentService = new EnvironmentService();

            try
            {
                if (string.IsNullOrEmpty(environmentService.DatabasePath))
                {
                    _logger.Warn("No database path found. Load the audit file first.");
                    _dialog.ShowInfo("No database is available. Load the audit file first.");
                    return;
                }

                DataContext = _viewModel = new ColumnMappingViewModel();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to initialize ColumnMappingControl.");
                _dialog.ShowError("An unexpected error occurred while initializing the control.");
            }
        }
    }
}
