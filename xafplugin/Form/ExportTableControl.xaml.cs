using Microsoft.Office.Interop.Excel;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Threading;
using xafplugin.Database;
using xafplugin.Helpers;
using xafplugin.Interfaces;
using xafplugin.Modules;
using xafplugin.ViewModels;
using UserControl = System.Windows.Controls.UserControl;

namespace xafplugin.Form
{
    public partial class ExportTableControl : UserControl
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly ExportTableViewModel _viewModel;
        private readonly DispatcherTimer _editCheckTimer;
        private readonly IMessageBoxService _dialog = new MessageBoxService();
        private readonly IEnvironmentService _env = new EnvironmentService();
        private System.Windows.Window _tableWindow;

        public ExportTableControl()
        {
            InitializeComponent();

            Unloaded += ExportTableControl_Unloaded;
            Loaded += ExportTableControl_Loaded;

            _env = new EnvironmentService();

            if (string.IsNullOrEmpty(_env.DatabasePath))
            {
                logger.Warn("No database path found. Load the audit file first.");
                _dialog.ShowInfo("No database is available. Load the audit file first.");
                return;
            }

            DataContext = _viewModel = new ExportTableViewModel();

            _editCheckTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            _editCheckTimer.Tick += (s, e) =>
            {
                try
                {
                    var bars = Globals.ThisAddIn.Application.CommandBars as Microsoft.Office.Core.CommandBars;
                    if (bars.GetEnabledMso("FileNewDefault"))
                    {
                        ExportButton.IsEnabled = true;
                        ExportLocationButton.IsEnabled = true;
                        ExportStatusText.Visibility = Visibility.Collapsed;
                        return;
                    }
                    if (string.IsNullOrEmpty(_viewModel.ExportLocation))
                    {
                        ExportButton.IsEnabled = false;
                        ExportLocationButton.IsEnabled = false;
                        ExportStatusText.Visibility = Visibility.Visible;
                    }
                }
                catch
                {
                    ExportButton.IsEnabled = true;
                    ExportLocationButton.IsEnabled = true;
                    ExportStatusText.Visibility = Visibility.Collapsed;
                }
            };
        }

        private void ButtonFilterToevoegen_click(object sender, RoutedEventArgs e)
        {
            logger.Info("Add Filter menu opening.");

            if (_viewModel == null)
            {
                _dialog.ShowWarning("_viewModel is null. Cannot add filter.");
                logger.Error("_viewModel null in ButtonFilterToevoegen_click");
                return;
            }

            if (_viewModel.SelectedTable == null)
            {
                _dialog.ShowWarning("No table selected. Select a table first.");
                logger.Warn("No table selected in ButtonFilterToevoegen_click");
                return;
            }

            var menu = new System.Windows.Controls.ContextMenu();
            menu.Items.Add(new System.Windows.Controls.MenuItem { Header = "Simple" });
            menu.Items.Add(new System.Windows.Controls.MenuItem { Header = "Advanced" });

            // Simple
            ((System.Windows.Controls.MenuItem)menu.Items[0]).Click += (s, args) =>
            {
                try
                {
                    var columns = _viewModel?.SelectedTable?.SelectedColumns?
                        .Where(c => c != null)
                        .Select(c => c.Column)
                        .ToList() ?? new List<string>();

                    var vm = new WizardFilterViewModel(columns);
                    var wnd = new WizardFilter(vm);
                    new WindowInteropHelper(wnd) { Owner = new IntPtr(Globals.ThisAddIn.Application.Hwnd) };
                    var ok = wnd.ShowDialog();
                    if (ok == true && _viewModel != null)
                    {
                        var filterItem = new FilterItem
                        {
                            Name = vm.FilterName,
                            Expression = vm.SqlText
                        };
                        _viewModel.Filters.Add(filterItem);

                        var sqlDbPath = _env.DatabasePath;
                        using (var databaseService = new DatabaseService(sqlDbPath))
                        {
                            var sqlCTEQuery = ExportDefinitionHelper.SqlQueryWithCte(_viewModel.SelectedTable, _viewModel.Filters.ToList());
                            if (!databaseService.IsValidAgainstDb(sqlCTEQuery) || !databaseService.QueryHasRowsAny(sqlCTEQuery))
                            {
                                _viewModel.Filters.Remove(filterItem);
                                _dialog.ShowError("No results from filter or validation failed against the database.");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Error while opening simple filter wizard.");
                    _dialog.ShowError("An error occurred while opening the filter window.");
                }
            };

            // Advanced
            ((System.Windows.Controls.MenuItem)menu.Items[1]).Click += (s, args) =>
            {
                try
                {
                    var columns = _viewModel?.SelectedTable?.SelectedColumns?
                        .Where(c => c != null)
                        .Select(c => c.Column)
                        .ToList() ?? new List<string>();

                    var vm = new WizardFilterSQLViewModel(columns);
                    var wnd = new WizardFilterSQL(vm);
                    new WindowInteropHelper(wnd) { Owner = new IntPtr(Globals.ThisAddIn.Application.Hwnd) };
                    var ok = wnd.ShowDialog();
                    if (ok == true && _viewModel != null)
                    {
                        var filterItem = new FilterItem
                        {
                            Name = vm.Name,
                            Expression = vm.ResultSql
                        };
                        _viewModel.Filters.Add(filterItem);

                        var sqlDbPath = _env.DatabasePath;
                        using (var databaseService = new DatabaseService(sqlDbPath))
                        {
                            var sqlCTEQuery = ExportDefinitionHelper.SqlQueryWithCte(_viewModel.SelectedTable, _viewModel.Filters.ToList());
                            if (!databaseService.IsValidAgainstDb(sqlCTEQuery) || !databaseService.QueryHasRowsAny(sqlCTEQuery))
                            {
                                _viewModel.Filters.Remove(filterItem);
                                _dialog.ShowError("Filter validation failed against the database.");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Error while opening advanced filter wizard.");
                    _dialog.ShowError("An error occurred while opening the filter window.");
                }
            };

            menu.PlacementTarget = sender as UIElement;
            menu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
            menu.IsOpen = true;
        }

        private void BtnSelectRange_Click(object sender, RoutedEventArgs e)
        {
            var parentWindow = System.Windows.Window.GetWindow(this);

            try
            {
                var xl = Globals.ThisAddIn.Application;
                object result = xl.InputBox("Select a target cell", "Select Range", Type: 8);
                if (result is bool b && b == false) return;

                if (result is Range rng)
                {
                    if (_viewModel.setTargetCell(rng))
                    {
                        logger.Info($"Target cell set: {_viewModel.ExportLocation}");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error selecting range.");
                _dialog.ShowError("An error occurred while selecting the range.");
            }
            finally
            {
                parentWindow?.Show();
            }
        }

        private void ExportTableControl_Loaded(object sender, RoutedEventArgs e) => _editCheckTimer.Start();

        private void ExportTableControl_Unloaded(object sender, RoutedEventArgs e)
        {
            _editCheckTimer.Stop();
        }

        private void Export_Click(object sender, RoutedEventArgs e) => _viewModel.ExportSelectedTable();

        private void ManageTablesButton_Click(object sender, RoutedEventArgs e)
        {
            logger.Info("Opening table manager.");
            if (_tableWindow != null && _tableWindow.IsVisible)
            {
                _tableWindow.Activate();
                return;
            }

            try
            {
                _tableWindow = new System.Windows.Window
                {
                    Title = "Select Table/Columns",
                    Content = new TableControl(),
                    SizeToContent = SizeToContent.WidthAndHeight,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    ResizeMode = ResizeMode.NoResize,
                    ShowInTaskbar = false,
                    WindowStyle = WindowStyle.SingleBorderWindow,
                    Icon = BitmapImageCoverter.ByteArrayToIcon(Properties.Resources.empty16)
                };
                new WindowInteropHelper(_tableWindow)
                {
                    Owner = new IntPtr(Globals.ThisAddIn.Application.Hwnd)
                };
                _tableWindow.ShowDialog();

                _viewModel.LoadAvailableTables();
                logger.Info("Table manager closed and tables reloaded.");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error opening table manager window.");
                _dialog.ShowError("An error occurred while opening the window.");
            }
        }

        private void BtnVerwijderFilter_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.Tag is FilterItem filterItem)
            {
                if (DataContext is ExportTableViewModel vm)
                {
                    var result = _dialog.Show(
                        $"Are you sure you want to remove filter '{filterItem.Name}'?",
                        "Confirm",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (result == DialogResult.Yes)
                    {
                        vm.Filters.Remove(filterItem);
                    }
                }
            }
        }
    }
}

