using Microsoft.VisualBasic;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using xafplugin.Database;
using xafplugin.Helpers;
using xafplugin.Modules;
using xafplugin.ViewModels;
using MessageBox = System.Windows.Forms.MessageBox;

namespace xafplugin.Form
{
    public partial class TableRelationPadControl : Window
    {
        private Window _relationWindow;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly EnvironmentService _env = new EnvironmentService();
        public TableRelationPadViewModel _viewModel { get; set; }
        public ExportDefinition Result { get; private set; }

        public TableRelationPadControl()
        {
            InitializeComponent();

            var settings = new SettingsProvider();
            var environmentService = new EnvironmentService();
            var dialog = new MessageBoxService();

            if (string.IsNullOrEmpty(environmentService.DatabasePath))
            {
                _logger.Warn("Geen databasepad gevonden. Auditfile moet eerst worden ingelezen.");
                MessageBox.Show("Er is geen database beschikbaar. Zorg ervoor dat de auditfile is ingelezen.");
                return;
            }

            this.DataContext = _viewModel = new TableRelationPadViewModel();

            //TODO dit verplaatsen naar een betere locatie en juiste verwerking. waarom niet één keer de settings voorbereiden en hierna gewoon verder werken. Dit scheelt ook een boel sql commands.
            var relationsviewmodel = new RelationsViewModel();
            relationsviewmodel.SaveRelationsToSettings();
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA1806:Do not ignore method results", Justification = "WindowInteropHelper constructor is intentionally used for its side-effect of setting the Owner property.")]
        private void BtnRelatieToevoegen_Click(object sender, RoutedEventArgs e)
        {
            _logger.Info("ButtonRelation_Click gestart – Relatiebeheer venster wordt geopend.");
            if (_relationWindow != null && _relationWindow.IsVisible)
            {
                _relationWindow.Activate();
                return;
            }
            try
            {
                _relationWindow = new Window
                {
                    Title = "Relations",
                    Content = new RelationsControl(),
                    SizeToContent = SizeToContent.WidthAndHeight,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    ResizeMode = ResizeMode.NoResize,
                    ShowInTaskbar = false,
                    WindowStyle = WindowStyle.SingleBorderWindow,
                    Icon = BitmapImageCoverter.ByteArrayToIcon(Properties.Resources.empty16)
                };

                new WindowInteropHelper(_relationWindow)
                {
                    Owner = new IntPtr(Globals.ThisAddIn.Application.Hwnd)
                };

                _relationWindow.ShowDialog();

                _logger.Info("Relatiebeheer venster succesvol weergegeven.");

                _viewModel.UpdateAvailableRelations();

                _logger.Info("Beschikbare relaties geupdated");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Fout bij openen van het relatiebeheer venster.");
                System.Windows.Forms.MessageBox.Show("Er trad een fout op bij het openen van het venster.", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void BtnAddCustomColumn_Click(object sender, RoutedEventArgs e)
        {
            _logger.Info("ButtonRelation_Click gestart – Relatiebeheer venster wordt geopend.");

            try
            {
                var selectedColumns = _viewModel.SelectedColumns.Select(c => c.Column);

                var wizardViewModel = new WizardSQLCaseViewModel(selectedColumns);
                var wnd = new WizardSQLCase(wizardViewModel);

                new WindowInteropHelper(wnd)
                {
                    Owner = new IntPtr(Globals.ThisAddIn.Application.Hwnd)
                };
                var ok = wnd.ShowDialog();

                //_viewModel.UpdateAvailableRelations();

                if (ok == true)
                {
                    // hier is je resultaat
                    var caseExpr = wizardViewModel.ResultSql;
                    _viewModel.CustomColumns.Add(caseExpr);
                    _viewModel.SelectedColumns.Add(new ColumnDescriptor
                    {
                        Column = caseExpr.Key,
                        Table = "Custom",
                        IsCustom = true
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Fout bij openen van het relatiebeheer venster.");
                System.Windows.Forms.MessageBox.Show("Er trad een fout op bij het openen van het venster.", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAddRelation_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.AddRelationToPath();
        }

        private void BtnBackRelation_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.RemoveLastRelation();
        }

        private void BtnAddColumn_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.AddColumns(ListAvailableColumns.SelectedItems);
        }

        private void BtnRemoveColumn_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.RemoveColumns(ListSelectedColumns.SelectedItems);
        }
        private void BtnOpslaan_Click(object sender, RoutedEventArgs e)
        {
            _logger.Info("Gebruiker start opslaan van exportdefinitie via BtnOpslaan_Click.");

            try
            {
                var vm = this.DataContext as TableRelationPadViewModel;
                var newTable = vm.CurrentExportDefinition;

                if (newTable == null)
                {
                    _logger.Warn("Geen tabel gevonden. Toevoegen afgebroken.");
                    MessageBox.Show("Geen tabel gevonden. Toevoegen afgebroken.", "Waarschuwing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (string.IsNullOrWhiteSpace(newTable.MainTable))
                {
                    _logger.Warn($"MainTable ontbreekt in definitie '{newTable.Name}'. Toevoegen afgebroken.");
                    MessageBox.Show("De hoofdtafel (MainTable) is verplicht.", "Waarschuwing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (newTable.SelectedColumns == null || !newTable.SelectedColumns.Any(c => !string.IsNullOrWhiteSpace(c.Column)))
                {
                    _logger.Warn($"Geen geldige kolommen opgegeven in definitie '{newTable.Name}'. Toevoegen afgebroken.");
                    MessageBox.Show("Er moeten minstens één of meer kolommen worden geselecteerd.", "Waarschuwing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string naam = Interaction.InputBox("Geef een naam op voor deze exportdefinitie:", "Naam export", vm?.CurrentExportDefinition?.Name ?? "");

                if (string.IsNullOrWhiteSpace(naam))
                {
                    _logger.Warn("Naam voor exportdefinitie werd niet opgegeven. Toevoegen afgebroken.");
                    return;
                }
                var settings = new SettingsProvider().Get(_env.FileHash);
                if (settings.ExportDefinitions.Any(def => def.Name == naam))
                {
                    _logger.Warn($"Er bestaat al een exportdefinitie met de naam '{naam}'. Toevoegen afgebroken.");
                    MessageBox.Show($"Een exportdefinitie met de naam '{naam}' bestaat al.", "Waarschuwing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var databasePath = new EnvironmentService().DatabasePath;
                using (var databaseService = new DatabaseService(databasePath))
                {
                    var sqlString = SqlQueryBuilder.BuildExportDefinitionQuery(newTable);
                    if (!databaseService.IsValidAgainstDb(sqlString))
                    {
                        _logger.Warn($"De exportdefinitie '{naam}' is ongeldig tegen de huidige database. Toevoegen afgebroken.");
                        MessageBox.Show("De exportdefinitie is ongeldig tegen de huidige database. Controleer of alle tabellen en kolommen nog bestaan");
                        return;
                    }

                    if (!databaseService.QueryHasRowsAny(sqlString))
                    {
                        _logger.Warn($"De exportdefinitie '{naam}' is geldig, maar levert geen resultaten op. Toevoegen afgebroken.");
                        MessageBox.Show("De exportdefinitie is geldig, maar levert geen resultaten op. Pas de definitie aan zodat er wel resultaten worden opgehaald.");
                        return;
                    }
                }

                vm.Name = naam;
                this.Result = vm.CurrentExportDefinition;
                this.DialogResult = true;
                _logger.Info($"Exportdefinitie '{naam}' succesvol opgeslagen.");
                this.Close();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Fout tijdens het opslaan van de exportdefinitie.");
                MessageBox.Show("Er trad een fout op tijdens het opslaan van de exportdefinitie.\n\n" + ex.Message, "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnMoveUp_Click(object sender, RoutedEventArgs e)
        {
            _logger.Info("BtnMoveUp_Click gestart.");
            
            if (ListSelectedColumns.SelectedItems.Count > 0)
            {
                // Create a copy of the selected items to maintain selection after moving
                var selectedColumns = ListSelectedColumns.SelectedItems.Cast<ColumnDescriptor>().ToList();
                
                // Move the columns up
                _viewModel.MoveColumnsUp(selectedColumns);
                
                // Restore selection
                ListSelectedColumns.SelectedItems.Clear();
                foreach (var column in selectedColumns)
                {
                    ListSelectedColumns.SelectedItems.Add(column);
                }
                
                // Focus the list again
                ListSelectedColumns.Focus();
            }
            else
            {
                _logger.Warn("Geen kolommen geselecteerd om omhoog te verplaatsen.");
            }
        }

        private void BtnMoveDown_Click(object sender, RoutedEventArgs e)
        {
            _logger.Info("BtnMoveDown_Click gestart.");
            
            if (ListSelectedColumns.SelectedItems.Count > 0)
            {
                // Create a copy of the selected items to maintain selection after moving
                var selectedColumns = ListSelectedColumns.SelectedItems.Cast<ColumnDescriptor>().ToList();
                
                // Move the columns down
                _viewModel.MoveColumnsDown(selectedColumns);
                
                // Restore selection
                ListSelectedColumns.SelectedItems.Clear();
                foreach (var column in selectedColumns)
                {
                    ListSelectedColumns.SelectedItems.Add(column);
                }
                
                // Focus the list again
                ListSelectedColumns.Focus();
            }
            else
            {
                _logger.Warn("Geen kolommen geselecteerd om omlaag te verplaatsen.");
            }
        }

    }
}