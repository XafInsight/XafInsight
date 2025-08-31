using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Input;
using xafplugin.Database;
using xafplugin.Helpers;
using xafplugin.Interfaces;
using xafplugin.Modules;

namespace xafplugin.ViewModels
{
    /// <summary>
    /// Manages renaming/mapping of columns per table with validation persisted to settings.
    /// </summary>
    public class ColumnMappingViewModel : ViewModelBase
    {
        public ObservableCollection<string> Tables { get; private set; }
        public ObservableCollection<ColumnRename> ColumnRenames { get; private set; }

        private readonly Dictionary<string, List<string>> _tableColumns = new Dictionary<string, List<string>>();

        private string _selectedTable;
        public string SelectedTable
        {
            get => _selectedTable;
            set
            {
                if (_selectedTable != value)
                {
                    _selectedTable = value;
                    OnPropertyChanged(nameof(SelectedTable));
                    LoadColumnsForTable(_selectedTable);
                }
            }
        }

        private string _errorMessage;
        public string ErrorMessage
        {
            get => _errorMessage;
            set
            {
                if (_errorMessage != value)
                {
                    _errorMessage = value;
                    OnPropertyChanged(nameof(ErrorMessage));
                }
            }
        }

        public ICommand ValidateCommand { get; private set; }

        public ColumnMappingViewModel(IEnvironmentService environment, ISettingsProvider settingsProvider, IMessageBoxService dialog)
            : base(environment, settingsProvider, dialog)
        {
            _logger.Info("Initializing ColumnMappingViewModel.");
            Initialize();
        }

        public ColumnMappingViewModel() : base()
        {
            _logger.Info("Initializing ColumnMappingViewModel.");
            Initialize();
        }

        private void Initialize()
        {
            Tables = new ObservableCollection<string>();
            ColumnRenames = new ObservableCollection<ColumnRename>();

            try
            {
                var databasePath = _env.DatabasePath;
                if (string.IsNullOrEmpty(databasePath))
                {
                    _logger.Warn("No database path found. Audit file must be loaded first.");
                    _dialog.ShowInfo("No database is available. Load the audit file first.");
                    return;
                }

                using (var db = new DatabaseService(_env.DatabasePath))
                {
                    var tables = db.GetTables();
                    if (tables != null && tables.Count > 0)
                    {
                        _logger.Debug($"Tables retrieved: {tables.Count}");
                        foreach (var t in tables)
                            Tables.Add(t);
                    }

                    var columnsByTable = db.GetTableColumns(Tables);
                    if (columnsByTable != null && columnsByTable.Count > 0)
                    {
                        foreach (var kv in columnsByTable)
                            _tableColumns[kv.Key] = kv.Value;

                        _logger.Debug("Columns per table loaded.");
                    }
                }

                if (Tables.Count > 0)
                {
                    SelectedTable = Tables[0];
                    _logger.Info($"Initial selected table: {SelectedTable}");
                }

                ValidateCommand = new RelayCommand(_ => ValidateAndSave(), _ => true);

                _logger.Info("ColumnMappingViewModel initialization completed.");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during initialization of ColumnMappingViewModel.");
                _dialog.Show("An error occurred while loading column information.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadColumnsForTable(string tableName)
        {
            _logger.Info($"Loading columns for table: {tableName}");
            ColumnRenames.Clear();

            try
            {
                var validRegex = new Regex(@"^[A-Za-z_][A-Za-z0-9_]*$", RegexOptions.None, TimeSpan.FromMilliseconds(100));
                var mapping = _settings.Get(_env.FileHash).ColumnMappings
                    .FirstOrDefault(m => m.TableName == tableName);

                var colMappings = mapping != null
                    ? mapping.Columns
                    : new Dictionary<string, string>();

                if (_tableColumns.TryGetValue(tableName, out var columns))
                {
                    foreach (var col in columns)
                    {
                        string mappedName;
                        if (colMappings.TryGetValue(col, out var mapped) && !string.IsNullOrEmpty(mapped))
                            mappedName = mapped;
                        else if (validRegex.IsMatch(col))
                            mappedName = col;
                        else
                            mappedName = string.Empty;

                        var colRename = new ColumnRename
                        {
                            Original = col,
                            NewName = mappedName
                        };
                        colRename.PropertyChanged += ColumnRename_PropertyChanged;
                        ColumnRenames.Add(colRename);
                    }
                }
                else
                {
                    _logger.Warn($"Table '{tableName}' not found.");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"Error loading columns for table '{tableName}'.");
            }
        }

        private void ColumnRename_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ColumnRename.NewName))
                ValidateAndSave();
        }

        private void ValidateAndSave()
        {
            _logger.Info("Validating column names.");

            foreach (var col in ColumnRenames)
                col.Error = null;

            var validRegex = new Regex(@"^[A-Za-z_][A-Za-z0-9_]*$", RegexOptions.None, TimeSpan.FromMilliseconds(100));

            var duplicates = ColumnRenames
                .GroupBy(c => (c.NewName ?? string.Empty).ToLowerInvariant())
                .Where(g => !string.IsNullOrWhiteSpace(g.Key) && g.Count() > 1)
                .SelectMany(g => g)
                .ToList();

            int validCount = 0;

            foreach (var col in ColumnRenames)
            {
                if (string.IsNullOrWhiteSpace(col.NewName))
                    col.Error = "Cannot be empty";
                else if (!validRegex.IsMatch(col.NewName))
                    col.Error = "Only letters, digits, underscore; must start with letter/underscore";
                else if (duplicates.Contains(col))
                    col.Error = "Duplicate name";
                else
                    validCount++;
            }

            ErrorMessage = ColumnRenames.Any(c => !string.IsNullOrEmpty(c.Error))
                ? "Some names are invalid."
                : null;

            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                _logger.Warn("Validation failed: " + ErrorMessage);
            }
            else
            {
                _logger.Info($"Validation succeeded. {validCount} columns valid. Saving.");
                try
                {
                    SaveRenamesToSettings();
                    _logger.Info("Column mappings saved.");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error saving column mappings.");
                    ErrorMessage = "An error occurred while saving.";
                }
            }
        }

        private void SaveRenamesToSettings()
        {
            _logger.Info($"Saving column mappings for table '{SelectedTable}'.");
            try
            {
                var mapping = new TableMapping
                {
                    TableName = SelectedTable,
                    Columns = new Dictionary<string, string>()
                };

                foreach (var col in ColumnRenames)
                    mapping.Columns[col.Original] = col.NewName;
                _settings.Set(_env.FileHash, settings =>
                {
                    var existing = settings.ColumnMappings.FirstOrDefault(m => m.TableName == SelectedTable);
                    if (existing != null)
                        settings.ColumnMappings.Remove(existing);

                    settings.ColumnMappings.Add(mapping);
                });

                _logger.Info($"Column mappings saved ({mapping.Columns.Count}) for file: {_env.FileHash}");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error saving column mappings.");
            }
        }
    }
}
