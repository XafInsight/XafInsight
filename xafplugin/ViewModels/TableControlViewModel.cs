using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using xafplugin.Database;
using xafplugin.Helpers;
using xafplugin.Interfaces;
using xafplugin.Modules;

namespace xafplugin.ViewModels
{
    /// <summary>
    /// Supplies export table definitions (name + columns) validated against the current temp database and stored settings.
    /// </summary>
    public class TableControlViewModel : ViewModelBase
    {
        public ObservableCollection<string> ExportTables { get; }
        public Dictionary<string, List<string>> ExportTableColumns { get; }
        public ObservableCollection<string> SelectedTableColumns { get; } = new ObservableCollection<string>();

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
                    UpdateColumns();
                }
            }
        }

        public TableControlViewModel() : base()
        {
            ExportTables = new ObservableCollection<string>();
            ExportTableColumns = new Dictionary<string, List<string>>();
            Initialize();
        }

        public TableControlViewModel(IEnvironmentService environment, ISettingsProvider settingsProvider, IMessageBoxService dialog)
            : base(environment, settingsProvider, dialog)
        {
            ExportTables = new ObservableCollection<string>();
            ExportTableColumns = new Dictionary<string, List<string>>();
            Initialize();
        }

        private void Initialize()
        {
            var tables = new ObservableCollection<string>();
            var tableColumns = new Dictionary<string, List<string>>();

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
                    var tablesTemp = db.GetTables();
                    if (tablesTemp != null && tablesTemp.Count > 0)
                    {
                        _logger.Debug($"Tables retrieved: {tablesTemp.Count}");
                        foreach (var t in tablesTemp)
                            tables.Add(t);
                    }

                    var columnsByTable = db.GetTableColumns(tables);
                    if (columnsByTable != null && columnsByTable.Count > 0)
                    {
                        foreach (var kv in columnsByTable)
                            tableColumns[kv.Key] = kv.Value;
                        _logger.Debug("Columns per table loaded.");
                    }
                }
                var settings = _settings.Get(_env.FileHash);
                if (settings.ExportDefinitions != null)
                {
                    foreach (var def in settings.ExportDefinitions)
                    {
                        if (string.IsNullOrWhiteSpace(def.MainTable) || !tableColumns.ContainsKey(def.MainTable))
                        {
                            _logger.Warn($"MainTable '{def.MainTable}' does not exist in database. Skipping definition '{def.Name}'.");
                            continue;
                        }

                        var validCols = def.SelectedColumns
                            .Where(col =>
                                col.IsCustom ||
                                (!string.IsNullOrWhiteSpace(col.Column) &&
                                 !string.IsNullOrWhiteSpace(col.Table) &&
                                 tableColumns.ContainsKey(col.Table) &&
                                 tableColumns[col.Table].Contains(col.Column)))
                            .Select(col => col.Column)
                            .Where(c => !string.IsNullOrWhiteSpace(c))
                            .Distinct()
                            .ToList();

                        if (validCols.Count == 0)
                        {
                            _logger.Warn($"Definition '{def.Name}' has no valid columns. Skipped.");
                            continue;
                        }

                        bool relationsValid = true;
                        if (def.Relations != null)
                        {
                            foreach (var rel in def.Relations)
                            {
                                bool ok =
                                    tableColumns.ContainsKey(rel.MainTable) &&
                                    tableColumns.ContainsKey(rel.RelatedTable) &&
                                    tableColumns[rel.MainTable].Contains(rel.MainTableColumn) &&
                                    tableColumns[rel.RelatedTable].Contains(rel.RelatedTableColumn);

                                if (!ok)
                                {
                                    _logger.Warn($"Invalid relation in definition '{def.Name}': {rel}. Skipped definition.");
                                    relationsValid = false;
                                    break;
                                }
                            }
                        }

                        if (!relationsValid)
                            continue;

                        ExportTables.Add(def.Name);
                        ExportTableColumns[def.Name] = validCols;
                        _logger.Info($"Definition '{def.Name}' added with {validCols.Count} columns.");
                    }
                }

                _logger.Info("TableControlViewModel initialization completed.");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error loading tables/columns.");
                _dialog.ShowError("Error loading tables/columns: " + ex.Message);
            }
        }

        private void UpdateColumns()
        {
            _logger.Info("Updating columns for: " + SelectedTable);
            SelectedTableColumns.Clear();

            if (!string.IsNullOrEmpty(SelectedTable) &&
                ExportTableColumns.TryGetValue(SelectedTable, out var columns))
            {
                foreach (var col in columns)
                    SelectedTableColumns.Add(col);

                _logger.Info($"Columns updated for '{SelectedTable}'.");
            }
            else
            {
                _logger.Warn("No valid table selected or table not found.");
            }
        }

        public void RemoveTable(string tableName)
        {
            _logger.Info($"Removing export definition: {tableName}");
            try
            {
                var settings = _settings.Get(_env.FileHash);
                var toRemove = settings.ExportDefinitions.FirstOrDefault(def => def.Name == tableName);
                if (toRemove != null)
                {
                    settings.ExportDefinitions.Remove(toRemove);
                    _logger.Info($"Definition '{tableName}' removed.");
                    ReloadFromSettings();
                }
                else
                {
                    _logger.Warn($"Definition '{tableName}' not found.");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"Error removing definition '{tableName}'.");
                _dialog.Show("Error removing table: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void AddTable(ExportDefinition newTable)
        {
            _logger.Info($"Adding export definition: '{newTable?.Name}'");
            try
            {
                if (newTable == null)
                    throw new ArgumentNullException(nameof(newTable), "Definition cannot be null.");

                if (string.IsNullOrWhiteSpace(newTable.MainTable))
                    throw new ArgumentException("MainTable cannot be empty.", nameof(newTable));

                if (newTable.SelectedColumns == null || !newTable.SelectedColumns.Any(c => !string.IsNullOrWhiteSpace(c.Column)))
                    throw new ArgumentException("At least one valid column is required.", nameof(newTable));
                var settings = _settings.Get(_env.FileHash);
                if (settings.ExportDefinitions.Any(def => def.Name == newTable.Name))
                    throw new InvalidOperationException($"A definition named '{newTable.Name}' already exists.");

                settings.ExportDefinitions.Add(newTable);
                _logger.Info($"Definition '{newTable.Name}' added.");
                ReloadFromSettings();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error adding export definition.");
                _dialog.Show("Error adding table: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ReloadFromSettings()
        {
            _logger.Info("Reloading export settings.");
            try
            {
                ExportTables.Clear();
                ExportTableColumns.Clear();
                Initialize();

                OnPropertyChanged(nameof(ExportTables));
                OnPropertyChanged(nameof(ExportTableColumns));
                OnPropertyChanged(nameof(SelectedTableColumns));

                _logger.Info("Reload completed.");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error reloading settings.");
                _dialog.Show("Error reloading settings: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}