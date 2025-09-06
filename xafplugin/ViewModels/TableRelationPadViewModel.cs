using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using xafplugin.Database;
using xafplugin.Interfaces;
using xafplugin.Modules;

namespace xafplugin.ViewModels
{
    public class TableRelationPadViewModel : ViewModelBase
    {
        public string Name { get; set; } = "Column1";

        public ObservableCollection<string> Tables { get; private set; }
        public Dictionary<string, List<string>> TableColumns { get; private set; }

        public ObservableCollection<TableRelation> AvailableRelations { get; } = new ObservableCollection<TableRelation>();
        public ObservableCollection<TableRelation> RelationPath { get; } = new ObservableCollection<TableRelation>();

        public List<TableRelation> UsedRelations { get; private set; } = new List<TableRelation>();

        public ObservableCollection<ColumnDescriptor> AvailableColumns { get; } = new ObservableCollection<ColumnDescriptor>();
        public ObservableCollection<ColumnDescriptor> SelectedColumns { get; } = new ObservableCollection<ColumnDescriptor>();

        public List<KeyValuePair<string, string>> CustomColumns { get; set; } = new List<KeyValuePair<string, string>>();


        private string _selectedMainTable;
        public string SelectedMainTable
        {
            get => _selectedMainTable;
            set
            {
                if (_selectedMainTable != value)
                {
                    SelectedColumns.Clear();
                    _selectedMainTable = value;
                    OnPropertyChanged(nameof(SelectedMainTable));
                    RelationPath.Clear();
                    UpdateAvailableRelations();
                    UpdateAvailableColumns();
                }
            }
        }

        private TableRelation _selectedRelation;
        public TableRelation SelectedRelation
        {
            get => _selectedRelation;
            set
            {
                if (_selectedRelation != value)
                {
                    _selectedRelation = value;
                    OnPropertyChanged(nameof(SelectedRelation));
                }
            }
        }

        public List<TableRelation> Relations { get; private set; } = new List<TableRelation>();

        public TableRelationPadViewModel() : base()
        {
            Initialize();
        }

        public TableRelationPadViewModel(IEnvironmentService environment, ISettingsProvider settingsProvider, IMessageBoxService dialog)
            : base(environment, settingsProvider, dialog)
        {
            Initialize();
        }

        private void Initialize()
        {
            Tables = new ObservableCollection<string>();
            TableColumns = new Dictionary<string, List<string>>();

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
                            TableColumns[kv.Key] = kv.Value;
                        _logger.Debug("Columns per table loaded.");
                    }

                    var relations = db.GetTableRelations(tables);
                    foreach (var r in relations)
                        Relations.Add(r);
                }

                UpdateRelations();
                _logger.Debug("Initial relations refreshed.");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error loading tables/columns.");
                _dialog.ShowError("Error loading tables/columns: " + ex.Message);
            }
        }

        public string CurrentTable => RelationPath.Count == 0 ? SelectedMainTable : RelationPath.Last().RelatedTable;

        private bool UpdateRelations()
        {
            try
            {
                var stored = _settings.Get(_env.FileHash).TableRelations;
                if (stored != null)
                {
                    foreach (var rel in stored)
                    {
                        bool valid =
                            TableColumns.ContainsKey(rel.MainTable) &&
                            TableColumns.ContainsKey(rel.RelatedTable) &&
                            TableColumns[rel.MainTable].Contains(rel.MainTableColumn) &&
                            TableColumns[rel.RelatedTable].Contains(rel.RelatedTableColumn);

                        if (valid)
                        {
                            _logger.Debug($"Valid relation: {rel}");
                            Relations.Add(rel);
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error updating relations.");
                _dialog.Show("An error occurred while updating relations: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public void AddRelationToPath()
        {
            _logger.Info("AddRelationToPath started.");
            if (SelectedRelation != null)
            {
                RelationPath.Add(SelectedRelation);
                UpdateAvailableRelations();
                UpdateAvailableColumns();
                SelectedRelation = null;
                _logger.Info("Relation added to path.");
            }
            else
            {
                _logger.Warn("No relation selected. Nothing added.");
            }
        }

        public void RemoveLastRelation()
        {
            _logger.Info("RemoveLastRelation started.");
            if (RelationPath.Count > 0)
            {
                var removed = RelationPath.Last();
                RelationPath.RemoveAt(RelationPath.Count - 1);
                _logger.Debug($"Removed relation: {removed.MainTable}.{removed.MainTableColumn} -> {removed.RelatedTable}.{removed.RelatedTableColumn}");
                UpdateAvailableRelations();
                UpdateAvailableColumns();
            }
            else
            {
                _logger.Warn("Relation path is empty.");
            }
        }

        public void UpdateAvailableRelations()
        {
            _logger.Info("Updating available relations.");
            AvailableRelations.Clear();
            Relations.Clear();
            UpdateRelations();

            var fromTable = CurrentTable;
            if (string.IsNullOrEmpty(fromTable))
            {
                _logger.Warn("CurrentTable is empty.");
                return;
            }

            var matches = Relations.Where(r => r.MainTable == fromTable).ToList();
            foreach (var rel in matches)
                AvailableRelations.Add(rel);

            _logger.Info("Available relations updated.");
        }

        public void UpdateAvailableColumns()
        {
            _logger.Info("Updating available columns.");
            AvailableColumns.Clear();

            if (string.IsNullOrEmpty(CurrentTable))
            {
                _logger.Warn("CurrentTable is empty. No columns.");
                return;
            }

            if (!TableColumns.TryGetValue(CurrentTable, out var allColumns))
            {
                _logger.Warn($"Table '{CurrentTable}' not found.");
                return;
            }

            var filtered = allColumns
                .Where(col => !SelectedColumns.Any(x => x.Table == CurrentTable && x.Column == col))
                .ToList();

            foreach (var col in filtered)
                AvailableColumns.Add(new ColumnDescriptor { Table = CurrentTable, Column = col });

            _logger.Info($"Available columns updated. {filtered.Count} columns exposed.");
        }

        public void AddColumns(System.Collections.IList items)
        {
            _logger.Info("Adding columns.");
            if (items == null || items.Count == 0)
            {
                _logger.Warn("No columns specified.");
                return;
            }

            var newColumns = new List<ColumnDescriptor>();
            foreach (ColumnDescriptor col in items)
            {
                if (!SelectedColumns.Any(x => x.Table == col.Table && x.Column == col.Column))
                {
                    newColumns.Add(new ColumnDescriptor
                    {
                        Table = col.Table,
                        Column = col.Column,
                        Path = RelationPath.ToList()
                    });
                }
            }

            var existing = SelectedColumns.ToList();
            SelectedColumns.Clear();

            foreach (var c in newColumns)
                SelectedColumns.Add(c);
            foreach (var c in existing)
                SelectedColumns.Add(c);

            UpdateAvailableColumns();
            UpdateUsedRelations();
            _logger.Info($"{newColumns.Count} new column(s) added.");
        }

        public void UpdateUsedRelations()
        {
            _logger.Info("Updating used relations.");
            var needed = new List<TableRelation>();

            foreach (var col in SelectedColumns)
            {
                foreach (var rel in col.Path)
                {
                    if (!needed.Any(x =>
                        x.MainTable == rel.MainTable &&
                        x.RelatedTable == rel.RelatedTable &&
                        x.MainTableColumn == rel.MainTableColumn &&
                        x.RelatedTableColumn == rel.RelatedTableColumn))
                    {
                        needed.Add(rel);
                    }
                }
            }

            UsedRelations = needed;
            _logger.Info($"Used relations count: {needed.Count}");
        }

        public void RemoveColumns(System.Collections.IList items)
        {
            _logger.Info("Removing columns.");
            if (items == null)
            {
                _logger.Warn("Items is null.");
                return;
            }

            var toRemove = new List<ColumnDescriptor>();
            foreach (ColumnDescriptor col in items)
            {
                var match = SelectedColumns.FirstOrDefault(x => x.Table == col.Table && x.Column == col.Column);
                if (match != null)
                {
                    if (CustomColumns.Any(c => c.Value.Contains(match.Column)))
                    {
                        _logger.Warn($"Column {col.Table}.{col.Column} is used in a custom expression.");
                        MessageBox.Show("This column cannot be removed because it is used in a custom column.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        continue;
                    }
                    toRemove.Add(match);
                }
            }

            foreach (var col in toRemove)
                SelectedColumns.Remove(col);

            UpdateAvailableColumns();
            UpdateUsedRelations();
            _logger.Info($"{toRemove.Count} column(s) removed.");
        }

        public bool MoveColumnUp(ColumnDescriptor column)
        {
            if (column == null) return false;
            int index = SelectedColumns.IndexOf(column);
            if (index <= 0) return false;
            SelectedColumns.Move(index, index - 1);
            return true;
        }

        public bool MoveColumnDown(ColumnDescriptor column)
        {
            if (column == null) return false;
            int index = SelectedColumns.IndexOf(column);
            if (index < 0 || index >= SelectedColumns.Count - 1) return false;
            SelectedColumns.Move(index, index + 1);
            return true;
        }

        public bool MoveColumnsUp(IList<ColumnDescriptor> columns)
        {
            if (columns == null || columns.Count == 0) return false;

            var ordered = columns.OrderBy(c => SelectedColumns.IndexOf(c)).ToList();
            bool moved = false;

            foreach (var column in ordered)
            {
                int index = SelectedColumns.IndexOf(column);
                if (index > 0)
                {
                    var prev = SelectedColumns[index - 1];
                    if (!columns.Contains(prev))
                    {
                        SelectedColumns.Move(index, index - 1);
                        moved = true;
                    }
                }
            }
            return moved;
        }

        public bool MoveColumnsDown(IList<ColumnDescriptor> columns)
        {
            if (columns == null || columns.Count == 0) return false;

            var ordered = columns.OrderByDescending(c => SelectedColumns.IndexOf(c)).ToList();
            bool moved = false;

            foreach (var column in ordered)
            {
                int index = SelectedColumns.IndexOf(column);
                if (index >= 0 && index < SelectedColumns.Count - 1)
                {
                    var next = SelectedColumns[index + 1];
                    if (!columns.Contains(next))
                    {
                        SelectedColumns.Move(index, index + 1);
                        moved = true;
                    }
                }
            }
            return moved;
        }

        public ExportDefinition CurrentExportDefinition
        {
            get
            {
                return new ExportDefinition
                {
                    Name = Name,
                    MainTable = SelectedMainTable,
                    SelectedColumns = SelectedColumns.ToList(),
                    Relations = UsedRelations,
                    CaseExpressions = CustomColumns.ToList()
                };
            }
        }
    }
}
