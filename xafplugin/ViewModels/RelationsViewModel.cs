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
    /// Manages table relations discovered from the temporary SQLite database and persisted settings.
    /// </summary>
    public class RelationsViewModel : ViewModelBase
    {
        public ObservableCollection<TableRelation> Relations { get; } = new ObservableCollection<TableRelation>();

        public List<string> Tables { get; private set; }
        public Dictionary<string, List<string>> TableColumns { get; private set; }

        private string _mainTable;
        public string MainTable
        {
            get => _mainTable;
            set
            {
                if (_mainTable != value)
                {
                    _mainTable = value;
                    MainTableColumn = null;
                    OnPropertyChanged(nameof(MainTable));
                    OnPropertyChanged(nameof(MainTableColumns));
                    OnPropertyChanged(nameof(MainTableColumn));
                }
            }
        }

        public IEnumerable<EJoinType> JoinTypes => Enum.GetValues(typeof(EJoinType)).Cast<EJoinType>();

        private EJoinType _joinType = EJoinType.LeftOuter;
        public EJoinType JoinType
        {
            get => _joinType;
            set
            {
                if (_joinType != value)
                {
                    _joinType = value;
                    OnPropertyChanged(nameof(JoinType));
                }
            }
        }

        private string _relatedTable;
        public string RelatedTable
        {
            get => _relatedTable;
            set
            {
                if (_relatedTable != value)
                {
                    _relatedTable = value;
                    RelatedTableColumn = null;
                    OnPropertyChanged(nameof(RelatedTable));
                    OnPropertyChanged(nameof(RelatedTableColumns));
                    OnPropertyChanged(nameof(RelatedTableColumn));
                }
            }
        }

        private string _mainTableColumn;
        public string MainTableColumn
        {
            get => _mainTableColumn;
            set
            {
                if (_mainTableColumn != value)
                {
                    _mainTableColumn = value;
                    OnPropertyChanged(nameof(MainTableColumns));
                }
            }
        }

        private string _relatedTableColumn;
        public string RelatedTableColumn
        {
            get => _relatedTableColumn;
            set
            {
                if (_relatedTableColumn != value)
                {
                    _relatedTableColumn = value;
                    OnPropertyChanged(nameof(RelatedTableColumns));
                }
            }
        }

        public List<string> MainTableColumns =>
            !string.IsNullOrEmpty(MainTable) && TableColumns.TryGetValue(MainTable, out var cols)
                ? cols
                : new List<string>();

        public List<string> RelatedTableColumns =>
            !string.IsNullOrEmpty(RelatedTable) && TableColumns.TryGetValue(RelatedTable, out var cols)
                ? cols
                : new List<string>();

        public RelationsViewModel() : base()
        {
            _logger.Info("Initializing RelationsViewModel.");
            Initialize();
        }

        public RelationsViewModel(IEnvironmentService environment, ISettingsProvider settingsProvider, IMessageBoxService dialog)
            : base(environment, settingsProvider, dialog)
        {
            _logger.Info("Initializing RelationsViewModel.");
            Initialize();
        }

        private void Initialize()
        {
            Tables = new List<string>();
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
                        Tables.AddRange(tables);
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
                var storedRelations = _settings.Get(_env.FileHash).TableRelations;
                foreach (var r in storedRelations)
                {
                    bool columnsExist =
                        TableColumns.ContainsKey(r.MainTable) &&
                        TableColumns[r.MainTable].Contains(r.MainTableColumn) &&
                        TableColumns.ContainsKey(r.RelatedTable) &&
                        TableColumns[r.RelatedTable].Contains(r.RelatedTableColumn);

                    if (!columnsExist)
                    {
                        _logger.Warn($"Skipped relation - columns missing: {r.MainTable}.{r.MainTableColumn} → {r.RelatedTable}.{r.RelatedTableColumn}");
                        continue;
                    }

                    bool alreadyExists = Relations.Any(existing =>
                        existing.MainTable == r.MainTable &&
                        existing.MainTableColumn == r.MainTableColumn &&
                        existing.RelatedTable == r.RelatedTable &&
                        existing.RelatedTableColumn == r.RelatedTableColumn &&
                        existing.JoinType == r.JoinType);

                    if (!alreadyExists)
                    {
                        Relations.Add(r);
                        _logger.Debug($"Relation added from settings: {r}");
                    }
                }

                _logger.Info("RelationsViewModel initialization completed.");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during initialization of RelationsViewModel.");
                _dialog.Show("An error occurred while loading relation information.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool AddRelation()
        {
            _logger.Info("Attempting to add relation.");

            if (string.IsNullOrEmpty(MainTable) ||
                string.IsNullOrEmpty(MainTableColumn) ||
                string.IsNullOrEmpty(RelatedTable) ||
                string.IsNullOrEmpty(RelatedTableColumn))
            {
                _logger.Warn("Not added: one or more required fields are empty.");
                _dialog.Show("One or more required fields are empty.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.Equals(MainTable, RelatedTable, StringComparison.OrdinalIgnoreCase))
            {
                _logger.Warn("Not added: cannot relate table to itself.");
                _dialog.Show("A relation cannot link the same table.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            var relation = new TableRelation
            {
                MainTable = MainTable,
                MainTableColumn = MainTableColumn,
                RelatedTable = RelatedTable,
                RelatedTableColumn = RelatedTableColumn,
                JoinType = JoinType
            };

            Relations.Add(relation);
            _logger.Info($"Relation added: {MainTable}.{MainTableColumn} → {RelatedTable}.{RelatedTableColumn}");

            ClearFields();
            return true;
        }

        public void ClearFields()
        {
            _logger.Debug("Clearing relation input fields.");
            MainTable = null;
            MainTableColumn = null;
            RelatedTable = null;
            RelatedTableColumn = null;

            OnPropertyChanged(nameof(MainTable));
            OnPropertyChanged(nameof(MainTableColumn));
            OnPropertyChanged(nameof(RelatedTable));
            OnPropertyChanged(nameof(RelatedTableColumn));
        }

        public void SaveRelationsToSettings()
        {
            _settings.Set(_env.FileHash, settings =>
            {
                settings.TableRelations = new ObservableCollection<TableRelation>(Relations);
            });
            _logger.Info("Relations saved to settings.");
        }
    }
}
