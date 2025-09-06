using NLog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using xafplugin.Database;
using xafplugin.Helpers;
using xafplugin.Interfaces;
using xafplugin.Modules;
using Excel = Microsoft.Office.Interop.Excel;

namespace xafplugin.ViewModels
{
    public class ExportTableViewModel : ViewModelBase
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public Visibility InvalidCell { get; set; } = Visibility.Collapsed;
        private List<ExportDefinition> _exportTables { get; set; }
        private ExportDefinition _selectedTable { get; set; }

        // Change the Filters property to use ObservableCollection<FilterItem>
        public ObservableCollection<FilterItem> Filters { get; set; } = new ObservableCollection<FilterItem>();


        private Excel.Range _targetCell;

        public ExportTableViewModel(IEnvironmentService environment, ISettingsProvider settingsProvider, IMessageBoxService dialog) : base(environment, settingsProvider, dialog)
        {
            _exportTables = new List<ExportDefinition>();

            LoadAvailableTables();
        }

        public ExportTableViewModel() : base()
        {
            _exportTables = new List<ExportDefinition>();

            LoadAvailableTables();
        }

        public string ExportLocation
        {
            get
            {
                try
                {

                    if (_targetCell == null)
                        return string.Empty;

                    // Pak de bladnaam
                    string sheetName = _targetCell.Worksheet.Name;

                    // Als er spaties of speciale tekens zijn in de naam → zet 'eromheen
                    if (sheetName.Contains(" "))
                        sheetName = $"'{sheetName}'";

                    // Geef het volledige adres in A1-notatie zonder $
                    string address = _targetCell.Address[false, false];

                    return $"{sheetName}!{address}";
                }
                catch
                {
                    return "";
                }
            }
        }

        public List<ExportDefinition> ExportTables
        {
            get => _exportTables;
            set
            {
                _exportTables = value;
                OnPropertyChanged(nameof(ExportTables));
            }
        }

        public ExportDefinition SelectedTable
        {
            get => _selectedTable;
            set
            {
                if (_selectedTable != value)
                {
                    _selectedTable = value;
                    OnPropertyChanged(nameof(SelectedTable));

                    Filters.Clear();
                    OnPropertyChanged(nameof(Filters));
                }
            }
        }

        public void LoadAvailableTables()
        {
            var tables = new ObservableCollection<string>();
            var tableColumns = new Dictionary<string, List<string>>();
            var exportTables = new List<ExportDefinition>();

            try
            {
                var databasePath = _env.DatabasePath;
                if (string.IsNullOrEmpty(databasePath))
                {
                    _logger.Warn("Geen databasepad gevonden. Auditfile moet eerst worden ingelezen.");
                    _dialog.ShowInfo("Er is geen database beschikbaar. Zorg ervoor dat de auditfile is ingelezen.");
                    return;
                }

                using (var db = new DatabaseService(_env.DatabasePath))
                {
                    var tablesTemp = db.GetTables();

                    if (tablesTemp != null || tablesTemp.Count > 0)
                    {
                        base._logger.Debug($"Aantal tabellen opgehaald: {tablesTemp.Count}");
                        foreach (var t in tablesTemp)
                            tables.Add(t);
                    }

                    var columnsByTable = db.GetTableColumns(tables);
                    if (columnsByTable != null && columnsByTable.Count > 0)
                    {
                        foreach (var kv in columnsByTable)
                            tableColumns[kv.Key] = kv.Value;

                        base._logger.Debug("Kolommen per tabel succesvol geladen.");
                    }
                }
                FileSettings settings = _settings.Get(_env.FileHash);
                if (settings.ExportDefinitions != null)
                {
                    foreach (var item in settings.ExportDefinitions)
                    {

                        if (string.IsNullOrWhiteSpace(item.MainTable) || !tableColumns.ContainsKey(item.MainTable))
                        {
                            _logger.Warn($"MainTable '{item.MainTable}' bestaat niet in de database. Overslaan.");
                            continue;
                        }

                        var validCols = xafplugin.Helpers.ExportValidationHelper.FilterValidColumns(item.SelectedColumns, tableColumns);

                        item.SelectedColumns = validCols;

                        if (validCols.Count == 0)
                        {
                            _logger.Warn($"Tabel '{item.Name}' heeft geen geldige kolommen. Overslaan.");
                            continue;
                        }

                        var validTables = validCols.Select(vc => vc.Table).Distinct().ToHashSet();

                        var usedRelations = ExportValidationHelper.FilterRelationsByTables(item.Relations, validTables);


                        item.Relations = usedRelations;

                        bool relatiesGeldig = true;

                        foreach (var rel in usedRelations)
                        {
                            bool geldig = ExportValidationHelper.IsRelationValid(rel, tableColumns);

                            if (!geldig)
                            {
                                _logger.Warn($"Ongeldige relatie: {rel}. Overslaan van tabel '{item.Name}'.");
                                relatiesGeldig = false;
                                break;
                            }
                        }

                        if (!relatiesGeldig)
                        {
                            _logger.Warn($"Minstens één relatie van '{item.Name}' is ongeldig. Overslaan.");
                            continue;
                        }

                        exportTables.Add(item);

                        _logger.Info($"Tabel '{item.Name}' succesvol toegevoegd met {validCols.Count} kolommen.");
                    }
                }
                ExportTables = exportTables.ToList();
                _logger.Info("Initialisatie van ExportTables voltooid.");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Fout bij laden van tabellen/kolommen.");
                _dialog.ShowError("Fout bij laden tabellen/kolommen: " + ex.Message);
            }
        }

        public void ExportSelectedTable()
        {
            try
            {
                if (SelectedTable == null)
                {
                    _dialog.ShowWarning("Selecteer een tabel om te exporteren.");
                    return;
                }

                var sqlQueryWithCte = ExportDefinitionHelper.SqlQueryWithCte(SelectedTable, Filters.ToList());

                object[,] data = null;
                using (var db = new DatabaseService(_env.DatabasePath))
                {
                    data = db.Select2DArray(sqlQueryWithCte);
                }

                if (data == null)
                {
                    _dialog.ShowInfo("Geen gegevens om te exporteren.");
                    return;
                }
                var usedMappings = new HashSet<string>();

                var columnMappingList = _settings.Get(_env.FileHash)?.ColumnMappings;
                if (columnMappingList != null || columnMappingList.Count > 0)
                {
                    var colCount = data.GetLength(1);
                    for (int col = 0; col < colCount; col++)
                    {
                        var headerObj = data[0, col];
                        var header = headerObj?.ToString();
                        if (string.IsNullOrEmpty(header)) continue;

                        foreach (var table in columnMappingList)
                        {
                            if (table.Columns.TryGetValue(header, out var mapping) && !string.IsNullOrEmpty(mapping))
                            {
                                if (!string.Equals(header, mapping, StringComparison.OrdinalIgnoreCase))
                                {
                                    data[0, col] = mapping;
                                    usedMappings.Add(header + " -> " + mapping);
                                }
                                break;
                            }
                        }
                    }
                }

                var lines = new List<string>();

                if (_showExportSourceFile)
                {
                    var filePath = Globals.ThisAddIn.FilePath;
                    if (!string.IsNullOrEmpty(filePath))
                    {
                        lines.Add($"Data exported from file: {filePath}");
                        lines.Add("");
                    }
                }

                if (_showExportTables)
                {
                    lines.AddRange(getExportTablesLines());
                }

                if (_showExportColumns)
                {
                    lines.AddRange(getExportColumnLines());
                }

                if (_showExportColumnRelations)
                {
                    var relations = SelectedTable.Relations.ToList();
                    if (relations.Count > 0)
                    {
                        lines.Add("Relationships used:");
                        foreach (var item in relations)
                        {
                            lines.Add($"{item.ToString()}");
                        }
                        lines.Add("");
                    }
                }

                if (_showExportFilters && Filters.Count > 0)
                {
                    lines.Add("Applied filters:");
                    foreach (var filter in Filters)
                    {
                        lines.Add($"- {filter.Name}: {filter.Expression}");
                    }
                    lines.Add("");
                }

                if (_showExportMapping)
                {
                    lines.AddRange(GetExportMappingLines(usedMappings));
                }

                if (lines.Count > 0)
                {
                    lines.Add("");
                }

                data = Array2DHelpers.PrependTextLines(data, lines, fill: null);

                _logger.Info($"SQL-query voor export van '{SelectedTable} uitgevoerd, aantal regels terug': {data.Length}");

                try
                {
                    if (_targetCell == null)
                        throw new InvalidOperationException("Geen geldige cell geselecteerd.");

                    if (_targetCell.Worksheet == null)
                        throw new InvalidOperationException("Werkblad niet meer aanwezig.");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Fout bij het valideren van Excel componenten.");
                    _dialog.ShowError($"Fout bij het valideren van Excel componenten. Probeer opnieuw een cel te selecteren.");
                    return;
                }

                int rows = data.GetLength(0);
                int cols = data.GetLength(1);

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        if (data[r, c] == null || data[r, c] is DBNull)
                            data[r, c] = ""; // of een andere default waarde
                    }
                }
                if (rows == 0 || cols == 0)
                {
                    _dialog.ShowInfo("Geen rijen/kolommen om te exporteren.");
                    return;
                }

                if (!ExcelHelper.ValidateRangeSize(_targetCell, rows, cols, _dialog))
                {
                    return;
                }

                // Use cells from the new worksheet
                Excel.Range writeRange = _targetCell.Resize[rows, cols];

                //check if the range already has data
                if (ExcelHelper.RangeHasData(writeRange, _dialog))
                {
                    return;
                }

                // Write data to Excel
                writeRange.Value2 = data;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"Fout bij exporteren van tabel '{SelectedTable}'.");
                _dialog.ShowError($"Fout bij exporteren: {ex.Message}");
            }
        }

        private string[] getExportTablesLines()
        {
            if (SelectedTable == null || SelectedTable.SelectedColumns == null)
                return new string[0];

            var columns = SelectedTable.SelectedColumns.Select(c => c.Table).Distinct().ToList();

            // Calculate how many lines we need (10 columns per line)
            int lineCount = (columns.Count + 9) / 10; // Ceiling division

            // Add one more line for the empty row
            string[] lines = new string[lineCount + 1];

            for (int i = 0; i < lineCount; i++)
            {
                var chunk = columns.Skip(i * 10).Take(10);
                lines[i] = i == 0
                    ? $"Selected tables are {string.Join(", ", chunk)}."
                    : $"{string.Join(", ", chunk)}.";
            }

            // Add empty line at the end
            lines[lineCount] = "";

            return lines;
        }

        private string[] GetExportMappingLines(HashSet<string> usedMappings)
        {
            if (usedMappings == null || usedMappings.Count == 0)
                return Array.Empty<string>();

            // Zorg voor consistente volgorde
            var mappings = usedMappings.OrderBy(x => x).ToList();

            int lineCount = (mappings.Count + 9) / 10; // ceiling division
            var lines = new List<string>(lineCount + 1);

            for (int i = 0; i < lineCount; i++)
            {
                var chunk = mappings.Skip(i * 10).Take(10);
                string line = (i == 0)
                    ? $"Used mappings: {string.Join(", ", chunk)}."
                    : $"{string.Join(", ", chunk)}.";
                lines.Add(line);
            }

            // optioneel: lege regel toevoegen
            lines.Add(string.Empty);

            return lines.ToArray();
        }


        private string[] getExportColumnLines()
        {
            if (SelectedTable == null || SelectedTable.SelectedColumns == null)
                return new string[0];

            var columns = SelectedTable.SelectedColumns.Select(c => c.Display).ToList();

            int lineCount = (columns.Count + 9) / 10;

            string[] lines = new string[lineCount + 1];

            for (int i = 0; i < lineCount; i++)
            {
                var chunk = columns.Skip(i * 10).Take(10);
                lines[i] = i == 0
                    ? $"Selected columns are {string.Join(", ", chunk)}."
                    : $"{string.Join(", ", chunk)}.";
            }

            // Add empty line at the end
            lines[lineCount] = "";

            return lines;
        }

        public bool setTargetCell(Excel.Range target)
        {

            if (target == null)
            {
                return false;
            }

            try
            {
                if (!ExcelHelper.ValidateWorksheet(target.Worksheet, _dialog))
                    return false;

                _targetCell = target;
                OnPropertyChanged(nameof(ExportLocation));
                return true;
            }
            catch (Exception ex)
            {
                _targetCell = null;
                return false;
            }
        }

        private bool _showExportTables = false;
        public bool ShowExportTables
        {
            get => _showExportTables;
            set
            {
                if (_showExportTables != value)
                {
                    _showExportTables = value;
                    OnPropertyChanged(nameof(ShowExportTables));
                }
            }
        }

        private bool _showExportColumns = false;
        public bool ShowExportColumns
        {
            get => _showExportColumns;
            set
            {
                if (_showExportColumns != value)
                {
                    _showExportColumns = value;
                    OnPropertyChanged(nameof(ShowExportColumns));
                }
            }
        }

        private bool _showExportFilters = false;
        public bool ShowExportFilters
        {
            get => _showExportFilters;
            set
            {
                if (_showExportFilters != value)
                {
                    _showExportFilters = value;
                    OnPropertyChanged(nameof(ShowExportFilters));
                }
            }
        }

        private bool _showExportSourceFile = false;
        public bool ExportSourceFile
        {
            get => _showExportSourceFile;
            set
            {
                if (_showExportSourceFile != value)
                {
                    _showExportSourceFile = value;
                    OnPropertyChanged(nameof(ExportSourceFile));
                }
            }
        }

        private bool _showExportMapping = false;
        public bool ShowExportMapping
        {
            get => _showExportMapping;
            set
            {
                if (_showExportMapping != value)
                {
                    _showExportMapping = value;
                    OnPropertyChanged(nameof(ShowExportMapping));
                }
            }
        }


        private bool _showExportColumnRelations = false;
        public bool ExportColumnRelations
        {
            get => _showExportColumnRelations;
            set
            {
                if (_showExportColumnRelations != value)
                {
                    _showExportColumnRelations = value;
                    OnPropertyChanged(nameof(ExportColumnRelations));
                }
            }
        }
    }
}
