using NLog;
using NLog.Config;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using xafplugin.Database;
using xafplugin.Helpers;
using xafplugin.Modules;

namespace xafplugin.ViewModels
{
    public class OpenFileWindowViewModel : ViewModelBase
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private ObservableCollection<string> _entries;
        private ObservableCollection<string> _filtered;
        private string[] _selectedEntry;
        private string _filterText;
        private const string OpenFileFilter = "Audit/XML files (*.xaf;*.xml;*.csv;*.db;*.sqlite)|*.xaf;*.xml;*.csv;*.db;*.sqlite";

        private bool _isBusy;
        private string _busyText;

        public event EventHandler<bool> RequestClose;

        public ObservableCollection<string> FilteredEntries
        {
            get => _filtered;
            private set
            {
                _filtered = value;
                OnPropertyChanged(nameof(FilteredEntries));
            }
        }

        public string SelectedEntry
        {
            get
            {
                if (_selectedEntry != null && _selectedEntry.Length > 0)
                {
                    return _selectedEntry[0];
                }
                return null;
            }
            set
            {
                _selectedEntry = value != null ? new string[] { value } : null;
                OnPropertyChanged(nameof(SelectedEntry));
            }
        }

        public string FilterText
        {
            get => _filterText;
            set
            {
                if (_filterText != value)
                {
                    _filterText = value;
                    OnPropertyChanged(nameof(FilterText));
                    ApplyFilter();
                }
            }
        }

        public bool IsBusy
        {
            get => _isBusy;
            private set
            {
                if (_isBusy != value)
                {
                    _isBusy = value;
                    OnPropertyChanged(nameof(IsBusy));
                }
            }
        }

        public string BusyText
        {
            get => _busyText;
            private set
            {
                if (_busyText != value)
                {
                    _busyText = value;
                    OnPropertyChanged(nameof(BusyText));
                }
            }
        }

        public ICommand ConfirmCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand SelectFileCommand { get; }

        public OpenFileWindowViewModel(System.Collections.Generic.IEnumerable<string> entries) : base()
        {
            _entries = new ObservableCollection<string>(entries.Where(s => s != null));
            _filtered = new ObservableCollection<string>(_entries);
            ConfirmCommand = new RelayCommand(_ => CanConfirm(), async _ => await ConfirmAsync());
            CancelCommand = new RelayCommand(_ => true, _ => Cancel());
            SelectFileCommand = new RelayCommand(_ => true, async _ => await SelectFileAsync());
        }

        private bool CanConfirm()
        {
            if (_selectedEntry == null || _selectedEntry.Length <= 0)
            {
                return false;
            }
            return true;
        }

        private async Task ConfirmAsync()
        {
            if (!CanConfirm())
                return;

            _logger.Info("OpenFileWindow confirm met selectie: {0}", _selectedEntry);

            var app = Globals.ThisAddIn.Application;
            var wb = app?.ActiveWorkbook;
            if (wb == null)
            {
                _dialog.ShowWarning("No active workbook is open.");
                _logger.Warn("Connect aborted: no active workbook.");
                return;
            }

            await FileHandleAsync();
            RequestClose?.Invoke(this, true);
        }

        private async Task SelectFileAsync()
        {
            _logger.Info("User initiated file selection.");

            using (var dlg = new OpenFileDialog
            {
                Title = "Select file(s)",
                Filter = OpenFileFilter,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Multiselect = true,
                CheckFileExists = true
            })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    if (dlg.FileNames?.Length > 0)
                    {
                        if (dlg.FileNames.Length > 1)
                        {
                            _logger.Info("User selected {0} files. Current import pipeline uses the first only.", dlg.FileNames.Length);
                        }

                        _selectedEntry = dlg.FileNames;
                        _logger.Info("User selected file: {0}", _selectedEntry);
                    }
                }
                else
                {
                    _logger.Info("User canceled file selection.");
                }
            }

            if (!CanConfirm())
            {
                _logger.Info("No file chosen. Import canceled.");
                return;
            }
            await FileHandleAsync();
            RequestClose?.Invoke(this, true);
        }

        private async Task FileHandleAsync()
        {
            if (_selectedEntry == null || _selectedEntry.Length == 0)
                return;

            if (IsBusy)
            {
                _logger.Warn("FileHandleAsync invoked while busy.");
                return;
            }

            if (isDatabasefile(_selectedEntry[0]))
            {
                SetBusy("Connecting to database...");
                try
                {
                    if (ConnectSQLDatabase(_selectedEntry))
                    {
                        HistoryFileManager.Append(Globals.ThisAddIn.Config.ConfigPath, _selectedEntry[0], Globals.ThisAddIn.Config.HistoryFileAmount, out string FileSelectError);
                        return;
                    }
                    else
                    {
                        _logger.Info("Failed to connect to the selected database file(s). Import canceled.");
                        _dialog.ShowWarning("Failed to connect to the selected database file(s). Import canceled.");
                        return;
                    }
                }
                finally
                {
                    ClearBusy();
                }
            }

            if (_selectedEntry.Length == 1)
            {
                HistoryFileManager.Append(Globals.ThisAddIn.Config.ConfigPath, _selectedEntry[0], Globals.ThisAddIn.Config.HistoryFileAmount, out string FileSelectError);
                if (!string.IsNullOrEmpty(FileSelectError))
                {
                    _logger.Warn("Failed to update history file: {0}", FileSelectError);
                }
            }

            _logger.Info("Starting audit file import.");

            try
            {
                var databaseBasePath = Globals.ThisAddIn.Config.TempDatabasePath;
                var FileHash = UniqueFileName.GetFileHash(_selectedEntry);
                var dbName = "XafInsight_" + FileHash;
                var fullpath = UniqueFileName.CombinePathAndName(dbName, databaseBasePath, EFileType.db);

                if (tempDatabaseExists(fullpath))
                {
                    if (reopenFile(fullpath))
                    {
                        Globals.ThisAddIn.FilePath = fullpath;
                        _env.DatabasePath = fullpath;
                        _env.FileHash = FileHash;
                        _logger.Info("Reconnect database");
                        _dialog.Show("Audit file reconnected successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.None);
                        return;
                    }
                }

                SetBusy("Generating database...");
                bool generated = false;
                try
                {
                    if (isCSVfile(_selectedEntry[0]))
                    {
                        generated = await GenerateCSVAsync(fullpath, CancellationToken.None);
                    }
                    else
                    {
                        generated = await GenerateDatabaseAsync(fullpath, CancellationToken.None);
                    }
                }
                finally
                {
                    ClearBusy();
                }

                if (!generated)
                {
                    _logger.Warn("Failed to generate database.");
                    throw new InvalidOperationException("Failed to generate database.");
                }
                _logger.Info("Database generated at: {0}", fullpath);

                Globals.ThisAddIn.FilePath = fullpath;
                _env.DatabasePath = fullpath;
                _env.FileHash = FileHash;

                if (_selectedEntry.Length == 1)
                {
                    _logger.Info("One file selected, finish connection.");
                    _dialog.Show("Audit file imported successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.None);
                    return;
                }

                var result = _dialog.Show(
                     "Multiple files were selected.\n" +
                     "It is highly recommended to export the database for future use.\n" +
                     "Do you want to export the database?",
                     "Confirm",
                     MessageBoxButtons.YesNo,
                     MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    if (string.IsNullOrEmpty(fullpath) || !File.Exists(fullpath))
                    {
                        _logger.Warn("Temporary database file not found: {0}", fullpath);
                        throw new FileNotFoundException("Temporary database file not found.", fullpath);
                    }
                    try
                    {
                        using (var dlg = new SaveFileDialog
                        {
                            Title = "Export SQLite Database",
                            Filter = "SQLite Database (*.db;*.sqlite)|*.db;*.sqlite|All files (*.*)|*.*",
                            FileName = "exported_database.db",
                            InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                            OverwritePrompt = true,
                            AddExtension = true
                        })
                        {
                            if (dlg.ShowDialog() == DialogResult.OK)
                            {
                                File.Copy(fullpath, dlg.FileName, overwrite: false);
                                _logger.Info("Database exported to: {0}", dlg.FileName);

                                FileHash = UniqueFileName.GetFileHash(dlg.FileName);
                                HistoryFileManager.Append(Globals.ThisAddIn.Config.ConfigPath, dlg.FileName, Globals.ThisAddIn.Config.HistoryFileAmount, out string FileSelectError);

                                Globals.ThisAddIn.FilePath = dlg.FileName;
                                _env.DatabasePath = dlg.FileName;
                                _env.FileHash = FileHash;
                                _logger.Info("New database location set.");
                                _dialog.Show("Audit file imported successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.None);
                            }
                            else
                            {
                                _logger.Info("Database export canceled by user.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Failed to export database.");
                        _dialog.ShowError("An error occurred while exporting the database.");
                    }
                }
                else
                {
                    _dialog.Show("Audit file imported successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.None);
                }
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.FilePath = null;
                _env.DatabasePath = null;
                _env.FileHash = null;
                _logger.Error(ex, "Error importing audit file: {0}", _selectedEntry);
                _dialog.ShowError("An error occurred while importing the audit file. :\n" + ex.Message);
            }
        }



        private bool reopenFile(string fullpath)
        {
            try
            {
                var result = _dialog.Show(
                                $"A temporary database for this file already exists. Do you want to reconnect?\n" +
                                "Not reconnecting result in deleting current file",
                                "Confirm",
                                System.Windows.Forms.MessageBoxButtons.YesNo,
                                System.Windows.Forms.MessageBoxIcon.Warning);

                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    _logger.Info("File already exists, reconnect");
                    if (ConnectSQLDatabase(new string[] { fullpath }))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to handle existing temporary database.");
                _dialog.ShowError("An error occurred while handling the existing temporary database.");
                throw;
            }
        }

        private void Cancel()
        {
            SelectedEntry = null;
            _logger.Info("OpenFileWindow geannuleerd.");
            RequestClose?.Invoke(this, false);
        }

        private void ApplyFilter()
        {
            var term = (FilterText ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(term))
            {
                FilteredEntries = new ObservableCollection<string>(_entries);
                return;
            }

            var filtered = _entries
                .Where(e => e?.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            FilteredEntries = new ObservableCollection<string>(filtered);

            if (SelectedEntry != null && !FilteredEntries.Contains(SelectedEntry))
            {
                SelectedEntry = null;
            }
        }

        private async Task<bool> GenerateDatabaseAsync(string fullPath, CancellationToken ct)
        {
            if (!DeleteIfExists(fullPath))
            {
                throw new IOException("Unable to prepare temporary database file (could not delete existing file): " + fullPath);
            }

            return await Task.Run(() =>
            {
                try
                {
                    using (var SQLiteHelper = new SqliteHelper(fullPath))
                    {
                        using (var conn = SQLiteHelper.GetWriteConnection())
                        {
                            using (var importer = new DynamicXmlImporter(conn))
                            {
                                foreach (string file in _selectedEntry)
                                {
                                    ct.ThrowIfCancellationRequested();

                                    if (!File.Exists(file))
                                    {
                                        _logger.Warn("Selected file does not exist: {0}", file);
                                        throw new FileNotFoundException("Selected file does not exist.", file);
                                    }

                                    if (string.IsNullOrEmpty(fullPath))
                                    {
                                        throw new InvalidOperationException("Failed to determine database path.");
                                    }

                                    using (var xmlStream = File.OpenRead(file))
                                    {
                                        importer.Import(xmlStream);
                                    }
                                }
                            }
                        }
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    return true;
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error generating database.");
                    if (File.Exists(fullPath))
                    {
                        try
                        {
                            File.Delete(fullPath);
                        }
                        catch (Exception delEx)
                        {
                            _logger.Warn(delEx, "Failed to delete database file after generation error: {0}", fullPath);
                        }
                    }
                    throw;
                }
            }, ct);
        }

        private async Task<bool> GenerateCSVAsync(string fullPath, CancellationToken ct)
        {
            if (!DeleteIfExists(fullPath))
            {
                throw new IOException("Unable to prepare temporary database file (could not delete existing file): " + fullPath);
            }

            return await Task.Run(() =>
            {

                using (var SQLiteHelper = new SqliteHelper(fullPath))
                {
                    using (var conn = SQLiteHelper.GetWriteConnection())
                    {
                        using (var csvImporter = new CsvImporter(conn, ';'))
                        {
                            foreach (string file in _selectedEntry)
                            {
                                ct.ThrowIfCancellationRequested();
                                if (!File.Exists(file))
                                {
                                    _logger.Warn("Selected file does not exist: {0}", file);
                                    throw new FileNotFoundException("Selected file does not exist.", file);
                                }

                                csvImporter.Import(file);
                            }
                        }
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    return true;
                }
            }, ct);
        }

        private bool DeleteIfExists(string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path))
                {
                    _logger.Warn("DeleteIfExists called with null/empty path.");
                    return false;
                }

                if (!File.Exists(path))
                    return true;

                try
                {
                    var attrs = File.GetAttributes(path);
                    if ((attrs & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    {
                        File.SetAttributes(path, attrs & ~FileAttributes.ReadOnly);
                    }
                }
                catch (Exception exAttr)
                {
                    _logger.Warn(exAttr, "Could not adjust file attributes before deletion: {0}", path);
                }

                const int maxAttempts = 3;
                for (int attempt = 1; attempt <= maxAttempts; attempt++)
                {
                    try
                    {
                        File.Delete(path);

                        if (!File.Exists(path))
                        {
                            if (attempt > 1)
                                _logger.Info("Deleted file after {0} attempt(s): {1}", attempt, path);
                            else
                                _logger.Info("Deleted existing file: {0}", path);
                            return true;
                        }
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        _logger.Warn(ex, "UnauthorizedAccess deleting file (attempt {0}/{1}): {2}", attempt, maxAttempts, path);
                        if (attempt == maxAttempts) return false;
                    }
                    catch (IOException ex)
                    {
                        _logger.Warn(ex, "IO error deleting file (attempt {0}/{1}): {2}", attempt, maxAttempts, path);
                        if (attempt == maxAttempts) return false;
                    }

                    Thread.Sleep(100 * attempt);
                }

                return !File.Exists(path);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Unexpected error while deleting file: {0}", path);
                return false;
            }
        }

        private class RelayCommand : ICommand
        {
            private readonly Predicate<object> _can;
            private readonly Action<object> _exec;

            public RelayCommand(Predicate<object> can, Action<object> exec)
            {
                _can = can ?? (_ => true);
                _exec = exec ?? throw new ArgumentNullException(nameof(exec));
            }

            public bool CanExecute(object parameter) => _can(parameter);
            public void Execute(object parameter) => _exec(parameter);
            public event EventHandler CanExecuteChanged
            {
                add { CommandManager.RequerySuggested += value; }
                remove { CommandManager.RequerySuggested -= value; }
            }
        }

        private bool isDatabasefile(string path)
        {
            return path.EndsWith(".sqlite", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".db", StringComparison.OrdinalIgnoreCase);
        }

        private bool isCSVfile(string path)
        {
            return path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".txt", StringComparison.OrdinalIgnoreCase);
        }

        private bool tempDatabaseExists(string dbPath)
        {
            return File.Exists(dbPath);
        }

        private bool ConnectSQLDatabase(string[] selectedPath)
        {
            try
            {
                _logger.Info("Selected file(s) contain a SQLite database. Please use the 'Import Database' option instead.");
                if (selectedPath.Length > 1)
                {
                    _logger.Info("Multiple databasefiles selected. Only the first file will be used.");
                    _dialog.ShowWarning("Multiple databasefiles selected. Only the first file will be used.");
                }
                var path = selectedPath[0];
                if (!File.Exists(path))
                {
                    _logger.Warn("Selected file does not exist: {0}", path);
                    _dialog.ShowWarning($"The selected file does not exist:\n{path}");
                    throw new FileNotFoundException("Selected file does not exist.", path);
                }
                if (!SqliteHelper.IsLikelySQLiteFile(path))
                {
                    _logger.Warn("The selected file is not compatible:", path);
                    _dialog.ShowWarning($"The selected file is not compatible:\n{path}");
                    throw new InvalidDataException("The selected file is not compatible.");
                }

                var FileHash = UniqueFileName.GetFileHash(selectedPath[0]);
                _env.FileHash = FileHash;
                Globals.ThisAddIn.FilePath = path;
                _env.DatabasePath = path;
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error connecting to database file: {0}", selectedPath);
                _dialog.ShowError(ex.Message);
                return false;
            }
        }

        private void SetBusy(string text)
        {
            BusyText = text;
            IsBusy = true;
        }

        private void ClearBusy()
        {
            IsBusy = false;
            BusyText = null;
        }
    }
}