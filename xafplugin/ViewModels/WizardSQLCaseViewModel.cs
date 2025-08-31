using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Input;
using xafplugin.Database;
using xafplugin.Interfaces;
using xafplugin.Modules;

namespace xafplugin.ViewModels
{
    /// <summary>
    /// _viewModel for creating a CASE expression that will be appended as a derived column.
    /// </summary>
    public class WizardSQLCaseViewModel : ViewModelBase
    {
        private string _customTableName;
        private readonly ObservableCollection<string> _columns;
        private readonly ObservableCollection<string> _suggestions;
        private bool _isSuggestionsVisible;
        private string _sqlText;
        private IMessageBoxService _localDialog; // fallback if base _dialog is null

        private IMessageBoxService Dialog => _dialog ?? (_localDialog ?? (_localDialog = new MessageBoxService()));

        public KeyValuePair<string, string> ResultSql { get; set; } = new KeyValuePair<string, string>();

        /// <summary>
        /// Raised when the window should be closed. Bool indicates confirm (true) or cancel (false).
        /// </summary>
        public event EventHandler<bool> RequestClose;

        public WizardSQLCaseViewModel(IEnumerable<string> columns = null) : base()
        {
            _columns = new ObservableCollection<string>(columns ?? Enumerable.Empty<string>());
            _suggestions = new ObservableCollection<string>();
            OkCommand = new RelayCommand(_ => ExecuteOk());
            CancelCommand = new RelayCommand(_ => OnRequestClose(false));
        }

        public WizardSQLCaseViewModel(IEnvironmentService environment, ISettingsProvider settingsProvider, IMessageBoxService dialog, IEnumerable<string> columns = null)
            : base(environment, settingsProvider, dialog)
        {
            _columns = new ObservableCollection<string>(columns ?? Enumerable.Empty<string>());
            _suggestions = new ObservableCollection<string>();
            OkCommand = new RelayCommand(_ => ExecuteOk());
            CancelCommand = new RelayCommand(_ => OnRequestClose(false));
        }

        public string CustomTableName
        {
            get => _customTableName;
            set
            {
                if (_customTableName != value)
                {
                    _customTableName = value;
                    OnPropertyChanged(nameof(CustomTableName));
                }
            }
        }

        public string SqlText
        {
            get => _sqlText;
            set
            {
                if (_sqlText != value)
                {
                    _sqlText = value;
                    OnPropertyChanged(nameof(SqlText));
                }
            }
        }

        public ObservableCollection<string> Columns => _columns;
        public ObservableCollection<string> Suggestions => _suggestions;

        public bool IsSuggestionsVisible
        {
            get => _isSuggestionsVisible;
            set
            {
                if (_isSuggestionsVisible != value)
                {
                    _isSuggestionsVisible = value;
                    OnPropertyChanged(nameof(IsSuggestionsVisible));
                }
            }
        }

        public ICommand OkCommand { get; }
        public ICommand CancelCommand { get; }

        private void OnRequestClose(bool result) => RequestClose?.Invoke(this, result);

        private void ExecuteOk()
        {
            if (IsValid())
            {
                ResultSql = new KeyValuePair<string, string>(_customTableName, BuildResultSql());
                OnRequestClose(true);
            }
        }

        private string BuildResultSql()
        {
            var text = _sqlText ?? string.Empty;

            var sql =
                "SELECT *,\r\n" +
                text + "\r\n" +
                $"END AS {_customTableName}\r\n" +
                "FROM PLACEHOLDER;";

            sql = sql.Replace("{optional}", "");
            sql = Regex.Replace(sql, @"""""(.*?)""""", "\"$1\"", RegexOptions.None, TimeSpan.FromMilliseconds(100));
            sql = Regex.Replace(sql, @"""(.*?)""", "'$1'", RegexOptions.None, TimeSpan.FromMilliseconds(100));
            sql = Regex.Replace(sql, @"^\s*\r?\n", "", RegexOptions.Multiline, TimeSpan.FromMilliseconds(100));
            return sql;
        }

        private bool IsValid()
        {
            if (string.IsNullOrWhiteSpace(_customTableName))
            {
                Dialog.ShowWarning("No column name specified.");
                return false;
            }

            var validRegex = new Regex(@"^[A-Za-z_][A-Za-z0-9_]*$", RegexOptions.None, TimeSpan.FromMilliseconds(100));
            if (!validRegex.IsMatch(_customTableName))
            {
                Dialog.ShowError("Invalid name. Only letters, digits and underscores are allowed and it must start with a letter or underscore.");
                return false;
            }

            if (_columns.Any(c => c.Equals(_customTableName, StringComparison.OrdinalIgnoreCase)))
            {
                Dialog.ShowError("The name already exists among existing columns.");
                return false;
            }

            var currentSql = _sqlText ?? string.Empty;
            if (currentSql.IndexOf("END", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Dialog.ShowError("END is added automatically.");
                return false;
            }

            int whenCount = CountOccurrences(currentSql, "WHEN");
            int thenCount = CountOccurrences(currentSql, "THEN");
            if (whenCount != thenCount)
            {
                Dialog.ShowError($"The number of 'WHEN' ({whenCount}) must equal the number of 'THEN' ({thenCount}).");
                return false;
            }

            if (!SqliteHelper.IsSyntaxValid(BuildResultSql()))
            {
                Dialog.ShowError("The SQL query is invalid. Check the syntax.");
                return false;
            }

            return true;
        }

        private int CountOccurrences(string text, string pattern)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(pattern))
                return 0;

            return Regex.Matches(text, Regex.Escape(pattern), RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(100)).Count;
        }
    }
}
