using NLog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Input;
using xafplugin.Database;
using xafplugin.Interfaces;
using xafplugin.Modules;

namespace xafplugin.ViewModels
{
    public class WizardFilterSQLViewModel : ViewModelBase
    {
        private ObservableCollection<string> _columns;
        private ObservableCollection<string> _suggestions;
        private bool _isSuggestionsVisible;
        private string _sqlText;
        private string _name;

        public string ResultSql { get; set; }

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                }
            }
        }

        /// <summary>
        /// Raised when the window should be closed.  The boolean result indicates
        /// whether the user confirmed (true) or cancelled (false).
        /// </summary>
        public event EventHandler<bool> RequestClose;

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlCaseQueryViewmodel"/> class.
        /// Initializes the collections and commands.  Columns can be populated by
        /// passing a list of strings via the constructor.
        /// </summary>
        /// <param name="columns">Optional list of column names to populate the
        /// Columns collection.</param>
        public WizardFilterSQLViewModel(IEnumerable<string> columns = null)
        {
            _columns = new ObservableCollection<string>(columns ?? Enumerable.Empty<string>());
            _suggestions = new ObservableCollection<string>();
            OkCommand = new RelayCommand(_ => ExecuteOk());
            CancelCommand = new RelayCommand(_ => OnRequestClose(false));
        }

        public WizardFilterSQLViewModel(IEnvironmentService environment, ISettingsProvider settingsProvider, IMessageBoxService dialog, IEnumerable<string> columns = null) : base(environment, settingsProvider, dialog)
        {
            _columns = new ObservableCollection<string>(columns ?? Enumerable.Empty<string>());
            _suggestions = new ObservableCollection<string>();
            OkCommand = new RelayCommand(_ => ExecuteOk());
            CancelCommand = new RelayCommand(_ => OnRequestClose(false));
        }

        /// <summary>
        /// Gets or sets the entire SQL expression as plain text.  This property is
        /// optional because the RichTextBox stores its own FlowDocument.  It can be
        /// used to persist the query if desired.
        /// </summary>
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

        /// <summary>
        /// Gets the collection of available column names.  Double‑clicking an entry
        /// inserts it into the editor at the caret position.
        /// </summary>
        public ObservableCollection<string> Columns => _columns;

        /// <summary>
        /// Gets the collection of current suggestions.  Suggestions are filtered in
        /// the code‑behind based on the word currently being typed.  The suggestions
        /// list is displayed only when this collection contains items.
        /// </summary>
        public ObservableCollection<string> Suggestions => _suggestions;

        /// <summary>
        /// Gets or sets a value indicating whether the suggestion list should be
        /// visible.  The XAML binds this property to the Visibility of the
        /// suggestion ListBox via a BooleanToVisibilityConverter.
        /// </summary>
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

        /// <summary>
        /// Command bound to the OK button.  Closes the dialog with a true result.
        /// </summary>
        public ICommand OkCommand { get; }

        /// <summary>
        /// Command bound to the Cancel button.  Closes the dialog with a false result.
        /// </summary>
        public ICommand CancelCommand { get; }

        private void OnRequestClose(bool result)
        {
            RequestClose?.Invoke(this, result);
        }

        private void ExecuteOk()
        {
            if (IsValid())
            {
                ResultSql = resultString();

                OnRequestClose(true);
            }
        }

        private string resultString()
        {
            string SQLCasestepstring =
                     "SELECT * FROM PLACEHOLDER\r\n" +
                     _sqlText + "\r\n";

            SQLCasestepstring = SQLCasestepstring.Replace("{optional}", "");
            // 1. Eerst dubbele dubbele quotes naar enkele dubbele quotes
            SQLCasestepstring = Regex.Replace(SQLCasestepstring, @"""""(.*?)""""", "\"$1\"", RegexOptions.None, TimeSpan.FromMilliseconds(100));

            // 2. Daarna enkele dubbele quotes naar enkele aanhalingstekens
            SQLCasestepstring = Regex.Replace(SQLCasestepstring, @"""(.*?)""", "'$1'", RegexOptions.None, TimeSpan.FromMilliseconds(100));

            SQLCasestepstring = Regex.Replace(SQLCasestepstring, @"^\s*\r?\n", "", RegexOptions.Multiline, TimeSpan.FromMilliseconds(100));

            return SQLCasestepstring;
        }

        private bool IsValid()
        {

            if (!SqliteHelper.IsSyntaxValid(resultString()))
            {
                MessageBox.Show("De SQL-query is ongeldig. Controleer de syntax.", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

    }
}
