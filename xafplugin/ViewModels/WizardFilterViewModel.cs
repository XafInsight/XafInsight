using NLog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using xafplugin.Database;
using xafplugin.Interfaces;

namespace xafplugin.ViewModels
{
    public class WizardFilterViewModel : ViewModelBase
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private ObservableCollection<string> _columns;
        public event EventHandler<bool> RequestClose;
        private string _sqlText;
        private string _filterName;

        public string FilterName
        {
            get => _filterName;
            set
            {
                if (_filterName != value)
                {
                    _filterName = value;
                }
            }
        }

        public string SqlText
        {
            get => _sqlText;
        }

        public ObservableCollection<string> Columns
        {
            get => _columns;
            set
            {
                if (_columns != value)
                {
                    _columns = value;
                    OnPropertyChanged(nameof(Columns));
                }
            }
        }


        public WizardFilterViewModel(IEnumerable<string> columns = null) : base()
        {
            _logger.Info("Initialisatie van ColumnMappingViewModel gestart.");
            _columns = new ObservableCollection<string>(columns ?? Enumerable.Empty<string>());

        }
        public WizardFilterViewModel(IEnvironmentService environment , ISettingsProvider settingsProvider, IMessageBoxService dialog, IEnumerable<string> columns = null) : base(environment, settingsProvider, dialog)
        {
            _logger.Info("Initialisatie van ColumnMappingViewModel gestart.");
            _columns = new ObservableCollection<string>(columns ?? Enumerable.Empty<string>());

        }


        public void OnRequestClose(bool result)
        {
            RequestClose?.Invoke(this, result);
        }

        public bool SetSQLSyntax(string filter)
        {
           
            var result = resultString(filter);

            if (!IsValid(result))
            {
                return false;
            }
            if (!SqliteHelper.IsSyntaxValid(result))
            {
                return false;
            }
          
            _sqlText = result;
            return true;
        }


        private bool IsValid(string sqlText)
        {
            if(sqlText.Contains("{") || sqlText.Contains("}"))
            {
                return false;
            }
            return true;
        }

        private string resultString(string filter)
        {
            string SQLCasestepstring =
                    "SELECT * FROM PLACEHOLDER\r\n" +
                    "WHERE " + filter + "\r\n";

            // 1. Eerst dubbele dubbele quotes naar enkele dubbele quotes
            SQLCasestepstring = Regex.Replace(SQLCasestepstring, @"""""(.*?)""""", "\"$1\"", RegexOptions.None, TimeSpan.FromMilliseconds(100));

            // 2. Daarna enkele dubbele quotes naar enkele aanhalingstekens
            SQLCasestepstring = Regex.Replace(SQLCasestepstring, @"""(.*?)""", "'$1'", RegexOptions.None, TimeSpan.FromMilliseconds(100));

            SQLCasestepstring = Regex.Replace(SQLCasestepstring, @"^\s*\r?\n", "", RegexOptions.Multiline, TimeSpan.FromMilliseconds(100));

            return SQLCasestepstring;
        }
    }




}
