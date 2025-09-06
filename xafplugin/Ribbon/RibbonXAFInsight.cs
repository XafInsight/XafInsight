using Newtonsoft.Json;
using NLog;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Interop;
using xafplugin.Form;
using xafplugin.Helpers;
using xafplugin.Modules;
using Office = Microsoft.Office.Core;

namespace xafplugin.Ribbon
{
    [ComVisible(true)]
    public class RibbonXAFInsight : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        public Office.IRibbonUI Ribbon { get { return _ribbon; } }
        public static RibbonXAFInsight Instance { get; private set; }

        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly MessageBoxService _dialog = new MessageBoxService();

        private Window _connectWindow;
        private Window _relationWindow;
        private Window _tableWindow;
        private Window _columnMappingWindow;
        private Window _exportWindow;

        private const string XmlPathPropertyName = "XmlPath";

        private static EnvironmentService _env = new EnvironmentService();

        public RibbonXAFInsight() { }

        #region Button click handlers

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060", Justification = "Office Ribbon callback signature")]
        public void ButtonConnect_Click(Office.IRibbonControl control)
        {
            _logger.Info("ButtonConnect_Click started");
            ShowOrActivate(_connectWindow,
               () => CreateWindow("Open file", new OpenFileWindow(), modeless: true, showInTaskbar: true),
               modal: true,
               w => _connectWindow = w);

            _ribbon?.Invalidate();
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060", Justification = "Office Ribbon callback signature")]
        public void ButtonDisconnect_Click(Office.IRibbonControl control)
        {
            _logger.Info("ButtonDisconnect_Click started");

            Globals.ThisAddIn.FilePath = null;
            _env.FileHash = null;
            Globals.ThisAddIn.TempDbPath = null;
            _ribbon?.Invalidate();
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060", Justification = "Office Ribbon callback signature")]
        public void ButtonExportTable_Click(Office.IRibbonControl control)
        {
            ShowOrActivate(_exportWindow,
                () => CreateWindow("Export Table", new ExportTableControl(), modeless: true, showInTaskbar: true),
                modal: false,
                w => _exportWindow = w);
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060", Justification = "Office Ribbon callback signature")]
        public void ButtonRelation_Click(Office.IRibbonControl control)
        {
            ShowOrActivate(_relationWindow,
                () => CreateWindow("Relations", new RelationsControl(), modeless: false),
                modal: true,
                w => _relationWindow = w);
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060", Justification = "Office Ribbon callback signature")]
        public void ButtonTabel_Click(Office.IRibbonControl control)
        {
            ShowOrActivate(_tableWindow,
                () => CreateWindow("Select Table/Column", new TableControl(), modeless: false),
                modal: true,
                w => _tableWindow = w);
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060", Justification = "Office Ribbon callback signature")]
        public void ButtonColumnMapping_Click(Office.IRibbonControl control)
        {
            ShowOrActivate(_columnMappingWindow,
                () => CreateWindow("Rename Columns", new ColumnMappingControl(), modeless: false),
                modal: true,
                w => _columnMappingWindow = w);
        }

        public void ButtonExportSettings_Click(Office.IRibbonControl control)
        {
            var settingsProvider = new SettingsProvider();

            var settings = settingsProvider.Get(_env.FileHash);
            if (settings == null)
            {
                _logger.Warn("No settings found for current file.");
                _dialog.ShowWarning("No settings available for the current file.");
                return;
            }

            string json = JsonConvert.SerializeObject(settings, Formatting.Indented);

            try
            {
                // Suggest a filename based on the current file (if any) or the settings name.
                var sourceFile = Globals.ThisAddIn.FilePath;
                var baseName = !string.IsNullOrWhiteSpace(sourceFile)
                    ? Path.GetFileNameWithoutExtension(sourceFile)
                    : (string.IsNullOrWhiteSpace(settings.Name) ? "settings" : settings.Name);

                // Sanitize filename
                foreach (var c in Path.GetInvalidFileNameChars())
                    baseName = baseName.Replace(c, '_');

                var defaultFileName = baseName + ".settings.json";

                using (var dlg = new SaveFileDialog
                {
                    Title = "Export Settings",
                    Filter = "Settings JSON (*.json)|*.json|All files (*.*)|*.*",
                    FileName = defaultFileName,
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    OverwritePrompt = true,
                    AddExtension = true
                })
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        File.WriteAllText(dlg.FileName, json);
                        _logger.Info("Settings exported to: {0}", dlg.FileName);
                        _dialog.Show("Settings exported successfully.", "Export", MessageBoxButtons.OK, MessageBoxIcon.None);
                    }
                    else
                    {
                        _logger.Info("Settings export canceled by user.");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to export settings.");
                _dialog.ShowError("An error occurred while exporting the settings.");
            }
        }

        public void ButtonImportSettings_Click(Office.IRibbonControl control)
        {
            try
            {
                using (var dlg = new OpenFileDialog
                {
                    Title = "Import Settings",
                    Filter = "Settings JSON (*.json)|*.json|All files (*.*)|*.*",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                })
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        string json = File.ReadAllText(dlg.FileName);
                        var fileSettings = JsonConvert.DeserializeObject<FileSettings>(json);
                        if (fileSettings == null)
                        {
                            _logger.Warn("Imported settings file is invalid or empty: {0}", dlg.FileName);
                            _dialog.ShowWarning("The selected settings file is invalid.");
                            return;
                        }
                        var settingsProvider = new SettingsProvider();
                        settingsProvider.Save(_env.FileHash, fileSettings);
                        _logger.Info("Settings imported from: {0}", dlg.FileName);
                        _dialog.Show("Settings imported successfully.", "Import", MessageBoxButtons.OK, MessageBoxIcon.None);
                    }
                    else
                    {
                        _logger.Info("Settings import canceled by user.");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to import settings.");
                _dialog.ShowError("An error occurred while importing the settings.");
            }
        }

        public void ButtonHelp_Click(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Process.Start("https://github.com/XafInsight/XafInsight/wiki");
            }
            catch (Exception ex)
            {
                _dialog.Show($"Unable to open help documentation: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ButtonReportIssue_Click(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Process.Start("https://github.com/XafInsight/XafInsight/issues");
            }
            catch (Exception ex)
            {
                _dialog.Show($"Unable to open issue reporting page: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ButtonExportSQLite_Click(Office.IRibbonControl control)
        {
            var dbPath = _env.DatabasePath;
            if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath))
            {
                _logger.Warn("No active database to export.");
                _dialog.ShowWarning("No active database to export.");
                return;
            }
            try
            {
                using (var dlg = new SaveFileDialog
                {
                    Title = "Export SQLite Database",
                    Filter = "SQLite Database (*.sqlite;*.db)|*.sqlite;*.db|All files (*.*)|*.*",
                    FileName = "exported_database.sqlite",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    OverwritePrompt = true,
                    AddExtension = true
                })
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        File.Copy(dbPath, dlg.FileName, overwrite: false);
                        _logger.Info("Database exported to: {0}", dlg.FileName);
                        _dialog.Show("Database exported successfully.", "Export", MessageBoxButtons.OK, MessageBoxIcon.None);
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

        #endregion

        #region Ribbon UI Callbacks

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060", Justification = "Office Ribbon callback signature")]
        public bool ButtonConnect_GetEnabled(Office.IRibbonControl control)
        {
            var dbPath = _env.DatabasePath;
            return string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath);
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060", Justification = "Office Ribbon callback signature")]
        public bool ButtonDisconnect_GetEnabled(Office.IRibbonControl control)
        {
            var dbPath = _env.DatabasePath;
            return !string.IsNullOrEmpty(dbPath) && File.Exists(dbPath);
        }

        public bool ButtonExport_GetEnabled(Office.IRibbonControl control)
        {
            var dbPath = _env.DatabasePath;
            return !string.IsNullOrEmpty(dbPath) && File.Exists(dbPath);
        }

        #endregion

        #region Button Images
        /// <summary>
        /// deze methodes geven de afbeeldingen voor de knoppen in de ribbon.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "Office Ribbon callback signature")]
        public stdole.IPictureDisp GetConnectButtonImage(Office.IRibbonControl control)
        {
            var bmp = Properties.Resources.LoadAuditfile;
            return PictureConverter.ImageToPictureDisp(bmp);
        }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "Office Ribbon callback signature")]
        public stdole.IPictureDisp GetCloseButtonImage(Office.IRibbonControl control)
        {
            var bmp = Properties.Resources.Disconnect;
            return PictureConverter.ImageToPictureDisp(bmp);
        }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "Office Ribbon callback signature")]
        public stdole.IPictureDisp GetRelationButtonImage(Office.IRibbonControl control)
        {
            var bmp = Properties.Resources.Relations;
            return PictureConverter.ImageToPictureDisp(bmp);
        }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "Office Ribbon callback signature")]
        public stdole.IPictureDisp GetTableButtonImage(Office.IRibbonControl control)
        {
            var bmp = Properties.Resources.Tables;
            return PictureConverter.ImageToPictureDisp(bmp);
        }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "Office Ribbon callback signature")]
        public stdole.IPictureDisp GetMappingButtonImage(Office.IRibbonControl control)
        {
            var bmp = Properties.Resources.ColumnMapping;
            return PictureConverter.ImageToPictureDisp(bmp);
        }
        public stdole.IPictureDisp GetExportButtonImage(Office.IRibbonControl control)
        {
            var bmp = Properties.Resources.ExportTable;
            return PictureConverter.ImageToPictureDisp(bmp);
        }

        public stdole.IPictureDisp GetExportSettingsImage(Office.IRibbonControl control)
        {
            var bmp = Properties.Resources.export;
            return PictureConverter.ImageToPictureDisp(bmp);
        }

        public stdole.IPictureDisp GetImportSettingsImage(Office.IRibbonControl control)
        {
            var bmp = Properties.Resources.import;
            return PictureConverter.ImageToPictureDisp(bmp);
        }

        public stdole.IPictureDisp GetExportSQLiteImage(Office.IRibbonControl control)
        {
            var bmp = Properties.Resources.export;
            return PictureConverter.ImageToPictureDisp(bmp);
        }

        public stdole.IPictureDisp GetHelpButtonImage(Office.IRibbonControl control)
        {
            var bmp = Properties.Resources.help;
            return PictureConverter.ImageToPictureDisp(bmp);
        }

        public stdole.IPictureDisp GetReportIssueButtonImage(Office.IRibbonControl control)
        {
            var bmp = Properties.Resources.issue;
            return PictureConverter.ImageToPictureDisp(bmp);
        }
        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            var xml = GetResourceText("xafplugin.Ribbon.RibbonXAFInsight.xml");
            if (xml == null)
            {
                throw new InvalidOperationException("Ribbon XML resource not found.");
            }
            return xml;
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            Instance = this;
            _logger.Info("RibbonXAFInsight loaded.");
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var match = asm
                .GetManifestResourceNames()
                .FirstOrDefault(n => string.Equals(resourceName, n, StringComparison.OrdinalIgnoreCase));

            if (match == null) return null;

            var stream = asm.GetManifestResourceStream(match);
            if (stream == null) return null;

            using (var reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        private Window CreateWindow(string title, object content, bool modeless, bool showInTaskbar = false)
        {
            var window = new Window
            {
                Title = title,
                Content = content,
                SizeToContent = SizeToContent.WidthAndHeight,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = showInTaskbar,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Icon = BitmapImageCoverter.ByteArrayToIcon(Properties.Resources.empty16)
            };

            new WindowInteropHelper(window) { Owner = new IntPtr(Globals.ThisAddIn.Application.Hwnd) };

            return window;
        }

        private void ShowOrActivate(Window currentWindow, Func<Window> factory, bool modal, Action<Window> assign)
        {
            try
            {
                if (currentWindow != null && currentWindow.IsVisible)
                {
                    currentWindow.Activate();
                    return;
                }

                var window = factory();
                assign(window);
                window.Closed += (s, e) => assign(null);

                if (modal)
                {
                    window.ShowDialog();
                }
                else
                {
                    ElementHost.EnableModelessKeyboardInterop(window);
                    window.Show();
                }

                _logger.Info("Window '{0}' displayed (modal={1}).", window.Title, modal);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error displaying window.");
                _dialog.ShowError("An error occurred while opening the window.");
            }
        }




        #endregion
    }
}
