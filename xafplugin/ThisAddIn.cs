using NLog;
using SQLitePCL;
using System;
using System.IO;
using System.Windows;
using xafplugin.Helpers;
using xafplugin.Modules;
using xafplugin.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace xafplugin
{

    public partial class ThisAddIn
    {
        private AppConfig _config;
        public AppConfig Config
        {
            get { return _config; }
        }

        public string TempDbPath { get; set; }
        public string FilePath { get; set; }

        public string FileHash { get; set; }



        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Deze functie wordt aangeroepen wanneer de add-in wordt gestart.
        /// De functie configureert de _logger, selt een tijdelijk bestandspad in voor de database, koppelt een event handler voor het wisselen van vensters in Excel,
        /// daarnaast verwijdert het oude tijdelijke bestanden die mogelijk zijn achtergebleven.
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Batteries_V2.Init();
                _config = ConfigLoader.Load();
                LoggerSetup.Configure();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fout bij het laden van de configuratie of het instellen van de logger: " + ex.Message + "\nneem contact op met applicatie beheerder", "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            _logger.Info("Plugin gestart.");

            // Koppel een event handler om de ribbon te verversen
            this.Application.WindowActivate += Application_WindowActivate;
            this.Application.SheetSelectionChange += Application_SheetSelectionChange;
            this.Application.SheetActivate += Application_SheetActivate;
            this.Application.WorkbookActivate += Application_WorkbookActivate;
            this.Application.WorkbookDeactivate += Application_WorkbookDeactivate;



            // verwijder dit naar window. 
            string tempDir = Path.GetTempPath();
            string tempFile = Path.Combine(tempDir, "XafInsight_" + Guid.NewGuid().ToString() + ".sqlite");
            TempDbPath = tempFile;
        }

        /// <summary>
        /// deze functie maakt een RibbonExtensibility object aan dat de RibbonXAFInsight klasse instantieert.
        /// </summary>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonXAFInsight();
        }

        /// <summary>
        /// deze functie wordt aangeroepen wanneer het actieve werkboek of venster in Excel verandert.
        /// </summary>
        private void Application_WindowActivate(Excel.Workbook wb, Excel.Window wn)
        {
            RibbonXAFInsight.Instance?.Ribbon?.Invalidate();
        }

        /// <summary>
        /// deze functie wordt aangeroepen wanneer de add-in wordt afgesloten.
        /// De functie verwijdert het tijdelijke bestand dat is aangemaakt bij de start van de add-in.
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            TempDatabaseClean.CleanOldTempDatabases(_config.TempDatabasePath, _config.RemoveTempDatabaseAfterDays);
        }

        private void Application_SheetSelectionChange(object sh, Excel.Range target)
        {
            RibbonXAFInsight.Instance?.Ribbon?.Invalidate();
        }

        private void Application_SheetActivate(object sh)
        {
            RibbonXAFInsight.Instance?.Ribbon?.Invalidate();
        }

        private void Application_WorkbookActivate(Excel.Workbook wb)
        {
            RibbonXAFInsight.Instance?.Ribbon?.Invalidate();
        }

        private void Application_WorkbookDeactivate(Excel.Workbook wb)
        {
            RibbonXAFInsight.Instance?.Ribbon?.Invalidate();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
