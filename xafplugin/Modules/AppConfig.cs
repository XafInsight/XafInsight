using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xafplugin.Modules
{
    public sealed class AppConfig
    {
        public string LogPath { get; set; }
        public string TempDatabasePath { get; set; }
        public string ConfigPath { get; set; }
        public string SettingsPath { get; set; }
        public int RemoveTempDatabaseAfterDays { get; set; } = 14;
        public int HistoryFileAmount { get; set; } = 15;
        public int ConfigVersion { get; set; } = 1;


    }
}
