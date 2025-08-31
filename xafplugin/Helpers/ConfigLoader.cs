using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xafplugin.Modules;

namespace xafplugin.Helpers
{
    public static class ConfigLoader
    {
        private const string Company = "XAFInsight";
        private const string Product = "ExcelPlugin";
        private static readonly string EnvVar = (Company + "_" + Product + "_CONFIG_PATH").ToUpperInvariant();

        public static AppConfig Load()
        {
            var s = Defaults();

            // 1) ENV
            var envPath = Environment.GetEnvironmentVariable(EnvVar, EnvironmentVariableTarget.User)
                        ?? Environment.GetEnvironmentVariable(EnvVar, EnvironmentVariableTarget.Process);
            MergeIfExists(s, envPath);

            // 2) HKCU
            MergeIfExists(s, ReadRegString(RegistryHive.CurrentUser, $@"Software\{Company}\{Product}", "ConfigPath"));

            // 3) User roaming JSON
            MergeIfExists(s, UserConfigPath());

            // Normaliseer + zorg voor schrijfbare paden in het user-profiel
            s.LogPath = EnsureWriteableDir(s.LogPath, DefaultLogPath());
            s.TempDatabasePath = EnsureWriteableDir(s.TempDatabasePath, Path.GetTempPath());
            s.ConfigPath = EnsureWriteableDir(s.ConfigPath, DefaultConfigPath());
            s.SettingsPath = EnsureWriteableDir(s.SettingsPath, DefaultSettingsPath());

            return s;
        }

        public static string UserConfigPath()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData); // Roaming
            return Path.Combine(appData, Company, Product, "config.json");
        }

        private static AppConfig Defaults() => new AppConfig
        {
            LogPath = DefaultLogPath(),
            TempDatabasePath = Path.GetTempPath(),
            ConfigPath = DefaultConfigPath(),
            SettingsPath = DefaultSettingsPath()

        };

        private static string DefaultLogPath()
        {
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(local, Company, Product, "Logs");
        }

        private static string DefaultConfigPath()
        {
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(local, Company, Product, "Config");
        }

        private static string DefaultSettingsPath()
        {
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(local, Company, Product, "Settings");
        }


        private static void MergeIfExists(AppConfig target, string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
                var json = File.ReadAllText(path);
                var incoming = JsonConvert.DeserializeObject<AppConfig>(json);
                if (incoming == null) return;

                if (!string.IsNullOrWhiteSpace(incoming.LogPath)) target.LogPath = incoming.LogPath;
                if (!string.IsNullOrWhiteSpace(incoming.TempDatabasePath)) target.TempDatabasePath = incoming.TempDatabasePath;
                if (!string.IsNullOrWhiteSpace(incoming.ConfigPath)) target.ConfigPath = incoming.ConfigPath;
                if (!string.IsNullOrWhiteSpace(incoming.SettingsPath)) target.SettingsPath = incoming.SettingsPath;
                if (incoming.ConfigVersion > 0) target.ConfigVersion = incoming.ConfigVersion;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"Config merge failed for '{path}': {ex}");
            }
        }

        private static string ReadRegString(RegistryHive hive, string subKey, string name)
        {
            try
            {
                // Prefer 64-bit view on 64-bit OS, else 32-bit
                var firstView = Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;
                using (var baseKey = RegistryKey.OpenBaseKey(hive, firstView))
                using (var key = baseKey.OpenSubKey(subKey))
                {
                    var val = key?.GetValue(name) as string;
                    if (!string.IsNullOrEmpty(val)) return val;
                }

                // Fallback: if we first tried 64-bit, also check 32-bit
                if (Environment.Is64BitOperatingSystem && firstView == RegistryView.Registry64)
                {
                    using (var baseKey32 = RegistryKey.OpenBaseKey(hive, RegistryView.Registry32))
                    using (var key32 = baseKey32.OpenSubKey(subKey))
                    {
                        return key32?.GetValue(name) as string;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine("ReadRegString failed: " + ex);
            }
            return null;
        }

        private static string EnsureWriteableDir(string requested, string fallback)
        {
            var p = string.IsNullOrWhiteSpace(requested) ? fallback : requested;
            try { Directory.CreateDirectory(p); }
            catch { p = fallback; Directory.CreateDirectory(p); }
            return p;
        }
    }
}
