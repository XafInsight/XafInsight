using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.ObjectModel;
using System.IO;
using xafplugin.Interfaces;
using xafplugin.Modules;

namespace xafplugin.Helpers
{
    /// <summary>
    /// Beheert instellingen per bestand via afzonderlijke JSON-bestanden in een gebruikersmap.
    /// </summary>
    public class SettingsProvider : ISettingsProvider
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly string _settingsFolder;

        public SettingsProvider()
        {
            _settingsFolder = ResolveStorageFolder();
        }

        private string ResolveStorageFolder()
        {
            var settingsdir = Globals.ThisAddIn?.Config?.SettingsPath;
            if (string.IsNullOrWhiteSpace(settingsdir))
            {
                throw new InvalidOperationException("Instellingenpad is niet geconfigureerd.");
            }

            return settingsdir;
        }

        public string GetFileKey(string filePath)
        {
            return UniqueFileName.GetFileHash(filePath);
        }

        public FileSettings Get(string fileKey)
        {
            try
            {
                if (!IsValidFileKey(fileKey))
                {
                    throw new InvalidOperationException($"Ongeldige file-key: {fileKey}");
                }

                string path = UniqueFileName.CombinePathAndName(fileKey, _settingsFolder, EFileType.Json);
                if (!File.Exists(path))
                {
                    _logger.Info($"Geen instellingenbestand voor {fileKey}, nieuw object aangemaakt.");
                    return new FileSettings();
                }

                string json = File.ReadAllText(path);
                var settings = JsonConvert.DeserializeObject<FileSettings>(json) ?? new FileSettings();

                if (settings.ExportDefinitions is ObservableCollection<ExportDefinition> coll)
                {
                    coll.CollectionChanged -= (s, e) => Save(fileKey, settings);
                    coll.CollectionChanged += (s, e) => Save(fileKey, settings);
                }

                _logger.Debug($"Instellingen geladen voor: {fileKey}");
                return settings;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"Fout bij laden van instellingen voor {fileKey}");
                return new FileSettings();
            }
        }

        public void Set(string fileKey, Action<FileSettings> updateAction)
        {
            if (!IsValidFileKey(fileKey))
            {
                throw new InvalidOperationException($"Ongeldige file-key: {fileKey}");
            }

            var settings = Get(fileKey);
            updateAction(settings);
            Save(fileKey, settings);
            _logger.Info($"Instellingen bijgewerkt voor: {fileKey}");
        }

        public void Save(string fileKey, FileSettings settings)
        {
            try
            {
                if (!IsValidFileKey(fileKey))
                {
                    throw new InvalidOperationException($"Ongeldige file-key: {fileKey}");
                }

                string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                string path = UniqueFileName.CombinePathAndName(fileKey, _settingsFolder, EFileType.Json);
                File.WriteAllText(path, json);
                _logger.Info($"Instellingen opgeslagen voor: {fileKey} -> {path}");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"Fout bij opslaan van instellingen voor {fileKey}");
            }
        }

        public void Reset(string fileKey)
        {
            try
            {
                if (!IsValidFileKey(fileKey))
                {
                    throw new InvalidOperationException($"Ongeldige file-key: {fileKey}");
                }

                string path = UniqueFileName.CombinePathAndName(_settingsFolder, fileKey, EFileType.Json);
                if (File.Exists(path))
                {
                    File.Delete(path);
                    _logger.Info($"Instellingen verwijderd voor: {fileKey}");
                }
                else
                {
                    _logger.Warn($"Geen instellingenbestand om te verwijderen voor: {fileKey}");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"Fout bij verwijderen van instellingenbestand voor: {fileKey}");
            }
        }

        public void ResetAll()
        {
            try
            {
                foreach (var file in Directory.GetFiles(_settingsFolder, "*.json"))
                {
                    File.Delete(file);
                }
                _logger.Warn("Alle instellingen zijn gewist.");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Fout bij resetten van alle instellingen.");
            }
        }

        private bool IsValidFileKey(string fileKey) =>
    !string.IsNullOrEmpty(fileKey) &&
    fileKey.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) < 0 &&
    System.Text.RegularExpressions.Regex.IsMatch(
        fileKey,
        @"^[a-f0-9]{64}$",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(100)
    );

    }
}
