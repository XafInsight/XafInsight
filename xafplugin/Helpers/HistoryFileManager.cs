using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NLog;

namespace xafplugin.Helpers
{
    /// <summary>
    /// Provides safe read/write operations for a History.txt file inside a given folder.
    /// Appends always puts newest entries at the top. If an entry already exists it is
    /// removed from its previous position and re-added at the top (de-duplication).
    /// When maxLines &gt; 0 the file is trimmed (oldest lines at the bottom are removed).
    /// </summary>
    public static class HistoryFileManager
    {
        private const string HistoryFileName = "History.txt";
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private static readonly object _syncRoot = new object();

        /// <summary>
        /// Gets the full path to History.txt for the specified folder (does not create anything).
        /// Returns null if the folderPath is invalid.
        /// </summary>
        public static string GetHistoryFilePath(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
                return null;

            try
            {
                var fullFolder = Path.GetFullPath(folderPath);
                return Path.Combine(fullFolder, HistoryFileName);
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, "Failed to combine history file path for folder: {0}", folderPath);
                return null;
            }
        }

        /// <summary>
        /// Adds a single line at the top (removing any existing duplicate first).
        /// </summary>
        public static bool Append(string folderPath, string line, int maxLines, out string error)
        {
            if (line == null) line = string.Empty;
            return Append(folderPath, new[] { line }, maxLines, out error);
        }

        /// <summary>
        /// Adds multiple lines at the top (first element ends up as the first line).
        /// Existing duplicates (case-insensitive) are removed before insertion.
        /// </summary>
        public static bool Append(string folderPath, IEnumerable<string> lines, int maxLines, out string error)
        {
            error = null;
            if (!ValidateLines(ref lines, ref error))
                return false;

            if (!EnsureDirectoryAndFile(folderPath, out var fullPath, out error))
                return false;

            try
            {
                lock (_syncRoot)
                {
                    // Normalize incoming lines to non-null and trim ending CR/LF artifacts.
                    var incoming = lines
                        .Select(l => (l ?? string.Empty).TrimEnd('\r', '\n'))
                        .ToList();

                    // Read existing list (ignore failures - start fresh).
                    List<string> existing;
                    try
                    {
                        existing = File.Exists(fullPath)
                            ? File.ReadAllLines(fullPath).Select(s => s ?? string.Empty).ToList()
                            : new List<string>();
                    }
                    catch (Exception readEx)
                    {
                        _logger.Warn(readEx, "Failed reading existing history; starting with empty. Path: {0}", fullPath);
                        existing = new List<string>();
                    }

                    // Remove duplicates of incoming lines from existing (case-insensitive).
                    if (incoming.Count > 0 && existing.Count > 0)
                    {
                        var set = new HashSet<string>(incoming, StringComparer.OrdinalIgnoreCase);
                        // Keep only those not in incoming
                        existing = existing.Where(l => !set.Contains(l)).ToList();
                    }

                    // Combine: newest (incoming) first, then the filtered existing.
                    var combined = new List<string>(incoming.Count + existing.Count);
                    combined.AddRange(incoming);
                    combined.AddRange(existing);

                    // Enforce maximum (keep top portion).
                    if (maxLines > 0 && combined.Count > maxLines)
                    {
                        combined = combined.Take(maxLines).ToList();
                    }

                    // Atomic-ish write.
                    var tempPath = fullPath + ".tmp";
                    File.WriteAllLines(tempPath, combined);
                    File.Copy(tempPath, fullPath, true);
                    File.Delete(tempPath);
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "Failed to append to history file.";
                _logger.Error(ex, "{0} Path: {1}", error, fullPath);
                return false;
            }
        }

        /// <summary>
        /// Reads all lines (newest first). Succeeds with empty list if file does not exist.
        /// </summary>
        public static bool ReadAll(string folderPath, out List<string> lines, out string error)
        {
            lines = new List<string>();
            error = null;

            var fullPath = GetHistoryFilePath(folderPath);
            if (fullPath == null)
            {
                error = "Invalid folder path.";
                return false;
            }

            if (!File.Exists(fullPath))
                return true;

            try
            {
                lock (_syncRoot)
                {
                    lines = File.ReadAllLines(fullPath).ToList();
                }
                return true;
            }
            catch (Exception ex)
            {
                error = "Failed to read history file.";
                _logger.Error(ex, "{0} Path: {1}", error, fullPath);
                return false;
            }
        }

        /// <summary>
        /// Replaces entire file content.
        /// </summary>
        public static bool ReplaceAll(string folderPath, IEnumerable<string> lines, out string error)
        {
            error = null;
            if (!ValidateLines(ref lines, ref error))
                return false;

            if (!EnsureDirectoryAndFile(folderPath, out var fullPath, out error))
                return false;

            try
            {
                lock (_syncRoot)
                {
                    File.WriteAllLines(fullPath, lines.Select(l => l ?? string.Empty));
                }
                return true;
            }
            catch (Exception ex)
            {
                error = "Failed to replace history file content.";
                _logger.Error(ex, "{0} Path: {1}", error, fullPath);
                return false;
            }
        }

        /// <summary>
        /// Clears file content (creates file if missing).
        /// </summary>
        public static bool Clear(string folderPath, out string error)
        {
            error = null;
            if (!EnsureDirectoryAndFile(folderPath, out var fullPath, out error))
                return false;

            try
            {
                lock (_syncRoot)
                {
                    File.WriteAllText(fullPath, string.Empty);
                }
                return true;
            }
            catch (Exception ex)
            {
                error = "Failed to clear history file.";
                _logger.Error(ex, "{0} Path: {1}", error, fullPath);
                return false;
            }
        }

        private static bool EnsureDirectoryAndFile(string folderPath, out string fullPath, out string error)
        {
            error = null;
            fullPath = null;

            if (string.IsNullOrWhiteSpace(folderPath))
            {
                error = "Folder path is null or empty.";
                return false;
            }

            try
            {
                var fullFolder = Path.GetFullPath(folderPath);
                if (!Directory.Exists(fullFolder))
                    Directory.CreateDirectory(fullFolder);

                fullPath = Path.Combine(fullFolder, HistoryFileName);

                if (!File.Exists(fullPath))
                {
                    using (var fs = new FileStream(fullPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                    {
                        // create empty file
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                error = "Failed to prepare history file.";
                _logger.Error(ex, "{0} Folder: {1}", error, folderPath);
                return false;
            }
        }

        private static bool ValidateLines(ref IEnumerable<string> lines, ref string error)
        {
            if (lines == null)
                lines = Enumerable.Empty<string>();
            return true;
        }
    }
}