using System;
using System.Collections.Generic;
using System.IO;
using NLog;

namespace xafplugin.Helpers
{
    public static class TempDatabaseClean
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Deletes temporary SQLite databases (pattern: XafInsight_*.sqlite) whose last activity
        /// (max of creation, last write and optionally last access) is older than retentionDays.
        /// </summary>
        /// <param name="tempDatabasePath">Directory containing temp databases.</param>
        /// <param name="retentionDays">Days to retain; values &lt; 0 coerced to 1.</param>
        /// <param name="useLastAccess">
        /// When true attempts to include LastAccessTimeUtc in the activity heuristic.
        /// Ignored if the value looks invalid or clearly not updating.
        /// </param>
        /// <param name="skipLocked">
        /// When true, skips files that appear locked (in use) to avoid deleting an active database.
        /// </param>
        public static IList<string> CleanOldTempDatabases(
            string tempDatabasePath,
            int retentionDays,
            bool useLastAccess = true,
            bool skipLocked = true)
        {
            var deleted = new List<string>();
            if (string.IsNullOrWhiteSpace(tempDatabasePath))
            {
                _logger.Warn("CleanOldTempDatabases: Provided path is null or empty.");
                return deleted;
            }

            try
            {
                if (!Directory.Exists(tempDatabasePath))
                {
                    _logger.Debug("CleanOldTempDatabases: Directory does not exist: {0}", tempDatabasePath);
                    return deleted;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "CleanOldTempDatabases: Failed to validate directory existence for path {0}", tempDatabasePath);
                return deleted;
            }

            if (retentionDays < 0)
            {
                _logger.Warn("CleanOldTempDatabases: Negative retentionDays ({0}) coerced to 1.", retentionDays);
                retentionDays = 1;
            }

            DateTime threshold = DateTime.UtcNow.AddDays(-retentionDays);

            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(tempDatabasePath, "XafInsight_*.sqlite", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "CleanOldTempDatabases: Failed to enumerate files in {0}", tempDatabasePath);
                return deleted;
            }

            foreach (var file in files)
            {
                try
                {
                    DateTime creationUtc;
                    DateTime lastWriteUtc;
                    DateTime? lastAccessUtc = null;

                    try
                    {
                        creationUtc = File.GetCreationTimeUtc(file);
                        lastWriteUtc = File.GetLastWriteTimeUtc(file);

                        if (useLastAccess)
                        {
                            var la = File.GetLastAccessTimeUtc(file);
                            // Filter out obviously invalid or default values
                            if (la.Year > 1980 && la <= DateTime.UtcNow.AddMinutes(5))
                                lastAccessUtc = la;
                        }
                    }
                    catch (FileNotFoundException)
                    {
                        continue;
                    }
                    catch (UnauthorizedAccessException uae)
                    {
                        _logger.Debug(uae, "CleanOldTempDatabases: Access denied reading timestamps for {0}", file);
                        continue;
                    }

                    // Heuristic last activity
                    DateTime lastActivityUtc = creationUtc > lastWriteUtc ? creationUtc : lastWriteUtc;
                    if (lastAccessUtc.HasValue && lastAccessUtc.Value > lastActivityUtc)
                        lastActivityUtc = lastAccessUtc.Value;

                    if (lastActivityUtc >= threshold)
                        continue;

                    if (skipLocked && IsFileLocked(file))
                    {
                        _logger.Debug("CleanOldTempDatabases: Skipped locked file {0}", file);
                        continue;
                    }

                    try
                    {
                        File.Delete(file);
                        deleted.Add(file);
                        _logger.Info("CleanOldTempDatabases: Deleted temp database {0} (last activity {1:u})", file, lastActivityUtc);
                    }
                    catch (FileNotFoundException)
                    {
                        // Race
                    }
                    catch (UnauthorizedAccessException uae)
                    {
                        _logger.Warn(uae, "CleanOldTempDatabases: Unauthorized to delete {0}", file);
                    }
                    catch (IOException ioex)
                    {
                        _logger.Debug(ioex, "CleanOldTempDatabases: IO error deleting {0}", file);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "CleanOldTempDatabases: Unexpected error processing file {0}", file);
                }
            }

            return deleted;
        }

        private static bool IsFileLocked(string path)
        {
            FileStream stream = null;
            try
            {
                // Try to open with exclusive access. If it succeeds, it's not locked.
                stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
                return false;
            }
            catch (IOException)
            {
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                // Treat as locked/inaccessible
                return true;
            }
            finally
            {
                if (stream != null)
                {
                    try { stream.Dispose(); } catch { /* ignore */ }
                }
            }
        }
    }
}
