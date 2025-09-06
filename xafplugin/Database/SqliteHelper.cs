using System;
using System.Data;
using System.IO;
using Microsoft.Data.Sqlite;

namespace xafplugin.Database
{
    /// <summary>
    /// SQLite helper:
    /// - Ensures database file exists.
    /// - Provides a long-lived read-only connection.
    /// - Creates short-lived write connections.
    /// - Can (lightly) validate single-statement SQL syntax without side effects.
    /// </summary>
    public sealed class SqliteHelper : IDisposable
    {
        private readonly string _dbPath;
        private SqliteConnection _readConnection;
        private SqliteConnection _writeConnection;
        private bool _disposed;

        /// <summary>
        /// Read-only connection (opened lazily). Caller must NOT dispose it.
        /// </summary>
        public SqliteConnection Connection
        {
            get
            {
                ThrowIfDisposed();
                if (_readConnection == null)
                {
                    _readConnection = CreateConnection(readOnly: true);
                }
                return _readConnection;
            }
        }

        public SqliteHelper(string dbPath)
        {
            if (string.IsNullOrWhiteSpace(dbPath))
                throw new ArgumentException("Database path required.", nameof(dbPath));

            _dbPath = Path.GetFullPath(dbPath);
            EnsureDatabaseExists();
        }

        /// <summary>
        /// Creates and opens a writable connection. Caller must dispose it.
        /// </summary>
        public SqliteConnection GetWriteConnection(bool enableForeignKeys = true)
        {
            ThrowIfDisposed();
            if (_writeConnection == null)
            {
                _writeConnection = CreateConnection(readOnly: false, enableForeignKeys: enableForeignKeys);
            }
            return _writeConnection;
        }

        /// <summary>
        /// Quick syntax validation (single statement). Returns false only for parse errors.
        /// </summary>
        public static bool IsSyntaxValid(string sql)
        {
            string ignoredMessage;
            return TryValidateSyntax(sql, out ignoredMessage);
        }

        /// <summary>
        /// Attempts to validate syntax. True = syntactically OK (even if runtime objects missing).
        /// </summary>
        public static bool TryValidateSyntax(string sql, out string message)
        {
            message = null;

            if (string.IsNullOrWhiteSpace(sql))
            {
                message = "Empty SQL.";
                return false;
            }

            if (!IsSingleStatement(sql))
            {
                message = "Only single statements are allowed.";
                return false;
            }

            try
            {
                using (var conn = new SqliteConnection("Data Source=:memory:"))
                {
                    conn.Open();

                    var trimmed = sql.TrimStart();

                    if (LooksLikeSelect(trimmed))
                    {
                        using (var cmd = new SqliteCommand(trimmed, conn))
                        using (cmd.ExecuteReader(CommandBehavior.SchemaOnly))
                        {
                            return true;
                        }
                    }
                    else
                    {
                        using (var cmd = new SqliteCommand("EXPLAIN " + trimmed, conn))
                        using (cmd.ExecuteReader())
                        {
                            return true;
                        }
                    }
                }
            }
            catch (SqliteException ex)
            {
                if (IsSQLiteSyntaxError(ex))
                {
                    message = ex.Message;
                    return false;
                }
                // Non-syntax database errors still mean the SQL parsed.
                message = ex.Message;
                return true;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return false;
            }
        }

        public static bool IsLikelySQLiteFile(string path)
        {
            try
            {
                // SQLite files start with: "SQLite format 3\0"
                var expected = "SQLite format 3\0";
                var buffer = new byte[16];
                using (var fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    if (fs.Length < buffer.Length) return false;
                    var read = fs.Read(buffer, 0, buffer.Length);
                    if (read < buffer.Length) return false;
                }
                var header = System.Text.Encoding.ASCII.GetString(buffer);
                return string.Equals(header, expected, StringComparison.Ordinal);
            }
            catch
            {
                return false;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            if (_readConnection != null)
            {
                try { _readConnection.Dispose(); } catch { }
                _readConnection = null;
            }
            if (_writeConnection != null)
            {
                try { _writeConnection.Dispose(); } catch { }
                _writeConnection = null;
            }
            GC.SuppressFinalize(this);
        }

        #region internal
        private SqliteConnection CreateConnection(bool readOnly, bool enableForeignKeys = true)
        {
            var csb = new SqliteConnectionStringBuilder
            {
                DataSource = _dbPath,
                Mode = readOnly ? SqliteOpenMode.ReadOnly : SqliteOpenMode.ReadWriteCreate
            };

            var conn = new SqliteConnection(csb.ToString());
            conn.Open();

            if (enableForeignKeys)
            {
                // Ensure pragma actually applied.
                using (var cmd = new SqliteCommand("PRAGMA foreign_keys = ON;", conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }

            return conn;
        }

        private void EnsureDatabaseExists()
        {
            var dir = Path.GetDirectoryName(_dbPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            if (!File.Exists(_dbPath))
            {
                // Create empty file so ReadWriteCreate can open it without native helpers.
                try
                {
                    using (File.Create(_dbPath)) { }
                }
                catch (Exception ex)
                {
                    throw new IOException("Failed to create SQLite database file.", ex);
                }
            }
        }

        private static bool IsSQLiteSyntaxError(SqliteException ex)
        {
            var msg = ex.Message.ToLowerInvariant();
            return
                msg.Contains("syntax error") ||
                msg.Contains("incomplete input") ||
                msg.Contains("unrecognized token") ||
                msg.Contains("unexpected") ||
                msg.Contains("malformed") ||
                msg.Contains("unclosed") ||
                msg.Contains("parse error");
        }

        private static bool LooksLikeSelect(string sql)
        {
            if (sql == null) return false;
            var s = sql.TrimStart();
            return s.StartsWith("select", StringComparison.OrdinalIgnoreCase) ||
                   s.StartsWith("with", StringComparison.OrdinalIgnoreCase);
        }

        // Simple heuristic: allow at most one terminating semicolon.
        private static bool IsSingleStatement(string sql)
        {
            var trimmed = sql.Trim();
            int semicolons = 0;
            for (int i = 0; i < trimmed.Length; i++)
                if (trimmed[i] == ';') semicolons++;

            if (semicolons == 0) return true;
            return semicolons == 1 && trimmed.EndsWith(";", StringComparison.Ordinal);
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(SqliteHelper));
        }
        #endregion
    }
}

