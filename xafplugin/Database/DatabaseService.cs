using Microsoft.Data.Sqlite;
using NLog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using xafplugin.Interfaces;
using xafplugin.Modules;

namespace xafplugin.Database
{
    public sealed class DatabaseService : IDatabaseService, IDisposable
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        private readonly SqliteHelper _helper;
        private readonly SqliteConnection _externalConnection;
        private readonly bool _ownsHelper;
        private bool _disposed;

        // Owns its own helper (typical production use)
        public DatabaseService(string databasePath)
        {
            if (string.IsNullOrWhiteSpace(databasePath))
                throw new ArgumentException("Database path required.", nameof(databasePath));
            _helper = new SqliteHelper(databasePath);
            _ownsHelper = true;
        }

        // Uses externally created helper (caller manages helper lifetime)
        public DatabaseService(SqliteHelper helper)
        {
            _helper = helper ?? throw new ArgumentNullException(nameof(helper));
            _ownsHelper = false;
        }

        // Legacy: direct external connection (caller manages connection lifetime)
        public DatabaseService(SqliteConnection connection)
        {
            _externalConnection = connection ?? throw new ArgumentNullException(nameof(connection));
        }

        private SqliteConnection Connection
        {
            get
            {
                ThrowIfDisposed();
                if (_externalConnection != null)
                {
                    if (_externalConnection.State != ConnectionState.Open)
                        throw new InvalidOperationException("Injected SqliteConnection is not open.");
                    return _externalConnection;
                }
                return _helper.Connection;
            }
        }

        public List<string> GetTables()
        {
            logger.Info("GetTables start.");
            TryEnsureUsable();
            const string sql = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';";

            return ExecuteList(sql,
                map: r => r.GetString(0),
                onDone: list => logger.Info("Retrieved {0} tables.", list.Count),
                context: "GetTables");
        }

        public Dictionary<string, List<string>> GetTableColumns(IEnumerable<string> tables)
        {
            if (tables == null) return new Dictionary<string, List<string>>();
            logger.Info("GetTableColumns start.");
            TryEnsureUsable();

            var result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var table in tables)
            {
                if (string.IsNullOrWhiteSpace(table))
                    continue;

                var safeName = QuoteIdentifier(table);
                string pragma = "PRAGMA table_info(" + safeName + ");";

                var columns = ExecuteList(pragma,
                    map: r => r["name"].ToString(),
                    context: "GetTableColumns:" + table);

                result[table] = columns;
                logger.Info("Table '{0}' has {1} columns.", table, columns.Count);
            }

            return result;
        }

        public List<TableRelation> GetTableRelations(IEnumerable<string> tables)
        {
            if (tables == null) return new List<TableRelation>();
            logger.Info("GetTableRelations start.");
            TryEnsureUsable();

            var relations = new List<TableRelation>();

            foreach (var table in tables)
            {
                if (string.IsNullOrWhiteSpace(table))
                    continue;

                var safeName = QuoteIdentifier(table);
                string pragma = "PRAGMA foreign_key_list(" + safeName + ");";

                Execute(pragma, reader =>
                {
                    while (reader.Read())
                    {
                        var rel = new TableRelation
                        {
                            MainTable = table,
                            MainTableColumn = reader["from"].ToString(),
                            RelatedTable = reader["table"].ToString(),
                            RelatedTableColumn = reader["to"].ToString(),
                            JoinType = EJoinType.LeftOuter
                        };

                        if (!relations.Any(r =>
                            r.MainTable == rel.MainTable &&
                            r.MainTableColumn == rel.MainTableColumn &&
                            r.RelatedTable == rel.RelatedTable &&
                            r.RelatedTableColumn == rel.RelatedTableColumn))
                        {
                            relations.Add(rel);
                            logger.Debug("Relation: {0}.{1} -> {2}.{3}",
                                rel.MainTable, rel.MainTableColumn, rel.RelatedTable, rel.RelatedTableColumn);
                        }
                    }
                }, "GetTableRelations:" + table);
            }

            logger.Info("Retrieved {0} unique relations.", relations.Count);
            return relations;
        }

        public object[,] Select2DArray(string sql, Dictionary<string, object> parameters = null)
        {
            if (string.IsNullOrWhiteSpace(sql))
                throw new ArgumentException("SQL required.", nameof(sql));

            TryEnsureUsable();

            logger.Debug("Select2DArray executing.");
            using (var cmd = new SqliteCommand(sql, Connection))
            {
                AddParameters(cmd, parameters);

                using (var reader = cmd.ExecuteReader())
                {
                    int columns = reader.FieldCount;
                    var rows = new List<object[]>();

                    // header
                    var header = new object[columns];
                    for (int i = 0; i < columns; i++)
                        header[i] = reader.GetName(i);
                    rows.Add(header);

                    while (reader.Read())
                    {
                        var row = new object[columns];
                        for (int c = 0; c < columns; c++)
                            row[c] = reader.GetValue(c);
                        rows.Add(row);
                    }

                    var totalRows = rows.Count;
                    var result = new object[totalRows, columns];
                    for (int r = 0; r < totalRows; r++)
                        for (int c = 0; c < columns; c++)
                            result[r, c] = rows[r][c];

                    logger.Info("Select2DArray returned {0} data rows (plus header).", totalRows - 1);
                    return result;
                }
            }
        }

        #region Public validation methods
        public bool IsValidAgainstDb(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql)) return false;
            if (!IsConnectionOpen()) return false;

            try
            {
                var trimmed = TrimLeadingSemicolons(sql.TrimStart());
                var isSelect = LooksLikeSelect(trimmed);
                var explain = (isSelect ? "EXPLAIN QUERY PLAN\n" : "EXPLAIN\n") + trimmed;

                using (var cmd = new SqliteCommand(explain, Connection))
                using (cmd.ExecuteReader())
                {
                    // success -> parsed & planned
                }
                return true;
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "IsValidAgainstDb failed.");
                return false;
            }
        }

        public bool QueryHasRowsAny(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql)) return false;
            if (!IsConnectionOpen()) return false;

            try
            {
                var trimmed = TrimTrailingSemicolons(TrimLeadingSemicolons(sql.TrimStart()));
                if (!LooksLikeSelect(trimmed)) return false;

                // Faster EXISTS wrapping; LIMIT 1 avoids scanning entire result.
                var existsSql = "SELECT EXISTS(SELECT 1 FROM (" + trimmed + ") AS sub LIMIT 1);";
                using (var cmd = new SqliteCommand(existsSql, Connection))
                {
                    var val = Convert.ToInt32(cmd.ExecuteScalar());
                    return val == 1;
                }
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "QueryHasRowsAny failed.");
                return false;
            }
        }

        #endregion

        #region IDisposable Support

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            if (_ownsHelper && _helper != null)
            {
                try { _helper.Dispose(); } catch { }
            }
            GC.SuppressFinalize(this);
        }
        #endregion

        #region internal functions

        private void Execute(string sql, Action<SqliteDataReader> readerAction, string context)
        {
            try
            {
                using (var cmd = new SqliteCommand(sql, Connection))
                using (var reader = cmd.ExecuteReader())
                {
                    readerAction(reader);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failure executing '{0}'", context);
                throw;
            }
        }

        private List<T> ExecuteList<T>(string sql, Func<SqliteDataReader, T> map, Action<List<T>> onDone = null, string context = null)
        {
            var list = new List<T>();
            Execute(sql, reader =>
            {
                while (reader.Read())
                {
                    list.Add(map(reader));
                }
            }, context ?? sql);
            onDone?.Invoke(list);
            return list;
        }

        private static void AddParameters(SqliteCommand cmd, IDictionary<string, object> parameters)
        {
            if (parameters == null) return;
            foreach (var kv in parameters)
            {
                var p = cmd.CreateParameter();
                p.ParameterName = kv.Key;
                p.Value = kv.Value ?? DBNull.Value;
                cmd.Parameters.Add(p);
            }
        }

        private bool IsConnectionOpen()
        {
            try
            {
                return Connection.State == ConnectionState.Open;
            }
            catch
            {
                return false;
            }
        }

        private void TryEnsureUsable()
        {
            if (!IsConnectionOpen())
                throw new InvalidOperationException("Database connection is not open.");
        }

        private static string QuoteIdentifier(string name)
        {
            // Basic defensive quoting with square brackets (SQLite accepts these).
            return "[" + name.Replace("]", "]] ") + "]";
        }

        private static bool LooksLikeSelect(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            var t = s.TrimStart();
            return t.StartsWith("select", StringComparison.OrdinalIgnoreCase) ||
                   t.StartsWith("with", StringComparison.OrdinalIgnoreCase);
        }

        private static string TrimLeadingSemicolons(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            int i = 0;
            while (i < s.Length && s[i] == ';') i++;
            return s.Substring(i);
        }

        private static string TrimTrailingSemicolons(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            int i = s.Length - 1;
            while (i >= 0 && (s[i] == ';' || char.IsWhiteSpace(s[i]))) i--;
            // If we stopped on whitespace, advance to include one space block
            return s.Substring(0, i + 1);
        }

        private void ThrowIfDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(DatabaseService));
        }

        List<TableRelation> IDatabaseService.GetTableRelations(IEnumerable<string> tables)
        {
            throw new NotImplementedException();
        }
        #endregion
    }
}
