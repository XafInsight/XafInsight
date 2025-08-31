using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace xafplugin.Database
{
    public class DynamicXmlImporter : IDisposable
    {
        private readonly SQLiteConnection _conn;

        private SQLiteTransaction _currentTransaction;
        private int _insertCount = 0;
        private readonly int _batchSize;
        private bool _disposed = false;

        // Caches
        private readonly HashSet<string> _existingTables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, HashSet<string>> _tableCols = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, long> _tableIdCounters = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        
        // Track relationships for later foreign key creation
        private readonly Dictionary<string, List<Tuple<string, string>>> _pendingRelationships = 
            new Dictionary<string, List<Tuple<string, string>>>(StringComparer.OrdinalIgnoreCase);
        
        private sealed class StackEntry
        {
            public string LocalName;
            public string NamespaceUri;
            public string Path;   // joined path with '/'
            public string Table;  // normalized table name (from local name)
            public string RowId;  // tablename_number, empty means not materialized yet
            public bool HadElementChildren;
            public bool HadAttributes;
            public bool Deferred; // true until we actually insert a row
        }

        private static readonly Regex _identCleaner = new Regex(@"[^A-Za-z0-9_]", RegexOptions.Compiled, TimeSpan.FromMilliseconds(100));

        public DynamicXmlImporter(SQLiteConnection conn, int batchSize = 25000)
        {
            _conn = conn;
            _batchSize = batchSize;
        }

        public void Import(Stream xmlStream)
        {
            EnsurePragmas();

            // Temporarily disable foreign keys during import
            ExecuteNonQuery("PRAGMA foreign_keys=OFF;");
            
            Stack<StackEntry> stack = new Stack<StackEntry>();
            _insertCount = 0;
            _currentTransaction = _conn.BeginTransaction();

            try
            {
                XmlReaderSettings settings = new XmlReaderSettings
                {
                    IgnoreWhitespace = true,
                    IgnoreComments = true,
                    DtdProcessing = DtdProcessing.Ignore,
                    ValidationType = ValidationType.None,
                    XmlResolver = null
                };

                using (XmlReader reader = XmlReader.Create(xmlStream, settings))
                {
                    List<string> pathParts = new List<string>();
                    StringBuilder valueBuffer = new StringBuilder();

                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.Element)
                        {
                            // mark that current parent has element children
                            if (stack.Count > 0)
                                stack.Peek().HadElementChildren = true;

                            string local = reader.LocalName;            // ignore prefix
                            string ns = reader.NamespaceURI ?? string.Empty;

                            pathParts.Add(local);
                            string path = string.Join("/", pathParts);
                            string table = NormalizeTableName(local);

                            // Attributes
                            Dictionary<string, string> attrs = new Dictionary<string, string>();
                            bool hadAttributes = false;
                            if (reader.HasAttributes)
                            {
                                for (int i = 0; i < reader.AttributeCount; i++)
                                {
                                    reader.MoveToAttribute(i);
                                    if (string.Equals(reader.Prefix, "xmlns", StringComparison.OrdinalIgnoreCase) ||
                                        string.Equals(reader.Name, "xmlns", StringComparison.OrdinalIgnoreCase))
                                        continue;

                                    attrs[reader.LocalName] = reader.Value;
                                    hadAttributes = true;
                                }
                                reader.MoveToElement();
                            }

                            // Ensure parent is materialized as soon as it gains a child
                            if (stack.Count > 0)
                            {
                                StackEntry parent = stack.Peek();
                                if (string.IsNullOrEmpty(parent.RowId)) // not materialized yet
                                {
                                    string gpTable = stack.Count > 1 ? stack.ElementAt(1).Table : null;
                                    string gpRowId = stack.Count > 1 ? stack.ElementAt(1).RowId : null;

                                    // Make sure parent table exists
                                    EnsureTable(parent.Table);

                                    Dictionary<string, object> parentData = new Dictionary<string, object>();
                                    if (!string.IsNullOrEmpty(gpRowId))
                                    {
                                        parentData["_ParentId"] = gpRowId;
                                        
                                        // Track relationship for later foreign key creation
                                        if (!string.IsNullOrEmpty(gpTable))
                                        {
                                            TrackRelationship(parent.Table, gpTable, "_ParentId");
                                        }
                                    }
                                    else
                                    {
                                        parentData["_ParentId"] = DBNull.Value;
                                    }
                                    
                                    parentData["_Path"] = parent.Path;
                                    parentData["_NS"] = parent.NamespaceUri;

                                    string newParentId = InsertRow(parent.Table, parentData);
                                    parent.RowId = newParentId;
                                    parent.Deferred = false;

                                    // rewrite the parent on the stack
                                    stack.Pop();
                                    stack.Push(parent);
                                }
                            }

                            // Now handle the current element
                            string parentRowId = stack.Count > 0 ? stack.Peek().RowId : null;
                            string parentTable = stack.Count > 0 ? stack.Peek().Table : null;

                            string rowId = string.Empty;
                            bool deferred = !hadAttributes; // if no attributes, defer creation

                            if (hadAttributes)
                            {
                                // Create table
                                EnsureTable(table);
                                EnsureColumns(table, attrs.Keys);

                                Dictionary<string, object> baseData = new Dictionary<string, object>();
                                if (!string.IsNullOrEmpty(parentRowId))
                                {
                                    baseData["_ParentId"] = parentRowId;
                                    
                                    // Track relationship for later foreign key creation
                                    if (!string.IsNullOrEmpty(parentTable))
                                    {
                                        TrackRelationship(table, parentTable, "_ParentId");
                                    }
                                }
                                else
                                {
                                    baseData["_ParentId"] = DBNull.Value;
                                }
                                
                                baseData["_Path"] = path;
                                baseData["_NS"] = ns;
                                
                                foreach (KeyValuePair<string, string> kv in attrs)
                                    baseData[NormalizeName(kv.Key)] = kv.Value;

                                rowId = InsertRow(table, baseData);
                            }

                            StackEntry entry = new StackEntry
                            {
                                LocalName = local,
                                NamespaceUri = ns,
                                Path = path,
                                Table = table,
                                RowId = rowId,                 // empty if deferred
                                HadElementChildren = false,
                                HadAttributes = hadAttributes,
                                Deferred = deferred
                            };
                            stack.Push(entry);

                            if (reader.IsEmptyElement)
                            {
                                FinalizeCurrentNode(stack, valueBuffer);
                                pathParts.RemoveAt(pathParts.Count - 1);
                            }
                        }
                        else if (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.CDATA)
                        {
                            valueBuffer.Append(reader.Value);
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement)
                        {
                            FinalizeCurrentNode(stack, valueBuffer);
                            if (pathParts.Count > 0)
                                pathParts.RemoveAt(pathParts.Count - 1);
                        }
                    }
                }

                if (_currentTransaction != null)
                {
                    _currentTransaction.Commit();
                    _currentTransaction.Dispose();
                    _currentTransaction = null;
                }
                
                // Add foreign key constraints in a second pass
                AddForeignKeyConstraints();
                
                // Re-enable foreign keys after import
                ExecuteNonQuery("PRAGMA foreign_keys=ON;");
            }
            catch (Exception ex)
            {
                try
                {
                    if (_currentTransaction != null)
                    {
                        _currentTransaction.Rollback();
                        _currentTransaction.Dispose();
                        _currentTransaction = null;
                    }
                }
                catch { /* ignore */ }

                // Make sure to restore foreign keys setting
                ExecuteNonQuery("PRAGMA foreign_keys=ON;");

                throw new Exception("Fout tijdens XML import: " + ex.Message, ex);
            }
        }

        private void TrackRelationship(string childTable, string parentTable, string columnName)
        {
            if (!_pendingRelationships.TryGetValue(childTable, out var list))
            {
                list = new List<Tuple<string, string>>();
                _pendingRelationships[childTable] = list;
            }
            
            // Only add if not already tracked
            if (!list.Any(r => r.Item1 == parentTable && r.Item2 == columnName))
            {
                list.Add(Tuple.Create(parentTable, columnName));
            }
        }
        
        private void AddForeignKeyConstraints()
        {
            // Create a new transaction for the structure changes
            using (var transaction = _conn.BeginTransaction())
            {
                try
                {
                    // Process each table that needs foreign keys
                    foreach (var entry in _pendingRelationships)
                    {
                        string childTable = entry.Key;
                        var relationships = entry.Value;
                        
                        // Skip if no relationships to add
                        if (relationships.Count == 0)
                            continue;
                            
                        // Get list of all columns in the table
                        List<string> columnDefinitions = new List<string>();
                        List<string> columnNames = new List<string>();
                        Dictionary<string, string> columnTypes = new Dictionary<string, string>();
                        
                        using (var cmd = new SQLiteCommand($"PRAGMA table_info([{childTable}])", _conn, transaction))
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string colName = reader["name"].ToString();
                                string colType = reader["type"].ToString();
                                int notNull = Convert.ToInt32(reader["notnull"]);
                                string defaultValue = reader["dflt_value"]?.ToString();
                                int isPk = Convert.ToInt32(reader["pk"]);
                                
                                columnNames.Add(colName);
                                columnTypes[colName] = colType;
                                
                                string colDef = $"[{colName}] {colType}";
                                if (notNull == 1)
                                    colDef += " NOT NULL";
                                    
                                if (!string.IsNullOrEmpty(defaultValue))
                                    colDef += $" DEFAULT {defaultValue}";
                                    
                                if (isPk == 1)
                                    colDef += " PRIMARY KEY";
                                    
                                columnDefinitions.Add(colDef);
                            }
                        }
                        
                        // Add foreign key definitions
                        foreach (var rel in relationships)
                        {
                            string parentTable = rel.Item1;
                            string columnName = rel.Item2;
                            
                            columnDefinitions.Add($"FOREIGN KEY ([{columnName}]) REFERENCES [{parentTable}]([_Id]) ON DELETE CASCADE");
                        }
                        
                        // Create temporary table with foreign keys
                        string tempTableName = $"{childTable}_temp";
                        string createTempSql = $"CREATE TABLE [{tempTableName}] ({string.Join(", ", columnDefinitions)});";
                        ExecuteNonQuery(createTempSql, transaction);
                        
                        // Copy data
                        string columns = string.Join(", ", columnNames.Select(c => $"[{c}]"));
                        string copyDataSql = $"INSERT INTO [{tempTableName}] ({columns}) SELECT {columns} FROM [{childTable}];";
                        ExecuteNonQuery(copyDataSql, transaction);
                        
                        // Drop old table
                        string dropSql = $"DROP TABLE [{childTable}];";
                        ExecuteNonQuery(dropSql, transaction);
                        
                        // Rename new table
                        string renameSql = $"ALTER TABLE [{tempTableName}] RENAME TO [{childTable}];";
                        ExecuteNonQuery(renameSql, transaction);
                        
                        // Recreate indexes
                        string indexSql = $"CREATE INDEX IF NOT EXISTS [ix_{childTable}__parent] ON [{childTable}] ([_ParentId]);";
                        ExecuteNonQuery(indexSql, transaction);
                    }
                    
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    throw new Exception("Failed to add foreign key constraints: " + ex.Message, ex);
                }
            }
        }

        private void FinalizeCurrentNode(Stack<StackEntry> stack, StringBuilder valueBuffer)
        {
            if (stack.Count == 0) return;

            StackEntry cur = stack.Pop();
            string text = valueBuffer.Length > 0 ? valueBuffer.ToString() : null;
            valueBuffer.Clear();

            bool isScalarLeaf = !cur.HadElementChildren && !cur.HadAttributes;

            if (isScalarLeaf)
            {
                // Pure leaf: never materialize a row/table. Promote as column on parent (if any).
                if (text != null && stack.Count > 0)
                {
                    StackEntry parent = stack.Peek();
                    EnsureColumns(parent.Table, new[] { cur.LocalName });

                    Dictionary<string, object> updates = new Dictionary<string, object>();
                    updates[NormalizeName(cur.LocalName)] = text;
                    UpdateColumns(parent.Table, parent.RowId, updates);
                }
                CommitIfNeeded();
                return;
            }

            // Non-leaf (has children and/or attributes): ensure it's materialized
            if (string.IsNullOrEmpty(cur.RowId))
            {
                string parentTable = stack.Count > 0 ? stack.Peek().Table : null;
                string parentRowId = stack.Count > 0 ? stack.Peek().RowId : null;

                // Create table
                EnsureTable(cur.Table);

                Dictionary<string, object> baseData = new Dictionary<string, object>();
                if (!string.IsNullOrEmpty(parentRowId))
                {
                    baseData["_ParentId"] = parentRowId;
                    
                    // Track relationship for later foreign key creation
                    if (!string.IsNullOrEmpty(parentTable))
                    {
                        TrackRelationship(cur.Table, parentTable, "_ParentId");
                    }
                }
                else
                {
                    baseData["_ParentId"] = DBNull.Value;
                }
                
                baseData["_Path"] = cur.Path;
                baseData["_NS"] = cur.NamespaceUri;

                cur.RowId = InsertRow(cur.Table, baseData);
                cur.Deferred = false;
            }

            // Save inner text (if present) on _Value
            if (!string.IsNullOrEmpty(text))
            {
                Dictionary<string, object> updates = new Dictionary<string, object>();
                updates["_Value"] = text;
                UpdateColumns(cur.Table, cur.RowId, updates);
            }

            CommitIfNeeded();
        }

        private void ExecuteNonQuery(string sql)
        {
            using (var cmd = new SQLiteCommand(sql, _conn))
            {
                cmd.ExecuteNonQuery();
            }
        }
        
        private void ExecuteNonQuery(string sql, SQLiteTransaction transaction)
        {
            using (var cmd = new SQLiteCommand(sql, _conn, transaction))
            {
                cmd.ExecuteNonQuery();
            }
        }

        private void EnsurePragmas()
        {
            ExecuteNonQuery("PRAGMA journal_mode=WAL;");
            ExecuteNonQuery("PRAGMA synchronous=NORMAL;");
        }

        private void EnsureTable(string table)
        {
            if (_existingTables.Contains(table))
                return;

            List<string> cols = new List<string>();
            cols.Add("[_Id] TEXT PRIMARY KEY");
            cols.Add("[_ParentId] TEXT NULL");
            cols.Add("[_Path] TEXT NOT NULL");
            cols.Add("[_NS] TEXT NULL");
            cols.Add("[_Value] TEXT NULL");

            string createSql = $"CREATE TABLE IF NOT EXISTS [{table}] ({string.Join(", ", cols)});";
            ExecuteNonQuery(createSql);

            string indexSql = $"CREATE INDEX IF NOT EXISTS [ix_{table}__parent] ON [{table}] ([_ParentId]);";
            ExecuteNonQuery(indexSql);

            _existingTables.Add(table);
            _tableCols[table] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "_id", "_parentid", "_path", "_ns", "_value"
            };
        }

        private void EnsureColumns(string table, IEnumerable<string> names)
        {
            if (!_tableCols.TryGetValue(table, out var set))
            {
                set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                using (var cmd = new SQLiteCommand($"PRAGMA table_info([{table}]);", _conn))
                using (var rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                        set.Add(Convert.ToString(rdr["name"]));
                }
                _tableCols[table] = set;
            }

            if (names == null) return;

            foreach (var raw in names)
            {
                string col = NormalizeName(raw);
                if (string.IsNullOrWhiteSpace(col)) continue;
                if (set.Contains(col)) continue;

                string alterSql = $"ALTER TABLE [{table}] ADD COLUMN [{col}] TEXT;";
                ExecuteNonQuery(alterSql);
                
                set.Add(col);
            }
        }

        private string InsertRow(string table, Dictionary<string, object> data)
        {
            if (data == null || data.Count == 0)
            {
                data = new Dictionary<string, object>();
                data["_Path"] = "";
                data["_NS"] = DBNull.Value;
            }

            // Generate a new ID in the format tablename_number
            long counter = GetNextIdCounter(table);
            string newId = $"{table}_{counter}";

            List<string> cols = new List<string>();
            List<string> pars = new List<string>();

            using (var cmd = new SQLiteCommand())
            {
                cmd.Connection = _conn;
                cmd.Transaction = _currentTransaction;

                // Add the ID column explicitly
                cols.Add("[_Id]");
                pars.Add("@_Id");
                cmd.Parameters.AddWithValue("@_Id", newId);

                foreach (KeyValuePair<string, object> kv in data)
                {
                    string col = NormalizeName(kv.Key);
                    cols.Add($"[{col}]");
                    string p = $"@{col}";
                    pars.Add(p);
                    cmd.Parameters.AddWithValue(p, kv.Value ?? DBNull.Value);
                }

                string sql = $"INSERT INTO [{table}] ({string.Join(", ", cols)}) VALUES ({string.Join(", ", pars)});";
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
            }

            return newId;
        }

        private long GetNextIdCounter(string table)
        {
            if (!_tableIdCounters.ContainsKey(table))
            {
                _tableIdCounters[table] = 1;
            }
            
            return _tableIdCounters[table]++;
        }

        private void UpdateColumns(string table, string rowId, Dictionary<string, object> updates)
        {
            if (string.IsNullOrEmpty(rowId) || updates == null || updates.Count == 0) return;

            List<string> sets = new List<string>();

            using (var cmd = new SQLiteCommand())
            {
                cmd.Connection = _conn;
                cmd.Transaction = _currentTransaction;

                foreach (KeyValuePair<string, object> kv in updates)
                {
                    string col = NormalizeName(kv.Key);
                    sets.Add($"[{col}] = @{col}");
                    cmd.Parameters.AddWithValue($"@{col}", kv.Value ?? DBNull.Value);
                }

                string sql = $"UPDATE [{table}] SET {string.Join(", ", sets)} WHERE [_Id] = @_id;";
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue("@_id", rowId);

                cmd.ExecuteNonQuery();
            }
        }

        private void DeleteRow(string table, string rowId)
        {
            if (string.IsNullOrEmpty(rowId)) return;
            using (var del = new SQLiteCommand($"DELETE FROM [{table}] WHERE [_Id] = @_id;", _conn, _currentTransaction))
            {
                del.Parameters.AddWithValue("@_id", rowId);
                del.ExecuteNonQuery();
            }
        }

        private void CommitIfNeeded()
        {
            _insertCount++;
            if (_insertCount >= _batchSize)
            {
                _currentTransaction.Commit();
                _currentTransaction.Dispose();
                _currentTransaction = _conn.BeginTransaction();
                _insertCount = 0;
            }
        }

        private static string NormalizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "_";

            // Don't add n_ prefix to system columns that already start with underscore
            if (name.StartsWith("_"))
                return name;

            string cleaned = _identCleaner.Replace(name, "_");
            cleaned = Regex.Replace(cleaned, "_+", "_", RegexOptions.None, TimeSpan.FromMilliseconds(100));
            if (!char.IsLetter(cleaned[0])) cleaned = "n_" + cleaned;
            return cleaned;
        }

        private string NormalizeTableName(string localName)
        {
            // Use just the local name for the table
            string normalized = NormalizeName(localName);
            return normalized;
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Dispose managed resources
                    if (_currentTransaction != null)
                    {
                        try
                        {
                            _currentTransaction.Dispose();
                        }
                        catch (Exception ex)
                        {
                            // Log exception but don't throw from Dispose
                            System.Diagnostics.Debug.WriteLine($"Error disposing transaction: {ex.Message}");
                        }
                        _currentTransaction = null;
                    }

                    // Clear collections - doing this thoroughly to prevent memory leaks
                    if (_existingTables != null)
                    {
                        _existingTables.Clear();
                    }
                    
                    if (_tableCols != null)
                    {
                        // Clear each nested HashSet before clearing the dictionary
                        foreach (var set in _tableCols.Values)
                        {
                            set?.Clear();
                        }
                        _tableCols.Clear();
                    }
                    
                    if (_tableIdCounters != null)
                    {
                        _tableIdCounters.Clear();
                    }
                    
                    if (_pendingRelationships != null)
                    {
                        // Clear each nested List before clearing the dictionary
                        foreach (var list in _pendingRelationships.Values)
                        {
                            list?.Clear();
                        }
                        _pendingRelationships.Clear();
                    }
                    
                    // Force garbage collection to reclaim memory immediately
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

                // Set disposed flag
                _disposed = true;
            }
        }
    }
}
