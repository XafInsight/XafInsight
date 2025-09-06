using Microsoft.Data.Sqlite;
using NLog;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace xafplugin.Database
{
    public class CsvImporter : IDisposable
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly SqliteConnection _connection;
        private readonly char _delimiter;
        private readonly bool _hasHeaders;
        private readonly Encoding _encoding;
        private bool _disposed = false;

        public CsvImporter(SqliteConnection connection, char delimiter = ';', bool hasHeaders = true, Encoding encoding = null)
        {
            _connection = connection ?? throw new ArgumentNullException(nameof(connection));
            _delimiter = delimiter;
            _hasHeaders = hasHeaders;
            _encoding = encoding ?? Encoding.UTF8;
        }

        public string Import(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("CSV file not found", filePath);
            }

            using (var stream = File.OpenRead(filePath))
            {
                return Import(stream, Path.GetFileNameWithoutExtension(filePath));
            }
        }

        public string Import(Stream csvStream, string tableName = null)
        {
            _logger.Info("Importing CSV into SQLite database");

            if (csvStream == null)
                throw new ArgumentNullException(nameof(csvStream));

            // Generate a unique table name if none provided
            if (string.IsNullOrEmpty(tableName))
            {
                tableName = $"CSVImport_{DateTime.Now:yyyyMMdd_HHmmss}";
            }

            // Sanitize table name (remove invalid characters)
            tableName = SanitizeTableName(tableName);

            _logger.Info($"Using table name: {tableName}");

            try
            {
                // Parse CSV data
                var data = ParseCsvStream(csvStream);
                if (data.Rows.Count == 0)
                {
                    _logger.Warn("CSV file contains no data");
                    return tableName;
                }

                // Create table in database
                CreateTable(tableName, data);

                // Import data
                ImportData(tableName, data);

                _logger.Info($"Successfully imported CSV data into table '{tableName}'");
                return tableName;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"Error importing CSV data: {ex.Message}");
                throw;
            }
        }

        private DataTable ParseCsvStream(Stream csvStream)
        {
            _logger.Debug("Parsing CSV stream");
            var data = new DataTable();

            using (var reader = new StreamReader(csvStream, _encoding, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true))
            {
                string line;
                int lineCount = 0;
                List<string> headers = null;

                // Read headers
                if (_hasHeaders)
                {
                    while ((line = reader.ReadLine()) != null)
                    {
                        line = line.Trim();
                        if (string.IsNullOrWhiteSpace(line) || IsPageBreak(line))
                            continue;

                        headers = ParseCsvLine(line);
                        for (int i = 0; i < headers.Count; i++)
                        {
                            var columnName = SanitizeColumnName(headers[i]) ?? $"Column{i}";

                            // Ensure unique column names
                            if (data.Columns.Contains(columnName))
                            {
                                columnName = $"{columnName}_{i}";
                            }

                            data.Columns.Add(columnName, typeof(string));
                        }
                        lineCount++;
                        break;
                    }
                }

                // If no headers or failed to read headers
                if (data.Columns.Count == 0)
                {
                    _logger.Debug("No headers found, creating default columns");
                    line = reader.ReadLine();
                    while (line != null && (string.IsNullOrWhiteSpace(line) || IsPageBreak(line)))
                    {
                        line = reader.ReadLine();
                    }

                    if (line != null)
                    {
                        var values = ParseCsvLine(line);
                        for (int i = 0; i < values.Count; i++)
                        {
                            data.Columns.Add($"Column{i}", typeof(string));
                        }

                        // Add first line as data if we're not using headers
                        if (!_hasHeaders)
                        {
                            var row = data.NewRow();
                            for (int i = 0; i < values.Count && i < data.Columns.Count; i++)
                            {
                                row[i] = values[i];
                            }
                            data.Rows.Add(row);
                            lineCount++;
                        }
                    }
                }

                // Read data
                while ((line = reader.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (string.IsNullOrWhiteSpace(line) || IsPageBreak(line))
                        continue;

                    var values = ParseCsvLine(line);
                    var row = data.NewRow();

                    // Ensure we don't exceed column count
                    for (int i = 0; i < values.Count && i < data.Columns.Count; i++)
                    {
                        row[i] = values[i];
                    }

                    data.Rows.Add(row);
                    lineCount++;
                }

                _logger.Debug($"Parsed {lineCount} lines from CSV");
            }

            return data;
        }

        private void CreateTable(string tableName, DataTable data)
        {
            _logger.Debug($"Creating table '{tableName}'");

            var createTableSql = new StringBuilder();
            createTableSql.Append($"CREATE TABLE IF NOT EXISTS [{tableName}] (");

            for (int i = 0; i < data.Columns.Count; i++)
            {
                string columnName = data.Columns[i].ColumnName;
                createTableSql.Append($"[{columnName}] TEXT");

                if (i < data.Columns.Count - 1)
                    createTableSql.Append(", ");
            }

            createTableSql.Append(")");

            using (var cmd = _connection.CreateCommand())
            {
                cmd.CommandText = createTableSql.ToString();
                cmd.ExecuteNonQuery();
            }
        }

        private void ImportData(string tableName, DataTable data)
        {
            _logger.Debug($"Importing {data.Rows.Count} rows into table '{tableName}'");

            // Begin transaction for faster imports
            using (var transaction = _connection.BeginTransaction())
            {
                try
                {
                    // Prepare the insert statement
                    var insertSql = new StringBuilder();
                    insertSql.Append($"INSERT INTO [{tableName}] (");

                    for (int i = 0; i < data.Columns.Count; i++)
                    {
                        insertSql.Append($"[{data.Columns[i].ColumnName}]");

                        if (i < data.Columns.Count - 1)
                            insertSql.Append(", ");
                    }

                    insertSql.Append(") VALUES (");

                    for (int i = 0; i < data.Columns.Count; i++)
                    {
                        insertSql.Append($"@p{i}");

                        if (i < data.Columns.Count - 1)
                            insertSql.Append(", ");
                    }

                    insertSql.Append(")");

                    // Insert each row
                    using (var cmd = _connection.CreateCommand())
                    {
                        cmd.CommandText = insertSql.ToString();
                        cmd.Transaction = transaction;

                        // Create parameters
                        for (int i = 0; i < data.Columns.Count; i++)
                        {
                            var p = cmd.CreateParameter();
                            p.ParameterName = $"@p{i}";
                            cmd.Parameters.Add(p);
                        }

                        // Insert each row
                        foreach (DataRow row in data.Rows)
                        {
                            for (int i = 0; i < data.Columns.Count; i++)
                            {
                                cmd.Parameters[$"@p{i}"].Value = row[i] ?? DBNull.Value;
                            }

                            cmd.ExecuteNonQuery();
                        }
                    }

                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error during data import, rolling back transaction");
                    transaction.Rollback();
                    throw;
                }
            }
        }

        private List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();

            // Check if we're dealing with a simple CSV without quoted values
            if (!line.Contains("\""))
            {
                return line.Split(_delimiter).ToList();
            }

            // Handle more complex CSV with quoted values
            bool inQuote = false;
            var currentValue = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    // Check for escaped quotes
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        currentValue.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuote = !inQuote;
                    }
                }
                else if (c == _delimiter && !inQuote)
                {
                    result.Add(currentValue.ToString());
                    currentValue.Clear();
                }
                else
                {
                    currentValue.Append(c);
                }
            }

            // Add the last value
            result.Add(currentValue.ToString());

            return result;
        }

        private bool IsPageBreak(string line)
        {
            // Common page break indicators
            return line.Contains("\f") || line.Contains("\u000C") || line.Contains("PAGE") ||
                   Regex.IsMatch(line, @"^\s*-{3,}\s*$") || Regex.IsMatch(line, @"^\s*={3,}\s*$", RegexOptions.None, TimeSpan.FromMilliseconds(100));
        }

        private string SanitizeTableName(string name)
        {
            // Remove invalid characters and sanitize for SQLite
            var sanitized = Regex.Replace(name, @"[^\w]", "_", RegexOptions.None, TimeSpan.FromMilliseconds(100));

            // Ensure it doesn't start with a number
            if (char.IsDigit(sanitized[0]))
                sanitized = "t_" + sanitized;

            return sanitized;
        }

        private string SanitizeColumnName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;

            // Remove invalid characters and sanitize for SQLite
            var sanitized = Regex.Replace(name.Trim(), @"[^\w]", "_", RegexOptions.None, TimeSpan.FromMilliseconds(100));

            // Ensure it doesn't start with a number
            if (sanitized.Length > 0 && char.IsDigit(sanitized[0]))
                sanitized = "c_" + sanitized;

            return sanitized;
        }

        // Implement IDisposable pattern
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                // Dispose managed resources
                _connection?.Dispose();
            }

            // Free unmanaged resources

            _disposed = true;
        }
    }
}