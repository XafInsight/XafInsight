using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using xafplugin.Helpers;
using xafplugin.Modules;

namespace xafplugin.Database
{
    public static class SqlQueryBuilder
    {
        /// <summary>
        /// Bouwt een geldige SQL-query met één of meerdere CTE-stappen.
        /// Sluit altijd af met SELECT * FROM [LaatsteStap];
        /// </summary>
        public static string BuildCteSteps(List<string> steps, List<string> tabelOrder)
        {
            if (steps == null || steps.Count == 0)
                throw new ArgumentException("Er moet minimaal één stap zijn.", nameof(steps));

            var sb = new StringBuilder();
            sb.AppendLine(";WITH");

            for (int i = 0; i < steps.Count; i++)
            {
                var name = $"step{i}";
                var sqlBody = (steps[i] ?? string.Empty).Trim().TrimEnd(';');

                bool ContainsPlaceHolder = sqlBody.Contains("PLACEHOLDER");

                if (ContainsPlaceHolder)
                {
                    if (i == 0)
                    {
                        throw new ArgumentException("Stap 0 mist een geldige FROM-clausule. Gebruik niet de placeholder", nameof(steps));
                    }
                    var prevName = $"step{i - 1}";
                    sqlBody = sqlBody.Replace("PLACEHOLDER", $"{prevName}");
                }

                sb.AppendLine($"{name} AS (");
                sb.AppendLine(sqlBody);
                sb.Append(i < steps.Count - 1 ? ")," : ")");
            }

            // Altijd een eindselect op de laatste step
            var lastStepName = $"step{steps.Count - 1}";
            sb.AppendLine();

            string selectList = (tabelOrder != null && tabelOrder.Count > 0)
                ? string.Join(", ", tabelOrder.Select(c =>
                    (c == "*" || c == $"{lastStepName}.*")
                        ? $"{QuoteIdent(lastStepName)}.*"
                        : QuoteIdent(c)))
                : "*";

            sb.AppendLine($"SELECT {selectList} FROM {QuoteIdent(lastStepName)};");

            return sb.ToString();
        }

        // Voorbeeld-implementatie; pas aan aan jouw quotelogica (SQL Server/Postgres etc.)
        private static string QuoteIdent(string ident)
        {
            if (string.IsNullOrWhiteSpace(ident)) return ident;
            // Quote niet als het al gequalificeerd is met .* (dat doen we hoger al)
            if (ident.EndsWith(".*", StringComparison.Ordinal)) return ident;
            // Simpele SQL Server-style quoting (pas aan indien nodig)
            return $"[{ident}]";
        }

        public static string BuildExportDefinitionQuery(ExportDefinition export)
        {
            if (export == null)
                throw new ArgumentNullException(nameof(export));

            var selectColumns = export.SelectedColumns
                .Where(c => c.IsCustom != true) 
                .Select(c => $"[{c.Table}].[{c.Column}]")
                .ToList();

            var sb = new StringBuilder();
            sb.AppendLine("SELECT " + string.Join(", ", selectColumns));
            sb.AppendLine($"FROM [{export.MainTable}]");

            foreach (var rel in export.Relations ?? Enumerable.Empty<TableRelation>())
            {
                sb.AppendLine($"{rel.JoinType.ToSqlQueryString()} [{rel.RelatedTable}] ON " +
                    $"[{rel.MainTable}].[{rel.MainTableColumn}] = " +
                    $"[{rel.RelatedTable}].[{rel.RelatedTableColumn}]");


            }

            return sb.ToString().Trim();
        }
    }
}
