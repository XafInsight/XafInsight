using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xafplugin.Modules;

namespace xafplugin.Helpers
{
    public static class ExportValidationHelper
    {
        public static List<ColumnDescriptor> FilterValidColumns(IEnumerable<ColumnDescriptor> columns, Dictionary<string, List<string>> tableColumns)
        {
            if (columns == null || tableColumns == null)
                return new List<ColumnDescriptor>();

            return columns
               .Where(col =>
                   !string.IsNullOrWhiteSpace(col?.Column) // altijd verplicht
                   && (
                       col.IsCustom == true
                       || (
                           !string.IsNullOrWhiteSpace(col?.Table)
                           && tableColumns.TryGetValue(col.Table, out var tableCols)
                           && tableCols.Contains(col.Column)
                       )
                   )
               )
               .ToList();

        }

        public static List<TableRelation> FilterRelationsByTables(
        IEnumerable<TableRelation> relations,
        HashSet<string> validTables)
        {
            if (relations == null || validTables == null)
                return new List<TableRelation>();

            return relations
                .Where(r =>
                    !string.IsNullOrWhiteSpace(r?.MainTable) &&
                    !string.IsNullOrWhiteSpace(r?.RelatedTable) &&
                    validTables.Contains(r.MainTable) &&
                    validTables.Contains(r.RelatedTable))
                .ToList();
        }
    public static bool IsRelationValid(TableRelation relation, Dictionary<string, List<string>> tableColumns)
        {
            if (relation == null || tableColumns == null)
                return false;

            return
                !string.IsNullOrWhiteSpace(relation.MainTable) &&
                !string.IsNullOrWhiteSpace(relation.RelatedTable) &&
                !string.IsNullOrWhiteSpace(relation.MainTableColumn) &&
                !string.IsNullOrWhiteSpace(relation.RelatedTableColumn) &&
                tableColumns.ContainsKey(relation.MainTable) &&
                tableColumns.ContainsKey(relation.RelatedTable) &&
                tableColumns[relation.MainTable].Contains(relation.MainTableColumn) &&
                tableColumns[relation.RelatedTable].Contains(relation.RelatedTableColumn);
        }
    }
}
