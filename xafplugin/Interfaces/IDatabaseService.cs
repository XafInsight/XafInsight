using System.Collections.Generic;
using xafplugin.Modules; // TableRelation
// If TableRelation lives elsewhere adjust the using accordingly.

namespace xafplugin.Interfaces
{
    /// <summary>
    /// Read-only schema/data inspection operations for a SQLite database.
    /// Implementations should not mutate schema or data.
    /// </summary>
    public interface IDatabaseService
    {
        /// <summary>
        /// Returns user table names (excludes internal sqlite_ tables).
        /// </summary>
        List<string> GetTables();

        /// <summary>
        /// Returns a mapping of table name -> ordered list of column names.
        /// Only tables present in the input enumeration are queried.
        /// </summary>
        Dictionary<string, List<string>> GetTableColumns(IEnumerable<string> tables);

        /// <summary>
        /// Returns discovered foreign key style relations between the given tables.
        /// </summary>
        List<TableRelation> GetTableRelations(IEnumerable<string> tables);

        /// <summary>
        /// Executes a (typically SELECT) statement and returns a 2D array
        /// where row 0 contains the column headers.
        /// </summary>
        object[,] Select2DArray(string sql, Dictionary<string, object> parameters = null);
    }
}
