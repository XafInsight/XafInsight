using System;
using System.Collections.Generic;

namespace xafplugin.Modules
{
    [Serializable]
    public class SqlCaseExpression
    {
        public string Name { get; set; }                 // Name for the column created by this CASE expression
        public string SourceColumn { get; set; }         // The column the CASE operates on
        public string SourceTable { get; set; }          // The table containing the source column
        public List<SqlCaseWhen> WhenClauses { get; set; } = new List<SqlCaseWhen>();
        public string ElseValue { get; set; }            // Value if no WHEN conditions match
    }

    [Serializable]
    public class SqlCaseWhen
    {
        public string WhenThenClause { get; set; }       // Complete "WHEN ... THEN ..." clause
    }
}