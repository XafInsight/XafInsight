using System.Collections.Concurrent;
using System.Collections.Generic;

namespace xafplugin.Modules
{
    public class ExportDefinition
    {
        public string Name { get; set; }
        public string MainTable { get; set; }
        public List<ColumnDescriptor> SelectedColumns { get; set; }
        public List<TableRelation> Relations { get; set; }
        public List<KeyValuePair<string, string>> CaseExpressions { get; set; } = new List<KeyValuePair<string, string>>();
    }
}
