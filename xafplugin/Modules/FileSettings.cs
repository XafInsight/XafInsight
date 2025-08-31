using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace xafplugin.Modules
{
    public class FileSettings
    {
        public FileSettings()
        {
            ColumnMappings = new ObservableCollection<TableMapping>();
            DefinitionMappings = new ObservableCollection<TableMapping>();
            TableRelations = new ObservableCollection<TableRelation>();
            ExportDefinitions = new ObservableCollection<ExportDefinition>();
        }
        public string Name { get; set; }
        public string AuditFileVersion { get; set; }
        public ObservableCollection<TableMapping> ColumnMappings { get; set; }
        public ObservableCollection<TableMapping> DefinitionMappings { get; set; }
        public ObservableCollection<TableRelation> TableRelations { get; set; }
        public ObservableCollection<ExportDefinition> ExportDefinitions { get; set; }
    }
    public class TableMapping
    {
        public string TableName { get; set; }
        public Dictionary<string, string> Columns { get; set; }
    }
}
