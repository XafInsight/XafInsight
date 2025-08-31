using System.Collections.Generic;

namespace xafplugin.Modules
{
    public class ColumnDescriptor
    {
        public string Table { get; set; }
        public string Column { get; set; }
        public List<TableRelation> Path { get; set; } = new List<TableRelation>();
        public bool IsCustom { get; set; } = false;
        public string Display => $"{Table}.{Column}";
        public override string ToString() => Display;
    }

}
