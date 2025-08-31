using System;

namespace xafplugin.Modules
{
    [Serializable]
    public class TableRelation
    {
        public string MainTable { get; set; }
        public string RelatedTable { get; set; }
        public string MainTableColumn { get; set; }
        public string RelatedTableColumn { get; set; }
        public EJoinType JoinType { get; set; } = EJoinType.LeftOuter;

        public override string ToString()
        {
            return $"{MainTable} → {RelatedTable} ({JoinType})";
        }
    }
}
