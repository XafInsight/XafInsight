using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xafplugin.Modules;

namespace xafplugin.Helpers
{
    public static class EJoinTypeExtensions
    {
        public static string ToSqlQueryString(this EJoinType joinType)
        {
            switch (joinType)
            {
                case EJoinType.Inner:
                    return "INNER JOIN";
                case EJoinType.LeftOuter:
                    return "LEFT JOIN";
                case EJoinType.RightOuter:
                    return "RIGHT JOIN";
                case EJoinType.FullOuter:
                    return "FULL OUTER JOIN";
                default:
                    throw new ArgumentOutOfRangeException(nameof(joinType), joinType, null);
            }
        }

    }
}
