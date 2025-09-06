using System.Collections.Generic;
using System.Linq;
using xafplugin.Database;
using xafplugin.Modules;

namespace xafplugin.Helpers
{
    public class ExportDefinitionHelper
    {
        public static string SqlQueryWithCte(ExportDefinition export, List<FilterItem> filters = null)
        {
            try
            {
                var exportDef = export;
                var CteSteps = new List<string>();
                string sqlQuery = SqlQueryBuilder.BuildExportDefinitionQuery(exportDef);
                int stepIndex = 0;
                CteSteps.Add(sqlQuery);

                foreach (var caseExpr in exportDef.CaseExpressions)
                {
                    CteSteps.Add(caseExpr.Value);
                    stepIndex++;
                }

                if (filters != null)
                {
                    foreach (var filter in filters)
                    {
                        CteSteps.Add(filter.Expression);
                        stepIndex++;
                    }
                }

                var order = exportDef.SelectedColumns.Select(c => c.Column).ToList();

                string sqlQueryWithCte = SqlQueryBuilder.BuildCteSteps(CteSteps, order);
                return sqlQueryWithCte;
            }
            catch
            {
                return null;
            }
        }
    }
}
