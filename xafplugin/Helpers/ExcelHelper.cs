using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xafplugin.Interfaces;
using Excel = Microsoft.Office.Interop.Excel;

namespace xafplugin.Helpers
{
    public static class ExcelHelper
    {
        // Excel worksheet limits
        private const int MaxExcelRows = 1_048_576;     // 2^20
        private const int MaxExcelColumns = 16_384;     // 2^14

        public static bool RangeHasData(Excel.Range range, IMessageBoxService dialog)
        {
            if (range == null) return false;

            try
            {
                // Check for values - much faster than SpecialCells for large ranges
                object[,] values = range.Value2 as object[,];
                if (values != null)
                {
                    for (int i = 1; i <= values.GetLength(0); i++)
                    {
                        for (int j = 1; j <= values.GetLength(1); j++)
                        {
                            if (values[i, j] != null && !string.IsNullOrEmpty(values[i, j].ToString()))
                            {
                                dialog.ShowWarning("Het bereik bevat al gegevens. Schrijven is geannuleerd.");
                                return true;
                            }
                        }
                    }
                }

                // Check for formulas
                object[,] formulas = range.Formula as object[,];
                if (formulas != null)
                {
                    for (int i = 1; i <= formulas.GetLength(0); i++)
                    {
                        for (int j = 1; j <= formulas.GetLength(1); j++)
                        {
                            // Only real formulas start with =
                            if (formulas[i, j] != null && formulas[i, j].ToString().StartsWith("="))
                            {
                                dialog.ShowWarning("Het bereik bevat al formules. Schrijven is geannuleerd.");
                                return true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                dialog.ShowError($"Fout bij controleren van gegevens: {ex.Message}");
            }

            return false;
        }

        public static (Excel.Application app, Excel.Workbook workbook, Excel.Worksheet sheet) GetActiveExcelComponents(IMessageBoxService dialog)
        {
            Excel.Application app = Globals.ThisAddIn.Application;

            if (app?.Workbooks == null || app.Workbooks.Count == 0)
            {
                dialog.ShowWarning("Er is geen werkboek geopend. Open of maak eerst een werkboek.");
                return (null, null, null);
            }

            Excel.Workbook workbook = app.ActiveWorkbook
                ?? app.Workbooks.Cast<Excel.Workbook>().First();

            Excel.Worksheet sheet = app.ActiveSheet as Excel.Worksheet;

            if (sheet == null)
            {
                dialog.ShowError("Kan het actieve werkblad niet gebruiken. Selecteer een geldig werkblad.");
                return (app, workbook, null);
            }

            return (app, workbook, sheet);
        }

        public static bool ValidateWorkbook(Excel.Workbook workbook, IMessageBoxService dialog)
        {
            if (workbook.ProtectStructure)
            {
                dialog.ShowError("De werkmapstructuur is beveiligd. Hef de beveiliging op om een werkblad toe te voegen.");
                return false;
            }
            if (workbook.ReadOnly)
            {
                dialog.ShowError("De werkmap is alleen-lezen. Sla op als schrijfbaar of maak een kopie.");
                return false;
            }
            return true;
        }

        public static bool ValidateWorksheet(Excel.Worksheet sheet, IMessageBoxService dialog)
        {
            if (sheet.ProtectContents)
            {
                dialog.ShowError("Het werkblad is beveiligd. Hef de beveiliging op om gegevens te plaatsen.");
                return false;
            }

            return true;
        }

        public static Excel.Range GetTargetCell(Excel.Worksheet sheet, IMessageBoxService dialog)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range cellToUse = app.ActiveCell as Excel.Range;

            if (cellToUse == null || cellToUse.Worksheet != sheet)
            {
                dialog.ShowWarning("Selecteer eerst een cel waar de gegevens geplaatst moeten worden.");
                return null;
            }

            // Activeer het blad en selecteer de cel
            sheet.Activate();
            cellToUse.Select();

            // Bepaal positie
            int r = cellToUse.Row;
            int c = cellToUse.Column;

            // Vertaal die positie naar het doel-werkblad
            Excel.Range cellOnTarget = (Excel.Range)sheet.Cells[r, c];

            if (cellOnTarget == null)
            {
                dialog.ShowWarning("Selecteer eerst een cel waar de gegevens geplaatst moeten worden.");
                return null;
            }

            cellOnTarget.Select();
            return cellOnTarget;
        }

        public static bool ValidateRangeSize(Excel.Range startCell, int rows, int cols, IMessageBoxService dialog)
        {
            if (startCell.Row + rows - 1 > MaxExcelRows || startCell.Column + cols - 1 > MaxExcelColumns)
            {
                dialog.ShowError("Bereik past niet op het werkblad.");
                return false;
            }
            return true;
        }
    }
}
