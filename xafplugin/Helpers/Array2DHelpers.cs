using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xafplugin.Helpers
{
    public static class Array2DHelpers
    {
        /// <summary>
        /// Creates a new 2D object array by inserting the provided text lines as new rows at the top.
        /// Each line is written into the first column of its row; remaining columns (if any) are filled with the specified fill value (null by default).
        /// The original data is copied unchanged below the newly added rows.
        /// </summary>
        /// <param name="data">The existing 2D object array to prepend to (must not be null).</param>
        /// <param name="lines">The text lines to insert at the top (null is treated as an empty sequence).</param>
        /// <param name="fill">Optional value used to fill the other columns in the newly created header rows (default null).</param>
        /// <returns>A new 2D object array containing the prepended text lines followed by the original data.</returns>
        /// <exception cref="ArgumentNullException">Thrown when data is null.</exception>
        public static object[,] PrependTextLines(object[,] data, IEnumerable<string> lines, object fill = null)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (lines == null) lines = Array.Empty<string>();

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            var list = lines.ToList();
            int extra = list.Count;

            var result = new object[rows + extra, cols];

            for (int r = 0; r < extra; r++)
            {
                if (cols > 0) result[r, 0] = list[r];
                for (int c = 1; c < cols; c++)
                    result[r, c] = fill;
            }

            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    result[extra + r, c] = data[r, c];

            return result;
        }

        /// <summary>
        /// Creates a new 2D object array by inserting the provided object[] rows at the top.
        /// Validates that each supplied row has exactly the same column count as the existing array.
        /// The original rows are copied unchanged below the newly inserted rows.
        /// </summary>
        /// <param name="data">The existing 2D object array to prepend to (must not be null).</param>
        /// <param name="rows">The sequence of rows (object arrays) to insert (null is treated as empty). Each row length must match the column count of data.</param>
        /// <returns>A new 2D object array containing the new rows followed by the original data.</returns>
        /// <exception cref="ArgumentNullException">Thrown when data is null.</exception>
        /// <exception cref="ArgumentException">Thrown when any provided row has a different number of columns than the data array.</exception>
        public static object[,] PrependRows(object[,] data, IEnumerable<object[]> rows)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (rows == null) rows = Array.Empty<object[]>();

            int rowsOld = data.GetLength(0);
            int cols = data.GetLength(1);
            var rowList = rows.ToList();

            foreach (var r in rowList)
                if (r.Length != cols)
                    throw new ArgumentException($"Elke rij moet {cols} kolommen hebben (gekregen: {r.Length}).");

            var result = new object[rowList.Count + rowsOld, cols];

            for (int r = 0; r < rowList.Count; r++)
                for (int c = 0; c < cols; c++)
                    result[r, c] = rowList[r][c];

            for (int r = 0; r < rowsOld; r++)
                for (int c = 0; c < cols; c++)
                    result[rowList.Count + r, c] = data[r, c];

            return result;
        }
    }
}
