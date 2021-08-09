using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace GoogleSheetsWrapper
{
    public class SheetRangeParser
    {
        public bool IsValidA1Notation(string rangeValue)
        {
            // https://developers.google.com/sheets/api/guides/concepts

            return Regex.IsMatch(
                rangeValue,
                "([A-Z]+!)?[A-Z]+[0-9]+:[A-Z]+([0-9]+)?",
                RegexOptions.IgnoreCase);
        }

        public bool IsValidR1C1Notation(string rangeValue)
        {
            // https://developers.google.com/sheets/api/guides/concepts

            return Regex.IsMatch(
                rangeValue,
                "([A-Z]+!)?R[0-9]+C[0-9]+(:R[0-9]+C[0-9]+)?",
                RegexOptions.IgnoreCase);
        }

        public bool HasTabName(string rangeValue)
        {
            return Regex.IsMatch(rangeValue, "[A-Z]+!", RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// Parse the Tab Name out of the sheet range value
        /// </summary>
        /// <param name="rangeValue"></param>
        /// <returns></returns>
        public string GetTabName(string rangeValue)
        {
            if (this.HasTabName(rangeValue))
            {
                var firstMatch = Regex.Match(rangeValue, "[A-Z]+!", RegexOptions.IgnoreCase);

                return firstMatch.Value.Replace("!", "");
            }

            return null;
        }

        public SheetRange ConvertFromR1C1Notation(string rangeValue)
        {
            if (!this.IsValidR1C1Notation(rangeValue))
            {
                throw new ArgumentException($"The range value: {rangeValue} is not a proper R1C1 Notation.");
            }

            var tabName = this.GetTabName(rangeValue);

            var rangeWithoutTabName = rangeValue.Replace(tabName?.ToString() + "!", "");

            var rows = Regex.Matches(rangeWithoutTabName, "R[0-9]+", RegexOptions.IgnoreCase);
            var cols = Regex.Matches(rangeWithoutTabName, "C[0-9]+", RegexOptions.IgnoreCase);

            int firstRow = int.Parse(rows[0].Value.ToUpper().Replace("R", ""));
            int? lastRow;

            if (rows.Count > 1)
            {
                lastRow = int.Parse(rows[1].Value.ToUpper().Replace("R", ""));
            }
            else
            {
                lastRow = null;
            }

            int firstCol = int.Parse(cols[0].Value.ToUpper().Replace("C", ""));
            int? lastCol;

            if (cols.Count > 1)
            {
                lastCol = int.Parse(cols[1].Value.ToUpper().Replace("C", ""));
            }
            else
            {
                lastCol = null;
            }

            return new SheetRange(tabName, firstCol, firstRow, lastCol, lastRow);
        }

        public SheetRange ConvertFromA1Notation(string rangeValue)
        {
            if (!this.IsValidA1Notation(rangeValue))
            {
                throw new ArgumentException($"The range value: {rangeValue} is not a proper A1 Notation.");
            }

            var tabName = this.GetTabName(rangeValue);

            var rangeWithoutTabName = rangeValue.Replace(tabName?.ToString() + "!", "");

            var columns = Regex.Matches(rangeWithoutTabName, "[A-Z]+", RegexOptions.IgnoreCase);

            var rows = Regex.Matches(rangeWithoutTabName, "[0-9]+");

            var firstColumn = SheetRange.GetColumnIDFromLetters(columns[0].Value);
            var lastColumn = SheetRange.GetColumnIDFromLetters(columns[1].Value);

            var firstRow = int.Parse(rows[0].Value);

            int? lastRow;

            if (rows.Count > 1)
            {
                lastRow = int.Parse(rows[1].Value);
            }
            else
            {
                lastRow = null;
            }

            return new SheetRange(tabName, firstColumn, firstRow, lastColumn, lastRow);
        }
    }
}
