using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace GoogleSheetsWrapper
{
    /// <summary>
    /// SheetRangeParse utility class dealing with sheet range text representations
    /// </summary>
    public class SheetRangeParser
    {
        /// <summary>
        /// Is the string range value valid A1 notation?
        /// </summary>
        /// <param name="rangeValue"></param>
        /// <returns></returns>
        public static bool IsValidA1Notation(string rangeValue)
        {
            // https://developers.google.com/sheets/api/guides/concepts

            return Regex.IsMatch(
                rangeValue,
                "([A-Z]+!)?[A-Z]+[0-9]+:[A-Z]+([0-9]+)?",
                RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// Is the range value text valid R1CI notation?
        /// </summary>
        /// <param name="rangeValue"></param>
        /// <returns></returns>
        public static bool IsValidR1C1Notation(string rangeValue)
        {
            // https://developers.google.com/sheets/api/guides/concepts

            return Regex.IsMatch(
                rangeValue,
                "([A-Z]+!)?R[0-9]+C[0-9]+(:R[0-9]+C[0-9]+)?",
                RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// Does the range value text have a tab name included?
        /// </summary>
        /// <param name="rangeValue"></param>
        /// <returns></returns>
        public static bool HasTabName(string rangeValue)
        {
            return Regex.IsMatch(rangeValue, "[A-Z]+!", RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// Parse the Tab Name out of the sheet range value
        /// </summary>
        /// <param name="rangeValue"></param>
        /// <returns></returns>
        public static string GetTabName(string rangeValue)
        {
            if (HasTabName(rangeValue))
            {
                var firstMatch = Regex.Match(rangeValue, "[A-Z]+!", RegexOptions.IgnoreCase);

                return firstMatch.Value.Replace("!", "");
            }

            return null;
        }

        /// <summary>
        /// Parses R1C1 notation into a SheetRange object
        /// </summary>
        /// <param name="rangeValue"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static SheetRange ConvertFromR1C1Notation(string rangeValue)
        {
            if (!IsValidR1C1Notation(rangeValue))
            {
                throw new ArgumentException($"The range value: {rangeValue} is not a proper R1C1 Notation.");
            }

            var tabName = GetTabName(rangeValue);

            var rangeWithoutTabName = rangeValue.Replace(tabName?.ToString() + "!", "");

            var rows = Regex.Matches(rangeWithoutTabName, "R[0-9]+", RegexOptions.IgnoreCase);
            var cols = Regex.Matches(rangeWithoutTabName, "C[0-9]+", RegexOptions.IgnoreCase);

            var firstRow = int.Parse(rows[0].Value.ToUpper(CultureInfo.CurrentCulture).Replace("R", ""), CultureInfo.CurrentCulture);
            int? lastRow;

            if (rows.Count > 1)
            {
                lastRow = int.Parse(rows[1].Value.ToUpper(CultureInfo.CurrentCulture).Replace("R", ""), CultureInfo.CurrentCulture);
            }
            else
            {
                lastRow = null;
            }

            var firstCol = int.Parse(cols[0].Value.ToUpper(CultureInfo.CurrentCulture).Replace("C", ""), CultureInfo.CurrentCulture);
            int? lastCol;

            if (cols.Count > 1)
            {
                lastCol = int.Parse(cols[1].Value.ToUpper(CultureInfo.CurrentCulture).Replace("C", ""), CultureInfo.CurrentCulture);
            }
            else
            {
                lastCol = null;
            }

            return new SheetRange(tabName, firstCol, firstRow, lastCol, lastRow);
        }

        /// <summary>
        /// Parses A1 notation into a SheetRange object
        /// </summary>
        /// <param name="rangeValue"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static SheetRange ConvertFromA1Notation(string rangeValue)
        {
            if (!IsValidA1Notation(rangeValue))
            {
                throw new ArgumentException($"The range value: {rangeValue} is not a proper A1 Notation.");
            }

            var tabName = GetTabName(rangeValue);

            var rangeWithoutTabName = rangeValue.Replace(tabName?.ToString() + "!", "");

            var columns = Regex.Matches(rangeWithoutTabName, "[A-Z]+", RegexOptions.IgnoreCase);

            var rows = Regex.Matches(rangeWithoutTabName, "[0-9]+");

            var firstColumn = SheetRange.GetColumnIDFromLetters(columns[0].Value);
            var lastColumn = SheetRange.GetColumnIDFromLetters(columns[1].Value);

            var firstRow = int.Parse(rows[0].Value, CultureInfo.CurrentCulture);

            int? lastRow;

            if (rows.Count > 1)
            {
                lastRow = int.Parse(rows[1].Value, CultureInfo.CurrentCulture);
            }
            else
            {
                lastRow = null;
            }

            return new SheetRange(tabName, firstColumn, firstRow, lastColumn, lastRow);
        }
    }
}