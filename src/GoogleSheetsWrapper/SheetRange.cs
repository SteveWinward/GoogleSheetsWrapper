using System;
using System.Collections.Generic;
using System.Linq;

namespace GoogleSheetsWrapper
{
    /// <summary>
    /// SheetRange object representing a Google Sheet range representation
    /// </summary>
    public class SheetRange : IEquatable<SheetRange>
    {
        /// <summary>
        /// Is this A1 notation?
        /// </summary>
        public string A1Notation { get; private set; }

        /// <summary>
        /// Is this R1C1 notation?
        /// </summary>
        public string R1C1Notation { get; private set; }

        /// <summary>
        /// Can this range support A1 notation?
        /// </summary>
        public bool CanSupportA1Notation { get; private set; }

        /// <summary>
        /// Is this a single cell?
        /// </summary>
        public bool IsSingleCellRange { get; private set; }

        /// <summary>
        /// StartColumn value (1 based index)
        /// </summary>
        public int StartColumn { get; set; }

        /// <summary>
        /// StartRow value (1 based index)
        /// </summary>
        public int StartRow { get; set; }

        /// <summary>
        /// EndColumn value (1 based index)
        /// </summary>
        public int? EndColumn { get; set; }

        /// <summary>
        /// EndRow value (1 based index)
        /// </summary>
        public int? EndRow { get; set; }

        /// <summary>
        /// Tab name in Google Sheets
        /// </summary>
        public string TabName { get; set; }

        private static readonly List<string> aToZ
            = Enumerable.Range('A', 26)
                .Select(x => (char)x + "")
                .ToList();

        /// <summary>
        /// Row and column numbers are 1 based indexes
        /// </summary>
        /// <param name="tabName"></param>
        /// <param name="startColumn"></param>
        /// <param name="endColumn"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        public SheetRange(string tabName, int startColumn, int startRow, int? endColumn = null, int? endRow = null)
        {
            if (endRow.HasValue && endColumn.HasValue)
            {
                R1C1Notation = $"R{startRow}C{startColumn}:R{endRow}C{endColumn}";
            }
            else
            {
                R1C1Notation = $"R{startRow}C{startColumn}";
                IsSingleCellRange = true;
            }

            if (!string.IsNullOrEmpty(tabName))
            {
                R1C1Notation = $"{tabName}!{R1C1Notation}";
            }

            if (endColumn.HasValue)
            {
                CanSupportA1Notation = true;
                IsSingleCellRange = !endRow.HasValue;

                var startLetters = GetLettersFromColumnID(startColumn);
                var endLetters = GetLettersFromColumnID(endColumn.Value);

                A1Notation = $"{startLetters}{startRow}:{endLetters}{endRow}";

                if (!string.IsNullOrEmpty(tabName))
                {
                    A1Notation = $"{tabName}!{A1Notation}";
                }
            }

            StartColumn = startColumn;
            StartRow = startRow;
            EndRow = endRow;
            EndColumn = endColumn;
            TabName = tabName ?? string.Empty;
        }

        /// <summary>
        /// Create a SheetRange from an A1 notation or an R1C1 notation
        /// </summary>
        /// <param name="rangeValue"></param>
        public SheetRange(string rangeValue)
        {
            SheetRange range;

            if (SheetRangeParser.IsValidR1C1Notation(rangeValue))
            {
                range = SheetRangeParser.ConvertFromR1C1Notation(rangeValue);
            }
            else if (SheetRangeParser.IsValidA1Notation(rangeValue))
            {
                range = SheetRangeParser.ConvertFromA1Notation(rangeValue);
            }
            else
            {
                throw new ArgumentException($"rangeValue: {rangeValue} is not a valid range!");
            }

            TabName = range.TabName;
            StartRow = range.StartRow;
            EndRow = range.EndRow;
            StartColumn = range.StartColumn;
            EndColumn = range.EndColumn;
            A1Notation = range.A1Notation;
            R1C1Notation = range.R1C1Notation;
            CanSupportA1Notation = range.CanSupportA1Notation;
            IsSingleCellRange = range.IsSingleCellRange;
        }

        /// <summary>
        /// columnId is a 1 based index
        /// </summary>
        /// <param name="columnID"></param>
        /// <returns></returns>
        public static string GetLettersFromColumnID(int columnID)
        {
            var block = columnID - 1;

            var columnLettersNotation = "";

            while (block >= 0)
            {
                columnLettersNotation += aToZ[block % 26];

                block = (block / 26) - 1;
            }

            columnLettersNotation = new string(columnLettersNotation.ToCharArray().Reverse().ToArray());

            return columnLettersNotation;
        }

        /// <summary>
        /// The resulting column id is on a 1 based index (i.e. A => 1)
        /// </summary>
        /// <param name="letters"></param>
        /// <returns></returns>
        public static int GetColumnIDFromLetters(string letters)
        {
            var result = 0;

            letters = letters.ToUpper(System.Globalization.CultureInfo.CurrentCulture);

            for (var i = 0; i < letters.Length; i++)
            {
                var currentLetter = letters[i];
                var currentLetterNumber = (int)currentLetter;

                result *= 26;
                result += currentLetterNumber - 'A' + 1;
            }

            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static int GetHashCode(SheetRange obj)
        {
            return HashCode.Combine(obj.StartRow, obj.StartColumn, obj.EndColumn, obj.EndRow, obj.TabName);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(SheetRange other)
        {
            return
                A1Notation == other.A1Notation &&
                CanSupportA1Notation == other.CanSupportA1Notation &&
                EndColumn == other.EndColumn &&
                EndRow == other.EndRow &&
                IsSingleCellRange == other.IsSingleCellRange &&
                R1C1Notation == other.R1C1Notation &&
                StartColumn == other.StartColumn &&
                StartRow == other.StartRow &&
                TabName == other.TabName;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            return Equals((SheetRange)obj);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return GetHashCode(this);
        }
    }
}