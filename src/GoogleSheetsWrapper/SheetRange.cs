using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GoogleSheetsWrapper
{
    public class SheetRange : IEquatable<SheetRange>
    {
        public string A1Notation { get; private set; }

        public string R1C1Notation { get; private set; }

        public bool CanSupportA1Notation { get; private set; }

        public bool IsSingleCellRange { get; private set; }

        public int StartColumn { get; set; }

        public int StartRow { get; set; }

        public int? EndColumn { get; set; }

        public int? EndRow { get; set; }

        public string TabName { get; set; }

        private static readonly List<string> aToZ
            = Enumerable.Range('A', 26)
                .Select(x => (char)x + "")
                .ToList();

        /// <summary>
        /// Row and column numbers are 1 based indexs
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
                this.R1C1Notation = $"R{startRow}C{startColumn}:R{endRow}C{endColumn}";
            }
            else
            {
                this.R1C1Notation = $"R{startRow}C{startColumn}";
                this.IsSingleCellRange = true;
            }

            if (!string.IsNullOrEmpty(tabName))
            {
                this.R1C1Notation = $"{tabName}!{this.R1C1Notation}";
            }

            if (endColumn.HasValue)
            {
                this.CanSupportA1Notation = true;
                this.IsSingleCellRange = !endRow.HasValue;

                var startLetters = GetLettersFromColumnID(startColumn);
                var endLetters = GetLettersFromColumnID(endColumn.Value);

                this.A1Notation = $"{startLetters}{startRow}:{endLetters}{endRow}";

                if (!string.IsNullOrEmpty(tabName))
                {
                    this.A1Notation = $"{tabName}!{this.A1Notation}";
                }
            }

            this.StartColumn = startColumn;
            this.StartRow = startRow;
            this.EndRow = endRow;
            this.EndColumn = endColumn;
            this.TabName = tabName == null ? string.Empty : tabName;
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
                columnLettersNotation += aToZ[(block % 26)];

                block = (block / 26) - 1;
            }

            columnLettersNotation = new string(columnLettersNotation.ToCharArray().Reverse().ToArray());

            return columnLettersNotation;
        }

        /// <summary>
        /// The resulting column id is on a 1 based index (ie A => 1)
        /// </summary>
        /// <param name="letters"></param>
        /// <returns></returns>
        public static int GetColumnIDFromLetters(string letters)
        {
            int result = 0;

            letters = letters.ToUpper();

            for (int i = 0; i < letters.Count(); i++)
            {
                var currentLetter = letters[i];
                var currentLetterNumber = (int)currentLetter;

                result *= 26;
                result += currentLetterNumber - 'A' + 1;
            }

            return result;
        }

        public int GetHashCode(SheetRange obj)
        {
            return HashCode.Combine(obj.StartRow, obj.StartColumn, obj.EndColumn, obj.EndRow, obj.TabName);
        }

        public bool Equals(SheetRange other)
        {
            return
                this.A1Notation == other.A1Notation &&
                this.CanSupportA1Notation == other.CanSupportA1Notation &&
                this.EndColumn == other.EndColumn &&
                this.EndRow == other.EndRow &&
                this.IsSingleCellRange == other.IsSingleCellRange &&
                this.R1C1Notation == other.R1C1Notation &&
                this.StartColumn == other.StartColumn &&
                this.StartRow == other.StartRow &&
                this.TabName == other.TabName;
        }

        public override bool Equals(object other)
        {
            return this.Equals((SheetRange)other);
        }

        public override int GetHashCode()
        {
            return this.GetHashCode(this);
        }
    }
}
