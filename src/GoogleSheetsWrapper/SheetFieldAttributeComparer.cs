using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace GoogleSheetsWrapper
{
    /// <summary>
    /// IComparer implementation for SheetFieldAttribute class
    /// </summary>
    public class SheetFieldAttributeComparer : IComparer<SheetFieldAttribute>
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public int Compare([AllowNull] SheetFieldAttribute x, [AllowNull] SheetFieldAttribute y)
        {
#pragma warning disable IDE0046 // IF statement can be simplified

            if ((x != null) && (y != null))
            {
                return x.ColumnID.CompareTo(y.ColumnID);
            }
            else
            {
                return 0;
            }

#pragma warning restore IDE0046 // IF statement can be simplified
        }
    }
}