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
        /// IComparer interface implementation.  
        /// This allows us to use the == and != operators to compare one SheetFieldAttribute to another SheetFieldAttribute.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public int Compare([AllowNull] SheetFieldAttribute x, [AllowNull] SheetFieldAttribute y)
        {
            // If either of the objects are null, return 0
            if (x == null || y == null)
            {
                return 0;
            }
            // Otherwise, return the actual CompareTo value
            else
            {
                return x.ColumnID.CompareTo(y.ColumnID);
            }
        }
    }
}