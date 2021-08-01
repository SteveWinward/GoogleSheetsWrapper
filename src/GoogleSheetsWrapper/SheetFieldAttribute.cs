using System;
using System.Collections.Generic;
using System.Text;

namespace GoogleSheetsWrapper
{
    public class SheetFieldAttribute : Attribute
    {
        public string DisplayName { get; set; }

        public int ColumnID { get; set; }

        public SheetFieldType FieldType { get; set; }
    }

    public enum SheetFieldType
    {
        String,
        DateTime,
        Currency,
        PhoneNumber
    }
}
