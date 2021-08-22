using GoogleSheetsWrapper;
using System;
using System.Collections.Generic;
using System.Text;

namespace GoogleSheetsWrapper.Tests.TestObjects
{
    public class TestRecordOffset : BaseRecord
    {
        [SheetField(
            DisplayName = "Name",
            ColumnID = 11,
            FieldType = SheetFieldType.String)]
        public string Name { get; set; }

        [SheetField(
            DisplayName = "Number",
            ColumnID = 12,
            FieldType = SheetFieldType.PhoneNumber)]
        public long PhoneNumber { get; set; }

        [SheetField(
            DisplayName = "Price Amount",
            ColumnID = 13,
            FieldType = SheetFieldType.Currency)]
        public double PriceAmount { get; set; }

        [SheetField(
            DisplayName = "Date",
            ColumnID = 14,
            FieldType = SheetFieldType.DateTime)]
        public DateTime DateTime { get; set; }

        [SheetField(
            DisplayName = "Quantity",
            ColumnID = 15,
            FieldType = SheetFieldType.Number)]
        public double Quantity { get; set; }


        public TestRecordOffset() { }

        public TestRecordOffset(IList<object> row, int rowId, int minColumnId = 1)
            : base(row, rowId, minColumnId)
        {
        }
    }
}
