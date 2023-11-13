using System;
using System.Collections.Generic;

namespace GoogleSheetsWrapper.Tests.TestObjects
{
    public class TestRecordNonNullableDateTime : BaseRecord
    {
        [SheetField(
            DisplayName = "Name",
            ColumnID = 1,
            FieldType = SheetFieldType.String)]
        public string Name { get; set; }

        [SheetField(
            DisplayName = "Number",
            ColumnID = 2,
            FieldType = SheetFieldType.PhoneNumber)]
        public long PhoneNumber { get; set; }

        [SheetField(
            DisplayName = "Price Amount",
            ColumnID = 3,
            FieldType = SheetFieldType.Currency)]
        public double PriceAmount { get; set; }

        [SheetField(
            DisplayName = "Date",
            ColumnID = 4,
            FieldType = SheetFieldType.DateTime)]
        public DateTime DateTime { get; set; }

        [SheetField(
            DisplayName = "Quantity",
            ColumnID = 5,
            FieldType = SheetFieldType.Number)]
        public double Quantity { get; set; }


        public TestRecordNonNullableDateTime() { }

        public TestRecordNonNullableDateTime(IList<object> row, int rowId, int minColumnId = 1)
            : base(row, rowId, minColumnId)
        {
        }
    }
}