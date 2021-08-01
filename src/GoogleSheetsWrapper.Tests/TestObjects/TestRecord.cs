using GoogleSheetsWrapper;
using System;
using System.Collections.Generic;
using System.Text;

namespace GoogleSheetsWrapper.Tests.TestObjects
{
    public class TestRecord : BaseRecord
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
            DisplayName = "Previous Donation Amount",
            ColumnID = 3,
            FieldType = SheetFieldType.Currency)]
        public double DonationAmount { get; set; }

        public TestRecord() { }

        public TestRecord(IList<object> row, int rowId) : base(row, rowId) { }
    }
}
