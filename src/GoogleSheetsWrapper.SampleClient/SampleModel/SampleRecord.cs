using System;
using System.Collections.Generic;

namespace GoogleSheetsWrapper.SampleClient.SampleModel
{
    public class SampleRecord : BaseRecord
    {
        [SheetField(
                DisplayName = "Task",
                ColumnID = 1,
                FieldType = SheetFieldType.String)]
        public string TaskName { get; set; }

        [SheetField(
                DisplayName = "Result",
                ColumnID = 2,
                FieldType = SheetFieldType.Boolean)]
        public bool Result { get; set; }

        [SheetField(
                DisplayName = "Error",
                ColumnID = 3,
                FieldType = SheetFieldType.String)]
        public string ErrorMessage { get; set; }

        [SheetField(
                DisplayName = "DateExecuted",
                ColumnID = 4,
                FieldType = SheetFieldType.DateTime)]
        public DateTime DateExecuted { get; set; }

        public SampleRecord() { }

        public SampleRecord(IList<object> row, int rowId, int minColumnId = 1)
            : base(row, rowId, minColumnId)
        {
        }
    }
}
