using System;

namespace GoogleSheetsWrapper.SampleClient
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
    }
}
