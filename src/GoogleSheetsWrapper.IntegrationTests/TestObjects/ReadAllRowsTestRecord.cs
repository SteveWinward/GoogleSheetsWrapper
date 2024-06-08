namespace GoogleSheetsWrapper.IntegrationTests
{
    public class ReadAllRowsTestRecord : BaseRecord
    {
        [SheetField(
                DisplayName = "Task",
                ColumnID = 1,
                FieldType = SheetFieldType.String)]
        public string Task { get; set; } = "";

        [SheetField(
            DisplayName = "Value",
            ColumnID = 2,
            FieldType = SheetFieldType.String)]
        public string Value { get; set; } = "";


        public ReadAllRowsTestRecord() { }

        public ReadAllRowsTestRecord(IList<object> row, int rowId, int minColumnId = 1)
            : base(row, rowId, minColumnId)
        {
        }
    }
}
