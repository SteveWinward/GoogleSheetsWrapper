namespace GoogleSheetsWrapper.IntegrationTests
{
    public class ReadAllRowsTests
    {
        public EnvironmentVariableConfig Config { get; set; }

        [SetUp]
        public void Setup()
        {
            Config = new EnvironmentVariableConfig();
        }

        [Test]
        public void ReadAllRowsHasFiveRows()
        {
            // Create a new SheetHelper class
            var sheetHelper = new SheetHelper<ReadAllRowsTestRecord>(
                Config.GoogleSpreadsheetId,
                Config.GoogleServiceAccount,
                "ReadAllRows");

            sheetHelper.Init(Config.JsonCredentials);

            var repository = new ReadAllRowsTestRepository(sheetHelper, hasHeaderRow: false);

            // Get the total row count for the existing sheet
            var rows = sheetHelper.GetRows(new SheetRange("", 1, 1, 1));

            // Delete all of the rows
            _ = sheetHelper.DeleteRows(1, rows.Count);

            // Add 5 records manually before doing the query
            var records = new List<ReadAllRowsTestRecord>();

            repository.WriteHeaders();

            for (var i = 0; i < 5; i++)
            {
                records.Add(new ReadAllRowsTestRecord()
                {
                    Task = $"Task {i + 1}",
                    Value = $"Value {i + 1}"
                });
            }

            _ = repository.AddRecords(records);

            // Create new repository class to query the records
            repository = new ReadAllRowsTestRepository(sheetHelper, hasHeaderRow: true);

            // Get all the records
            records = repository.GetAllRecords();

            // Validate there are 5 records
            Assert.That(records, Has.Count.EqualTo(5));
        }
    }
}