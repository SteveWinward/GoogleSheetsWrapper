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

            var respository = new ReadAllRowsTestRepository(sheetHelper);

            var rows = respository.GetAllRecords();

            Assert.That(rows, Has.Count.EqualTo(5));
        }
    }
}