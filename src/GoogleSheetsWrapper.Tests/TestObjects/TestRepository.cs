namespace GoogleSheetsWrapper.Tests.TestObjects
{
    public class TestRepository : BaseRepository<TestRecord>
    {
        public TestRepository() { }

        public TestRepository(SheetHelper<TestRecord> sheetsHelper)
            : base(sheetsHelper) { }

        public TestRepository(string spreadsheetID, string serviceAccountEmail, string tabName, string jsonCredentials)
            : base(spreadsheetID, serviceAccountEmail, tabName, jsonCredentials) { }
    }
}