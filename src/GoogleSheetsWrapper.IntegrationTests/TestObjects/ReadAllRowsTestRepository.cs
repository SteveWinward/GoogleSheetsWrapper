namespace GoogleSheetsWrapper.IntegrationTests
{
    public class ReadAllRowsTestRepository : BaseRepository<ReadAllRowsTestRecord>
    {
        public ReadAllRowsTestRepository() { }

        public ReadAllRowsTestRepository(SheetHelper<ReadAllRowsTestRecord> sheetsHelper, bool hasHeaderRow)
            : base(sheetsHelper, hasHeaderRow) { }
    }
}
