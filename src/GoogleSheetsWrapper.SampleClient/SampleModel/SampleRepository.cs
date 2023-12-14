using System;

namespace GoogleSheetsWrapper.SampleClient
{
    public class SampleRepository : BaseRepository<SampleRecord>
    {
        public SampleRepository() { }

        public SampleRepository(SheetHelper<SampleRecord> sheetsHelper)
            : base(sheetsHelper) {}
    }
}
