using GoogleSheetsWrapper;
using System;
using System.Collections.Generic;
using System.Text;

namespace GoogleSheetsWrapper.Tests.TestObjects
{
    public class TestRepository : BaseRepository<TestRecord>
    {
        public TestRepository() { }

        public TestRepository(SheetHelper<TestRecord> sheetsHelper)
            : base(sheetsHelper) { }
    }
}
