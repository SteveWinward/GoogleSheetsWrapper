using GoogleSheetsWrapper;
using NUnit.Framework;

namespace CAR.Core.Tests
{
    public class SheetRangeTests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void SheetRange_GetLettersFromColumnID_Tests()
        {
            this.AssertLettersFromColumnID(1, "A");
            this.AssertLettersFromColumnID(26, "Z");
            this.AssertLettersFromColumnID(27, "AA");
            this.AssertLettersFromColumnID(54, "BB");
            this.AssertLettersFromColumnID(752, "ABX");
        }

        private void AssertLettersFromColumnID(int columnID, string expectedLetters)
        {
            var result = SheetRange.GetLettersFromColumnID(columnID);

            Assert.AreEqual(expectedLetters, result);
        }
    }
}