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

        [Test]
        public void SheetRange_NoTabName_A1Notation_Formatted_Correctly()
        {
            var range = new SheetRange("", 3, 1, 5, 6);

            Assert.AreEqual("C1:E6", range.A1Notation);

            Assert.AreEqual(3, range.StartColumn);
            Assert.AreEqual(1, range.StartRow);
            Assert.AreEqual(5, range.EndColumn);
            Assert.AreEqual(6, range.EndRow);
        }

        [Test]
        public void SheetRange_With_TabName_A1Notation_Formatted_Correctly()
        {
            var range = new SheetRange("MyCustomTabName", 3, 1, 5, 6);

            Assert.AreEqual("MyCustomTabName!C1:E6", range.A1Notation);
        }

        [Test]
        public void SheetRange_NoTabName_SingleCell_A1Notation_Is_Null()
        {
            var range = new SheetRange("", 3, 1);

            Assert.IsFalse(range.CanSupportA1Notation);
            Assert.IsTrue(range.IsSingleCellRange);
        }

        [Test]
        public void SheetRange_NoTabName_R1C1_Formatted_Correctly()
        {
            var range = new SheetRange("", 3, 1, 5, 5);

            Assert.AreEqual("R1C3:R5C5", range.R1C1Notation);
        }

        [Test]
        public void SheetRange_With_TabName_R1C1_Formatted_Correctly()
        {
            var range = new SheetRange("MyCustomTab", 3, 1, 5, 5);

            Assert.AreEqual("MyCustomTab!R1C3:R5C5", range.R1C1Notation);
        }

        [Test]
        public void SheetRange_NoTabName_SingleCell_R1C1_Formatted_Correctly()
        {
            var range = new SheetRange("", 3, 1);

            Assert.IsTrue(range.IsSingleCellRange);

            Assert.AreEqual("R1C3", range.R1C1Notation);
        }

        private void AssertLettersFromColumnID(int columnID, string expectedLetters)
        {
            var result = SheetRange.GetLettersFromColumnID(columnID);

            Assert.AreEqual(expectedLetters, result);
        }
    }
}