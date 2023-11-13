using NUnit.Framework;

namespace GoogleSheetsWrapper.Tests
{
    public class SheetRangeTests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void SheetRangeGetLettersFromColumnIDTests()
        {
            AssertLettersFromColumnID(1, "A");
            AssertLettersFromColumnID(26, "Z");
            AssertLettersFromColumnID(27, "AA");
            AssertLettersFromColumnID(54, "BB");
            AssertLettersFromColumnID(752, "ABX");
        }

        [Test]
        public void SheetRangeGetColumnIDFromLettersTests()
        {
            AssertColumnIDFromLetters("A", 1);
            AssertColumnIDFromLetters("Z", 26);
            AssertColumnIDFromLetters("AA", 27);
            AssertColumnIDFromLetters("BB", 54);
            AssertColumnIDFromLetters("ABX", 752);
        }

        [Test]
        public void SheetRangeNoTabNameA1NotationFormattedCorrectly()
        {
            var range = new SheetRange("", 3, 1, 5, 6);

            Assert.AreEqual("C1:E6", range.A1Notation);

            Assert.AreEqual(3, range.StartColumn);
            Assert.AreEqual(1, range.StartRow);
            Assert.AreEqual(5, range.EndColumn);
            Assert.AreEqual(6, range.EndRow);

            var newRange = SheetRangeParser.ConvertFromA1Notation("C1:E6");

            Assert.IsTrue(range.Equals(newRange));
        }

        [Test]
        public void SheetRangeWithTabNameA1NotationFormattedCorrectly()
        {
            var range = new SheetRange("MyCustomTabName", 3, 1, 5, 6);

            Assert.AreEqual("MyCustomTabName!C1:E6", range.A1Notation);

            var newRange = SheetRangeParser.ConvertFromA1Notation("MyCustomTabName!C1:E6");

            Assert.IsTrue(range.Equals(newRange));
        }

        [Test]
        public void SheetRangeNoTabNameSingleCellA1NotationIsNull()
        {
            var range = new SheetRange("", 3, 1);

            Assert.IsFalse(range.CanSupportA1Notation);
            Assert.IsTrue(range.IsSingleCellRange);
            Assert.AreEqual("R1C3", range.R1C1Notation);

            var newRange = SheetRangeParser.ConvertFromR1C1Notation("R1C3");

            Assert.IsTrue(range.Equals(newRange));
        }

        [Test]
        public void SheetRangeNoTabNameR1C1FormattedCorrectly()
        {
            var range = new SheetRange("", 3, 1, 5, 5);

            Assert.AreEqual("R1C3:R5C5", range.R1C1Notation);

            var newRange = SheetRangeParser.ConvertFromR1C1Notation("R1C3:R5C5");

            Assert.IsTrue(range.Equals(newRange));
        }

        [Test]
        public void SheetRangeWithTabNameR1C1FormattedCorrectly()
        {
            var range = new SheetRange("MyCustomTab", 3, 1, 5, 5);

            Assert.AreEqual("MyCustomTab!R1C3:R5C5", range.R1C1Notation);

            var newRange = SheetRangeParser.ConvertFromR1C1Notation("MyCustomTab!R1C3:R5C5");

            Assert.IsTrue(range.Equals(newRange));
        }

        [Test]
        public void SheetRangeNoTabNameSingleCellR1C1FormattedCorrectly()
        {
            var range = new SheetRange("", 3, 1);

            Assert.IsTrue(range.IsSingleCellRange);

            Assert.AreEqual("R1C3", range.R1C1Notation);

            var newRange = SheetRangeParser.ConvertFromR1C1Notation("R1C3");

            Assert.IsTrue(range.Equals(newRange));
        }

        private static void AssertLettersFromColumnID(int columnID, string expectedLetters)
        {
            var result = SheetRange.GetLettersFromColumnID(columnID);

            Assert.AreEqual(expectedLetters, result);
        }

        private static void AssertColumnIDFromLetters(string letters, int expectedColumnID)
        {
            var result = SheetRange.GetColumnIDFromLetters(letters);

            Assert.AreEqual(expectedColumnID, result);
        }
    }
}