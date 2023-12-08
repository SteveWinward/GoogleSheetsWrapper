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

            Assert.That(range.A1Notation, Is.EqualTo("C1:E6"));

            Assert.That(range.StartColumn, Is.EqualTo(3));
            Assert.That(range.StartRow, Is.EqualTo(1));
            Assert.That(range.EndColumn, Is.EqualTo(5));
            Assert.That(range.EndRow, Is.EqualTo(6));

            var newRange = SheetRangeParser.ConvertFromA1Notation("C1:E6");

            Assert.That(newRange, Is.EqualTo(range));
        }

        [Test]
        public void SheetRangeWithTabNameA1NotationFormattedCorrectly()
        {
            var range = new SheetRange("MyCustomTabName", 3, 1, 5, 6);

            Assert.That(range.A1Notation, Is.EqualTo("MyCustomTabName!C1:E6"));

            var newRange = SheetRangeParser.ConvertFromA1Notation("MyCustomTabName!C1:E6");

            Assert.That(newRange, Is.EqualTo(range));
        }

        [Test]
        public void SheetRangeWithTabNameNotationIncrememtStartColumnIsCorrect()
        {
            var range = new SheetRange("MyCustomTabName", 3, 1, 5, 6);

            Assert.That(range.A1Notation, Is.EqualTo("MyCustomTabName!C1:E6"));
            Assert.That(range.R1C1Notation, Is.EqualTo("MyCustomTabName!R1C3:R6C5"));

            range.StartColumn = 4;

            Assert.That(range.A1Notation, Is.EqualTo("MyCustomTabName!D1:E6"));
            Assert.That(range.R1C1Notation, Is.EqualTo("MyCustomTabName!R1C4:R6C5"));
        }

        [Test]
        public void SheetRangeWithTabNameNotationIncrememtStartRowIsCorrect()
        {
            var range = new SheetRange("MyCustomTabName", 3, 1, 5, 6);

            Assert.That(range.A1Notation, Is.EqualTo("MyCustomTabName!C1:E6"));
            Assert.That(range.R1C1Notation, Is.EqualTo("MyCustomTabName!R1C3:R6C5"));

            range.StartRow = 2;

            Assert.That(range.A1Notation, Is.EqualTo("MyCustomTabName!C2:E6"));
            Assert.That(range.R1C1Notation, Is.EqualTo("MyCustomTabName!R2C3:R6C5"));
        }

        [Test]
        public void SheetRangeWithTabNameNotationIncrememtEndColumnIsCorrect()
        {
            var range = new SheetRange("MyCustomTabName", 3, 1, 5, 6);

            Assert.That(range.A1Notation, Is.EqualTo("MyCustomTabName!C1:E6"));
            Assert.That(range.R1C1Notation, Is.EqualTo("MyCustomTabName!R1C3:R6C5"));

            range.EndColumn = 6;

            Assert.That(range.A1Notation, Is.EqualTo("MyCustomTabName!C1:F6"));
            Assert.That(range.R1C1Notation, Is.EqualTo("MyCustomTabName!R1C3:R6C6"));
        }

        [Test]
        public void SheetRangeWithTabNameNotationIncrememtEndRowIsCorrect()
        {
            var range = new SheetRange("MyCustomTabName", 3, 1, 5, 6);

            Assert.That(range.A1Notation, Is.EqualTo("MyCustomTabName!C1:E6"));
            Assert.That(range.R1C1Notation, Is.EqualTo("MyCustomTabName!R1C3:R6C5"));

            range.EndRow = 7;

            Assert.That(range.A1Notation, Is.EqualTo("MyCustomTabName!C1:E7"));
            Assert.That(range.R1C1Notation, Is.EqualTo("MyCustomTabName!R1C3:R7C5"));
        }

        [Test]
        public void SheetRangeWithTabNameNotationChangeTabNameIsCorrect()
        {
            var range = new SheetRange("MyCustomTabName", 3, 1, 5, 6);

            Assert.That(range.A1Notation, Is.EqualTo("MyCustomTabName!C1:E6"));
            Assert.That(range.R1C1Notation, Is.EqualTo("MyCustomTabName!R1C3:R6C5"));

            range.TabName = "NewTabName";

            Assert.That(range.A1Notation, Is.EqualTo("NewTabName!C1:E6"));
            Assert.That(range.R1C1Notation, Is.EqualTo("NewTabName!R1C3:R6C5"));
        }

        [Test]
        public void SheetRangeNoTabNameSingleCellA1NotationIsNull()
        {
            var range = new SheetRange("", 3, 1);

            Assert.That(range.CanSupportA1Notation, Is.False);
            Assert.That(range.IsSingleCellRange, Is.True);
            Assert.That(range.R1C1Notation, Is.EqualTo("R1C3"));

            var newRange = SheetRangeParser.ConvertFromR1C1Notation("R1C3");

            Assert.That(newRange, Is.EqualTo(range));
        }

        [Test]
        public void SheetRangeNoTabNameR1C1FormattedCorrectly()
        {
            var range = new SheetRange("", 3, 1, 5, 5);

            Assert.That(range.R1C1Notation, Is.EqualTo("R1C3:R5C5"));

            var newRange = SheetRangeParser.ConvertFromR1C1Notation("R1C3:R5C5");

            Assert.That(newRange, Is.EqualTo(range));
        }

        [Test]
        public void SheetRangeWithTabNameR1C1FormattedCorrectly()
        {
            var range = new SheetRange("MyCustomTab", 3, 1, 5, 5);

            Assert.That(range.R1C1Notation, Is.EqualTo("MyCustomTab!R1C3:R5C5"));

            var newRange = SheetRangeParser.ConvertFromR1C1Notation("MyCustomTab!R1C3:R5C5");

            Assert.That(newRange, Is.EqualTo(range));
        }

        [Test]
        public void SheetRangeNoTabNameSingleCellR1C1FormattedCorrectly()
        {
            var range = new SheetRange("", 3, 1);

            Assert.That(range.IsSingleCellRange, Is.True);

            Assert.That(range.R1C1Notation, Is.EqualTo("R1C3"));

            var newRange = SheetRangeParser.ConvertFromR1C1Notation("R1C3");

            Assert.That(newRange, Is.EqualTo(range));
        }

        private static void AssertLettersFromColumnID(int columnID, string expectedLetters)
        {
            var result = SheetRange.GetLettersFromColumnID(columnID);

            Assert.That(result, Is.EqualTo(expectedLetters));
        }

        private static void AssertColumnIDFromLetters(string letters, int expectedColumnID)
        {
            var result = SheetRange.GetColumnIDFromLetters(letters);

            Assert.That(result, Is.EqualTo(expectedColumnID));
        }
    }
}