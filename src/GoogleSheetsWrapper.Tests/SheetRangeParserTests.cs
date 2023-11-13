using NUnit.Framework;

namespace GoogleSheetsWrapper.Tests
{
    public class SheetRangeParserTests
    {
        [SetUp]
        public void Setup()
        {

        }

        [Test]
        public void SheetRangeParserGetTabNameHappyPath()
        {
            var parser = new SheetRangeParser();

            var tabName = parser.GetTabName("MySheet!A1:D4");

            Assert.AreEqual("MySheet", tabName);
        }

        [Test]
        public void SheetRangeParserIsValidR1C1NotationHappyPathWithTab()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidR1C1Notation("MySheet!R1C1:R2C2");

            Assert.IsTrue(result);
        }

        [Test]
        public void SheetRangeParserIsValidR1C1NotationHappyPathWithoutTab()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidR1C1Notation("R1C1:R2C2");

            Assert.IsTrue(result);
        }

        [Test]
        public void SheetRangeParserIsValidR1C1NotationHappyPathSingleCellWithoutTab()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidR1C1Notation("R1C1");

            Assert.IsTrue(result);
        }

        [Test]
        public void SheetRangeParserIsValidR1C1NotationFailsWithA1Notation()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidR1C1Notation("A1:D4");

            Assert.IsFalse(result);
        }

        [Test]
        public void SheetRangeParserIsValidA1NotationHappyPathWithTab()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidA1Notation("MySheet!A1:D4");

            Assert.IsTrue(result);
        }

        [Test]
        public void SheetRangeParserIsValidA1NotationHappyPathWithTabWithoutLastRow()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidA1Notation("MySheet!A1:D");

            Assert.IsTrue(result);
        }

        [Test]
        public void SheetRangeParserIsValidA1NotationHappyPathWithoutTab()
        {
            var parser = new SheetRangeParser();

            var result = parser.IsValidA1Notation("A1:D4");

            Assert.IsTrue(result);
        }
    }
}