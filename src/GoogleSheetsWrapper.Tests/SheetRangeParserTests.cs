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
            var tabName = SheetRangeParser.GetTabName("MySheet!A1:D4");

            Assert.That(tabName, Is.EqualTo("MySheet"));
        }

        [Test]
        public void SheetRangeParserIsValidR1C1NotationHappyPathWithTab()
        {
            var result = SheetRangeParser.IsValidR1C1Notation("MySheet!R1C1:R2C2");

            Assert.That(result, Is.True);
        }

        [Test]
        public void SheetRangeParserIsValidR1C1NotationHappyPathWithoutTab()
        {
            var result = SheetRangeParser.IsValidR1C1Notation("R1C1:R2C2");

            Assert.That(result, Is.True);
        }

        [Test]
        public void SheetRangeParserIsValidR1C1NotationHappyPathSingleCellWithoutTab()
        {
            var result = SheetRangeParser.IsValidR1C1Notation("R1C1");

            Assert.That(result, Is.True);
        }

        [Test]
        public void SheetRangeParserIsValidR1C1NotationFailsWithA1Notation()
        {
            var result = SheetRangeParser.IsValidR1C1Notation("A1:D4");

            Assert.That(result, Is.False);
        }

        [Test]
        public void SheetRangeParserIsValidA1NotationHappyPathWithTab()
        {
            var result = SheetRangeParser.IsValidA1Notation("MySheet!A1:D4");

            Assert.That(result, Is.True);
        }

        [Test]
        public void SheetRangeParserIsValidA1NotationHappyPathWithTabWithoutLastRow()
        {
            var result = SheetRangeParser.IsValidA1Notation("MySheet!A1:D");

            Assert.That(result, Is.True);
        }

        [Test]
        public void SheetRangeParserIsValidA1NotationHappyPathWithoutTab()
        {
            var result = SheetRangeParser.IsValidA1Notation("A1:D4");

            Assert.That(result, Is.True);
        }
    }
}